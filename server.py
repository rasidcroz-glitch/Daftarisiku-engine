"""Server API DaftarIsiKu — siap hosting (tanpa login).

Model produk (keputusan user 17 Nov 2026):
  - TANPA login Google. Alur: Upload -> Pratinjau gratis -> Bayar
    Rp15.000 (QRIS/Midtrans) -> Download. Link download berlaku 24 jam.
  - Upload + PRATINJAU daftar isi di layar = GRATIS.
  - DOWNLOAD file Word hasil = Rp15.000 per dokumen.

Endpoint:
  GET  /api/health          -> {"status": "ok", ...}
  POST /api/toc             -> multipart form, field "file" (.docx,
                               maks 10 MB). Field opsional
                               "include_lists" ("0" = jangan buat
                               DAFTAR GAMBAR/TABEL; default buat).
                               Balasan JSON: {job, toc_text,
                               entries, daftar_gambar, daftar_tabel,
                               pages_in, pages_out, verification,
                               download_url}. Pratinjau gratis.
  POST /api/pay/mock/<job>  -> PROTOTIPE SAJA: menandai job sudah dibayar.
                               Diganti webhook Midtrans saat rilis.
  GET  /api/download/<job>  -> file .docx hasil; 402 bila belum dibayar.

Keamanan/batas:
  - Hanya file .docx, maks 10 MB (413 bila lebih).
  - CORS hanya untuk origin di env ALLOWED_ORIGINS (dipisah koma).
    Default: tidak ada origin yang diizinkan (same-origin saja).
  - Nama file dibersihkan (secure_filename); job diidentifikasi dari
    id acak 12 hex, bukan dari nama file.
  - Hasil job + catatan bayar yang berumur > 24 jam dihapus otomatis
    (saat startup dan sebelum tiap request).

Jalankan lokal:  gunicorn -c gunicorn.conf.py server:app
  atau:           python3 server.py   (dev server Flask, port 5000)
"""

import json
import os
import shutil
import tempfile
import threading
import time
import uuid

from flask import Flask, jsonify, request, send_file
from werkzeug.utils import secure_filename

from toc import process

app = Flask(__name__)

# --- konfigurasi (env, dengan default aman) -------------------------
RESULTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "results")
os.makedirs(RESULTS_DIR, exist_ok=True)
PAID_FILE = os.path.join(RESULTS_DIR, "paid.json")

MAX_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", 10 * 1024 * 1024))
app.config["MAX_CONTENT_LENGTH"] = MAX_BYTES
JOB_TTL_SECONDS = int(os.environ.get("JOB_TTL_SECONDS", 24 * 3600))

ALLOWED_ORIGINS = {
    o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "").split(",")
    if o.strip()
}
VERSION = "0.6.6"


# --- CORS manual (tanpa dependensi tambahan) ------------------------
@app.after_request
def apply_cors(resp):
    origin = request.headers.get("Origin")
    if origin and origin in ALLOWED_ORIGINS:
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Vary"] = "Origin"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
    return resp


@app.route("/api/toc", methods=["OPTIONS"])
@app.route("/api/download/<job>", methods=["OPTIONS"])
@app.route("/api/pay/mock/<job>", methods=["OPTIONS"])
def _preflight(job=None):
    return ("", 204)


# --- catatan bayar (file JSON sederhana; ganti DB saat skala naik) ---
def load_paid():
    try:
        with open(PAID_FILE) as f:
            return set(json.load(f))
    except (FileNotFoundError, json.JSONDecodeError):
        return set()


def save_paid(paid):
    tmp = PAID_FILE + ".tmp"
    with open(tmp, "w") as f:
        json.dump(sorted(paid), f)
    os.replace(tmp, PAID_FILE)


# --- bersih-bersih job lama (> 24 jam) ------------------------------
_cleanup_lock = threading.Lock()
_last_cleanup = 0.0


def cleanup_old_jobs():
    """Hapus file hasil + entri bayar yang lebih tua dari JOB_TTL_SECONDS,
    dan direktori temp toc-* yang sudah tua. Aman dipanggil berkali-kali."""
    global _last_cleanup
    now = time.time()
    with _cleanup_lock:
        if now - _last_cleanup < 3600:  # paling sering 1x/jam
            return
        _last_cleanup = now
    cutoff = now - JOB_TTL_SECONDS
    try:
        for name in os.listdir(RESULTS_DIR):
            path = os.path.join(RESULTS_DIR, name)
            if name == "paid.json":
                continue
            try:
                if os.path.getmtime(path) < cutoff:
                    os.remove(path)
            except OSError:
                pass
        # entri bayar untuk job yang filenya sudah tidak ada -> buang
        paid = load_paid()
        alive = {j for j in paid
                 if os.path.exists(os.path.join(
                     RESULTS_DIR, f"{j}-hasil-daftar-isi.docx"))}
        if alive != paid:
            save_paid(alive)
    except OSError:
        pass
    tmpdir = tempfile.gettempdir()
    try:
        for name in os.listdir(tmpdir):
            if not name.startswith("toc-"):
                continue
            path = os.path.join(tmpdir, name)
            try:
                if os.path.getmtime(path) < cutoff:
                    shutil.rmtree(path, ignore_errors=True)
            except OSError:
                pass
    except OSError:
        pass


@app.before_request
def _maybe_cleanup():
    cleanup_old_jobs()


@app.errorhandler(413)
def _too_large(_exc):
    return jsonify({"error": f"file terlalu besar (maks {MAX_BYTES // (1024 * 1024)} MB)"}), 413


# --- endpoint ---------------------------------------------------------
@app.get("/api/health")
def health():
    return jsonify({"status": "ok", "app": "DaftarIsiKu",
                    "version": VERSION})


@app.post("/api/toc")
def make_toc():
    # TANPA login: pratinjau GRATIS untuk siapa saja. Pembayaran hanya
    # mengunci DOWNLOAD (cek di /api/download). Tanpa akun, hasil tidak
    # tersimpan per pengguna — link download berlaku 24 jam.
    if "file" not in request.files:
        return jsonify({"error": "field 'file' (.docx) wajib diisi"}), 400
    f = request.files["file"]
    filename = secure_filename(f.filename or "")
    if not filename.lower().endswith(".docx"):
        return jsonify({"error": "hanya file .docx yang didukung"}), 400

    job = uuid.uuid4().hex[:12]
    work = tempfile.mkdtemp(prefix=f"toc-{job}-")
    src = os.path.join(work, filename)
    include_lists = (request.form.get("include_lists", "1") or "1").strip()
    include_lists = include_lists not in ("0", "false", "False", "no")
    try:
        f.save(src)
        rep = process(src, work, include_lists=include_lists)
    except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna
        shutil.rmtree(work, ignore_errors=True)
        return jsonify({"error": f"gagal memproses dokumen: {exc}"}), 422

    out_path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")
    shutil.move(rep["output_docx"], out_path)
    shutil.rmtree(work, ignore_errors=True)
    return jsonify({
        "job": job,
        "toc_text": rep["toc_text"],
        "entries": [{"text": e["text"], "level": e["level"],
                     "page": e["page"]} for e in rep["entries"]],
        "daftar_gambar": [{"text": e["text"], "page": e["page"]}
                          for e in rep["daftar_gambar"]],
        "daftar_tabel": [{"text": e["text"], "page": e["page"]}
                         for e in rep["daftar_tabel"]],
        "pages_in": rep["pages_in"],
        "pages_out": rep["pages_out"],
        "verification": "lolos" if not rep["mismatches"] else rep["mismatches"],
        "download_url": f"/api/download/{job}",
        "download_expires_hours": JOB_TTL_SECONDS // 3600,
    })


@app.post("/api/pay/mock/<job>")
def mock_pay(job):
    """PROTOTIPE SAJA — simulasi konfirmasi bayar Rp15.000.
    Diganti webhook Midtrans (QRIS) saat rilis; endpoint ini dihapus."""
    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")
    if not os.path.exists(path):
        return jsonify({"error": "hasil tidak ditemukan"}), 404
    paid = load_paid()
    paid.add(job)
    save_paid(paid)
    return jsonify({"job": job, "paid": True, "amount": 15000,
                    "mock": True})


@app.get("/api/download/<job>")
def download(job):
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")
    if not os.path.exists(path):
        return jsonify({"error": "hasil tidak ditemukan atau sudah kedaluwarsa (24 jam)"}), 404
    if job not in load_paid():
        return jsonify({"error": "belum dibayar",
                        "amount": 15000,
                        "hint": "bayar Rp15.000 via QRIS dulu, "
                                "lalu download lagi"}), 402
    return send_file(path, as_attachment=True,
                     download_name="skripsi-daftar-isi.docx")


if __name__ == "__main__":
    cleanup_old_jobs()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
