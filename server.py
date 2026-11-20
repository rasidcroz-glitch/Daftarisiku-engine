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
                               Tetap ada untuk tes, tapi website sudah
                               tidak memakainya.
  POST /api/pay/qris/<job>  -> Buat tagihan QRIS Midtrans (Rp15.000)
                               untuk job ini. Balasan: {job, order_id,
                               amount, qr_url, expiry_time,
                               expires_in_seconds, status}. 503 bila
                               MIDTRANS_SERVER_KEY belum disetel; 409
                               bila job sudah dibayar. Charge yang
                               masih berlaku dipakai ulang (QR sama);
                               charge kedaluwarsa -> order_id baru.
  GET  /api/qr/<job>        -> Gambar PNG kode QRIS charge yang sedang
                               aktif (dibuat server dari qr_string
                               Midtrans). 404 bila belum ada charge.
  GET  /api/status/<job>     -> {job, paid, amount, filename}; 404 bila
                               job tidak ada/kedaluwarsa. Dipakai website
                               untuk mengecek status bayar (mis. setelah
                               refresh) sebelum membuka tombol download.
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

import base64
import io
import json
import os
import shutil
import tempfile
import threading
import time
import urllib.error
import urllib.request
import uuid
from datetime import datetime, timedelta, timezone

from flask import Flask, jsonify, redirect, request, send_file
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
VERSION = "0.6.13"  # v0.6.13: (1) sub-heading pra-BAB yang yatim dipindah ke setelah judul BAB (revisi nifas __1: "1.1 Latar Belakang" tidak lagi sendiri di hal. 7); (2) footer BAB I diverifikasi dari render akhir dan dipaksa jadi "1" bila masih Romawi; (3) batas section daftar lama hanya dipertahankan bila kolom section berbeda; (4) _strip_page_field kini membuang hasil-cache field PAGE (tidak ada lagi angka ganda "11" di footer)
# (titlePg + first-page footer tengah; nextPage break hanya sebelum
# BAB yang mulai di halaman baru). QRIS v0.6.9 tetap ada.

# --- konfigurasi Midtrans (kunci HANYA dari env, tidak pernah di kode) --
MIDTRANS_SERVER_KEY = os.environ.get("MIDTRANS_SERVER_KEY", "")
MIDTRANS_IS_PRODUCTION = os.environ.get("MIDTRANS_IS_PRODUCTION", "") == "1"
# Override khusus untuk tes lokal (mis. server Midtrans palsu);
# kosong = pakai endpoint resmi sandbox/produksi.
MIDTRANS_API_BASE_URL = os.environ.get("MIDTRANS_API_BASE_URL", "")
PRICE_IDR = 15000
QRIS_EXPIRY_SECONDS = 15 * 60  # default Midtrans untuk QRIS: 15 menit


# --- CORS manual (tanpa dependensi tambahan) ------------------------
@app.after_request
def apply_cors(resp):
    origin = request.headers.get("Origin")
    if origin and origin in ALLOWED_ORIGINS:
        resp.headers["Access-Control-Allow-Origin"] = origin
        resp.headers["Vary"] = "Origin"
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"
        resp.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
    return resp


@app.route("/api/toc", methods=["OPTIONS"])
@app.route("/api/status/<job>", methods=["OPTIONS"])
@app.route("/api/download/<job>", methods=["OPTIONS"])
@app.route("/api/pay/mock/<job>", methods=["OPTIONS"])
@app.route("/api/pay/qris/<job>", methods=["OPTIONS"])
@app.route("/api/qr/<job>", methods=["OPTIONS"])
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


# --- Midtrans QRIS (Core API v2; kunci hanya dari env) -----------------
# Alur: POST /api/pay/qris/<job> -> charge QRIS (order_id unik per
# percobaan) -> website menampilkan PNG dari GET /api/qr/<job> ->
# GET /api/status/<job> menanyakan status transaksi ke Midtrans dan
# menandai job lunas bila settlement/capture. Tanpa webhook: polling
# status dari website sudah cukup untuk Rp15.000/dokumen.
class MidtransError(Exception):
    pass


def _midtrans_base():
    if MIDTRANS_API_BASE_URL:
        return MIDTRANS_API_BASE_URL.rstrip("/")
    if MIDTRANS_IS_PRODUCTION:
        return "https://api.midtrans.com"
    return "https://api.sandbox.midtrans.com"


def _midtrans_auth():
    if not MIDTRANS_SERVER_KEY:
        return None
    token = base64.b64encode(
        f"{MIDTRANS_SERVER_KEY}:".encode()).decode()
    return {"Authorization": f"Basic {token}",
            "Accept": "application/json",
            "Content-Type": "application/json"}


def _midtrans_request(method, path, body=None, timeout=15):
    """Panggil Core API Midtrans; kembalikan dict JSON.
    Melempar MidtransError bila kunci belum disetel, jaringan gagal,
    atau Midtrans menjawab error (pesan aslinya ikut dibawa)."""
    headers = _midtrans_auth()
    if headers is None:
        raise MidtransError("pembayaran QRIS belum diaktifkan di server")
    url = _midtrans_base() + path
    data = json.dumps(body).encode() if body is not None else None
    req = urllib.request.Request(url, data=data, headers=headers,
                                 method=method)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode() or "{}")
    except urllib.error.HTTPError as exc:
        try:
            detail = json.loads(exc.read().decode() or "{}")
            msg = detail.get("status_message") or detail.get("error") \
                or f"HTTP {exc.code}"
        except (ValueError, UnicodeDecodeError):
            msg = f"HTTP {exc.code}"
        raise MidtransError(f"Midtrans: {msg} (HTTP {exc.code})") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise MidtransError(
            "tidak bisa menghubungi Midtrans, coba lagi sebentar lagi"
        ) from exc


def _pay_path(job):
    return os.path.join(RESULTS_DIR, f"{job}.pay.json")


def load_pay(job):
    try:
        with open(_pay_path(job)) as f:
            return json.load(f) or {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_pay(job, rec):
    tmp = _pay_path(job) + ".tmp"
    with open(tmp, "w") as f:
        json.dump(rec, f)
    os.replace(tmp, _pay_path(job))


WIB = timezone(timedelta(hours=7))  # Midtrans mengembalikan waktu WIB


def _parse_expiry(expiry_time, fallback_epoch):
    """'2026-11-19 23:40:00' (WIB) -> epoch; gagal parse -> fallback."""
    if expiry_time:
        try:
            dt = datetime.strptime(expiry_time, "%Y-%m-%d %H:%M:%S")
            return int(dt.replace(tzinfo=WIB).timestamp())
        except ValueError:
            pass
    return fallback_epoch


def _qris_info(job, rec):
    """Ringkasan charge untuk website (tanpa qr_string mentah)."""
    if not rec or not rec.get("order_id"):
        return None
    now = int(time.time())
    expires_at = rec.get("expires_at") or now
    return {"order_id": rec["order_id"],
            "status": rec.get("status", "pending"),
            "expiry_time": rec.get("expiry_time"),
            "expires_in_seconds": max(0, expires_at - now),
            "qr_url": f"/api/qr/{job}"}


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
    # v0.6.7: simpan nama file asli supaya nama download mengikuti
    # file yang di-upload (bukan "skripsi-daftar-isi.docx" terus).
    try:
        with open(os.path.join(RESULTS_DIR, f"{job}.meta.json"),
                  "w") as mf:
            json.dump({"filename": filename}, mf)
    except OSError:
        pass
    shutil.rmtree(work, ignore_errors=True)
    return jsonify({
        "job": job,
        "toc_text": rep["toc_text"],
        "entries": [{"text": e["text"], "display": e.get("display") or e["text"],
                     "level": e["level"], "page": e["page"]} for e in rep["entries"]],
        "daftar_gambar": [{"text": e["text"], "display": e.get("display") or e["text"],
                           "page": e["page"]} for e in rep["daftar_gambar"]],
        "daftar_tabel": [{"text": e["text"], "display": e.get("display") or e["text"],
                          "page": e["page"]} for e in rep["daftar_tabel"]],
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


def _charge_qris(order_id):
    """Buat satu tagihan QRIS di Midtrans; kembalikan dict respons."""
    body = {"payment_type": "qris",
            "transaction_details": {"order_id": order_id,
                                    "gross_amount": PRICE_IDR},
            "qris": {"acquirer": "gopay"}}
    return _midtrans_request("POST", "/v2/charge", body)


def _check_midtrans_status(order_id):
    """Tanya status transaksi ke Midtrans; kembalikan dict respons."""
    return _midtrans_request("GET", f"/v2/{order_id}/status")


def _mark_paid(job, rec):
    """Tandai job lunas di paid.json + catat status di pay.json."""
    paid = load_paid()
    paid.add(job)
    save_paid(paid)
    if rec is not None:
        rec["status"] = "settlement"
        save_pay(job, rec)


@app.post("/api/pay/qris/<job>")
def qris_pay(job):
    """Buat (atau pakai ulang) tagihan QRIS Rp15.000 untuk sebuah job.
    order_id harus unik di Midtrans: percobaan pertama
    'daftarisiku-<job>', percobaan berikut 'daftarisiku-<job>-2', dst.
    Charge yang masih berlaku dipakai ulang supaya pelanggan tidak
    dibebani dua tagihan untuk dokumen yang sama."""
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")
    if not os.path.exists(path):
        return jsonify({"error": "hasil tidak ditemukan"}), 404
    if job in load_paid():
        return jsonify({"error": "dokumen ini sudah dibayar"}), 409
    if not MIDTRANS_SERVER_KEY:
        return jsonify(
            {"error": "pembayaran QRIS belum diaktifkan di server"}), 503

    now = int(time.time())
    rec = load_pay(job)
    # Charge lama masih berlaku -> pakai ulang (QR yang sama).
    if rec.get("order_id") and rec.get("status") == "pending" \
            and (rec.get("expires_at") or 0) > now and rec.get("qr_string"):
        info = _qris_info(job, rec)
        return jsonify({"job": job, "order_id": rec["order_id"],
                        "amount": PRICE_IDR, "qr_url": info["qr_url"],
                        "expiry_time": rec.get("expiry_time"),
                        "expires_in_seconds": info["expires_in_seconds"],
                        "status": "pending", "reused": True}), 201
    # Charge lama kedaluwarsa: cek sekali ke Midtrans, siapa tahu
    # pelanggan sempat bayar di detik-detik terakhir.
    if rec.get("order_id") and rec.get("status") == "pending":
        try:
            st = _check_midtrans_status(rec["order_id"])
            if st.get("transaction_status") in ("settlement", "capture"):
                _mark_paid(job, rec)
                return jsonify({"job": job, "paid": True,
                                "amount": PRICE_IDR}), 200
        except MidtransError:
            pass  # lanjut buat charge baru; status dicek lagi nanti

    attempt = int(rec.get("attempt") or 0) + 1
    order_id = f"daftarisiku-{job}" if attempt == 1 \
        else f"daftarisiku-{job}-{attempt}"
    try:
        resp = _charge_qris(order_id)
    except MidtransError as exc:
        # order_id bentrok (mis. catatan lokal hilang, tapi Midtrans
        # masih menyimpannya): coba sekali lagi dengan suffix baru.
        if "order_id" in str(exc).lower() and "already" in str(exc).lower():
            attempt += 1
            order_id = f"daftarisiku-{job}-{attempt}"
            try:
                resp = _charge_qris(order_id)
            except MidtransError as exc2:
                return jsonify({"error": str(exc2)}), 502
        else:
            return jsonify({"error": str(exc)}), 502
    qr_string = resp.get("qr_string") or ""
    if not qr_string:
        return jsonify(
            {"error": "Midtrans tidak mengembalikan kode QRIS"}), 502
    # URL gambar Midtrans (cadangan bila library qrcode tidak ada).
    qr_image_url = ""
    for a in resp.get("actions") or []:
        if a.get("name") == "generate-qr-code":
            qr_image_url = a.get("url") or ""
            break
    created = int(time.time())
    expires_at = _parse_expiry(resp.get("expiry_time"),
                               created + QRIS_EXPIRY_SECONDS)
    rec = {"order_id": order_id, "attempt": attempt,
           "status": "pending", "qr_string": qr_string,
           "qr_image_url": qr_image_url,
           "expiry_time": resp.get("expiry_time"),
           "expires_at": expires_at, "created_at": created,
           "last_checked": 0}
    save_pay(job, rec)
    info = _qris_info(job, rec)
    return jsonify({"job": job, "order_id": order_id,
                    "amount": PRICE_IDR, "qr_url": info["qr_url"],
                    "expiry_time": rec.get("expiry_time"),
                    "expires_in_seconds": info["expires_in_seconds"],
                    "status": "pending", "reused": False}), 201


@app.get("/api/qr/<job>")
def qris_image(job):
    """Gambar PNG kode QRIS charge yang sedang aktif untuk job ini.
    PNG dibuat server dari qr_string Midtrans (pakai library qrcode),
    jadi website tidak butuh library QR apa pun. Bila library qrcode
    tidak tersedia, alihkan ke URL gambar Midtrans (actions)."""
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    rec = load_pay(job)
    qr_string = rec.get("qr_string") or ""
    if not qr_string or rec.get("status") not in ("pending", None):
        return jsonify({"error": "tidak ada kode QRIS aktif"}), 404
    try:
        import qrcode
        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M,
                           box_size=8, border=2)
        qr.add_data(qr_string)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return send_file(buf, mimetype="image/png",
                         download_name=f"qris-{job}.png")
    except ImportError:
        pass
    # Cadangan: URL gambar Midtrans (bila library qrcode tidak ada).
    if rec.get("qr_image_url"):
        return redirect(rec["qr_image_url"])
    return jsonify({"error": "gambar QR tidak tersedia di server"}), 503


@app.get("/api/status/<job>")
def status(job):
    """Cek status bayar sebuah job tanpa men-download.
    Dipakai website setelah refresh halaman (job tersimpan di
    localStorage), sebelum tombol download dibuka, dan sebagai
    polling saat menunggu pembayaran QRIS: bila ada charge QRIS
    yang masih pending, status transaksi ditanyakan ke Midtrans
    (paling sering 1x/10 detik) dan job ditandai lunas bila
    settlement/capture. Bila transaksi kedaluwarsa, balasan membawa
    qris_expired=true supaya website menawarkan QR baru."""
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")
    if not os.path.exists(path):
        return jsonify({"error": "hasil tidak ditemukan atau sudah kedaluwarsa (24 jam)"}), 404
    filename = None
    try:
        with open(os.path.join(RESULTS_DIR, f"{job}.meta.json")) as mf:
            filename = (json.load(mf) or {}).get("filename")
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass
    paid = job in load_paid()
    rec = load_pay(job)
    qris_expired = False
    if not paid and rec.get("order_id") \
            and rec.get("status") == "pending":
        now = int(time.time())
        # Sudah lunas di Midtrans tapi catatan lokal belum? Cek sekali.
        # Dibatasi 1x/10 detik supaya polling website (tiap ~5 detik)
        # tidak membanjiri Midtrans.
        if MIDTRANS_SERVER_KEY and now - int(rec.get("last_checked") or 0) >= 10:
            rec["last_checked"] = now
            save_pay(job, rec)
            try:
                st = _check_midtrans_status(rec["order_id"])
                tx = st.get("transaction_status")
                if tx in ("settlement", "capture"):
                    _mark_paid(job, rec)
                    paid = True
                elif tx in ("expire", "cancel", "deny"):
                    rec["status"] = tx
                    save_pay(job, rec)
                    qris_expired = True
                elif tx == "pending":
                    pass  # masih menunggu; last_checked sudah disimpan
            except MidtransError as exc:
                # 404 = order_id tidak dikenal Midtrans -> anggap hangus.
                if "HTTP 404" in str(exc):
                    rec["status"] = "expire"
                    save_pay(job, rec)
                    qris_expired = True
                # Error lain (jaringan/dll): jangan gagalkan status;
                # kembalikan info charge terakhir yang tersimpan.
                pass
        elif rec.get("status") == "pending" \
                and (rec.get("expires_at") or 0) <= now:
            qris_expired = True
    elif not paid and rec.get("status") in ("expire", "cancel", "deny"):
        qris_expired = True
    return jsonify({"job": job, "paid": paid,
                    "amount": PRICE_IDR, "filename": filename,
                    "qris": None if paid else _qris_info(job, rec),
                    "qris_expired": qris_expired})


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
    # Nama download sama persis dengan file yang di-upload
    # ("Laporan.docx" -> "Laporan.docx"); job lama tanpa meta
    # tetap memakai nama lama.
    dl_name = "skripsi.docx"
    try:
        with open(os.path.join(RESULTS_DIR, f"{job}.meta.json")) as mf:
            orig = (json.load(mf) or {}).get("filename") or ""
        dl_name = secure_filename(orig) or "skripsi.docx"
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass
    return send_file(path, as_attachment=True,
                     download_name=dl_name)


if __name__ == "__main__":
    cleanup_old_jobs()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
