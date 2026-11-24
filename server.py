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
                               "page_pos" ("right" | "center" |
                               "top-right"; posisi nomor halaman isi,
                               default "right"; halaman depan +
                               halaman pertama BAB tetap tengah).
                               "left" (Kiri bawah) dihapus dari
                               website v0.6.33 tapi tetap diterima
                               mesin bila ada job lama yang
                               mengirimnya.
                               ASINKRON (v0.6.17): langsung balas 202
                               {job, status: "queued"}; proses jalan
                               di thread latar (satu dokumen dalam
                               satu waktu) dan hasilnya dicek lewat
                               GET /api/status/<job>. Alasannya:
                               dokumen berat (4 MB, 147 hal) butuh
                               ~46 detik, sementara hosting memutus
                               request yang lebih dari ~20 detik
                               (503) — alur sinkron lama mentok di
                               website untuk file sebesar itu.
  POST /api/pay/qris/<job>  -> Buat tagihan QRIS Midtrans (Rp15.000)
                               untuk job ini. Balasan: {job, order_id,
                               amount, qr_url, expiry_time,
                               expires_in_seconds, status}. 503 bila
                               MIDTRANS_SERVER_KEY belum disetel; 409
                               bila job sudah dibayar. Charge yang
                               masih berlaku dipakai ulang (QR sama);
                               charge kedaluwarsa -> order_id baru.
  POST /api/pay/simulate/<job> -> Simulasi bayar (v0.6.34, SANDBOX
                               saja, hanya pemilik/Muse): tandai job
                               lunas tanpa Midtrans. Wajib header
                               X-Reports-Token = REPORTS_TOKEN (sama
                               seperti /api/reports); tanpa token
                               yang benar -> 404 (seolah endpoint
                               tidak ada). Mati total bila
                               MIDTRANS_IS_PRODUCTION=1 (-> 404).
                               404 bila hasil belum ada; 409 bila
                               sudah dibayar. TIDAK dipakai website;
                               pelanggan tetap lewat QRIS.
  GET  /api/qr/<job>        -> Gambar PNG kode QRIS charge yang sedang
                               aktif (dibuat server dari qr_string
                               Midtrans). 404 bila belum ada charge.
  GET  /api/status/<job>     -> {job, status, paid, amount, filename,
                               qris, qris_expired}; status =
                               "queued" | "processing" | "done" |
                               "error". Bila "done", balasan juga
                               membawa hasil pratinjau (entries,
                               daftar_gambar, daftar_tabel, pages_in,
                               pages_out, verification, download_url)
                               supaya website bisa menampilkan
                               pratinjau dari polling ini. 404 bila
                               job tidak ada/kedaluwarsa. Dipakai
                               website untuk mengecek status proses
                               dan status bayar (mis. setelah refresh)
                               sebelum membuka tombol download.
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
import queue
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
VERSION = "0.6.35"  # v0.6.35: footer Romawi dokumen 1 section — sectPr continuous di paragraf BAB I bila tepat 1 section + _has_front + BAB I mulai di halaman baru (makalah: depan Romawi, bab1_footer=1). v0.6.34: simulasi bayar bertoken POST /api/pay/simulate/<job> (sandbox saja; wajib header X-Reports-Token = REPORTS_TOKEN, tanpa token -> 404; mati total bila MIDTRANS_IS_PRODUCTION=1; pengganti /api/pay/mock yang ditutup v0.6.28 — mock lama TETAP tertutup, tidak dibuka lagi). v0.6.33: posisi "Kanan atas" (page_pos: top-right) — nomor halaman isi di header kanan atas, footer isi dikosongkan; "Kiri bawah" (left) dihapus dari website. Halaman depan + halaman pertama BAB tetap tengah. Perbaiki juga pemanggilan _ensure_bab_first_page_footer 2-argumen (TypeError laten v0.6.32). v0.6.32: pilihan posisi nomor halaman isi (page_pos: right/center/left; halaman depan + halaman pertama BAB tetap tengah). v0.6.31: anti-DoS + strip rel eksternal (audit keamanan 23 Nov): antrean maks 5 (penuh -> 429), rate limit upload 5/jam/IP (env MAX_UPLOADS_PER_HOUR), tolak .docx yang mengembang >100 MB total / >50 MB per file (zip bomb), hapus relasi eksternal dari salinan + hasil sebelum di-render LibreOffice. v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: '1.1' bertitik, 'I.' Romawi, 'A.' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab '1' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya '1)' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator '1)' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> ('Page Numbers' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).
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
# v0.6.30: batas percobaan charge QRIS per job; token baca /api/reports
# (dari env, di-set di drop-in VPS; JANGAN tulis nilainya di kode/chat).
MAX_QRIS_ATTEMPTS = 5
REPORTS_TOKEN = os.environ.get("REPORTS_TOKEN", "")


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
    # v0.6.30: security headers untuk semua respons engine.
    resp.headers["Strict-Transport-Security"] = "max-age=63072000"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Referrer-Policy"] = "same-origin"
    return resp


@app.route("/api/toc", methods=["OPTIONS"])
@app.route("/api/status/<job>", methods=["OPTIONS"])
@app.route("/api/download/<job>", methods=["OPTIONS"])
@app.route("/api/pay/qris/<job>", methods=["OPTIONS"])
@app.route("/api/pay/simulate/<job>", methods=["OPTIONS"])
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


# --- antrean proses asinkron (v0.6.17) -------------------------------
# /api/toc sinkron mentok di hosting gratis: request yang lebih dari
# ~20 detik diputus platform (503), sementara dokumen berat butuh
# 46 detik lebih (UJI_COBA_2: 4 MB, 147 hal, 92 tabel). Jadi upload
# hanya mengantrekan pekerjaan dan langsung membalas 202; satu worker
# thread latar memproses antrean FIFO satu dokumen dalam satu waktu
# (dua LibreOffice bersamaan langsung melewati 512 MB). Status +
# hasil tersimpan di file {job}.job.json (tulis atomik) supaya
# GET /api/status bisa menjawab walau worker sempat di-recycle.
JOB_QUEUE = queue.Queue(maxsize=5)  # v0.6.31: antrean berbatas (T-1);
# upload saat penuh -> 429 (put_nowait di make_toc). 1 worker, jadi
# maks 5 menunggu + 1 sedang diproses.
_UPLOAD_LIMITS = {}  # v0.6.31: rate limit upload per IP (T-1)
_UPLOAD_LIMIT_LOCK = threading.Lock()
MAX_UPLOADS_PER_HOUR = int(os.environ.get("MAX_UPLOADS_PER_HOUR", "5"))


def _client_ip():
    """IP klien: X-Forwarded-For pertama (Vercel + Caddy menulis IP
    asli di depan), fallback remote_addr. Header ini bisa dipalsukan
    pemanggil langsung — hanya dipakai untuk rate limit; batas
    kerasnya tetap antrean maxsize 5."""
    fwd = request.headers.get("X-Forwarded-For") or ""
    if fwd:
        return fwd.split(",")[0].strip()
    return request.remote_addr or "unknown"


def _check_upload_rate(ip):
    """True bila IP masih boleh upload (maks N/jam, sliding window
    1 jam, in-memory; reset saat restart — bisa diterima)."""
    now = time.time()
    with _UPLOAD_LIMIT_LOCK:
        stamps = [t for t in _UPLOAD_LIMITS.get(ip, [])
                  if now - t < 3600]
        if len(stamps) >= MAX_UPLOADS_PER_HOUR:
            return False
        stamps.append(now)
        _UPLOAD_LIMITS[ip] = stamps
        return True
_JOB_STATE_LOCK = threading.Lock()
_JOB_STALE_SECONDS = 10 * 60  # "processing" lebih tua dari ini = terhenti


def _job_state_path(job):
    return os.path.join(RESULTS_DIR, f"{job}.job.json")


def _save_job_state(job, state):
    tmp = _job_state_path(job) + ".tmp"
    try:
        with open(tmp, "w") as f:
            json.dump(state, f)
        os.replace(tmp, _job_state_path(job))
    except OSError:
        pass


def _load_job_state(job):
    try:
        with open(_job_state_path(job)) as f:
            return json.load(f) or {}
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def _update_job_state(job, **fields):
    with _JOB_STATE_LOCK:
        state = _load_job_state(job)
        state.update(fields)
        state["updated_at"] = int(time.time())
        _save_job_state(job, state)
        return state


def _queue_position(job):
    """Posisi job di antrean (0 = paling depan, giliran berikutnya)."""
    try:
        items = list(JOB_QUEUE.queue)
    except Exception:  # noqa: BLE001 - antrean tidak bisa dibaca
        return None
    for i, item in enumerate(items):
        if item.get("job") == job:
            return i
    return None


def _toc_result_payload(rep, job):
    """Bentuk hasil pratinjau yang sama seperti balasan /api/toc lama."""
    return {
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
        # v0.6.22: AI Document Checker — peringatan struktur dokumen.
        "checker_warnings": rep.get("checker_warnings") or [],
        "download_url": f"/api/download/{job}",
        # v0.6.20: satu bayar Rp15.000 membuka Word DAN PDF.
        "download_pdf_url": (f"/api/download/{job}?format=pdf"
                             if rep.get("output_pdf") else None),
        "download_expires_hours": JOB_TTL_SECONDS // 3600,
    }


def _job_worker():
    """Satu worker latar: ambil job dari antrean, proses, simpan hasil."""
    while True:
        item = JOB_QUEUE.get()
        job = item["job"]
        try:
            _update_job_state(job, status="processing")
            rep = process(item["src"], item["work"],
                          include_lists=item["include_lists"],
                          body_align=item.get("page_pos", "right"))
            out_path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")
            shutil.move(rep["output_docx"], out_path)
            # v0.6.20: simpan juga PDF hasil akhir (render yang cocok
            # dengan docx final) supaya satu bayar membuka Word DAN PDF.
            pdf_src = rep.get("output_pdf")
            if pdf_src and os.path.exists(pdf_src):
                try:
                    shutil.move(pdf_src, os.path.join(
                        RESULTS_DIR, f"{job}-hasil-daftar-isi.pdf"))
                except OSError:
                    pass
            # v0.6.7: simpan nama file asli supaya nama download mengikuti
            # file yang di-upload (bukan "skripsi-daftar-isi.docx" terus).
            try:
                with open(os.path.join(RESULTS_DIR, f"{job}.meta.json"),
                          "w") as mf:
                    json.dump({"filename": item["filename"]}, mf)
            except OSError:
                pass
            _update_job_state(job, status="done",
                              result=_toc_result_payload(rep, job))
        except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna
            # v0.6.30: detail error (berisi path server) hanya ke log,
            # ke pelanggan hanya pesan umum.
            print(f"worker gagal job {job}: {exc}", flush=True)
            _update_job_state(job, status="error",
                              error="gagal memproses dokumen.")
        finally:
            shutil.rmtree(item["work"], ignore_errors=True)
            JOB_QUEUE.task_done()


_worker_started = False
_worker_lock = threading.Lock()


def _ensure_worker():
    global _worker_started
    with _worker_lock:
        if not _worker_started:
            t = threading.Thread(target=_job_worker, daemon=True,
                                 name="toc-worker")
            t.start()
            _worker_started = True


_ensure_worker()


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
    # ASINKRON (v0.6.17): hanya terima + antrekan, langsung balas 202.
    # Proses berat jalan di worker latar; hasilnya dicek lewat
    # GET /api/status/<job> (website polling tiap ~2 detik).
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
    # v0.6.33: posisi nomor halaman isi ("right"|"center"|"top-right";
    # "left" tetap diterima mesin untuk job lama).
    page_pos = (request.form.get("page_pos", "right") or "right").strip().lower()
    if page_pos not in ("right", "center", "left", "top-right"):
        page_pos = "right"
    # v0.6.31: rate limit per IP sebelum file disimpan.
    if not _check_upload_rate(_client_ip()):
        shutil.rmtree(work, ignore_errors=True)
        return jsonify(
            {"error": "terlalu banyak upload dari alamat ini, "
                      "coba lagi nanti"}), 429
    try:
        f.save(src)
    except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna
        shutil.rmtree(work, ignore_errors=True)
        print(f"gagal menyimpan upload: {exc}", flush=True)
        return jsonify({"error": "gagal menyimpan file."}), 422
    now = int(time.time())
    _save_job_state(job, {"job": job, "status": "queued",
                          "filename": filename,
                          "include_lists": include_lists,
                          "page_pos": page_pos,
                          "created_at": now, "updated_at": now})
    _ensure_worker()
    try:
        JOB_QUEUE.put_nowait({"job": job, "src": src, "work": work,
                              "filename": filename,
                              "include_lists": include_lists,
                              "page_pos": page_pos})
    except queue.Full:
        # v0.6.31: antrean penuh — bersihkan state + file, tolak 429.
        try:
            os.remove(_job_state_path(job))
        except OSError:
            pass
        shutil.rmtree(work, ignore_errors=True)
        return jsonify(
            {"error": "server sedang sibuk, coba lagi sebentar"}), 429
    return jsonify({"job": job, "status": "queued"}), 202


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


def _amount_matches(st):
    """True bila gross_amount respons Midtrans == PRICE_IDR (v0.6.30).
    Midtrans mengirim string, mis. "15000" atau "15000.00". Bila
    nominal tidak ada / tidak bisa dibaca / tidak sama -> False
    (job JANGAN ditandai lunas)."""
    try:
        return float(str(st.get("gross_amount"))) == PRICE_IDR
    except (TypeError, ValueError):
        return False


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
            if st.get("transaction_status") in ("settlement", "capture") \
                    and _amount_matches(st):
                _mark_paid(job, rec)
                return jsonify({"job": job, "paid": True,
                                "amount": PRICE_IDR}), 200
            if st.get("transaction_status") in ("settlement", "capture"):
                print(f"PERINGATAN nominal Midtrans tidak cocok "
                      f"job {job}: {st.get('gross_amount')}", flush=True)
        except MidtransError:
            pass  # lanjut buat charge baru; status dicek lagi nanti

    # v0.6.30: batasi percobaan charge per job (anti banjir order_id).
    if int(rec.get("attempt") or 0) >= MAX_QRIS_ATTEMPTS:
        return jsonify(
            {"error": "terlalu banyak percobaan bayar untuk dokumen ini"}), 429
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


@app.post("/api/pay/simulate/<job>")
def simulate_pay(job):
    """Simulasi bayar (v0.6.34): tandai job lunas TANPA Midtrans.
    Hanya untuk pemilik/Muse saat Sandbox — dipakai mengetes jalur
    download selagi review bisnis Midtrans belum lolos (perintah user
    24 Nov 2026: jangan pakai Midtrans dulu, pakai simulasi yang awal
    saja; nanti bila review selesai, kembali ke Midtrans). BUKAN
    untuk pelanggan dan TIDAK dipakai website.
    Pengaman (semua wajib, fail closed):
      - MIDTRANS_IS_PRODUCTION=1 -> 404 (endpoint tidak ada).
      - Header X-Reports-Token harus sama dengan REPORTS_TOKEN
        (pemeriksaan yang sama seperti /api/reports); hilang/salah
        atau env belum disetel -> 404, bukan 401, supaya keberadaan
        endpoint tidak terungkap dari luar.
      - File hasil job harus ada dan job belum dibayar (409 bila
        sudah dibayar). Penandaan lunas memakai _mark_paid yang sama
        seperti settlement Midtrans; catatan pay ditandai
        simulated=True supaya jejaknya beda dari lunas Midtrans.
    Endpoint /api/pay/mock yang lama (v0.6.28, tanpa token — siapa
    pun yang tahu link job bisa menandai lunas gratis) TETAP dihapus
    dan tidak dibuka lagi; simulasi ini penggantinya yang bertoken."""
    if MIDTRANS_IS_PRODUCTION:
        return jsonify({"error": "tidak ditemukan"}), 404
    if not REPORTS_TOKEN \
            or request.headers.get("X-Reports-Token") != REPORTS_TOKEN:
        return jsonify({"error": "tidak ditemukan"}), 404
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")
    if not os.path.exists(path):
        return jsonify({"error": "hasil tidak ditemukan"}), 404
    if job in load_paid():
        return jsonify({"error": "dokumen ini sudah dibayar"}), 409
    rec = load_pay(job)
    rec["simulated"] = True
    _mark_paid(job, rec)
    return jsonify({"job": job, "paid": True, "simulated": True,
                    "amount": PRICE_IDR}), 200


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
    """Cek status proses + status bayar sebuah job.
    Dipakai website (v0.6.17) sebagai polling saat menunggu proses
    selesai (tiap ~2 detik): status "queued"/"processing" berarti
    masih jalan; "done" membawa hasil pratinjau (entries,
    daftar_gambar, daftar_tabel, pages_in, dst.) supaya pratinjau
    bisa dirender dari polling ini; "error" membawa pesan gagal.
    Dipakai juga setelah refresh halaman (job tersimpan di
    localStorage), sebelum tombol download dibuka, dan sebagai
    polling saat menunggu pembayaran QRIS: bila ada charge QRIS
    yang masih pending, status transaksi ditanyakan ke Midtrans
    (paling sering 1x/10 detik) dan job ditandai lunas bila
    settlement/capture. Bila transaksi kedaluwarsa, balasan membawa
    qris_expired=true supaya website menawarkan QR baru."""
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    state = _load_job_state(job)
    out_path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")

    # Job masih dalam antrean/proses (v0.6.17): laporkan statusnya.
    if state.get("status") in ("queued", "processing"):
        # Worker terhenti (mis. worker di-recycle gunicorn saat proses
        # berjalan): status tidak diperbarui >10 menit -> anggap gagal.
        if state["status"] == "processing" and \
                int(time.time()) - int(state.get("updated_at") or 0) \
                > _JOB_STALE_SECONDS:
            _update_job_state(
                job, status="error",
                error="proses terhenti di server, silakan upload ulang dokumenmu.")
            return jsonify({
                "job": job, "status": "error",
                "error": "proses terhenti di server, silakan upload ulang dokumenmu.",
                "paid": False, "amount": PRICE_IDR,
                "filename": state.get("filename"),
                "qris": None, "qris_expired": False})
        pos = _queue_position(job) if state["status"] == "queued" else None
        return jsonify({"job": job, "status": state["status"],
                        "queue_position": pos,
                        "paid": False, "amount": PRICE_IDR,
                        "filename": state.get("filename"),
                        "qris": None, "qris_expired": False})
    if state.get("status") == "error":
        return jsonify({"job": job, "status": "error",
                        "error": state.get("error") or "gagal memproses dokumen.",
                        "paid": False, "amount": PRICE_IDR,
                        "filename": state.get("filename"),
                        "qris": None, "qris_expired": False})

    # Job selesai (atau job lama tanpa state): butuh file hasil.
    if not os.path.exists(out_path):
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
                if tx in ("settlement", "capture") and _amount_matches(st):
                    _mark_paid(job, rec)
                    paid = True
                elif tx in ("settlement", "capture"):
                    # v0.6.30: lunas tapi nominal tidak cocok -> JANGAN
                    # tandai lunas; catat dan biarkan charge kedaluwarsa.
                    print(f"PERINGATAN nominal Midtrans tidak cocok "
                          f"job {job}: {st.get('gross_amount')}", flush=True)
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
    resp = {"job": job, "status": "done", "paid": paid,
            "amount": PRICE_IDR, "filename": filename,
            "qris": None if paid else _qris_info(job, rec),
            "qris_expired": qris_expired}
    # Hasil pratinjau (v0.6.17) supaya website bisa merender pratinjau
    # dari polling status; job lama tanpa state tidak membawanya dan
    # website memakai pratinjau tersimpan di localStorage.
    resp.update(state.get("result") or {})
    return jsonify(resp)


@app.post("/api/report/<job>")
def report(job):
    """Laporkan hasil salah (v0.6.21): pelanggan di pratinjau menandai
    hasil bermasalah supaya pemilik bisa memperbaiki mesin. Laporan
    TIDAK membuka download (bayar tetap Rp15.000), satu laporan per job.
    Disimpan sebagai {job}.report.json di RESULTS_DIR (ikut kedaluwarsa
    24 jam seperti hasil; Muse mengecek berkala lewat /api/reports)."""
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    if (not os.path.exists(os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx"))
            and not os.path.exists(_job_state_path(job))):
        return jsonify({"error": "job tidak ditemukan"}), 404
    rep_path = os.path.join(RESULTS_DIR, f"{job}.report.json")
    if os.path.exists(rep_path):
        return jsonify({"error": "laporan sudah dikirim untuk dokumen ini",
                        "reported": True}), 409
    try:
        data = request.get_json(silent=True) or {}
    except Exception:
        data = {}
    msg = (data.get("message") or "")
    if not isinstance(msg, str):
        msg = ""
    msg = msg.strip()[:500]
    try:
        with open(os.path.join(RESULTS_DIR, f"{job}.meta.json")) as mf:
            filename = (json.load(mf) or {}).get("filename") or ""
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        filename = ""
    rec = {"job": job, "filename": filename, "message": msg,
           "created_at": int(time.time())}
    try:
        with open(rep_path, "w") as f:
            json.dump(rec, f)
    except OSError as exc:
        print(f"gagal menyimpan laporan job {job}: {exc}", flush=True)
        return jsonify({"error": "gagal menyimpan laporan."}), 500
    return jsonify({"reported": True, "job": job})


@app.get("/api/reports")
def list_reports():
    """Daftar laporan hasil salah (untuk pemilik/Muse mengecek).
    v0.6.30: wajib token rahasia di header X-Reports-Token (env
    REPORTS_TOKEN); tanpa token yang benar -> 401. Bila env belum
    disetel, endpoint tertutup total (fail closed)."""
    if not REPORTS_TOKEN \
            or request.headers.get("X-Reports-Token") != REPORTS_TOKEN:
        return jsonify({"error": "tidak diizinkan"}), 401
    out = []
    try:
        for name in sorted(os.listdir(RESULTS_DIR)):
            if not name.endswith(".report.json"):
                continue
            try:
                with open(os.path.join(RESULTS_DIR, name)) as f:
                    out.append(json.load(f))
            except (OSError, json.JSONDecodeError):
                continue
    except OSError:
        pass
    return jsonify({"reports": out})


@app.get("/api/download/<job>")
def download(job):
    if not job.isalnum() or len(job) > 32:
        return jsonify({"error": "job tidak valid"}), 400
    # v0.6.20: ?format=pdf mengunduh PDF (satu bayar membuka keduanya).
    want_pdf = (request.args.get("format", "") or "").lower() == "pdf"
    suffix = "pdf" if want_pdf else "docx"
    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.{suffix}")
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
    if want_pdf:
        dl_name = (dl_name[:-5] + ".pdf") if dl_name.lower().endswith(".docx") else (dl_name + ".pdf")
    return send_file(path, as_attachment=True,
                     download_name=dl_name)


if __name__ == "__main__":
    cleanup_old_jobs()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
