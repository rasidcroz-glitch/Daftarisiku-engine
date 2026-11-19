# Gunicorn — DaftarIsiKu engine.
# Satu worker sinkron saja: LibreOffice headless berat RAM, dan tiap
# request memproses satu dokumen penuh (render PDF 2x). Worker lebih
# dari satu di mesin kecil berisiko OOM. Timeout panjang karena
# konversi LibreOffice dokumen besar bisa makan 1-3 menit.
import os

bind = "0.0.0.0:" + os.environ.get("PORT", "5000")
workers = int(os.environ.get("GUNICORN_WORKERS", "1"))
# v0.6.2: 1 thread saja — dua request /api/toc bersamaan berarti dua
# LibreOffice (~325 MB tiap proses) + Python, langsung melewati
# 512 MB. Satu request jalan, yang lain antre.
threads = int(os.environ.get("GUNICORN_THREADS", "1"))
timeout = int(os.environ.get("GUNICORN_TIMEOUT", "300"))
graceful_timeout = 30
keepalive = 5
loglevel = os.environ.get("LOG_LEVEL", "info")
# Restart worker tiap 50 request agar memori LibreOffice yang bocor
# pelan-pelan tidak menumpuk.
max_requests = int(os.environ.get("GUNICORN_MAX_REQUESTS", "50"))
max_requests_jitter = 10
