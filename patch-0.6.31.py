#!/usr/bin/env python3
"""Patch v0.6.31 untuk engine DaftarIsiKu di VPS Awan.

Anti-DoS + strip relasi eksternal (audit keamanan 23 Nov):
  - antrean proses maks 5 (upload saat penuh -> 429 "server sedang
    sibuk, coba lagi sebentar"); 1 worker, jadi maks 5 menunggu +
    1 sedang diproses.
  - rate limit upload 5/jam/IP (env MAX_UPLOADS_PER_HOUR, default 5;
    in-memory, reset saat restart). IP dari X-Forwarded-For pertama
    (Vercel + Caddy), fallback remote_addr.
  - tolak .docx yang mengembang >100 MB total / >50 MB per file
    (zip bomb) SEBELUM parse apa pun; pesan detail hanya ke log
    server (ke pelanggan pesan umum, aturan v0.6.30).
  - hapus relasi eksternal (TargetMode=External) dari salinan
    _bersih.docx dan dari hasil sebelum di-render LibreOffice,
    supaya server tidak pernah GET ke URL penyerang. Uji canary
    lokal: dokumen berisi gambar eksternal diproses tuntas
    (32 entri) dan server canary menerima NOL request.

Jalankan SETELAH v0.6.30 ter-deploy (patch ini mengharapkan server.py
v0.6.30 + toc.py v0.6.29 di server).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.31.py -o /tmp/patch-0.6.31.py && python3 /tmp/patch-0.6.31.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.30. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

TOC = "/opt/daftarisiku-engine/toc.py"
SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    (TOC,
     '    return warnings\n\n\ndef process(docx_path, work_dir, include_lists=True):\n    """Pipeline lengkap. Kembalikan dict laporan + path hasil.\n    include_lists=False: DAFTAR GAMBAR / DAFTAR TABEL tidak dibuat\n    (daftar isi tetap jalan seperti biasa)."""\n    os.makedirs(work_dir, exist_ok=True)\n    entries = extract_structure(docx_path)\n    assign_display(entries)\n    if not entries:\n',
     '    return warnings\n\n\n# v0.6.31: batas dekompresi (anti zip/XML bomb, T-2) + strip relasi\n# eksternal (anti SSRF LibreOffice, S-2).\nMAX_DECOMPRESSED_TOTAL = 100 * 1024 * 1024  # 100 MB total isi .docx\nMAX_DECOMPRESSED_FILE = 50 * 1024 * 1024    # 50 MB per satu file\n\n\ndef _check_decompressed_size(docx_path):\n    """Tolak .docx yang mengembang terlalu besar (zip bomb).\n    Dipanggil di awal process(), SEBELUM parse apa pun: jumlahkan\n    ZipInfo.file_size semua isi. Pesan hanya sampai ke log server\n    (ke pelanggan hanya pesan umum, aturan v0.6.30)."""\n    import zipfile\n    total = 0\n    with zipfile.ZipFile(docx_path) as z:\n        for info in z.infolist():\n            if info.file_size > MAX_DECOMPRESSED_FILE:\n                raise ValueError(\n                    f"isi dokumen terlalu besar: {info.filename} "\n                    f"({info.file_size // (1024 * 1024)} MB)")\n            total += info.file_size\n            if total > MAX_DECOMPRESSED_TOTAL:\n                raise ValueError(\n                    "dokumen mengembang terlalu besar setelah dibuka "\n                    "(lebih dari 100 MB)")\n\n\ndef _strip_external_rels(docx_path):\n    """Hapus relasi eksternal (TargetMode=External) dari SEMUA part\n    .docx — linked image, INCLUDEPICTURE, dsb. — supaya LibreOffice\n    tidak pernah GET ke URL penyerang saat me-render (S-2).\n    Kembalikan jumlah rel yang dihapus; 0 = tidak ada perubahan\n    dan file tidak ditulis ulang."""\n    doc = Document(docx_path)\n    dropped = 0\n    for part in doc.part.package.iter_parts():\n        for rId, rel in list(part.rels.items()):\n            if rel.is_external:\n                part.drop_rel(rId)\n                dropped += 1\n    if dropped:\n        doc.save(docx_path)\n    return dropped\n\n\ndef process(docx_path, work_dir, include_lists=True):\n    """Pipeline lengkap. Kembalikan dict laporan + path hasil.\n    include_lists=False: DAFTAR GAMBAR / DAFTAR TABEL tidak dibuat\n    (daftar isi tetap jalan seperti biasa)."""\n    os.makedirs(work_dir, exist_ok=True)\n    # v0.6.31: tolak zip bomb SEBELUM parse apa pun.\n    _check_decompressed_size(docx_path)\n    entries = extract_structure(docx_path)\n    assign_display(entries)\n    if not entries:\n'),
    (TOC,
     '    del _clean\n    gc.collect()\n    _trim_memory()\n    pdf_in = render_pdf(clean_path, work_dir)\n    pages_in = pdf_pages_text(pdf_in)\n    total_in = len(pages_in)\n',
     '    del _clean\n    gc.collect()\n    _trim_memory()\n    # v0.6.31: buang rel eksternal dari salinan sebelum di-render.\n    _strip_external_rels(clean_path)\n    pdf_in = render_pdf(clean_path, work_dir)\n    pages_in = pdf_pages_text(pdf_in)\n    total_in = len(pages_in)\n'),
    (TOC,
     '    build_output_docx(docx_path, entries, out_docx,\n                      gambar=gambar, tabel=tabel,\n                      normalize_pre_bab=normalize_pre_bab)\n\n    # --- verifikasi: nomor Arab relatif terhadap BAB I, jadi penyisipan\n    # TOC sebelum BAB I TIDAK mengubah nomor (BAB I dan semua heading\n',
     '    build_output_docx(docx_path, entries, out_docx,\n                      gambar=gambar, tabel=tabel,\n                      normalize_pre_bab=normalize_pre_bab)\n    # v0.6.31: buang rel eksternal dari hasil sebelum di-render.\n    _strip_external_rels(out_docx)\n\n    # --- verifikasi: nomor Arab relatif terhadap BAB I, jadi penyisipan\n    # TOC sebelum BAB I TIDAK mengubah nomor (BAB I dan semua heading\n'),
    (SRV,
     '    o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "").split(",")\n    if o.strip()\n}\nVERSION = "0.6.30"  # v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n# (titlePg + first-page footer tengah; nextPage break hanya sebelum\n# BAB yang mulai di halaman baru). QRIS v0.6.9 tetap ada.\n\n',
     '    o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "").split(",")\n    if o.strip()\n}\nVERSION = "0.6.31"  # v0.6.31: anti-DoS + strip rel eksternal (audit keamanan 23 Nov): antrean maks 5 (penuh -> 429), rate limit upload 5/jam/IP (env MAX_UPLOADS_PER_HOUR), tolak .docx yang mengembang >100 MB total / >50 MB per file (zip bomb), hapus relasi eksternal dari salinan + hasil sebelum di-render LibreOffice. v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n# (titlePg + first-page footer tengah; nextPage break hanya sebelum\n# BAB yang mulai di halaman baru). QRIS v0.6.9 tetap ada.\n\n'),
    (SRV,
     '# (dua LibreOffice bersamaan langsung melewati 512 MB). Status +\n# hasil tersimpan di file {job}.job.json (tulis atomik) supaya\n# GET /api/status bisa menjawab walau worker sempat di-recycle.\nJOB_QUEUE = queue.Queue()\n_JOB_STATE_LOCK = threading.Lock()\n_JOB_STALE_SECONDS = 10 * 60  # "processing" lebih tua dari ini = terhenti\n\n',
     '# (dua LibreOffice bersamaan langsung melewati 512 MB). Status +\n# hasil tersimpan di file {job}.job.json (tulis atomik) supaya\n# GET /api/status bisa menjawab walau worker sempat di-recycle.\nJOB_QUEUE = queue.Queue(maxsize=5)  # v0.6.31: antrean berbatas (T-1);\n# upload saat penuh -> 429 (put_nowait di make_toc). 1 worker, jadi\n# maks 5 menunggu + 1 sedang diproses.\n_UPLOAD_LIMITS = {}  # v0.6.31: rate limit upload per IP (T-1)\n_UPLOAD_LIMIT_LOCK = threading.Lock()\nMAX_UPLOADS_PER_HOUR = int(os.environ.get("MAX_UPLOADS_PER_HOUR", "5"))\n\n\ndef _client_ip():\n    """IP klien: X-Forwarded-For pertama (Vercel + Caddy menulis IP\n    asli di depan), fallback remote_addr. Header ini bisa dipalsukan\n    pemanggil langsung — hanya dipakai untuk rate limit; batas\n    kerasnya tetap antrean maxsize 5."""\n    fwd = request.headers.get("X-Forwarded-For") or ""\n    if fwd:\n        return fwd.split(",")[0].strip()\n    return request.remote_addr or "unknown"\n\n\ndef _check_upload_rate(ip):\n    """True bila IP masih boleh upload (maks N/jam, sliding window\n    1 jam, in-memory; reset saat restart — bisa diterima)."""\n    now = time.time()\n    with _UPLOAD_LIMIT_LOCK:\n        stamps = [t for t in _UPLOAD_LIMITS.get(ip, [])\n                  if now - t < 3600]\n        if len(stamps) >= MAX_UPLOADS_PER_HOUR:\n            return False\n        stamps.append(now)\n        _UPLOAD_LIMITS[ip] = stamps\n        return True\n_JOB_STATE_LOCK = threading.Lock()\n_JOB_STALE_SECONDS = 10 * 60  # "processing" lebih tua dari ini = terhenti\n\n'),
    (SRV,
     '    src = os.path.join(work, filename)\n    include_lists = (request.form.get("include_lists", "1") or "1").strip()\n    include_lists = include_lists not in ("0", "false", "False", "no")\n    try:\n        f.save(src)\n    except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna\n',
     '    src = os.path.join(work, filename)\n    include_lists = (request.form.get("include_lists", "1") or "1").strip()\n    include_lists = include_lists not in ("0", "false", "False", "no")\n    # v0.6.31: rate limit per IP sebelum file disimpan.\n    if not _check_upload_rate(_client_ip()):\n        shutil.rmtree(work, ignore_errors=True)\n        return jsonify(\n            {"error": "terlalu banyak upload dari alamat ini, "\n                      "coba lagi nanti"}), 429\n    try:\n        f.save(src)\n    except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna\n'),
    (SRV,
     '                          "include_lists": include_lists,\n                          "created_at": now, "updated_at": now})\n    _ensure_worker()\n    JOB_QUEUE.put({"job": job, "src": src, "work": work,\n                   "filename": filename, "include_lists": include_lists})\n    return jsonify({"job": job, "status": "queued"}), 202\n\n\n',
     '                          "include_lists": include_lists,\n                          "created_at": now, "updated_at": now})\n    _ensure_worker()\n    try:\n        JOB_QUEUE.put_nowait({"job": job, "src": src, "work": work,\n                              "filename": filename,\n                              "include_lists": include_lists})\n    except queue.Full:\n        # v0.6.31: antrean penuh — bersihkan state + file, tolak 429.\n        try:\n            os.remove(_job_state_path(job))\n        except OSError:\n            pass\n        shutil.rmtree(work, ignore_errors=True)\n        return jsonify(\n            {"error": "server sedang sibuk, coba lagi sebentar"}), 429\n    return jsonify({"job": job, "status": "queued"}), 202\n\n\n'),
]

def run(cmd):
    r = subprocess.run(cmd, shell=True, capture_output=True, text=True,
                       timeout=120)
    return r.returncode, (r.stdout + r.stderr).strip()


def main():
    for path, old, new in PAIRS:
        try:
            with open(path, encoding="utf-8") as f:
                src = f.read()
        except FileNotFoundError:
            print(f"GAGAL: file tidak ketemu: {path}")
            return 1
        n = src.count(old)
        if n != 1:
            print(f"GAGAL: pola ke-{PAIRS.index((path, old, new)) + 1} "
                  f"ketemu {n}x di {path} (harusnya 1x) — "
                  "patch DIBATALKAN, tidak ada yang diubah.")
            return 1
    # Semua pola cocok — cadangkan dulu, baru tulis.
    for path in {p for p, _, _ in PAIRS}:
        with open(path, encoding="utf-8") as f:
            src = f.read()
        bak = path + ".bak-0.6.30"
        with open(bak, "w", encoding="utf-8") as f:
            f.write(src)
        print(f"cadangan: {bak}")
    for path, old, new in PAIRS:
        with open(path, encoding="utf-8") as f:
            src = f.read()
        with open(path, "w", encoding="utf-8") as f:
            f.write(src.replace(old, new, 1))
    print(f"OK: {len(PAIRS)} penggantian diterapkan.")
    code, out = run(f"python3 -m py_compile {TOC} {SRV}")
    if code != 0:
        print("GAGAL: syntax error setelah patch:\n" + out)
        return 1
    print("OK: syntax valid.")
    code, out = run("systemctl restart daftarisiku")
    if code != 0:
        print("GAGAL restart service:\n" + out)
        return 1
    print("OK: service daftarisiku di-restart.")
    import time
    time.sleep(5)
    code, out = run("curl -s -o /dev/null -w '%{http_code}' "
                    "http://127.0.0.1:8080/api/health")
    print(f"health lokal: HTTP {out}")
    if out != "200":
        print("PERINGATAN: health bukan 200 — cek log: "
              "journalctl -u daftarisiku -n 50")
        return 1
    print("SELESAI: engine v0.6.31 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
