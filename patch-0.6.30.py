#!/usr/bin/env python3
"""Patch v0.6.30 untuk engine DaftarIsiKu di VPS Awan.

Pengerasan bayar + laporan (audit keamanan 23 Nov):
  - tandai lunas hanya bila gross_amount Midtrans == Rp15.000
    (sebelumnya hanya cek status settlement/capture).
  - batasi charge QRIS maks 5 percobaan per job (ke-6 -> 429).
  - /api/reports kini wajib token rahasia di header X-Reports-Token
    (env REPORTS_TOKEN, di-set di drop-in systemd SEBELUM patch ini
    jalan; tanpa token yang benar -> 401, tanpa env -> tertutup).
  - pesan error ke pelanggan dibuat generik (detail berisi path
    server hanya ditulis ke log server).
  - security headers di semua respons engine (HSTS, nosniff,
    X-Frame-Options DENY, Referrer-Policy same-origin).

Jalankan SETELAH v0.6.29 ter-deploy (patch ini mengharapkan server.py
v0.6.29 di server). Token REPORTS_TOKEN harus sudah ada di drop-in
systemd SEBELUM service di-restart patch ini.

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.30.py -o /tmp/patch-0.6.30.py && python3 /tmp/patch-0.6.30.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.29. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    (SRV,
     '    o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "").split(",")\n    if o.strip()\n}\nVERSION = "0.6.29"  # v0.6.29: caption terbelah 2 paragraf kini terdeteksi ("Gambar 1.1" + deskripsi di paragraf berikut, atau "Gambar" saja + "1.1 ..."; label yang diikuti sub-bab tidak digabung); varian baru v17, suite 17/17 + korpus 8/8. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n# (titlePg + first-page footer tengah; nextPage break hanya sebelum\n# BAB yang mulai di halaman baru). QRIS v0.6.9 tetap ada.\n\n',
     '    o.strip() for o in os.environ.get("ALLOWED_ORIGINS", "").split(",")\n    if o.strip()\n}\nVERSION = "0.6.30"  # v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n# (titlePg + first-page footer tengah; nextPage break hanya sebelum\n# BAB yang mulai di halaman baru). QRIS v0.6.9 tetap ada.\n\n'),
    (SRV,
     'MIDTRANS_API_BASE_URL = os.environ.get("MIDTRANS_API_BASE_URL", "")\nPRICE_IDR = 15000\nQRIS_EXPIRY_SECONDS = 15 * 60  # default Midtrans untuk QRIS: 15 menit\n\n\n# --- CORS manual (tanpa dependensi tambahan) ------------------------\n',
     'MIDTRANS_API_BASE_URL = os.environ.get("MIDTRANS_API_BASE_URL", "")\nPRICE_IDR = 15000\nQRIS_EXPIRY_SECONDS = 15 * 60  # default Midtrans untuk QRIS: 15 menit\n# v0.6.30: batas percobaan charge QRIS per job; token baca /api/reports\n# (dari env, di-set di drop-in VPS; JANGAN tulis nilainya di kode/chat).\nMAX_QRIS_ATTEMPTS = 5\nREPORTS_TOKEN = os.environ.get("REPORTS_TOKEN", "")\n\n\n# --- CORS manual (tanpa dependensi tambahan) ------------------------\n'),
    (SRV,
     '        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"\n        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"\n        resp.headers["Access-Control-Expose-Headers"] = "Content-Disposition"\n    return resp\n\n\n',
     '        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"\n        resp.headers["Access-Control-Allow-Headers"] = "Content-Type"\n        resp.headers["Access-Control-Expose-Headers"] = "Content-Disposition"\n    # v0.6.30: security headers untuk semua respons engine.\n    resp.headers["Strict-Transport-Security"] = "max-age=63072000"\n    resp.headers["X-Content-Type-Options"] = "nosniff"\n    resp.headers["X-Frame-Options"] = "DENY"\n    resp.headers["Referrer-Policy"] = "same-origin"\n    return resp\n\n\n'),
    (SRV,
     '            _update_job_state(job, status="done",\n                              result=_toc_result_payload(rep, job))\n        except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna\n            _update_job_state(job, status="error",\n                              error=f"gagal memproses dokumen: {exc}")\n        finally:\n            shutil.rmtree(item["work"], ignore_errors=True)\n            JOB_QUEUE.task_done()\n',
     '            _update_job_state(job, status="done",\n                              result=_toc_result_payload(rep, job))\n        except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna\n            # v0.6.30: detail error (berisi path server) hanya ke log,\n            # ke pelanggan hanya pesan umum.\n            print(f"worker gagal job {job}: {exc}", flush=True)\n            _update_job_state(job, status="error",\n                              error="gagal memproses dokumen.")\n        finally:\n            shutil.rmtree(item["work"], ignore_errors=True)\n            JOB_QUEUE.task_done()\n'),
    (SRV,
     '        f.save(src)\n    except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna\n        shutil.rmtree(work, ignore_errors=True)\n        return jsonify({"error": f"gagal menyimpan file: {exc}"}), 422\n    now = int(time.time())\n    _save_job_state(job, {"job": job, "status": "queued",\n                          "filename": filename,\n',
     '        f.save(src)\n    except Exception as exc:  # noqa: BLE001 - pesan error ke pengguna\n        shutil.rmtree(work, ignore_errors=True)\n        print(f"gagal menyimpan upload: {exc}", flush=True)\n        return jsonify({"error": "gagal menyimpan file."}), 422\n    now = int(time.time())\n    _save_job_state(job, {"job": job, "status": "queued",\n                          "filename": filename,\n'),
    (SRV,
     'def _check_midtrans_status(order_id):\n    """Tanya status transaksi ke Midtrans; kembalikan dict respons."""\n    return _midtrans_request("GET", f"/v2/{order_id}/status")\n\n\ndef _mark_paid(job, rec):\n',
     'def _check_midtrans_status(order_id):\n    """Tanya status transaksi ke Midtrans; kembalikan dict respons."""\n    return _midtrans_request("GET", f"/v2/{order_id}/status")\n\n\ndef _amount_matches(st):\n    """True bila gross_amount respons Midtrans == PRICE_IDR (v0.6.30).\n    Midtrans mengirim string, mis. "15000" atau "15000.00". Bila\n    nominal tidak ada / tidak bisa dibaca / tidak sama -> False\n    (job JANGAN ditandai lunas)."""\n    try:\n        return float(str(st.get("gross_amount"))) == PRICE_IDR\n    except (TypeError, ValueError):\n        return False\n\n\ndef _mark_paid(job, rec):\n'),
    (SRV,
     '    if rec.get("order_id") and rec.get("status") == "pending":\n        try:\n            st = _check_midtrans_status(rec["order_id"])\n            if st.get("transaction_status") in ("settlement", "capture"):\n                _mark_paid(job, rec)\n                return jsonify({"job": job, "paid": True,\n                                "amount": PRICE_IDR}), 200\n        except MidtransError:\n            pass  # lanjut buat charge baru; status dicek lagi nanti\n\n    attempt = int(rec.get("attempt") or 0) + 1\n    order_id = f"daftarisiku-{job}" if attempt == 1 \\\n        else f"daftarisiku-{job}-{attempt}"\n',
     '    if rec.get("order_id") and rec.get("status") == "pending":\n        try:\n            st = _check_midtrans_status(rec["order_id"])\n            if st.get("transaction_status") in ("settlement", "capture") \\\n                    and _amount_matches(st):\n                _mark_paid(job, rec)\n                return jsonify({"job": job, "paid": True,\n                                "amount": PRICE_IDR}), 200\n            if st.get("transaction_status") in ("settlement", "capture"):\n                print(f"PERINGATAN nominal Midtrans tidak cocok "\n                      f"job {job}: {st.get(\'gross_amount\')}", flush=True)\n        except MidtransError:\n            pass  # lanjut buat charge baru; status dicek lagi nanti\n\n    # v0.6.30: batasi percobaan charge per job (anti banjir order_id).\n    if int(rec.get("attempt") or 0) >= MAX_QRIS_ATTEMPTS:\n        return jsonify(\n            {"error": "terlalu banyak percobaan bayar untuk dokumen ini"}), 429\n    attempt = int(rec.get("attempt") or 0) + 1\n    order_id = f"daftarisiku-{job}" if attempt == 1 \\\n        else f"daftarisiku-{job}-{attempt}"\n'),
    (SRV,
     '            try:\n                st = _check_midtrans_status(rec["order_id"])\n                tx = st.get("transaction_status")\n                if tx in ("settlement", "capture"):\n                    _mark_paid(job, rec)\n                    paid = True\n                elif tx in ("expire", "cancel", "deny"):\n                    rec["status"] = tx\n                    save_pay(job, rec)\n',
     '            try:\n                st = _check_midtrans_status(rec["order_id"])\n                tx = st.get("transaction_status")\n                if tx in ("settlement", "capture") and _amount_matches(st):\n                    _mark_paid(job, rec)\n                    paid = True\n                elif tx in ("settlement", "capture"):\n                    # v0.6.30: lunas tapi nominal tidak cocok -> JANGAN\n                    # tandai lunas; catat dan biarkan charge kedaluwarsa.\n                    print(f"PERINGATAN nominal Midtrans tidak cocok "\n                          f"job {job}: {st.get(\'gross_amount\')}", flush=True)\n                elif tx in ("expire", "cancel", "deny"):\n                    rec["status"] = tx\n                    save_pay(job, rec)\n'),
    (SRV,
     '        with open(rep_path, "w") as f:\n            json.dump(rec, f)\n    except OSError as exc:\n        return jsonify({"error": f"gagal menyimpan laporan: {exc}"}), 500\n    return jsonify({"reported": True, "job": job})\n\n\n@app.get("/api/reports")\ndef list_reports():\n    """Daftar laporan hasil salah (untuk pemilik/Muse mengecek)."""\n    out = []\n    try:\n        for name in sorted(os.listdir(RESULTS_DIR)):\n',
     '        with open(rep_path, "w") as f:\n            json.dump(rec, f)\n    except OSError as exc:\n        print(f"gagal menyimpan laporan job {job}: {exc}", flush=True)\n        return jsonify({"error": "gagal menyimpan laporan."}), 500\n    return jsonify({"reported": True, "job": job})\n\n\n@app.get("/api/reports")\ndef list_reports():\n    """Daftar laporan hasil salah (untuk pemilik/Muse mengecek).\n    v0.6.30: wajib token rahasia di header X-Reports-Token (env\n    REPORTS_TOKEN); tanpa token yang benar -> 401. Bila env belum\n    disetel, endpoint tertutup total (fail closed)."""\n    if not REPORTS_TOKEN \\\n            or request.headers.get("X-Reports-Token") != REPORTS_TOKEN:\n        return jsonify({"error": "tidak diizinkan"}), 401\n    out = []\n    try:\n        for name in sorted(os.listdir(RESULTS_DIR)):\n'),
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
        bak = path + ".bak-0.6.29"
        with open(bak, "w", encoding="utf-8") as f:
            f.write(src)
        print(f"cadangan: {bak}")
    for path, old, new in PAIRS:
        with open(path, encoding="utf-8") as f:
            src = f.read()
        with open(path, "w", encoding="utf-8") as f:
            f.write(src.replace(old, new, 1))
    print(f"OK: {len(PAIRS)} penggantian diterapkan.")
    code, out = run(f"python3 -m py_compile {SRV}")
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
    print("SELESAI: engine v0.6.30 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
