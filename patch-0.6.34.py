#!/usr/bin/env python3
"""Patch v0.6.34 untuk engine DaftarIsiKu di VPS Awan.

Simulasi bayar bertoken (pengganti /api/pay/mock yang ditutup v0.6.28):
  - server.py: endpoint baru POST /api/pay/simulate/<job> — tandai job
    lunas TANPA Midtrans, hanya saat Sandbox dan hanya dengan header
    X-Reports-Token = REPORTS_TOKEN (tanpa token yang benar -> 404;
    mati total bila MIDTRANS_IS_PRODUCTION=1). Penandaan lunas memakai
    _mark_paid yang sama seperti settlement Midtrans; catatan pay
    ditandai simulated=True. Endpoint /api/pay/mock yang lama TETAP
    tertutup dan tidak dibuka lagi. Website TIDAK berubah.
  - VERSION 0.6.33 -> 0.6.34.

Perintah user 24 Nov 2026: jangan pakai Midtrans dulu, pakai simulasi
yang awal saja; nanti bila review bisnis Midtrans selesai, kembali ke
Midtrans.

Jalankan SETELAH v0.6.33 ter-deploy (patch ini mengharapkan server.py
v0.6.33 di server).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.34.py -o /tmp/patch-0.6.34.py && python3 /tmp/patch-0.6.34.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.33. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    (SRV,
     '                               charge kedaluwarsa -> order_id baru.\n'
     '  GET  /api/qr/<job>        -> Gambar PNG kode QRIS charge yang sedang\n'
     ,
     '                               charge kedaluwarsa -> order_id baru.\n'
     '  POST /api/pay/simulate/<job> -> Simulasi bayar (v0.6.34, SANDBOX\n'
     '                               saja, hanya pemilik/Muse): tandai job\n'
     '                               lunas tanpa Midtrans. Wajib header\n'
     '                               X-Reports-Token = REPORTS_TOKEN (sama\n'
     '                               seperti /api/reports); tanpa token\n'
     '                               yang benar -> 404 (seolah endpoint\n'
     '                               tidak ada). Mati total bila\n'
     '                               MIDTRANS_IS_PRODUCTION=1 (-> 404).\n'
     '                               404 bila hasil belum ada; 409 bila\n'
     '                               sudah dibayar. TIDAK dipakai website;\n'
     '                               pelanggan tetap lewat QRIS.\n'
     '  GET  /api/qr/<job>        -> Gambar PNG kode QRIS charge yang sedang\n'
     ),
    (SRV,
     'VERSION = "0.6.33"  # v0.6.33: posisi "Kanan atas" (page_pos: top-right) — nomor halaman isi di header kanan atas, footer isi dikosongkan; "Kiri bawah" (left) dihapus dari website. Halaman depan + halaman pertama BAB tetap tengah. Perbaiki juga pemanggilan _ensure_bab_first_page_footer 2-argumen (TypeError laten v0.6.32). v0.6.32: pilihan posisi nomor halaman isi (page_pos: right/center/left; halaman depan + halaman pertama BAB tetap tengah). v0.6.31: anti-DoS + strip rel eksternal (audit keamanan 23 Nov): antrean maks 5 (penuh -> 429), rate limit upload 5/jam/IP (env MAX_UPLOADS_PER_HOUR), tolak .docx yang mengembang >100 MB total / >50 MB per file (zip bomb), hapus relasi eksternal dari salinan + hasil sebelum di-render LibreOffice. v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n'
     ,
     'VERSION = "0.6.34"  # v0.6.34: simulasi bayar bertoken POST /api/pay/simulate/<job> (sandbox saja; wajib header X-Reports-Token = REPORTS_TOKEN, tanpa token -> 404; mati total bila MIDTRANS_IS_PRODUCTION=1; pengganti /api/pay/mock yang ditutup v0.6.28 — mock lama TETAP tertutup, tidak dibuka lagi). v0.6.33: posisi "Kanan atas" (page_pos: top-right) — nomor halaman isi di header kanan atas, footer isi dikosongkan; "Kiri bawah" (left) dihapus dari website. Halaman depan + halaman pertama BAB tetap tengah. Perbaiki juga pemanggilan _ensure_bab_first_page_footer 2-argumen (TypeError laten v0.6.32). v0.6.32: pilihan posisi nomor halaman isi (page_pos: right/center/left; halaman depan + halaman pertama BAB tetap tengah). v0.6.31: anti-DoS + strip rel eksternal (audit keamanan 23 Nov): antrean maks 5 (penuh -> 429), rate limit upload 5/jam/IP (env MAX_UPLOADS_PER_HOUR), tolak .docx yang mengembang >100 MB total / >50 MB per file (zip bomb), hapus relasi eksternal dari salinan + hasil sebelum di-render LibreOffice. v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n'
     ),
    (SRV,
     '@app.route("/api/pay/qris/<job>", methods=["OPTIONS"])\n@app.route("/api/qr/<job>", methods=["OPTIONS"])\n'
     ,
     '@app.route("/api/pay/qris/<job>", methods=["OPTIONS"])\n'
     '@app.route("/api/pay/simulate/<job>", methods=["OPTIONS"])\n@app.route("/api/qr/<job>", methods=["OPTIONS"])\n'
     ),
    (SRV,
     '                    "status": "pending", "reused": False}), 201\n\n\n@app.get("/api/qr/<job>")\n'
     ,
     '                    "status": "pending", "reused": False}), 201\n\n\n@app.post("/api/pay/simulate/<job>")\n'
     'def simulate_pay(job):\n    """Simulasi bayar (v0.6.34): tandai job lunas TANPA Midtrans.\n'
     '    Hanya untuk pemilik/Muse saat Sandbox — dipakai mengetes jalur\n'
     '    download selagi review bisnis Midtrans belum lolos (perintah user\n'
     '    24 Nov 2026: jangan pakai Midtrans dulu, pakai simulasi yang awal\n'
     '    saja; nanti bila review selesai, kembali ke Midtrans). BUKAN\n'
     '    untuk pelanggan dan TIDAK dipakai website.\n    Pengaman (semua wajib, fail closed):\n'
     '      - MIDTRANS_IS_PRODUCTION=1 -> 404 (endpoint tidak ada).\n'
     '      - Header X-Reports-Token harus sama dengan REPORTS_TOKEN\n'
     '        (pemeriksaan yang sama seperti /api/reports); hilang/salah\n'
     '        atau env belum disetel -> 404, bukan 401, supaya keberadaan\n'
     '        endpoint tidak terungkap dari luar.\n      - File hasil job harus ada dan job belum dibayar (409 bila\n'
     '        sudah dibayar). Penandaan lunas memakai _mark_paid yang sama\n'
     '        seperti settlement Midtrans; catatan pay ditandai\n'
     '        simulated=True supaya jejaknya beda dari lunas Midtrans.\n'
     '    Endpoint /api/pay/mock yang lama (v0.6.28, tanpa token — siapa\n'
     '    pun yang tahu link job bisa menandai lunas gratis) TETAP dihapus\n'
     '    dan tidak dibuka lagi; simulasi ini penggantinya yang bertoken."""\n    if MIDTRANS_IS_PRODUCTION:\n'
     '        return jsonify({"error": "tidak ditemukan"}), 404\n    if not REPORTS_TOKEN \\\n'
     '            or request.headers.get("X-Reports-Token") != REPORTS_TOKEN:\n'
     '        return jsonify({"error": "tidak ditemukan"}), 404\n    if not job.isalnum() or len(job) > 32:\n'
     '        return jsonify({"error": "job tidak valid"}), 400\n'
     '    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")\n    if not os.path.exists(path):\n'
     '        return jsonify({"error": "hasil tidak ditemukan"}), 404\n    if job in load_paid():\n'
     '        return jsonify({"error": "dokumen ini sudah dibayar"}), 409\n    rec = load_pay(job)\n'
     '    rec["simulated"] = True\n    _mark_paid(job, rec)\n'
     '    return jsonify({"job": job, "paid": True, "simulated": True,\n'
     '                    "amount": PRICE_IDR}), 200\n\n\n@app.get("/api/qr/<job>")\n'
     ),
]

def run(cmd):
    r = subprocess.run(cmd, shell=True, capture_output=True, text=True,
                       timeout=120)
    return r.returncode, (r.stdout + r.stderr).strip()


def apply_pairs():
    """Verifikasi semua pola (harus cocok TEPAT SATU kali), cadangkan
    file asli, lalu terapkan. Kembalikan 0 bila sukses, 1 bila ada
    pola yang tidak cocok — tanpa mengubah apa pun."""
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
        bak = path + ".bak-0.6.33"
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
    return 0


def main():
    if apply_pairs() != 0:
        return 1
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
    print("SELESAI: engine v0.6.34 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
