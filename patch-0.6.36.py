#!/usr/bin/env python3
"""Patch v0.6.36 untuk engine DaftarIsiKu di VPS Awan.

Perbaiki footer Romawi dokumen 1 section yang BAB I-nya mulai di halaman
baru TANPA break eksplisit (SKRIPSI ACA, dilapor user 24 Nov 2026 dan
gagal di v0.6.35: footer hal. 2-7 tetap Arab "2"-"7", bukan Romawi ii-vii):
  - toc.py (add_footer_page_numbers): sectPr pindahan 1-section kini
    diubah jadi nextPage bila bab_new_page True (terverifikasi PDF oleh
    process(): BAB I baris isi pertama halamannya di render pra-footer,
    jadi break nextPage bersifat netral pagination — BAB I sudah di
    puncak halaman baru — dan LibreOffice baru menerapkan lowerRoman).
    Sebelumnya nextPage hanya bila ada break eksplisit (_stripped /
    pageBreakBefore); tanpa itu sectPr tetap continuous dan LibreOffice
    mengabaikan lowerRoman. Fallback continuous dipertahankan bila BAB I
    tidak terverifikasi mulai di halaman baru.
  - VERSION 0.6.35 -> 0.6.36.

Jalankan SETELAH v0.6.35 ter-deploy (patch ini mengharapkan toc.py
dan server.py v0.6.35 di server).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.36.py -o /tmp/patch-0.6.36.py && python3 /tmp/patch-0.6.36.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.35. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys
import time
import urllib.request
import json

SRV = "/opt/daftarisiku-engine/server.py"
TOC = "/opt/daftarisiku-engine/toc.py"

PAIRS = [
    (TOC,
     '                        # v0.6.35: ubah sectPr pindahan jadi nextPage.\n                        # LibreOffice MENGABAIKAN pgNumType\n                        # (lowerRoman) pada section yang ditutup\n                        # break continuous — footer depan tetap Arab\n                        # ("2","3") — tapi menerapkannya pada break\n                        # nextPage (dibuktikan eksperimen minimal:\n                        # continuous -> angka Arab, nextPage -> "ii").\n                        # Break-nya tetap TEPAT SATU: nextPage ini\n                        # MENGGANTIKAN satu break eksplisit yang\n                        # sudah ada — run break yang baru dibuang,\n                        # atau pageBreakBefore judul (dihapus).\n                        # Bila tidak ada break eksplisit (BAB I\n                        # kebetulan di halaman baru), biarkan\n                        # continuous supaya pagination tidak berubah\n                        # (depan tetap Arab — batasan).\n                        _mt = _bsp.find(qn("w:type"))\n                        if _mt is None:\n                            _mt = OxmlElement("w:type")\n                            _bsp.insert(0, _mt)\n                        if (_stripped\n                                or bab_para.paragraph_format\n                                .page_break_before):\n',
     '                        # v0.6.36: ubah sectPr pindahan jadi nextPage.\n                        # LibreOffice MENGABAIKAN pgNumType\n                        # (lowerRoman) pada section yang ditutup\n                        # break continuous — footer depan tetap Arab\n                        # ("2","3") — tapi menerapkannya pada break\n                        # nextPage (dibuktikan eksperimen minimal:\n                        # continuous -> angka Arab, nextPage -> "ii").\n                        # v0.6.35 hanya mengubah ke nextPage bila\n                        # ada break eksplisit (_stripped /\n                        # pageBreakBefore); SKRIPSI ACA gagal karena\n                        # BAB I-nya mulai di halaman baru HANYA\n                        # karena aliran pagination (tanpa break\n                        # eksplisit), sehingga sectPr pindahan tetap\n                        # continuous dan footer depan tetap Arab\n                        # "2"–"7" walau label TOC sudah Romawi.\n                        # bab_new_page di sini terverifikasi PDF\n                        # oleh process() (BAB I baris isi pertama\n                        # halamannya di render pra-footer), jadi\n                        # break nextPage sebelum judul bersifat\n                        # netral pagination — BAB I sudah di puncak\n                        # halaman baru — dan LibreOffice menerapkan\n                        # lowerRoman. Fallback continuous hanya bila\n                        # BAB I tidak terverifikasi mulai di halaman\n                        # baru (cabang else di bawah; tidak terjangkau\n                        # bila gate _moved sudah mensyaratkan\n                        # bab_new_page True — dipertahankan defensif).\n                        _mt = _bsp.find(qn("w:type"))\n                        if _mt is None:\n                            _mt = OxmlElement("w:type")\n                            _bsp.insert(0, _mt)\n                        if (_stripped\n                                or bab_para.paragraph_format\n                                .page_break_before\n                                or bab_new_page):\n'),
    (SRV,
     'VERSION = "0.6.35"  # v0.6.35: footer Romawi dokumen 1 section — sectPr continuous di paragraf BAB I bila tepat 1 section + _has_front + BAB I mulai di halaman baru (makalah: depan Romawi, bab1_footer=1). v0.6.34: simulasi bayar bertoken POST /api/pay/simulate/<job> (sandbox saja; wajib header X-Reports-Token = REPORTS_TOKEN, tanpa token -> 404; mati total bila MIDTRANS_IS_PRODUCTION=1; pengganti /api/pay/mock yang ditutup v0.6.28 — mock lama TETAP tertutup, tidak dibuka lagi). v0.6.33: posisi "Kanan atas" (page_pos: top-right) — nomor halaman isi di header kanan atas, footer isi dikosongkan; "Kiri bawah" (left) dihapus dari website. Halaman depan + halaman pertama BAB tetap tengah. Perbaiki juga pemanggilan _ensure_bab_first_page_footer 2-argumen (TypeError laten v0.6.32). v0.6.32: pilihan posisi nomor halaman isi (page_pos: right/center/left; halaman depan + halaman pertama BAB tetap tengah). v0.6.31: anti-DoS + strip rel eksternal (audit keamanan 23 Nov): antrean maks 5 (penuh -> 429), rate limit upload 5/jam/IP (env MAX_UPLOADS_PER_HOUR), tolak .docx yang mengembang >100 MB total / >50 MB per file (zip bomb), hapus relasi eksternal dari salinan + hasil sebelum di-render LibreOffice. v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n',
     'VERSION = "0.6.36"  # v0.6.36: footer Romawi dokumen 1 section yang BAB I-nya mulai di halaman baru TANPA break eksplisit (SKRIPSI ACA 24 Nov 2026: sectPr pindahan 1-section kini jadi nextPage bila bab_new_page True dari PDF pra-footer — break nextPage itu netral pagination karena BAB I sudah di puncak halaman baru, dan LibreOffice baru menerapkan lowerRoman; sebelumnya fallback continuous membuat footer depan tetap Arab "2"–"7" walau label TOC Romawi dan verifikasi "lolos"). v0.6.35: footer Romawi dokumen 1 section — sectPr continuous di paragraf BAB I bila tepat 1 section + _has_front + BAB I mulai di halaman baru (makalah: depan Romawi, bab1_footer=1). v0.6.34: simulasi bayar bertoken POST /api/pay/simulate/<job> (sandbox saja; wajib header X-Reports-Token = REPORTS_TOKEN, tanpa token -> 404; mati total bila MIDTRANS_IS_PRODUCTION=1; pengganti /api/pay/mock yang ditutup v0.6.28 — mock lama TETAP tertutup, tidak dibuka lagi). v0.6.33: posisi "Kanan atas" (page_pos: top-right) — nomor halaman isi di header kanan atas, footer isi dikosongkan; "Kiri bawah" (left) dihapus dari website. Halaman depan + halaman pertama BAB tetap tengah. Perbaiki juga pemanggilan _ensure_bab_first_page_footer 2-argumen (TypeError laten v0.6.32). v0.6.32: pilihan posisi nomor halaman isi (page_pos: right/center/left; halaman depan + halaman pertama BAB tetap tengah). v0.6.31: anti-DoS + strip rel eksternal (audit keamanan 23 Nov): antrean maks 5 (penuh -> 429), rate limit upload 5/jam/IP (env MAX_UPLOADS_PER_HOUR), tolak .docx yang mengembang >100 MB total / >50 MB per file (zip bomb), hapus relasi eksternal dari salinan + hasil sebelum di-render LibreOffice. v0.6.30: pengerasan bayar + laporan (audit keamanan 23 Nov): tandai lunas hanya bila gross_amount Midtrans == Rp15.000; batasi charge QRIS maks 5/job; /api/reports wajib token rahasia (env REPORTS_TOKEN); pesan error ke pelanggan generik (detail hanya di log server); security headers di semua respons engine. v0.6.28: endpoint /api/pay/mock DIHAPUS (celah keamanan: siapa pun yang tahu link job bisa menandai lunas gratis); website sudah tidak memakainya sejak PAY_MOCK=false, bayar hanya lewat QRIS Midtrans. v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26: daftar isi diletakkan sebelum judul BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25: gaya penomoran daftar isi mengikuti gaya dokumen sendiri — nomor dirender Word di isi dokumen (simulasi penghitung list Word atas semua paragraf: \'1.1\' bertitik, \'I.\' Romawi, \'A.\' huruf; jatuh ke awalan literal judul bila tanpa numPr, lalu format contoh dosen). v0.6.24: judul DAFTAR ISI yang membawa sectPr sendiri kini sectPr-nya dihapus (bila format nomor sama) + page_break_before, supaya judul+entri mulai di halaman baru (cacat UJI_COBA_2: judul yatim di halaman sendiri, footer jadi Arab "10"). v0.6.23: label TOC dari PDF hasil dihitung dari baris judul (bukan substring) + bab1 = halaman Arab \'1\' (bukan halaman judul BAB I — bila BAB I mulai di tengah halaman, halaman judulnya masih Romawi; label dihitung dari footer, diklem minimal 1) + strip nomor gaya \'1)\' + gabung judul 3 baris — perbaikan cacat UJI_COBA_2 (BAB II–V 5 halaman terlalu kecil karena paragraf Sistematika Penulisan menyebut semua judul BAB) dan Gusti Ayu (sub-judul indikator \'1)\' + BAB I tengah halaman). v0.6.22: AI Document Checker — check_document() di toc.py (aturan struktur, tanpa API, laporan saja): judul BAB tidak konsisten (Romawi vs Arab), penomoran sub-bab loncat / sub-sub tanpa induk, tabel/gambar tanpa keterangan bernomor, entri yang nomor halamannya tak terverifikasi. Hasilnya di payload "checker_warnings"; website menampilkan kotaknya di pratinjau. v0.6.21: tombol "Laporkan hasil salah" — POST /api/report/<job> (satu laporan per job, TIDAK membuka download; disimpan {job}.report.json, ikut kedaluwarsa 24 jam) + GET /api/reports (daftar laporan untuk pemilik/Muse). Website: tombol ghost di layar pratinjau + prompt deskripsi opsional + status konfirmasi. v0.6.20: pilihan download PDF atau Word — satu bayar Rp15.000 membuka keduanya (process() kini mengembalikan output_pdf = render akhir yang cocok dengan docx final; worker menyimpannya sebagai {job}-hasil-daftar-isi.pdf; /api/download/<job>?format=pdf mengunduh PDF dengan nama file asli + .pdf). v0.6.19: (1) nomor double — footer yang field PAGE-nya berada di dalam <w:sdt> (\'Page Numbers\' Word) kini terdeteksi, jadi mesin tidak menambah field kedua; bila ada >1 field, sisanya dibuang. (2) daftar isi pisah — section break nextPage tepat setelah judul DAFTAR ISI/GAMBAR/TABEL diubah jadi continuous (hanya bila fmt pgNumType kedua sisi sama) supaya judul+entri satu halaman. (3) hemat memori — Document sumber dan teks PDF masukan dibebaskan (del + gc + trim) sebelum tiap render LibreOffice, supaya total tidak melewati 512 MB container dan job tidak hilang (404) saat polling. v0.6.18: RE_CAPTION kini menerima nomor "Gambar 2. 1" (titik+spasi; dulu hanya "2.1"/"2 1") — caption UJI_COBA_2 yang 56 buah kini terdeteksi, Daftar Gambar/Tabel tidak lagi kosong; heading template daftar yang kosong tetap diberi page_break_before supaya tidak menempel di daftar isi. v0.6.17: /api/toc ASINKRON — upload langsung balas 202 {job, status:"queued"}; proses jalan di satu worker thread latar (antrean FIFO, satu dokumen dalam satu waktu) dan status dicek lewat /api/status (queued/processing/done/error + hasil pratinjau bila done).\n'),

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
        with open(path + ".bak-0.6.35", "w", encoding="utf-8") as f:
            f.write(src)
        for p, old, new in PAIRS:
            if p == path:
                src = src.replace(old, new)
        with open(path, "w", encoding="utf-8") as f:
            f.write(src)
        print(f"OK: {path} di-patch (backup: {path}.bak-0.6.35)")
    return 0


def main():
    if apply_pairs() != 0:
        return 1
    for path in (TOC, SRV):
        rc, out = run(f"python3 -m py_compile {path}")
        if rc != 0:
            print(f"GAGAL: py_compile {path}:\n{out}")
            return 1
    print("OK: py_compile lolos")
    rc, out = run("systemctl restart daftarisiku")
    if rc != 0:
        print(f"GAGAL: systemctl restart daftarisiku:\n{out}")
        return 1
    print("OK: service daftarisiku di-restart")
    for i in range(12):
        time.sleep(5)
        try:
            with urllib.request.urlopen(
                    "http://127.0.0.1:8080/api/health",
                    timeout=10) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception as e:  # noqa: BLE001 - retry sampai service bangun
            print(f"health coba {i + 1}: belum bisa ({e})")
            continue
        print(f"health: {data}")
        if data.get("version") == "0.6.36" and data.get("status") == "ok":
            print("OK: engine v0.6.36 LIVE")
            return 0
        print(f"health coba {i + 1}: versi/status belum cocok: {data}")
    print("GAGAL: /api/health tidak menunjukkan v0.6.36 setelah 60 detik")
    return 1


if __name__ == "__main__":
    sys.exit(main())
