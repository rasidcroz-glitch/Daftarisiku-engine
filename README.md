# DaftarIsiKu Engine

Mesin pembuat daftar isi otomatis untuk file Word (.docx) — nomor halaman
akurat dihitung dari hasil render PDF dokumen itu sendiri.

- Upload `.docx` (skripsi, makalah, laporan, proposal) → daftar isi rapi
  mengikuti format contoh dosen (halaman depan angka Romawi, BAB tebal
  huruf kapital, sub-bab A./B., sub-sub 1./2., sub-sub-sub a./b./c.)
- Daftar Gambar & Daftar Tabel ikut dibuat bila ada caption bernomor
- API: `POST /api/toc`, `POST /api/pay/mock/<job>`, `GET /api/download/<job>`
- Stack: Flask + python-docx + PyMuPDF + LibreOffice headless (Docker)

Deploy: Render Blueprint — file `render.yaml` (plan free, Docker).
Website: https://daftarisiku-web.vercel.app
