# Pembayaran Kasera Pay — DaftarIsiKu (engine v0.6.37)

Pengganti Midtrans (akun Midtrans DITOLAK 25 Nov 2026 untuk 4 metode
bayar). Alur sama seperti Midtrans dulu: pelanggan klik bayar ->
muncul QRIS Rp15.000 -> scan pakai DANA/GoPay/OVO/m-banking apa saja
-> bayar -> download .docx/.pdf langsung terbuka otomatis, tanpa cek
manual. Jalur Midtrans + simulasi TIDAK dihapus — Kasera hanya opsi
baru (default di website), Midtrans tetap bisa dipilih di layar bayar.

## Endpoint engine (server.py)

- `POST /api/pay/kasera/<job>` — buat pembayaran Kasera Rp15.000.
  Balasan: `{job, payment_id, amount: 15000, qr_url,
  checkout_url, expires_in_seconds, status: "pending",
  reused: bool}`. Pembayaran yang masih berlaku dipakai ulang (QR
  yang sama); `Idempotency-Key: daftarisiku-<job>-<attempt>` disimpan
  SEBELUM API dipanggil. 503 bila `KASERA_PAY_KEY` belum disetel;
  409 bila job sudah dibayar; 429 bila sudah 5x percobaan.
- `POST /api/pay/kasera/webhook` — webhook Kasera. Verifikasi
  `Kasera-Signature-V1` (HMAC-SHA256 `<unix>.<raw body>`, toleransi
  5 menit); tanda tangan salah -> 401; secret belum disetel -> 503.
  Hanya event `payment.paid` yang melunasi, dan hanya bila nominal
  == Rp15.000 dan `payment_request_id` cocok dengan pembayaran yang
  dibuat server ini. Dedupe id event (pengiriman at-least-once).
- `GET /api/pay/kasera/<job>/qr` — PNG kode QRIS Kasera yang aktif.
- `GET /api/status/<job>` — kini juga membawa `kasera`
  (`{payment_id, status, checkout_url, expires_in_seconds, qr_url}`)
  dan `kasera_expired`; saat Kasera masih pending, engine menanyakan
  status ke Kasera (dibatasi 1x/10 detik; Kasera membatasi baca
  1x/detik per transaksi) dan menandai lunas bila `succeeded`.

## Konfigurasi (VPS — JANGAN tulis key di kode/chat/git)

Di systemd drop-in service `daftarisiku` (atau env file mode 0600):

```
KASERA_PAY_KEY=kp_test_...            # sandbox; nanti kp_live_...
KASERA_PAY_WEBHOOK_SECRET=<secret>   # dari dashboard Kasera
```

Key dibuat di dashboard Kasera (pay.kasera.id) menu Developer.
Webhook didaftarkan di dashboard (mode test dan live punya endpoint
+ secret BERBEDA):

```
https://engine.daftarisiku.web.id/api/pay/kasera/webhook
```

## Uji sandbox (butuh kp_test_ + secret test)

1. Isi `KASERA_PAY_KEY` + `KASERA_PAY_WEBHOOK_SECRET` (mode test) di
   VPS, restart service.
2. Di website: upload dokumen uji -> layar bayar (metode default
   "QRIS Otomatis") -> QR muncul.
3. Ambil token dari `checkout_url` (bagian setelah `/p/`), lalu:
   `curl -X POST https://pay.kasera.id/api/v1/checkout/<token>/simulate-payment -H "Content-Type: application/json" -d '{"outcome": "succeeded"}'`
4. Webhook `payment.paid` terkirim ke engine -> download terbuka.
   Ulangi dengan `"outcome": "expired"` untuk jalur kedaluwarsa.

## Go live

1. Selesaikan verifikasi usaha di dashboard Kasera (KTP + rekening
   atas nama sendiri). Sebelum verifikasi selesai: akun hanya bisa
   menerima 10 pembayaran, total Rp1.000.000.
2. Ganti key ke `kp_live_` dan secret ke secret mode live. Tidak ada
   perubahan kode lain.
3. Pastikan webhook mode live menunjuk URL di atas.

## Biaya (per 25 Nov 2026)

QRIS 0,7% + Rp250 per transaksi -> Rp15.000 = fee Rp355,
diterima bersih Rp14.645. Pencairan ke rekening Rp3.000/pencairan.
Tidak ada biaya bulanan.

## Tes otomatis

`tests/test_kasera.py` — server Kasera tiruan di localhost + Flask
test client: create, reuse QR, QR PNG, webhook valid/duplikat/
tanda tangan salah/timestamp tua/nominal salah/payment asing,
polling status lewat `/api/status`, 503 tanpa key, 409 bila lunas,
404 bila job tidak ada.
