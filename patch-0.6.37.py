#!/usr/bin/env python3
"""Patch v0.6.37 untuk engine DaftarIsiKu di VPS Awan.

Integrasi Kasera Pay (pengganti Midtrans yang ditolak 25 Nov 2026):
  - POST /api/pay/kasera/<job> — buat pembayaran QRIS Rp15.000
    (Idempotency-Key daftarisiku-<job>-<attempt>, reuse QR pending,
    503 bila KASERA_PAY_KEY belum disetel, 409 bila sudah dibayar).
  - POST /api/pay/kasera/webhook — verifikasi HMAC-SHA256
    Kasera-Signature-V1, dedupe id event, hanya payment.paid yang
    melunasi (nominal harus Rp15.000); tanda tangan salah -> 401.
  - GET /api/pay/kasera/<job>/qr — PNG kode QRIS Kasera.
  - GET /api/status/<job> — kini membawa info kasera +
    kasera_expired dan polling status Kasera (1x/10 detik).
  - Helper _qr_png dipakai /api/qr/<job> (Midtrans) dan QR Kasera.
  - Midtrans + simulasi TIDAK diubah. VERSION 0.6.36 -> 0.6.37.

Jalankan SETELAH v0.6.36 ter-deploy (patch ini mengharapkan
server.py v0.6.36 di server).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.37.py -o /tmp/patch-0.6.37.py && python3 /tmp/patch-0.6.37.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.36. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys
import time
import urllib.request
import json

SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    (SRV,
     'import base64\nimport io\nimport json\n',
     'import base64\nimport hashlib\nimport hmac\nimport io\nimport json\n'),
    (SRV,
     'VERSION = "0.6.36"  # v0.6.36: ',
     'VERSION = "0.6.37"  # v0.6.37: integrasi Kasera Pay — POST /api/pay/kasera/<job> (QRIS Rp15.000 otomatis; kunci dari env KASERA_PAY_KEY, 503 bila belum disetel), webhook POST /api/pay/kasera/webhook (verifikasi HMAC-SHA256 Kasera-Signature-V1, dedupe id event, hanya payment.paid yang melunasi), GET /api/pay/kasera/<job>/qr (PNG QRIS Kasera), /api/status kini membawa info kasera + polling status Kasera. Midtrans + simulasi TIDAK diubah. v0.6.36: '),
    (SRV,
     'MAX_QRIS_ATTEMPTS = 5\nREPORTS_TOKEN = os.environ.get("REPORTS_TOKEN", "")\n',
     'MAX_QRIS_ATTEMPTS = 5\nREPORTS_TOKEN = os.environ.get("REPORTS_TOKEN", "")\n\n# --- konfigurasi Kasera Pay (v0.6.37; kunci HANYA dari env) --------\n# KASERA_PAY_KEY: kp_test_... (sandbox) atau kp_live_... (produksi),\n# dibuat di dashboard Kasera menu Developer. KASERA_PAY_WEBHOOK_SECRET:\n# signing secret endpoint webhook (berbeda per mode test/live).\n# Kosong = jalur Kasera mati (endpoint bayar menjawab 503, webhook 503).\nKASERA_PAY_KEY = os.environ.get("KASERA_PAY_KEY", "")\nKASERA_PAY_WEBHOOK_SECRET = os.environ.get(\n    "KASERA_PAY_WEBHOOK_SECRET", "")\nKASERA_PAY_BASE_URL = os.environ.get(\n    "KASERA_PAY_BASE_URL", "https://pay.kasera.id")\n'),
    (SRV,
     'def _pay_path(job):\n    return os.path.join(RESULTS_DIR, f"{job}.pay.json")',
     '# --- Kasera Pay (v0.6.37; QRIS otomatis pengganti Midtrans) -----------\n# Kontrak: https://pay.kasera.id/en/docs/api — POST /v1/transactions\n# (Bearer kp_test_/kp_live_, Idempotency-Key, amount rupiah utuh,\n# payment_methods ["qris"]) -> payment {type "qr", qr_string},\n# checkout_url, expires_at (ISO-8601 +07:00). Webhook payment.paid:\n# header Kasera-Signature-V1 "t=<unix>,v1=<hex>" = HMAC-SHA256\n# "<unix>.<raw body>" dengan signing secret endpoint; tolak bila\n# |now-t| > 300 detik; dedupe pada id event (at-least-once).\nclass KaseraPayError(Exception):\n    pass\n\n\ndef _kasera_request(method, path, body=None, idempotency_key=None,\n                    timeout=15):\n    """Panggil API Kasera Pay; kembalikan dict JSON.\n    Melempar KaseraPayError bila kunci belum disetel, jaringan gagal,\n    atau Kasera menjawab error."""\n    if not KASERA_PAY_KEY:\n        raise KaseraPayError("pembayaran Kasera belum diaktifkan di server")\n    url = KASERA_PAY_BASE_URL.rstrip("/") + path\n    headers = {"Authorization": f"Bearer {KASERA_PAY_KEY}",\n               "Accept": "application/json",\n               "Content-Type": "application/json"}\n    if idempotency_key:\n        headers["Idempotency-Key"] = idempotency_key\n    data = json.dumps(body).encode() if body is not None else None\n    req = urllib.request.Request(url, data=data, headers=headers,\n                                 method=method)\n    try:\n        with urllib.request.urlopen(req, timeout=timeout) as resp:\n            return json.loads(resp.read().decode() or "{}")\n    except urllib.error.HTTPError as exc:\n        try:\n            detail = json.loads(exc.read().decode() or "{}")\n            err = detail.get("error") or {}\n            msg = err.get("message") or err.get("code") \\\n                or f"HTTP {exc.code}"\n        except (ValueError, UnicodeDecodeError):\n            msg = f"HTTP {exc.code}"\n        raise KaseraPayError(f"Kasera Pay: {msg} (HTTP {exc.code})") \\\n            from exc\n    except (urllib.error.URLError, TimeoutError, OSError) as exc:\n        raise KaseraPayError(\n            "tidak bisa menghubungi Kasera Pay, coba lagi sebentar lagi"\n        ) from exc\n\n\ndef _parse_kasera_expiry(expires_at, fallback_epoch):\n    """\'2026-11-25T14:30:00+07:00\' (ISO-8601) -> epoch;\n    gagal parse -> fallback."""\n    if expires_at:\n        try:\n            dt = datetime.fromisoformat(\n                str(expires_at).replace("Z", "+00:00"))\n            return int(dt.timestamp())\n        except ValueError:\n            pass\n    return fallback_epoch\n\n\ndef _kasera_verify_signature(raw, header, secret):\n    """True bila Kasera-Signature-V1 sah: \'t=<unix>,v1=<hex>[,v1=<hex>]\',\n    v1 = HMAC-SHA256(\'<unix>.<raw body>\'); toleransi 5 menit; cukup\n    satu v1 yang cocok (dua v1 selama 24 jam setelah rotasi secret)."""\n    parts = (header or "").split(",")\n    if not parts or not parts[0].startswith("t="):\n        return False\n    try:\n        t = int(parts[0][2:])\n    except ValueError:\n        return False\n    if abs(time.time() - t) > 300:\n        return False\n    signed = f"{t}.".encode() + raw\n    expected = hmac.new(secret.encode(), signed,\n                        hashlib.sha256).hexdigest()\n    return any(hmac.compare_digest(p[3:], expected)\n               for p in parts[1:] if p.startswith("v1="))\n\n\ndef _kasera_info(job, rec):\n    """Ringkasan pembayaran Kasera untuk website (tanpa qr_string)."""\n    if not rec or not rec.get("kasera_payment_id"):\n        return None\n    now = int(time.time())\n    expires_at = rec.get("kasera_expires_at") or now\n    return {"payment_id": rec["kasera_payment_id"],\n            "status": rec.get("kasera_status", "pending"),\n            "checkout_url": rec.get("kasera_checkout_url"),\n            "expires_in_seconds": max(0, expires_at - now),\n            "qr_url": f"/api/pay/kasera/{job}/qr"}\n\n\ndef _kasera_find_job(payment_request_id, external_id):\n    """Cari job dari webhook: cocokkan payment_request_id yang disimpan\n    saat create (otoritatif; external_id boleh kembar). Fallback ke\n    external_id (= job) hanya bila catatannya membawa payment id yang\n    sama. None bila bukan pembayaran server ini."""\n    if payment_request_id:\n        try:\n            for name in os.listdir(RESULTS_DIR):\n                if not name.endswith(".pay.json"):\n                    continue\n                job = name[:-len(".pay.json")]\n                if load_pay(job).get("kasera_payment_id") \\\n                        == payment_request_id:\n                    return job\n        except OSError:\n            pass\n    if external_id and external_id.isalnum() and len(external_id) <= 32:\n        rec = load_pay(external_id)\n        if rec.get("kasera_payment_id") == payment_request_id:\n            return external_id\n    return None\n\n\ndef _pay_path(job):\n    return os.path.join(RESULTS_DIR, f"{job}.pay.json")'),
    (SRV,
     '@app.route("/api/pay/simulate/<job>", methods=["OPTIONS"])\n@app.route("/api/qr/<job>", methods=["OPTIONS"])',
     '@app.route("/api/pay/simulate/<job>", methods=["OPTIONS"])\n@app.route("/api/pay/kasera/<job>", methods=["OPTIONS"])\n@app.route("/api/pay/kasera/webhook", methods=["OPTIONS"])\n@app.route("/api/pay/kasera/<job>/qr", methods=["OPTIONS"])\n@app.route("/api/qr/<job>", methods=["OPTIONS"])'),
    (SRV,
     '@app.get("/api/qr/<job>")\ndef qris_image(job):',
     'def _qr_png(qr_string):\n    """PNG kode QR dari qr_string (EMV QRIS); None bila library qrcode\n    tidak tersedia."""\n    try:\n        import qrcode\n    except ImportError:\n        return None\n    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M,\n                       box_size=8, border=2)\n    qr.add_data(qr_string)\n    qr.make(fit=True)\n    img = qr.make_image(fill_color="black", back_color="white")\n    buf = io.BytesIO()\n    img.save(buf, format="PNG")\n    buf.seek(0)\n    return buf\n\n\n@app.get("/api/qr/<job>")\ndef qris_image(job):'),
    (SRV,
     '    try:\n        import qrcode\n        qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_M,\n                           box_size=8, border=2)\n        qr.add_data(qr_string)\n        qr.make(fit=True)\n        img = qr.make_image(fill_color="black", back_color="white")\n        buf = io.BytesIO()\n        img.save(buf, format="PNG")\n        buf.seek(0)\n        return send_file(buf, mimetype="image/png",\n                         download_name=f"qris-{job}.png")\n    except ImportError:\n        pass\n',
     '    buf = _qr_png(qr_string)\n    if buf is not None:\n        return send_file(buf, mimetype="image/png",\n                         download_name=f"qris-{job}.png")\n'),
    (SRV,
     '    rec = load_pay(job)\n    rec["simulated"] = True\n    _mark_paid(job, rec)\n    return jsonify({"job": job, "paid": True, "simulated": True,\n                    "amount": PRICE_IDR}), 200\n',
     '    rec = load_pay(job)\n    rec["simulated"] = True\n    _mark_paid(job, rec)\n    return jsonify({"job": job, "paid": True, "simulated": True,\n                    "amount": PRICE_IDR}), 200\n\n\n@app.post("/api/pay/kasera/<job>")\ndef kasera_pay(job):\n    """Buat (atau pakai ulang) pembayaran QRIS Kasera Rp15.000 untuk\n    sebuah job. Idempotency-Key \'daftarisiku-<job>-<attempt>\' dibuat\n    SEKALI per percobaan dan disimpan di pay.json SEBELUM API\n    dipanggil, jadi percobaan ulang setelah timeout tidak membuat\n    tagihan kedua. Pembayaran yang masih berlaku dipakai ulang (QR\n    yang sama); yang kedaluwarsa dicek sekali ke Kasera, siapa tahu\n    pelanggan sempat bayar di detik-detik terakhir. 503 bila\n    KASERA_PAY_KEY belum disetel; 409 bila job sudah dibayar."""\n    if not job.isalnum() or len(job) > 32:\n        return jsonify({"error": "job tidak valid"}), 400\n    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")\n    if not os.path.exists(path):\n        return jsonify({"error": "hasil tidak ditemukan"}), 404\n    if job in load_paid():\n        return jsonify({"error": "dokumen ini sudah dibayar"}), 409\n    if not KASERA_PAY_KEY:\n        return jsonify(\n            {"error": "pembayaran Kasera belum diaktifkan di server"}), 503\n\n    now = int(time.time())\n    rec = load_pay(job)\n    # Pembayaran Kasera masih berlaku -> pakai ulang (QR yang sama).\n    if rec.get("kasera_payment_id") \\\n            and rec.get("kasera_status") == "pending" \\\n            and (rec.get("kasera_expires_at") or 0) > now \\\n            and rec.get("kasera_qr_string"):\n        info = _kasera_info(job, rec)\n        return jsonify({"job": job, "payment_id": rec["kasera_payment_id"],\n                        "amount": PRICE_IDR,\n                        "qr_url": info["qr_url"],\n                        "checkout_url": rec.get("kasera_checkout_url"),\n                        "expires_in_seconds": info["expires_in_seconds"],\n                        "status": "pending", "reused": True}), 201\n    # Kedaluwarsa: cek sekali ke Kasera, siapa tahu sudah lunas.\n    if rec.get("kasera_payment_id") \\\n            and rec.get("kasera_status") == "pending":\n        try:\n            st = _kasera_request(\n                "GET", f"/v1/transactions/{rec[\'kasera_payment_id\']}")\n            if st.get("status") == "succeeded" \\\n                    and int(st.get("amount") or 0) == PRICE_IDR:\n                rec["kasera_status"] = "succeeded"\n                _mark_paid(job, rec)\n                return jsonify({"job": job, "paid": True,\n                                "amount": PRICE_IDR}), 200\n            if st.get("status") == "succeeded":\n                print(f"PERINGATAN nominal Kasera tidak cocok "\n                      f"job {job}: {st.get(\'amount\')}", flush=True)\n            elif st.get("status") in ("expired", "failed", "canceled"):\n                rec["kasera_status"] = st["status"]\n                save_pay(job, rec)\n        except KaseraPayError:\n            pass  # lanjut buat pembayaran baru; status dicek lagi nanti\n        except (TypeError, ValueError):\n            pass\n\n    if int(rec.get("kasera_attempt") or 0) >= MAX_QRIS_ATTEMPTS:\n        return jsonify(\n            {"error": "terlalu banyak percobaan bayar untuk dokumen ini"}), 429\n    attempt = int(rec.get("kasera_attempt") or 0) + 1\n    idem_key = f"daftarisiku-{job}-{attempt}"\n    rec["kasera_attempt"] = attempt\n    rec["kasera_idempotency_key"] = idem_key\n    save_pay(job, rec)  # kunci disimpan SEBELUM API dipanggil\n    try:\n        with open(os.path.join(RESULTS_DIR, f"{job}.meta.json")) as mf:\n            filename = (json.load(mf) or {}).get("filename") or job\n    except (FileNotFoundError, json.JSONDecodeError, OSError):\n        filename = job\n    body = {"amount": PRICE_IDR,\n            "description": f"DaftarIsiKu: {filename}"[:120],\n            "external_id": job,\n            "payment_methods": ["qris"]}\n    try:\n        resp = _kasera_request("POST", "/v1/transactions", body,\n                               idempotency_key=idem_key)\n    except KaseraPayError as exc:\n        return jsonify({"error": str(exc)}), 502\n    payment = resp.get("payment") or {}\n    qr_string = payment.get("qr_string") or ""\n    if payment.get("type") != "qr" or not qr_string:\n        return jsonify(\n            {"error": "Kasera Pay tidak mengembalikan kode QRIS"}), 502\n    created = int(time.time())\n    rec["kasera_payment_id"] = resp.get("id")\n    rec["kasera_status"] = resp.get("status") or "pending"\n    rec["kasera_qr_string"] = qr_string\n    rec["kasera_checkout_url"] = resp.get("checkout_url")\n    rec["kasera_expires_at"] = _parse_kasera_expiry(\n        resp.get("expires_at"), created + QRIS_EXPIRY_SECONDS)\n    rec["kasera_created_at"] = created\n    rec["kasera_last_checked"] = 0\n    save_pay(job, rec)\n    info = _kasera_info(job, rec)\n    return jsonify({"job": job, "payment_id": rec["kasera_payment_id"],\n                    "amount": PRICE_IDR, "qr_url": info["qr_url"],\n                    "checkout_url": rec.get("kasera_checkout_url"),\n                    "expires_in_seconds": info["expires_in_seconds"],\n                    "status": "pending", "reused": False}), 201\n\n\n@app.post("/api/pay/kasera/webhook")\ndef kasera_webhook():\n    """Terima webhook Kasera Pay. Hanya event payment.paid yang\n    melunasi job, dan hanya bila tanda tangan HMAC sah, nominal\n    == Rp15.000, dan payment_request_id cocok dengan pembayaran yang\n    dibuat server ini. Dedupe id event (pengiriman at-least-once).\n    Tanda tangan salah -> 401; secret belum disetel -> 503 (fail\n    closed, jangan pernah tandai lunas tanpa verifikasi). Event lain\n    (test.ping, payment.expired, payment.failed) dijawab 200 tanpa\n    mengubah apa pun."""\n    raw = request.get_data()\n    if not KASERA_PAY_WEBHOOK_SECRET:\n        return jsonify(\n            {"error": "webhook Kasera belum diaktifkan di server"}), 503\n    header = request.headers.get("Kasera-Signature-V1", "")\n    if not header or not _kasera_verify_signature(\n            raw, header, KASERA_PAY_WEBHOOK_SECRET):\n        return jsonify({"error": "tanda tangan tidak valid"}), 401\n    try:\n        event = json.loads(raw.decode())\n    except (ValueError, UnicodeDecodeError):\n        return jsonify({"error": "body bukan JSON"}), 400\n    if event.get("type") != "payment.paid":\n        return "", 200\n    data = event.get("data") or {}\n    job = _kasera_find_job(data.get("payment_request_id"),\n                           data.get("external_id"))\n    if job is None:\n        return "", 200  # bukan pembayaran server ini\n    rec = load_pay(job)\n    event_id = event.get("id") \\\n        or request.headers.get("Kasera-Event-Id", "")\n    seen = rec.get("kasera_events") or []\n    if event_id and event_id in seen:\n        return "", 200  # duplikat: sudah diproses\n    try:\n        amount = int(data.get("amount") or 0)\n    except (TypeError, ValueError):\n        amount = 0\n    if amount != PRICE_IDR:\n        print(f"PERINGATAN nominal Kasera tidak cocok "\n              f"job {job}: {data.get(\'amount\')}", flush=True)\n        return "", 200\n    if job in load_paid():\n        return "", 200\n    rec["kasera_status"] = "succeeded"\n    rec["kasera_paid_at"] = data.get("paid_at")\n    if event_id:\n        seen.append(event_id)\n        rec["kasera_events"] = seen[-20:]\n    _mark_paid(job, rec)\n    return "", 200\n\n\n@app.get("/api/pay/kasera/<job>/qr")\ndef kasera_qr_image(job):\n    """Gambar PNG kode QRIS Kasera yang sedang aktif untuk job ini."""\n    if not job.isalnum() or len(job) > 32:\n        return jsonify({"error": "job tidak valid"}), 400\n    rec = load_pay(job)\n    qr_string = rec.get("kasera_qr_string") or ""\n    if not qr_string or rec.get("kasera_status") not in ("pending", None):\n        return jsonify({"error": "tidak ada kode QRIS aktif"}), 404\n    buf = _qr_png(qr_string)\n    if buf is None:\n        return jsonify(\n            {"error": "gambar QR tidak tersedia di server"}), 503\n    return send_file(buf, mimetype="image/png",\n                     download_name=f"qris-kasera-{job}.png")\n'),
    (SRV,
     '    elif not paid and rec.get("status") in ("expire", "cancel", "deny"):\n        qris_expired = True\n    resp = {"job": job, "status": "done", "paid": paid,\n            "amount": PRICE_IDR, "filename": filename,\n            "qris": None if paid else _qris_info(job, rec),\n            "qris_expired": qris_expired}\n',
     '    elif not paid and rec.get("status") in ("expire", "cancel", "deny"):\n        qris_expired = True\n    # v0.6.37: pembayaran Kasera — polling status (dibatasi 1x/10 detik;\n    # Kasera membatasi baca 1x/detik per transaksi) bila masih pending.\n    kasera_expired = False\n    if not paid and rec.get("kasera_payment_id") \\\n            and rec.get("kasera_status") == "pending":\n        now = int(time.time())\n        if KASERA_PAY_KEY \\\n                and now - int(rec.get("kasera_last_checked") or 0) >= 10:\n            rec["kasera_last_checked"] = now\n            save_pay(job, rec)\n            try:\n                st = _kasera_request(\n                    "GET",\n                    f"/v1/transactions/{rec[\'kasera_payment_id\']}")\n                tx = st.get("status")\n                if tx == "succeeded" \\\n                        and int(st.get("amount") or 0) == PRICE_IDR:\n                    rec["kasera_status"] = "succeeded"\n                    _mark_paid(job, rec)\n                    paid = True\n                elif tx == "succeeded":\n                    print(f"PERINGATAN nominal Kasera tidak cocok "\n                          f"job {job}: {st.get(\'amount\')}", flush=True)\n                elif tx in ("expired", "failed", "canceled"):\n                    rec["kasera_status"] = tx\n                    save_pay(job, rec)\n                    kasera_expired = True\n            except KaseraPayError as exc:\n                # 404 = payment id tidak dikenal Kasera -> anggap hangus.\n                if "HTTP 404" in str(exc):\n                    rec["kasera_status"] = "expired"\n                    save_pay(job, rec)\n                    kasera_expired = True\n                # Error lain: kembalikan info terakhir yang tersimpan.\n                pass\n            except (TypeError, ValueError):\n                pass\n        elif (rec.get("kasera_expires_at") or 0) <= now:\n            kasera_expired = True\n    elif not paid and rec.get("kasera_status") in ("expired", "failed",\n                                                   "canceled"):\n        kasera_expired = True\n    resp = {"job": job, "status": "done", "paid": paid,\n            "amount": PRICE_IDR, "filename": filename,\n            "qris": None if paid else _qris_info(job, rec),\n            "qris_expired": qris_expired,\n            "kasera": None if paid else _kasera_info(job, rec),\n            "kasera_expired": kasera_expired}\n'),
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
        with open(path + ".bak-0.6.36", "w", encoding="utf-8") as f:
            f.write(src)
        for p, old, new in PAIRS:
            if p == path:
                src = src.replace(old, new)
        with open(path, "w", encoding="utf-8") as f:
            f.write(src)
        print(f"OK: {path} di-patch (backup: {path}.bak-0.6.36)")
    return 0


def main():
    if apply_pairs() != 0:
        return 1
    rc, out = run(f"python3 -m py_compile {SRV}")
    if rc != 0:
        print(f"GAGAL: py_compile {SRV}:\n{out}")
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
        if data.get("version") == "0.6.37" and data.get("status") == "ok":
            print("OK: engine v0.6.37 LIVE")
            return 0
        print(f"health coba {i + 1}: versi/status belum cocok: {data}")
    print("GAGAL: /api/health tidak menunjukkan v0.6.37 setelah 60 detik")
    return 1


if __name__ == "__main__":
    sys.exit(main())
