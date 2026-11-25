#!/usr/bin/env python3
"""Tes integrasi Kasera Pay (v0.6.37) dengan server Kasera tiruan.

Menjalankan Flask test client melawan server tiruan Kasera di
localhost: create -> qr_string palsu -> webhook payment.paid
bertanda tangan. Kasus: happy path, reuse QR pending, QR PNG,
webhook valid, duplikat webhook, tanda tangan salah, timestamp
kedaluwarsa, nominal salah, payment id asing, test.ping, polling
status lewat /api/status (lunas + kedaluwarsa), create tanpa key
(503), key salah (502), 409 bila sudah lunas, 404 bila job tidak ada.

Pakai: ~/workspace/venvs/toc/bin/python tests/test_kasera.py
Exit 0 = semua PASS.
"""
import hashlib
import hmac
import json
import os
import sys
import tempfile
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE_DIR = os.path.dirname(HERE)
sys.path.insert(0, ENGINE_DIR)

FAKE_QR = ("00020101021226570011ID.DANA.WWW0118936009153150123456"
           "02031501000000000000003580200001303UMI520454145303360"
           "5405150005802ID5914DaftarIsiKu6007Mataram6108362000"
           "703UMI6304ABCD")
TEST_SECRET = "whsec_test_rahasia"
TEST_KEY = "kp_test_dummy"


class FakeKaseraState:
    def __init__(self):
        self.create_calls = 0
        self.get_status = {}  # payment_id -> status
        self.last_idem = None


STATE = FakeKaseraState()


class FakeKasera(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if self.path != "/v1/transactions":
            return self._send(404, {"error": {"code": "not_found",
                                              "message": "not found"}})
        if self.headers.get("Authorization") != f"Bearer {TEST_KEY}":
            return self._send(401, {"error": {"code": "unauthorized",
                                              "message": "bad key"}})
        if not self.headers.get("Idempotency-Key"):
            return self._send(400, {"error": {"code": "invalid_idempotency_key",
                                              "message": "no key"}})
        length = int(self.headers.get("Content-Length") or 0)
        body = json.loads(self.rfile.read(length) or b"{}")
        if body.get("amount") != 15000 \
                or body.get("payment_methods") != ["qris"]:
            return self._send(422, {"error": {"code": "validation_failed",
                                              "message": "bad body"}})
        STATE.create_calls += 1
        STATE.last_idem = self.headers.get("Idempotency-Key")
        pid = f"payreq_test{STATE.create_calls:03d}"
        STATE.get_status[pid] = "pending"
        exp = time.strftime("%Y-%m-%dT%H:%M:%S+07:00",
                            time.gmtime(time.time() + 3600 + 7 * 3600))
        self._send(201, {"id": pid, "status": "pending",
                         "amount": 15000, "fee": 355, "net": 14645,
                         "currency": "IDR",
                         "description": body.get("description"),
                         "external_id": body.get("external_id"),
                         "payment": {"type": "qr", "qr_string": FAKE_QR},
                         "checkout_url": f"https://pay.kasera.id/p/{pid}",
                         "expires_at": exp, "livemode": False})

    def do_GET(self):
        parts = self.path.strip("/").split("/")
        if len(parts) == 3 and parts[0] == "v1" \
                and parts[1] == "transactions":
            pid = parts[2]
            if pid not in STATE.get_status:
                return self._send(404, {"error": {"code": "not_found",
                                                  "message": "no tx"}})
            if self.headers.get("Authorization") != f"Bearer {TEST_KEY}":
                return self._send(401, {"error": {"code": "unauthorized",
                                                  "message": "bad key"}})
            return self._send(200, {"id": pid,
                                    "status": STATE.get_status[pid],
                                    "amount": 15000})
        return self._send(404, {"error": {"code": "not_found",
                                          "message": "not found"}})


def sign(secret, raw, t=None):
    """Header Kasera-Signature-V1 untuk body mentah."""
    t = int(time.time()) if t is None else t
    sig = hmac.new(secret.encode(), f"{t}.".encode() + raw,
                   hashlib.sha256).hexdigest()
    return f"t={t},v1={sig}"


def main():
    tmp = tempfile.mkdtemp(prefix="kasera-test-")
    results = os.path.join(tmp, "results")
    os.makedirs(results)
    os.environ["KASERA_PAY_KEY"] = TEST_KEY
    os.environ["KASERA_PAY_WEBHOOK_SECRET"] = TEST_SECRET
    srv = ThreadingHTTPServer(("127.0.0.1", 0), FakeKasera)
    port = srv.server_address[1]
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    os.environ["KASERA_PAY_BASE_URL"] = f"http://127.0.0.1:{port}"

    import server as S  # noqa: E402 - env di atas dibaca saat import

    S.RESULTS_DIR = results
    S.PAID_FILE = os.path.join(results, "paid.json")
    S.KASERA_PAY_BASE_URL = f"http://127.0.0.1:{port}"
    client = S.app.test_client()

    passed, failed = [], []

    def check(name, ok, extra=""):
        (passed if ok else failed).append(name)
        print(f"  {'PASS' if ok else 'FAIL'}  {name}"
              + (f" — {extra}" if extra and not ok else ""))

    def make_job(job, filename="Skripsi_Test.docx"):
        with open(os.path.join(results, f"{job}-hasil-daftar-isi.docx"),
                  "wb") as f:
            f.write(b"PK\x03\x04palsu")
        with open(os.path.join(results, f"{job}.meta.json"), "w") as f:
            json.dump({"filename": filename}, f)
        S._save_job_state(job, {"job": job, "status": "done",
                                "filename": filename,
                                "created_at": int(time.time()),
                                "updated_at": int(time.time())})

    def webhook(job_pid, job, amount=15000, event_id="evt_1",
                secret=TEST_SECRET, t=None, etype="payment.paid"):
        body = json.dumps({"id": event_id, "type": etype,
                           "livemode": False,
                           "created_at": "2026-11-25T13:00:00+07:00",
                           "data": {"payment_request_id": job_pid,
                                    "external_id": job,
                                    "amount": amount, "currency": "IDR",
                                    "paid_at": "2026-11-25T13:00:05+07:00"}},
                          separators=(",", ":")).encode()
        return client.post("/api/pay/kasera/webhook", data=body,
                           headers={"Kasera-Signature-V1":
                                    sign(secret, body, t),
                                    "Kasera-Event-Id": event_id,
                                    "Content-Type": "application/json"})

    job1, job2 = "a1b2c3d4e5f6", "b2c3d4e5f6a7"
    make_job(job1)
    make_job(job2)

    # 1. create tanpa key -> 503 (jangan crash, jangan hubungi Kasera)
    old_key = S.KASERA_PAY_KEY
    S.KASERA_PAY_KEY = ""
    r = client.post(f"/api/pay/kasera/{job1}")
    check("create tanpa key -> 503", r.status_code == 503
          and "belum diaktifkan" in r.get_json().get("error", ""),
          f"dapat {r.status_code}")
    check("create tanpa key tidak memanggil Kasera",
          STATE.create_calls == 0)
    S.KASERA_PAY_KEY = old_key

    # 2. webhook tanpa secret -> 503 (fail closed)
    old_sec = S.KASERA_PAY_WEBHOOK_SECRET
    S.KASERA_PAY_WEBHOOK_SECRET = ""
    body = b'{"id":"evt_x","type":"payment.paid","data":{}}'
    r = client.post("/api/pay/kasera/webhook", data=body,
                    headers={"Kasera-Signature-V1": sign("x", body)})
    check("webhook tanpa secret -> 503", r.status_code == 503,
          f"dapat {r.status_code}")
    S.KASERA_PAY_WEBHOOK_SECRET = old_sec

    # 3. create happy path -> 201 + field lengkap
    r = client.post(f"/api/pay/kasera/{job1}")
    d = r.get_json() or {}
    check("create -> 201", r.status_code == 201, f"dapat {r.status_code}")
    check("create membawa qr_url + checkout_url",
          d.get("qr_url") == f"/api/pay/kasera/{job1}/qr"
          and d.get("checkout_url", "").startswith(
              "https://pay.kasera.id/p/")
          and d.get("amount") == 15000
          and d.get("status") == "pending"
          and d.get("expires_in_seconds", 0) > 0)
    pid1 = d.get("payment_id")
    rec = S.load_pay(job1)
    check("pay.json menyimpan payment id + idempotency key",
          rec.get("kasera_payment_id") == pid1
          and rec.get("kasera_idempotency_key") == f"daftarisiku-{job1}-1"
          and rec.get("kasera_qr_string") == FAKE_QR)

    # 4. create lagi saat pending -> reuse, Kasera tidak dipanggil lagi
    calls = STATE.create_calls
    r = client.post(f"/api/pay/kasera/{job1}")
    d = r.get_json() or {}
    check("create ulang pending -> reuse",
          r.status_code == 201 and d.get("reused") is True
          and d.get("payment_id") == pid1
          and STATE.create_calls == calls)

    # 5. QR PNG bisa diunduh
    r = client.get(f"/api/pay/kasera/{job1}/qr")
    check("GET qr -> 200 PNG", r.status_code == 200
          and r.content_type.startswith("image/png")
          and len(r.data) > 100, f"dapat {r.status_code}")

    # 6. /api/status membawa info kasera, belum lunas
    r = client.get(f"/api/status/{job1}")
    d = r.get_json() or {}
    check("status membawa kasera + belum lunas",
          d.get("paid") is False
          and (d.get("kasera") or {}).get("payment_id") == pid1
          and d.get("kasera_expired") is False)

    # 7. webhook valid -> 200 + job lunas
    r = webhook(pid1, job1)
    check("webhook valid -> 200", r.status_code == 200,
          f"dapat {r.status_code}")
    check("job lunas setelah webhook valid", job1 in S.load_paid())
    r = client.get(f"/api/status/{job1}")
    check("status paid=True setelah webhook",
          (r.get_json() or {}).get("paid") is True)

    # 8. webhook duplikat (id event sama) -> 200, tidak lunas 2x
    rec = S.load_pay(job1)
    n_events = len(rec.get("kasera_events") or [])
    r = webhook(pid1, job1)
    rec = S.load_pay(job1)
    check("webhook duplikat -> 200, event tidak dicatat 2x",
          r.status_code == 200
          and len(rec.get("kasera_events") or []) == n_events == 1)

    # 9. create setelah lunas -> 409
    r = client.post(f"/api/pay/kasera/{job1}")
    check("create setelah lunas -> 409", r.status_code == 409,
          f"dapat {r.status_code}")

    # 10. webhook tanda tangan salah -> 401, job2 TIDAK lunas
    r = client.post(f"/api/pay/kasera/{job2}")
    pid2 = (r.get_json() or {}).get("payment_id")
    check("create job2 -> 201", r.status_code == 201
          and pid2, f"dapat {r.status_code}")
    r = webhook(pid2, job2, event_id="evt_bad",
                secret="whsec_SALAH")
    check("webhook tanda tangan salah -> 401", r.status_code == 401,
          f"dapat {r.status_code}")
    check("job2 tetap belum lunas", job2 not in S.load_paid())

    # 11. webhook timestamp tua (>5 menit) -> 401
    r = webhook(pid2, job2, event_id="evt_old",
                t=int(time.time()) - 3600)
    check("webhook timestamp tua -> 401", r.status_code == 401,
          f"dapat {r.status_code}")
    check("job2 tetap belum lunas (timestamp)", job2 not in S.load_paid())

    # 12. webhook nominal salah -> 200 tapi TIDAK lunas
    r = webhook(pid2, job2, amount=10000, event_id="evt_nominal")
    check("webhook nominal salah -> 200", r.status_code == 200,
          f"dapat {r.status_code}")
    check("job2 tetap belum lunas (nominal)", job2 not in S.load_paid())

    # 13. webhook payment id asing -> 200, tidak ada yang lunas
    r = webhook("payreq_asing", job2, event_id="evt_asing")
    check("webhook payment id asing -> 200", r.status_code == 200,
          f"dapat {r.status_code}")
    check("job2 tetap belum lunas (asing)", job2 not in S.load_paid())

    # 14. test.ping bertanda tangan sah -> 200, tidak ada yang berubah
    r = webhook(pid2, job2, event_id="evt_ping", etype="test.ping")
    check("test.ping -> 200 tanpa efek", r.status_code == 200
          and job2 not in S.load_paid(), f"dapat {r.status_code}")

    # 15. polling /api/status: Kasera succeeded -> lunas via status
    STATE.get_status[pid2] = "succeeded"
    rec = S.load_pay(job2)
    rec["kasera_last_checked"] = 0
    S.save_pay(job2, rec)
    r = client.get(f"/api/status/{job2}")
    check("polling status Kasera succeeded -> lunas",
          (r.get_json() or {}).get("paid") is True
          and job2 in S.load_paid(), f"dapat {r.status_code}")

    # 16. polling /api/status: Kasera expired -> kasera_expired
    job3 = "c3d4e5f6a7b8"
    make_job(job3)
    r = client.post(f"/api/pay/kasera/{job3}")
    pid3 = (r.get_json() or {}).get("payment_id")
    STATE.get_status[pid3] = "expired"
    rec = S.load_pay(job3)
    rec["kasera_last_checked"] = 0
    S.save_pay(job3, rec)
    r = client.get(f"/api/status/{job3}")
    d = r.get_json() or {}
    check("polling status Kasera expired -> kasera_expired",
          d.get("paid") is False and d.get("kasera_expired") is True
          and job3 not in S.load_paid())

    # 17. create job tidak ada -> 404; key salah -> 502
    r = client.post("/api/pay/kasera/0123456789ab")
    check("create job tidak ada -> 404", r.status_code == 404,
          f"dapat {r.status_code}")
    S.KASERA_PAY_KEY = "kp_test_salah"
    job4 = "d4e5f6a7b8c9"
    make_job(job4)
    r = client.post(f"/api/pay/kasera/{job4}")
    check("create key salah -> 502", r.status_code == 502,
          f"dapat {r.status_code}")
    S.KASERA_PAY_KEY = old_key

    srv.shutdown()
    print(f"\n{len(passed)} PASS, {len(failed)} FAIL")
    if failed:
        print("GAGAL: " + ", ".join(failed))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
