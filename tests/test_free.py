#!/usr/bin/env python3
"""Tes kuota gratis v0.6.38: 10 dokumen pertama GRATIS (maks 1/IP).

Kasus: counter 0->1->2 lalu job ke-3 wajib bayar (FREE_LIMIT=2 di tes);
IP sama upload kedua wajib bayar meski kuota masih ada; IP beda tetap
gratis; download job gratis tanpa bayar (docx + pdf); job belum bayar
non-gratis tetap 402; job lunas tetap bisa download; /api/status dan
/api/health membawa angka gratis yang benar; job gratis TIDAK masuk
paid.json; free_quota.json tidak dihapus cleanup_old_jobs.

Upload memakai file .docx PALSU (bukan dokumen asli): yang diuji di sini
adalah keputusan klaim saat upload (endpoint + counter + IP), bukan
proses TOC-nya — proses TOC sudah dicakup run-suite + tes live.
Worker latar akan gagal memproses file palsu itu (status "error");
klaim sudah tercatat sebelum itu dan tidak terpengaruh.

Pakai: ~/workspace/venvs/toc/bin/python tests/test_free.py
Exit 0 = semua PASS.
"""
import io
import json
import os
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ENGINE_DIR = os.path.dirname(HERE)
sys.path.insert(0, ENGINE_DIR)

os.environ["FREE_LIMIT"] = "2"  # kuota kecil supaya batas ke-3 teruji

FAKE_DOCX = b"PK\x03\x04palsu"


def main():
    tmp = tempfile.mkdtemp(prefix="free-test-")
    results = os.path.join(tmp, "results")
    os.makedirs(results)

    import server as S  # noqa: E402 - env di atas dibaca saat import

    S.RESULTS_DIR = results
    S.PAID_FILE = os.path.join(results, "paid.json")
    S.FREE_QUOTA_FILE = os.path.join(results, "free_quota.json")
    S.MAX_UPLOADS_PER_HOUR = 100
    client = S.app.test_client()

    passed, failed = [], []

    def check(name, ok, extra=""):
        (passed if ok else failed).append(name)
        print(f"  {'PASS' if ok else 'FAIL'}  {name}"
              + (f" — {extra}" if extra and not ok else ""))

    def quota():
        try:
            with open(S.FREE_QUOTA_FILE) as f:
                return json.load(f) or {}
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            return {}

    def make_job(job, free_number=None, filename="Skripsi_Test.docx"):
        """Job done palsu (seperti test_kasera.make_job)."""
        with open(os.path.join(results, f"{job}-hasil-daftar-isi.docx"),
                  "wb") as f:
            f.write(b"PK\x03\x04palsu")
        with open(os.path.join(results, f"{job}-hasil-daftar-isi.pdf"),
                  "wb") as f:
            f.write(b"%PDF-palsu")
        with open(os.path.join(results, f"{job}.meta.json"), "w") as f:
            json.dump({"filename": filename}, f)
        state = {"job": job, "status": "done", "filename": filename,
                 "created_at": int(time.time()),
                 "updated_at": int(time.time())}
        if free_number is not None:
            state["free"] = True
            state["free_number"] = free_number
        S._save_job_state(job, state)

    def upload(ip):
        return client.post(
            "/api/toc",
            data={"file": (io.BytesIO(FAKE_DOCX), "Skripsi.docx")},
            headers={"X-Forwarded-For": ip},
            content_type="multipart/form-data")

    # 1. health membawa sisa kuota + batas
    r = client.get("/api/health")
    d = r.get_json() or {}
    check("health: free_left=2, free_limit=2, version 0.6.38",
          d.get("free_left") == 2 and d.get("free_limit") == 2
          and d.get("version") == "0.6.38"
          and d.get("status") == "ok", f"dapat {d}")

    # 2. upload pertama IP A -> gratis #1, counter 0->1
    r = upload("10.0.0.1")
    job1 = (r.get_json() or {}).get("job")
    st1 = S._load_job_state(job1) if job1 else {}
    q = quota()
    check("upload IP A pertama -> 202 + gratis #1",
          r.status_code == 202 and job1
          and st1.get("free") is True and st1.get("free_number") == 1
          and q.get("used") == 1 and q.get("ips") == ["10.0.0.1"],
          f"dapat {r.status_code} {st1} {q}")
    check("job gratis TIDAK masuk paid.json", job1 not in S.load_paid())

    # 3. IP sama upload kedua -> wajib bayar meski kuota masih ada
    r = upload("10.0.0.1")
    job2 = (r.get_json() or {}).get("job")
    st2 = S._load_job_state(job2) if job2 else {}
    check("upload IP A kedua -> bukan gratis",
          r.status_code == 202 and job2 and not st2.get("free")
          and quota().get("used") == 1,
          f"dapat {r.status_code} {st2} {quota()}")

    # 4. IP beda -> gratis #2, counter 1->2 (kuota habis)
    r = upload("10.0.0.2")
    job3 = (r.get_json() or {}).get("job")
    st3 = S._load_job_state(job3) if job3 else {}
    check("upload IP B -> gratis #2",
          r.status_code == 202 and job3
          and st3.get("free") is True and st3.get("free_number") == 2
          and quota().get("used") == 2,
          f"dapat {r.status_code} {st3} {quota()}")

    # 5. kuota habis -> IP baru wajib bayar
    r = upload("10.0.0.3")
    job4 = (r.get_json() or {}).get("job")
    st4 = S._load_job_state(job4) if job4 else {}
    check("upload IP C saat kuota habis -> bukan gratis",
          r.status_code == 202 and job4 and not st4.get("free")
          and quota().get("used") == 2,
          f"dapat {r.status_code} {st4} {quota()}")
    r = client.get("/api/health")
    check("health: free_left=0 setelah kuota habis",
          (r.get_json() or {}).get("free_left") == 0,
          f"dapat {r.get_json()}")

    # 6. download job gratis (docx + pdf) tanpa bayar
    free_job = "f1ee00000001"
    make_job(free_job, free_number=1)
    r = client.get(f"/api/download/{free_job}")
    check("download job gratis docx -> 200",
          r.status_code == 200, f"dapat {r.status_code}")
    r = client.get(f"/api/download/{free_job}?format=pdf")
    check("download job gratis pdf -> 200",
          r.status_code == 200, f"dapat {r.status_code}")
    check("job gratis tetap TIDAK masuk paid.json setelah download",
          free_job not in S.load_paid())

    # 7. download job belum bayar non-gratis -> 402
    plain_job = "b2c3d4e5f6a7"
    make_job(plain_job)
    r = client.get(f"/api/download/{plain_job}")
    check("download job belum bayar -> 402",
          r.status_code == 402, f"dapat {r.status_code}")

    # 8. job lunas tetap bisa download (alur bayar tidak berubah)
    paid_job = "c3d4e5f6a7b8"
    make_job(paid_job)
    S._mark_paid(paid_job, {})
    r = client.get(f"/api/download/{paid_job}")
    check("download job lunas -> 200",
          r.status_code == 200, f"dapat {r.status_code}")
    check("job lunas ada di paid.json, job gratis tidak",
          paid_job in S.load_paid()
          and free_job not in S.load_paid())

    # 9. status job gratis membawa free + angka benar, paid False
    r = client.get(f"/api/status/{free_job}")
    d = r.get_json() or {}
    check("status job gratis: free + free_number + sisa 0, paid False",
          d.get("free") is True and d.get("free_number") == 1
          and d.get("free_left") == 0 and d.get("free_limit") == 2
          and d.get("paid") is False, f"dapat {d}")

    # 10. status job non-gratis: free False
    r = client.get(f"/api/status/{plain_job}")
    d = r.get_json() or {}
    check("status job non-gratis: free False",
          d.get("free") is False and d.get("paid") is False,
          f"dapat {d}")

    # 11. cleanup_old_jobs tidak menghapus free_quota.json tua
    old = time.time() - 2 * 24 * 3600
    os.utime(S.FREE_QUOTA_FILE, (old, old))
    dummy = os.path.join(results, "jadul.json")
    with open(dummy, "w") as f:
        json.dump({}, f)
    os.utime(dummy, (old, old))
    S._last_cleanup = 0.0
    S.cleanup_old_jobs()
    check("cleanup tidak menghapus free_quota.json (counter utuh)",
          os.path.exists(S.FREE_QUOTA_FILE)
          and quota().get("used") == 2,
          f"ada={os.path.exists(S.FREE_QUOTA_FILE)} {quota()}")
    check("cleanup tetap menghapus file jadul lain",
          not os.path.exists(dummy))

    print(f"\n{len(passed)} PASS, {len(failed)} FAIL")
    if failed:
        print("GAGAL: " + ", ".join(failed))
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
