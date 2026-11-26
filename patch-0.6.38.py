#!/usr/bin/env python3
"""Patch v0.6.38 untuk engine DaftarIsiKu di VPS Awan.

10 dokumen pertama GRATIS (tanpa login):
  - Counter persisten free_quota.json (env FREE_LIMIT, default 10)
    + maks 1 gratis per IP. Dihitung saat job dibuat (POST /api/toc).
  - GET /api/download/<job> membuka job gratis tanpa bayar (402
    hanya untuk job belum bayar yang BUKAN gratis).
  - GET /api/status/<job> membawa free + free_number + free_left +
    free_limit supaya website bisa melewati layar bayar.
  - GET /api/health membawa free_left + free_limit (banner website).
  - Job gratis TIDAK masuk paid.json (status "free" di job state) —
    laporan harian + rekonsiliasi Kasera tetap benar.
  - Kasera, Midtrans, dan /api/pay/simulate TIDAK diubah. Alur bayar
    Rp15.000 untuk job berbayar tidak berubah. VERSION 0.6.37 -> 0.6.38.

Jalankan SETELAH v0.6.37 ter-deploy (patch ini mengharapkan
server.py v0.6.37 di server).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.38.py -o /tmp/patch-0.6.38.py && python3 /tmp/patch-0.6.38.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.37. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import json
import os
import subprocess
import sys
import time
import urllib.request

SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    (SRV,
     '  - DOWNLOAD file Word hasil = Rp15.000 per dokumen.\n',
     '  - DOWNLOAD file Word hasil = Rp15.000 per dokumen.\n'
     '  - v0.6.38: 10 dokumen pertama GRATIS (tanpa login; maks 1 per IP;\n'
     '    setelah kuota habis, download berbayar Rp15.000).\n'),
    (SRV,
     'VERSION = "0.6.37"  # v0.6.37: ',
     'VERSION = "0.6.38"  # v0.6.38: 10 dokumen pertama GRATIS — counter persisten free_quota.json (env FREE_LIMIT, default 10) + maks 1 gratis per IP; /api/download membuka job gratis tanpa bayar (status "free", TIDAK masuk paid.json), /api/status membawa free + free_number + free_left + free_limit, /api/health membawa free_left + free_limit. Kasera/Midtrans/simulate TIDAK diubah. v0.6.37: '),
    (SRV,
     'MAX_QRIS_ATTEMPTS = 5\nREPORTS_TOKEN = os.environ.get("REPORTS_TOKEN", "")\n',
     'MAX_QRIS_ATTEMPTS = 5\nREPORTS_TOKEN = os.environ.get("REPORTS_TOKEN", "")\n'
     '\n'
     '# --- dokumen gratis (v0.6.38) --------------------------------------\n'
     '# 10 dokumen pertama GRATIS (tanpa login): counter persisten di\n'
     '# free_quota.json + maks 1 gratis per IP. FREE_LIMIT dari env supaya\n'
     '# tes bisa set 1-2. Bila kuota habis -> alur bayar Rp15.000 biasa.\n'
     'FREE_QUOTA_FILE = os.path.join(RESULTS_DIR, "free_quota.json")\n'
     'FREE_LIMIT = int(os.environ.get("FREE_LIMIT", "10"))\n'),
    (SRV,
     'def save_paid(paid):\n'
     '    tmp = PAID_FILE + ".tmp"\n'
     '    with open(tmp, "w") as f:\n'
     '        json.dump(sorted(paid), f)\n'
     '    os.replace(tmp, PAID_FILE)\n',
     'def save_paid(paid):\n'
     '    tmp = PAID_FILE + ".tmp"\n'
     '    with open(tmp, "w") as f:\n'
     '        json.dump(sorted(paid), f)\n'
     '    os.replace(tmp, PAID_FILE)\n'
     '\n'
     '\n'
     '# --- kuota gratis (v0.6.38) -----------------------------------------\n'
     '# Pola file sama seperti paid.json: tulis atomik (tmp + os.replace).\n'
     '_FREE_QUOTA_LOCK = threading.Lock()\n'
     '\n'
     '\n'
     'def _load_free_quota():\n'
     '    """{"used": int, "ips": [ip, ...]}; file rusak/hilang -> kosong."""\n'
     '    try:\n'
     '        with open(FREE_QUOTA_FILE) as f:\n'
     '            d = json.load(f) or {}\n'
     '    except (FileNotFoundError, json.JSONDecodeError, OSError):\n'
     '        return {"used": 0, "ips": []}\n'
     '    return {"used": int(d.get("used") or 0),\n'
     '            "ips": [i for i in (d.get("ips") or [])\n'
     '                    if isinstance(i, str)]}\n'
     '\n'
     '\n'
     'def _save_free_quota(q):\n'
     '    tmp = FREE_QUOTA_FILE + ".tmp"\n'
     '    with open(tmp, "w") as f:\n'
     '        json.dump({"used": int(q.get("used") or 0),\n'
     '                   "ips": sorted(set(q.get("ips") or []))}, f)\n'
     '    os.replace(tmp, FREE_QUOTA_FILE)\n'
     '\n'
     '\n'
     'def _free_left():\n'
     '    """Sisa dokumen gratis (0 bila kuota habis)."""\n'
     '    return max(0, FREE_LIMIT - _load_free_quota()["used"])\n'
     '\n'
     '\n'
     'def _try_claim_free(ip):\n'
     '    """Klaim 1 slot gratis untuk IP ini; kembalikan nomor urut\n'
     '    (1..FREE_LIMIT) atau None bila kuota habis / IP sudah pakai.\n'
     '    Dihitung saat job dibuat. Kunci proses: gunicorn 1 worker."""\n'
     '    with _FREE_QUOTA_LOCK:\n'
     '        q = _load_free_quota()\n'
     '        if q["used"] >= FREE_LIMIT or ip in q["ips"]:\n'
     '            return None\n'
     '        q["used"] += 1\n'
     '        q["ips"].append(ip)\n'
     '        try:\n'
     '            _save_free_quota(q)\n'
     '        except OSError:\n'
     '            return None\n'
     '        return q["used"]\n'
     '\n'
     '\n'
     'def _is_free_job(job):\n'
     '    """True bila job ini ditandai gratis saat dibuat (job state)."""\n'
     '    return bool(_load_job_state(job).get("free"))\n'),
    (SRV,
     '            if name == "paid.json":\n                continue\n',
     '            # v0.6.38: free_quota.json TIDAK boleh dihapus — bila\n'
     '            # hilang, counter reset ke 0 dan 10 dokumen gratis baru\n'
     '            # muncul lagi. Sama seperti paid.json.\n'
     '            if name in ("paid.json", "free_quota.json"):\n'
     '                continue\n'),
    (SRV,
     '@app.get("/api/health")\n'
     'def health():\n'
     '    return jsonify({"status": "ok", "app": "DaftarIsiKu",\n'
     '                    "version": VERSION})\n',
     '@app.get("/api/health")\n'
     'def health():\n'
     '    return jsonify({"status": "ok", "app": "DaftarIsiKu",\n'
     '                    "version": VERSION,\n'
     '                    "free_left": _free_left(),\n'
     '                    "free_limit": FREE_LIMIT})\n'),
    (SRV,
     '    except queue.Full:\n'
     '        # v0.6.31: antrean penuh — bersihkan state + file, tolak 429.\n'
     '        try:\n'
     '            os.remove(_job_state_path(job))\n'
     '        except OSError:\n'
     '            pass\n'
     '        shutil.rmtree(work, ignore_errors=True)\n'
     '        return jsonify(\n'
     '            {"error": "server sedang sibuk, coba lagi sebentar"}), 429\n'
     '    return jsonify({"job": job, "status": "queued"}), 202\n',
     '    except queue.Full:\n'
     '        # v0.6.31: antrean penuh — bersihkan state + file, tolak 429.\n'
     '        try:\n'
     '            os.remove(_job_state_path(job))\n'
     '        except OSError:\n'
     '            pass\n'
     '        shutil.rmtree(work, ignore_errors=True)\n'
     '        return jsonify(\n'
     '            {"error": "server sedang sibuk, coba lagi sebentar"}), 429\n'
     '    # v0.6.38: klaim slot gratis SETELAH antrean menerima job, supaya\n'
     '    # job yang ditolak (antrean penuh) tidak memakan kuota.\n'
     '    free_number = _try_claim_free(_client_ip())\n'
     '    if free_number is not None:\n'
     '        _update_job_state(job, free=True, free_number=free_number)\n'
     '    return jsonify({"job": job, "status": "queued"}), 202\n'),
    (SRV,
     '        return jsonify({"job": job, "status": state["status"],\n'
     '                        "queue_position": pos,\n'
     '                        "paid": False, "amount": PRICE_IDR,\n'
     '                        "filename": state.get("filename"),\n'
     '                        "qris": None, "qris_expired": False})\n',
     '        return jsonify({"job": job, "status": state["status"],\n'
     '                        "queue_position": pos,\n'
     '                        "paid": False, "amount": PRICE_IDR,\n'
     '                        "filename": state.get("filename"),\n'
     '                        "free": bool(state.get("free")),\n'
     '                        "free_left": _free_left(),\n'
     '                        "free_limit": FREE_LIMIT,\n'
     '                        "qris": None, "qris_expired": False})\n'),
    (SRV,
     '    resp = {"job": job, "status": "done", "paid": paid,\n'
     '            "amount": PRICE_IDR, "filename": filename,\n'
     '            "qris": None if paid else _qris_info(job, rec),\n'
     '            "qris_expired": qris_expired,\n'
     '            "kasera": None if paid else _kasera_info(job, rec),\n'
     '            "kasera_expired": kasera_expired}\n',
     '    resp = {"job": job, "status": "done", "paid": paid,\n'
     '            "amount": PRICE_IDR, "filename": filename,\n'
     '            "free": bool(state.get("free")),\n'
     '            "free_number": state.get("free_number"),\n'
     '            "free_left": _free_left(),\n'
     '            "free_limit": FREE_LIMIT,\n'
     '            "qris": None if paid else _qris_info(job, rec),\n'
     '            "qris_expired": qris_expired,\n'
     '            "kasera": None if paid else _kasera_info(job, rec),\n'
     '            "kasera_expired": kasera_expired}\n'),
    (SRV,
     '    if job not in load_paid():\n'
     '        return jsonify({"error": "belum dibayar",\n'
     '                        "amount": 15000,\n'
     '                        "hint": "bayar Rp15.000 via QRIS dulu, "\n'
     '                                "lalu download lagi"}), 402\n',
     '    # v0.6.38: job gratis boleh download tanpa bayar; job lain yang\n'
     '    # belum lunas tetap 402. Gratis TIDAK masuk paid.json.\n'
     '    if job not in load_paid() and not _is_free_job(job):\n'
     '        return jsonify({"error": "belum dibayar",\n'
     '                        "amount": 15000,\n'
     '                        "hint": "bayar Rp15.000 via QRIS dulu, "\n'
     '                                "lalu download lagi"}), 402\n'),
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
        with open(path + ".bak-0.6.37", "w", encoding="utf-8") as f:
            f.write(src)
        for p, old, new in PAIRS:
            if p == path:
                src = src.replace(old, new)
        with open(path, "w", encoding="utf-8") as f:
            f.write(src)
        print(f"OK: {path} di-patch (backup: {path}.bak-0.6.37)")
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
        if data.get("version") == "0.6.38" and data.get("status") == "ok":
            print("OK: engine v0.6.38 LIVE")
            return 0
        print(f"health coba {i + 1}: versi/status belum cocok: {data}")
    print("GAGAL: /api/health tidak menunjukkan v0.6.38 setelah 60 detik")
    return 1


if __name__ == "__main__":
    sys.exit(main())
