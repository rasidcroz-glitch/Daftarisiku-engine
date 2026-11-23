#!/usr/bin/env python3
"""Patch v0.6.28 untuk engine DaftarIsiKu di VPS Awan.

Menghapus endpoint /api/pay/mock (celah keamanan: siapa pun yang tahu
link job bisa menandai job itu lunas TANPA bayar). Website sudah tidak
memakainya sejak PAY_MOCK=false (deploy Vercel 22 Nov 2026 ~18.30);
bayar hanya lewat QRIS Midtrans. Yang dihapus:
  1. Baris dokstring endpoint mock di server.py.
  2. Baris preflight OPTIONS /api/pay/mock.
  3. Route mock_pay() beserta fungsinya.
  4. Versi naik 0.6.27 -> 0.6.28.

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.28.py -o /tmp/patch-0.6.28.py && python3 /tmp/patch-0.6.28.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.27. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

TOC = "/opt/daftarisiku-engine/toc.py"
SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    # 1. Hapus baris dokstring endpoint mock.
    (SRV,
     '                               (503) — alur sinkron lama mentok di\n'
     '                               website untuk file sebesar itu.\n'
     '  POST /api/pay/mock/<job>  -> PROTOTIPE SAJA: menandai job sudah dibayar.\n'
     '                               Tetap ada untuk tes, tapi website sudah\n'
     '                               tidak memakainya.\n'
     '  POST /api/pay/qris/<job>  -> Buat tagihan QRIS Midtrans (Rp15.000)',
     '                               (503) — alur sinkron lama mentok di\n'
     '                               website untuk file sebesar itu.\n'
     '  POST /api/pay/qris/<job>  -> Buat tagihan QRIS Midtrans (Rp15.000)'),
    # 2. Hapus baris preflight OPTIONS /api/pay/mock.
    (SRV,
     '@app.route("/api/download/<job>", methods=["OPTIONS"])\n'
     '@app.route("/api/pay/mock/<job>", methods=["OPTIONS"])\n'
     '@app.route("/api/pay/qris/<job>", methods=["OPTIONS"])',
     '@app.route("/api/download/<job>", methods=["OPTIONS"])\n'
     '@app.route("/api/pay/qris/<job>", methods=["OPTIONS"])'),
    # 3. Hapus route mock_pay() beserta fungsinya.
    (SRV,
     '@app.post("/api/pay/mock/<job>")\n'
     'def mock_pay(job):\n'
     '    """PROTOTIPE SAJA — simulasi konfirmasi bayar Rp15.000.\n'
     '    Diganti webhook Midtrans (QRIS) saat rilis; endpoint ini dihapus."""\n'
     '    path = os.path.join(RESULTS_DIR, f"{job}-hasil-daftar-isi.docx")\n'
     '    if not os.path.exists(path):\n'
     '        return jsonify({"error": "hasil tidak ditemukan"}), 404\n'
     '    paid = load_paid()\n'
     '    paid.add(job)\n'
     '    save_paid(paid)\n'
     '    return jsonify({"job": job, "paid": True, "amount": 15000,\n'
     '                    "mock": True})\n'
     '\n'
     '\n'
     'def _charge_qris(order_id):',
     'def _charge_qris(order_id):'),
    # 4. Naikkan versi di server.py.
    (SRV,
     'VERSION = "0.6.27"  # v0.6.27:',
     'VERSION = "0.6.28"  # v0.6.28: endpoint /api/pay/mock DIHAPUS '
     '(celah keamanan: siapa pun yang tahu link job bisa menandai lunas '
     'gratis); website sudah tidak memakainya sejak PAY_MOCK=false, '
     'bayar hanya lewat QRIS Midtrans. v0.6.27:'),
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
        bak = path + ".bak-0.6.27"
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
    print("SELESAI: engine v0.6.28 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
