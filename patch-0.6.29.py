#!/usr/bin/env python3
"""Patch v0.6.29 untuk engine DaftarIsiKu di VPS Awan.

Caption gambar/tabel yang terbelah 2 paragraf kini terdeteksi:
  - "Gambar 1.1" di satu paragraf + deskripsi di paragraf berikut.
  - "Gambar" saja + "1.1 ..." di paragraf berikut.
Label yang diikuti sub-bab (bukan deskripsi) TIDAK digabung, supaya
tidak ada entri palsu. Teks gabungan dipakai sebagai entri; isi
dokumen TIDAK diubah. Bukti: skripsi truck di korpus kini mendeteksi
"Gambar 1 Kerangka Penelitian" yang sebelumnya hilang (n_gambar 0->1).

Jalankan SETELAH v0.6.28 ter-deploy (patch ini mengharapkan server.py
v0.6.28 + toc.py v0.6.27 di server).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.29.py -o /tmp/patch-0.6.29.py && python3 /tmp/patch-0.6.29.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.28. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

TOC = "/opt/daftarisiku-engine/toc.py"
SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    (TOC,
     '# Caption gambar/tabel pola standar: "Gambar 1.1 Peta Lokasi",\n# "Tabel 2.1 Data Responden", nomor tunggal "Gambar 1 Denah", dan\n# salah ketik pemisah spasi "Gambar 2 1 Kerangka Berpikir" (v0.4.0)\n# serta titik+spasi "Gambar 2. 1 Arsitektur LSTM" (v0.6.18: pemisah\n# antar angka kini [.\\s]+, bukan satu karakter).\nRE_CAPTION = re.compile(r"^(Gambar|Tabel)\\s+(\\d+(?:[.\\s]+\\d+)*)\\s+\\S",\n                         re.IGNORECASE)',
     '# Caption gambar/tabel pola standar: "Gambar 1.1 Peta Lokasi",\n# "Tabel 2.1 Data Responden", nomor tunggal "Gambar 1 Denah", dan\n# salah ketik pemisah spasi "Gambar 2 1 Kerangka Berpikir" (v0.4.0)\n# serta titik+spasi "Gambar 2. 1 Arsitektur LSTM" (v0.6.18: pemisah\n# antar angka kini [.\\s]+, bukan satu karakter).\nRE_CAPTION = re.compile(r"^(Gambar|Tabel)\\s+(\\d+(?:[.\\s]+\\d+)*)\\s+\\S",\n                         re.IGNORECASE)\n# Caption terbelah 2 paragraf (v0.6.29): "Gambar 1.1" di satu\n# paragraf, deskripsinya di paragraf berikut (HEAD); atau "Gambar"\n# saja lalu "1.1 ..." di paragraf berikut (BARE + NUM).\nRE_CAPTION_HEAD = re.compile(r"^(Gambar|Tabel)\\s+(\\d+(?:[.\\s]+\\d+)*)\\s*$",\n                             re.IGNORECASE)\nRE_CAPTION_BARE = re.compile(r"^(Gambar|Tabel)\\s*$", re.IGNORECASE)\nRE_CAPTION_NUM = re.compile(r"^(\\d+(?:[.\\s]+\\d+)*)\\s+\\S", re.IGNORECASE)'),
    (TOC,
     'def extract_captions(docx_path):\n    """Kembalikan (gambar, tabel): entri caption dari paragraf dokumen,\n    dalam urutan kemunculan. Hanya pola standar yang terdeteksi\n    (\'Gambar 1.1 ...\', \'Tabel 2 ...\', \'Gambar 1 ...\'); caption tanpa\n    nomor (mis. \'Gambar: Peta\') atau caption di dalam sel tabel\n    tidak terdeteksi. kind \'caption\' -> nomor halaman Arab relatif\n    terhadap BAB I, sama seperti heading isi.\n\n    Normalisasi tampilan (v0.4.0): pemisah spasi pada nomor\n    (\'Gambar 2 1 ...\', \'Gambar 2. 1 ...\') ditampilkan sebagai titik\n    (\'Gambar 2.1 ...\'). Isi dokumen TIDAK diubah; \'text\' tetap apa\n    adanya (untuk mencari halaman di PDF), \'display\' yang\n    ternormalisasi untuk daftar."""\n    doc = Document(docx_path)\n    gambar, tabel = [], []\n    for i, p in enumerate(doc.paragraphs):\n        text = norm(p.text)\n        if _is_old_toc_line(p):\n            # v0.6.1: baris daftar gambar/tabel LAMA ("Gambar 1.1 ...\\t5")\n            # bukan caption asli — abaikan supaya tidak jadi entri duplikat.\n            continue\n        m = RE_CAPTION.match(text)\n        if not m:\n            continue\n        kind = m.group(1)\n        num = ".".join(re.split(r"[.\\s]+", m.group(2)))\n        rest = text[m.end(2):].strip()\n        display = f"{kind} {num} {rest}"\n        entry = {"text": text, "level": 0, "kind": "caption",\n                 "list": kind.lower(), "para_index": i, "bold": False,\n                 "display": display, "post_anchor": False,\n                 "search": [("text", text)]}\n        (gambar if kind.lower() == "gambar" else tabel).append(entry)\n    return gambar, tabel',
     'def _split_caption_tail(paras, i, num_start=False):\n    """Paragraf lanjutan caption yang terbelah 2 paragraf (v0.6.29).\n    Kembalikan (indeks, teks_norm) atau None. Satu paragraf kosong\n    di antaranya boleh dilewati. Lanjutan harus teks biasa yang\n    pendek (<=120 karakter) tanpa titik di akhir — frasa deskripsi,\n    bukan kalimat isi — dan bukan heading, baris daftar lama,\n    caption lain, judul BAB/sub-bab, atau judul halaman depan,\n    supaya label caption yang berdiri sendiri tidak salah gabung\n    dengan isi berikutnya. num_start=True: lanjutan harus diawali\n    nomor + teks (\'1.1 ...\'), untuk bentuk \'Gambar\' saja di\n    paragraf kepala."""\n    j = i + 1\n    if j < len(paras) and not norm(paras[j].text):\n        j += 1\n    if j >= len(paras):\n        return None\n    q = paras[j]\n    t = norm(q.text)\n    if not t or _is_old_toc_line(q):\n        return None\n    if q.style.name in ("Heading 1", "Heading 2", "Heading 3"):\n        return None\n    if (RE_CAPTION.match(t) or RE_CAPTION_HEAD.match(t)\n            or RE_CAPTION_BARE.match(t) or RE_BAB.match(t)):\n        return None\n    if t.upper().replace(" ", "") in FRONT_BY_NOSPACE:\n        return None\n    if num_start:\n        if not RE_CAPTION_NUM.match(t):\n            return None\n    elif (RE_LETTER.match(t) or RE_NUMBER.match(t)\n            or RE_LOWER.match(t)):\n        return None\n    if len(t) > 120 or t.endswith("."):\n        return None\n    return j, t\n\n\ndef extract_captions(docx_path):\n    """Kembalikan (gambar, tabel): entri caption dari paragraf dokumen,\n    dalam urutan kemunculan. Hanya pola standar yang terdeteksi\n    (\'Gambar 1.1 ...\', \'Tabel 2 ...\', \'Gambar 1 ...\'); caption tanpa\n    nomor (mis. \'Gambar: Peta\') atau caption di dalam sel tabel\n    tidak terdeteksi. kind \'caption\' -> nomor halaman Arab relatif\n    terhadap BAB I, sama seperti heading isi.\n\n    Caption terbelah 2 paragraf (v0.6.29) ikut terdeteksi: \'Gambar\n    1.1\' di satu paragraf + deskripsi di paragraf berikut, atau\n    \'Gambar\' saja + \'1.1 ...\' di paragraf berikut. Teks gabungan\n    dipakai sebagai entri (untuk cari halaman di PDF, pencocokan\n    tanpa spasi sudah menangani baris yang terbelah); isi dokumen\n    TIDAK diubah.\n\n    Normalisasi tampilan (v0.4.0): pemisah spasi pada nomor\n    (\'Gambar 2 1 ...\', \'Gambar 2. 1 ...\') ditampilkan sebagai titik\n    (\'Gambar 2.1 ...\'). Isi dokumen TIDAK diubah; \'text\' tetap apa\n    adanya (untuk mencari halaman di PDF), \'display\' yang\n    ternormalisasi untuk daftar."""\n    doc = Document(docx_path)\n    gambar, tabel = [], []\n    paras = doc.paragraphs\n    i = 0\n    while i < len(paras):\n        p = paras[i]\n        text = norm(p.text)\n        if _is_old_toc_line(p):\n            # v0.6.1: baris daftar gambar/tabel LAMA ("Gambar 1.1 ...\\t5")\n            # bukan caption asli — abaikan supaya tidak jadi entri duplikat.\n            i += 1\n            continue\n        m = RE_CAPTION.match(text)\n        if m:\n            kind = m.group(1)\n            num = ".".join(re.split(r"[.\\s]+", m.group(2)))\n            rest = text[m.end(2):].strip()\n            display = f"{kind} {num} {rest}"\n            entry = {"text": text, "level": 0, "kind": "caption",\n                     "list": kind.lower(), "para_index": i, "bold": False,\n                     "display": display, "post_anchor": False,\n                     "search": [("text", text)]}\n            (gambar if kind.lower() == "gambar" else tabel).append(entry)\n            i += 1\n            continue\n        # v0.6.29: caption terbelah 2 paragraf.\n        merged = None\n        mh = RE_CAPTION_HEAD.match(text)\n        if mh:\n            tail = _split_caption_tail(paras, i)\n            if tail:\n                j, rest = tail\n                merged = (mh.group(1), mh.group(2), rest, j)\n        else:\n            mb = RE_CAPTION_BARE.match(text)\n            if mb:\n                tail = _split_caption_tail(paras, i, num_start=True)\n                if tail:\n                    j, rest = tail\n                    mn = RE_CAPTION_NUM.match(rest)\n                    merged = (mb.group(1), mn.group(1),\n                              rest[mn.end(1):].strip(), j)\n        if merged:\n            kind, num, rest, j = merged\n            num = ".".join(re.split(r"[.\\s]+", num))\n            full = f"{kind} {num} {rest}"\n            entry = {"text": full, "level": 0, "kind": "caption",\n                     "list": kind.lower(), "para_index": i, "bold": False,\n                     "display": full, "post_anchor": False,\n                     "search": [("text", full)]}\n            (gambar if kind.lower() == "gambar" else tabel).append(entry)\n            i = j + 1\n            continue\n        i += 1\n    return gambar, tabel'),
    (SRV,
     'VERSION = "0.6.28"  # v0.6.28:',
     'VERSION = "0.6.29"  # v0.6.29: caption terbelah 2 paragraf kini terdeteksi ("Gambar 1.1" + deskripsi di paragraf berikut, atau "Gambar" saja + "1.1 ..."; label yang diikuti sub-bab tidak digabung); varian baru v17, suite 17/17 + korpus 8/8. v0.6.28:'),
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
        bak = path + ".bak-0.6.28"
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
    print("SELESAI: engine v0.6.29 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
