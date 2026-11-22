#!/usr/bin/env python3
"""Patch v0.6.27 untuk engine DaftarIsiKu di VPS Awan.

Pengerasan umum (ditemukan 22 Nov 2026 oleh harness regresi: korpus
8 dokumen asli + 16 varian sintetis; v0.6.26 lolos 4/8 dan 9/16):
  1. Jeda halaman setelah tiap daftar kini berlaku umum tanpa memandang
     urutan heading di dokumen — dokumen YANG PUNYA heading DAFTAR ISI
     (makalah ppkn) dan Daftar Tabel SEBELUM Daftar Gambar (Gusti Ayu)
     tidak lagi menempelkan BAB I di halaman daftar.
  2. _needs_break_after kini mengenali break pada paragraf KOSONG
     (run break / sectPr nextPage) — jeda tambahan di atas break sumber
     jadi halaman kosong (UJI_COBA_2).
  3. Run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak
     ada isi di antara judul BAB I dan sectPr penutup (Gusti Ayu).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.27.py -o /tmp/patch-0.6.27.py && python3 /tmp/patch-0.6.27.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.26. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

TOC = "/opt/daftarisiku-engine/toc.py"
SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    # 1. _needs_break_after: break pada paragraf kosong + judul daftar berikut.
    (TOC,
     '        Dilewati bila paragraf isi berikutnya sudah punya\n        page_break_before atau break halaman eksplisit (mis. judul\n        BAB I) supaya tidak muncul halaman kosong. Paragraf kosong\n        dilewati saat melihat ke depan."""\n        idx = next((i for i, p in enumerate(doc.paragraphs)\n                    if p._p is last_para._p), None)\n        if idx is None:\n            return True\n        for nxt in doc.paragraphs[idx + 1:]:\n            if not nxt.text.strip():\n                continue\n            if nxt.paragraph_format.page_break_before:\n                return False\n            return not _para_has_page_break(nxt)\n        return True\n\n    def _break_after(last_para):\n',
     '        Dilewati bila paragraf isi berikutnya sudah punya\n        page_break_before atau break halaman eksplisit (mis. judul\n        BAB I) supaya tidak muncul halaman kosong. Paragraf kosong\n        dilewati saat melihat ke depan — KECUALI bila paragraf\n        kosong itu sendiri membawa break (run break halaman,\n        page_break_before, atau sectPr nextPage): dua break\n        berurutan = satu halaman kosong (UJI_COBA_2: paragraf\n        break + sectPr sumber sebelum BAB I sudah menjeda, jeda\n        tambahan dari mesin jadi halaman kosong). v0.6.27: dilewati\n        juga bila paragraf berikut adalah judul DAFTAR GAMBAR/TABEL\n        yang sudah ada di dokumen — judul itu mendapat\n        page_break_before di bagiannya sendiri, jadi jeda di sini\n        jadi halaman kosong."""\n        idx = next((i for i, p in enumerate(doc.paragraphs)\n                    if p._p is last_para._p), None)\n        if idx is None:\n            return True\n        for nxt in doc.paragraphs[idx + 1:]:\n            if nxt.paragraph_format.page_break_before:\n                return False\n            if _para_has_page_break(nxt):\n                return False\n            _pPr = nxt._p.find(qn("w:pPr"))\n            _sect = (_pPr.find(qn("w:sectPr"))\n                     if _pPr is not None else None)\n            if _sect is not None and _sectPr_break_is_new_page(_sect):\n                return False\n            if not nxt.text.strip():\n                continue\n            if norm(nxt.text).upper() in ("DAFTAR GAMBAR",\n                                          "DAFTAR TABEL"):\n                return False\n            return True\n        return True\n\n    def _break_after(last_para):\n'),
    # 2. Jeda setelah entri DAFTAR ISI terakhir berlaku untuk SEMUA jalur.
    (TOC,
     '        # baru (anchor kini paragraf sebelum BAB I, jadi tanpa ini\n        # judul menempel di akhir halaman depan).\n        _page_break_before(toc_made[0])\n        br = doc.add_paragraph()\n        br.add_run().add_break(WD_BREAK.PAGE)\n        _insert_after(last_toc, br)\n        last_toc = br\n    # v0.6.19: judul DAFTAR ISI jangan sampai terkurung sendirian di\n    # satu halaman oleh section break nextPage (lihat\n    # _relax_list_heading_break). Hanya bila heading itu dari dokumen\n',
     '        # baru (anchor kini paragraf sebelum BAB I, jadi tanpa ini\n        # judul menempel di akhir halaman depan).\n        _page_break_before(toc_made[0])\n    # v0.6.27: jeda halaman setelah entri daftar isi TERAKHIR — dulu\n    # hanya di jalur tanpa heading DAFTAR ISI (v0.6.26); bila dokumen\n    # PUNYA heading itu, jeda tidak ada sehingga daftar isi yang\n    # pendek membuat judul BAB I (+ isi bab) mengalir ke halaman\n    # yang sama (makalah ppkn: 12 entri). Kini berlaku untuk SEMUA\n    # jalur; _needs_break_after melewatinya bila paragraf berikut\n    # sudah mulai di halaman baru atau berupa judul daftar berikut.\n    if _needs_break_after(last_toc):\n        last_toc = _break_after(last_toc)\n    # v0.6.19: judul DAFTAR ISI jangan sampai terkurung sendirian di\n    # satu halaman oleh section break nextPage (lihat\n    # _relax_list_heading_break). Hanya bila heading itu dari dokumen\n'),
    # 3. Jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel.
    (TOC,
     '        for e in gambar:\n            g_made.append(add_entry(disp(e), e["page"], 0, False))\n        last_anchor = insert_after_seq(g_anchor, g_made)\n        # v0.6.26: jeda halaman setelah entri Daftar Gambar terakhir\n        # bila tidak ada Daftar Tabel sesudahnya — tanpa ini isi bab\n        # mengalir ke halaman yang sama dengan daftar (cacat 22 Nov\n        # 2026). Dilewati bila paragraf berikutnya sudah mulai di\n        # halaman baru.\n        if not tabel and _needs_break_after(last_anchor):\n            last_anchor = _break_after(last_anchor)\n    else:\n        # v0.6.18: daftar kosong tapi dokumen punya heading template\n',
     '        for e in gambar:\n            g_made.append(add_entry(disp(e), e["page"], 0, False))\n        last_anchor = insert_after_seq(g_anchor, g_made)\n        # v0.6.27: jeda halaman setelah entri Daftar Gambar terakhir —\n        # dulu hanya bila tidak ada Daftar Tabel (v0.6.26 mengasumsikan\n        # urutan Gambar lalu Tabel); padahal urutan daftar mengikuti\n        # heading di dokumen — bila Daftar Tabel SUDAH ada sebelum\n        # Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman\n        # Daftar Gambar. Kini selalu dicek; _needs_break_after\n        # melewatinya bila paragraf berikut sudah mulai di halaman\n        # baru atau berupa judul Daftar Tabel yang memang dijeda\n        # sendiri.\n        if _needs_break_after(last_anchor):\n            last_anchor = _break_after(last_anchor)\n    else:\n        # v0.6.18: daftar kosong tapi dokumen punya heading template\n'),
    # 4. Run break footer hanya bila tidak ada isi sebelum sectPr penutup.
    (TOC,
     '            _cols = _csp.find(qn("w:cols"))\n            if _cols is not None:\n                _csp.remove(_cols)\n            if _moved and _close_j + 1 < len(doc.paragraphs):\n                _nxt = doc.paragraphs[_close_j + 1]\n                _r = OxmlElement("w:r")\n                _br = OxmlElement("w:br")\n',
     '            _cols = _csp.find(qn("w:cols"))\n            if _cols is not None:\n                _csp.remove(_cols)\n            # v0.6.27: run break ini hanya bila TIDAK ada isi di antara\n            # judul BAB I dan sectPr penutup — bila ada (Gusti Ayu:\n            # sub-heading + 2 paragraf isi mengisi halaman judul),\n            # section break penutup itu sendiri sudah memisahkan isi\n            # lanjutannya ke halaman baru, dan run break tambahan\n            # justru membuat satu halaman kosong di antaranya. Run\n            # break hanya diperlukan bila section Arab baru hanya\n            # berisi judul (nifas): dua section break berurutan\n            # digabung LibreOffice sehingga isi naik ke halaman judul.\n            _has_content = any(\n                doc.paragraphs[_j].text.strip()\n                for _j in range(bab_j + 1, _close_j))\n            if (_moved and not _has_content\n                    and _close_j + 1 < len(doc.paragraphs)):\n                _nxt = doc.paragraphs[_close_j + 1]\n                _r = OxmlElement("w:r")\n                _br = OxmlElement("w:br")\n'),
    # 5. Naikkan versi di server.py.
    (SRV,
     'VERSION = "0.6.26"  # v0.6.26:',
     'VERSION = "0.6.27"  # v0.6.27: jeda halaman setelah tiap daftar berlaku umum tanpa memandang urutan heading di dokumen — (a) jeda setelah entri DAFTAR ISI terakhir juga di jalur dokumen YANG PUNYA heading DAFTAR ISI (makalah ppkn: BAB I + isi bab menempel di halaman daftar isi); (b) jeda setelah Daftar Gambar tidak lagi dilewati bila ada Daftar Tabel — urutan daftar mengikuti heading dokumen, dan bila Daftar Tabel sudah ada SEBELUM Daftar Gambar (Gusti Ayu), judul BAB I menempel di halaman Daftar Gambar; (c) _needs_break_after kini mengenali break pada paragraf KOSONG (run break / sectPr nextPage) — jeda tambahan di atas break sumber jadi halaman kosong (UJI_COBA_2); (d) run break footer v0.6.11 langkah (b) hanya disisipkan bila tidak ada isi di antara judul BAB I dan sectPr penutup — bila ada isi, section break penutup sudah memisahkan sendiri dan run break tambahan jadi halaman kosong (Gusti Ayu). v0.6.26:'),
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
        bak = path + ".bak-0.6.26"
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
    print("SELESAI: engine v0.6.27 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
