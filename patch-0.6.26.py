#!/usr/bin/env python3
"""Patch v0.6.26 untuk engine DaftarIsiKu di VPS Awan.

Memperbaiki 2 cacat (ditemukan 22 Nov 2026, dokumen pelanggan tanpa
heading 'DAFTAR ISI'):
  1. Daftar isi disisipkan SETELAH judul BAB I (harusnya SEBELUM).
  2. Tidak ada jeda halaman setelah Daftar Gambar/Tabel, sehingga isi
     bab mengalir ke halaman yang sama dengan daftar.

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.26.py -o /tmp/patch-0.6.26.py && python3 /tmp/patch-0.6.26.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.25. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

TOC = "/opt/daftarisiku-engine/toc.py"
SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    # 1. Helper _insert_before (setelah _insert_after).
    (TOC,
     'def _insert_after(ref_para, new_para):\n'
     '    ref_para._p.addnext(new_para._p)\n',
     'def _insert_after(ref_para, new_para):\n'
     '    ref_para._p.addnext(new_para._p)\n'
     '\n'
     '\n'
     'def _insert_before(ref_para, new_para):\n'
     '    ref_para._p.addprevious(new_para._p)\n'),
    # 2. Anchor: daftar isi SEBELUM judul BAB I (bukan sesudahnya).
    (TOC,
     '    toc_para = find_heading("DAFTAR ISI")\n'
     '    if toc_para is not None:\n'
     '        anchor = toc_para\n'
     '        insert_before_bab = False\n'
     '    else:\n'
     '        bab_idx = next((i for i, p in enumerate(paras)\n'
     '                        if RE_BAB.match(norm(p.text))), None)\n'
     '        anchor = paras[bab_idx] if bab_idx is not None else paras[-1]\n'
     '        insert_before_bab = True\n',
     '    toc_para = find_heading("DAFTAR ISI")\n'
     '    if toc_para is not None:\n'
     '        anchor = toc_para\n'
     '        insert_before_bab = False\n'
     '        insert_before = None\n'
     '    else:\n'
     '        bab_idx = next((i for i, p in enumerate(paras)\n'
     '                        if RE_BAB.match(norm(p.text))), None)\n'
     '        insert_before_bab = True\n'
     '        # v0.6.26: daftar isi diletakkan SEBELUM judul BAB I, bukan\n'
     '        # sesudahnya (cacat 22 Nov 2026: dokumen tanpa heading\n'
     '        # "DAFTAR ISI" membuat daftar isi muncul tepat setelah judul\n'
     '        # BAB I, dan isi bab menempel di halaman Daftar Tabel).\n'
     '        if bab_idx is None:\n'
     '            anchor = paras[-1]\n'
     '            insert_before = None\n'
     '        elif bab_idx > 0:\n'
     '            anchor = paras[bab_idx - 1]\n'
     '            insert_before = None\n'
     '        else:  # BAB I paragraf pertama dokumen\n'
     '            anchor = None\n'
     '            insert_before = paras[0]\n'),
    # 3. Helper insert_before_seq + _needs_break_after + _break_after.
    (TOC,
     '    def insert_after_seq(anchor_para, new_paras):\n'
     '        """Sisipkan new_paras berurutan setelah anchor_para.\n'
     '        Kembalikan paragraf terakhir yang disisipkan."""\n'
     '        cursor = anchor_para\n'
     '        for p in new_paras:\n'
     '            _insert_after(cursor, p)\n'
     '            cursor = p\n'
     '        return cursor\n',
     '    def insert_after_seq(anchor_para, new_paras):\n'
     '        """Sisipkan new_paras berurutan setelah anchor_para.\n'
     '        Kembalikan paragraf terakhir yang disisipkan."""\n'
     '        cursor = anchor_para\n'
     '        for p in new_paras:\n'
     '            _insert_after(cursor, p)\n'
     '            cursor = p\n'
     '        return cursor\n'
     '\n'
     '    def insert_before_seq(ref_para, new_paras):\n'
     '        """Sisipkan new_paras berurutan sebelum ref_para (v0.6.26).\n'
     '        Kembalikan paragraf terakhir yang disisipkan."""\n'
     '        cursor = ref_para\n'
     '        for p in reversed(new_paras):\n'
     '            _insert_before(cursor, p)\n'
     '            cursor = p\n'
     '        return new_paras[-1]\n'
     '\n'
     '    def _needs_break_after(last_para):\n'
     '        """True bila halaman harus dijeda setelah last_para: paragraf\n'
     '        isi berikutnya belum tentu mulai di halaman baru (v0.6.26).\n'
     '        Dilewati bila paragraf isi berikutnya sudah punya\n'
     '        page_break_before atau break halaman eksplisit (mis. judul\n'
     '        BAB I) supaya tidak muncul halaman kosong. Paragraf kosong\n'
     '        dilewati saat melihat ke depan."""\n'
     '        idx = next((i for i, p in enumerate(doc.paragraphs)\n'
     '                    if p._p is last_para._p), None)\n'
     '        if idx is None:\n'
     '            return True\n'
     '        for nxt in doc.paragraphs[idx + 1:]:\n'
     '            if not nxt.text.strip():\n'
     '                continue\n'
     '            if nxt.paragraph_format.page_break_before:\n'
     '                return False\n'
     '            return not _para_has_page_break(nxt)\n'
     '        return True\n'
     '\n'
     '    def _break_after(last_para):\n'
     '        """Sisipkan jeda halaman setelah last_para, kembalikan\n'
     '        paragraf jeda itu (v0.6.26)."""\n'
     '        br = doc.add_paragraph()\n'
     '        br.add_run().add_break(WD_BREAK.PAGE)\n'
     '        _insert_after(last_para, br)\n'
     '        return br\n'),
    # 4. Sisipkan sebelum BAB I bila perlu + judul mulai di halaman baru.
    (TOC,
     '    last_toc = insert_after_seq(anchor, toc_made)\n'
     '    if insert_before_bab:\n'
     '        br = doc.add_paragraph()\n'
     '        br.add_run().add_break(WD_BREAK.PAGE)\n'
     '        _insert_after(last_toc, br)\n'
     '        last_toc = br\n',
     '    if insert_before is not None:\n'
     '        last_toc = insert_before_seq(insert_before, toc_made)\n'
     '    else:\n'
     '        last_toc = insert_after_seq(anchor, toc_made)\n'
     '    if insert_before_bab:\n'
     '        # v0.6.26: judul DAFTAR ISI yang baru dibuat mulai di halaman\n'
     '        # baru (anchor kini paragraf sebelum BAB I, jadi tanpa ini\n'
     '        # judul menempel di akhir halaman depan).\n'
     '        _page_break_before(toc_made[0])\n'
     '        br = doc.add_paragraph()\n'
     '        br.add_run().add_break(WD_BREAK.PAGE)\n'
     '        _insert_after(last_toc, br)\n'
     '        last_toc = br\n'),
    # 5. Jeda halaman setelah Daftar Gambar terakhir (bila tanpa Tabel).
    (TOC,
     '        for e in gambar:\n'
     '            g_made.append(add_entry(disp(e), e["page"], 0, False))\n'
     '        last_anchor = insert_after_seq(g_anchor, g_made)\n',
     '        for e in gambar:\n'
     '            g_made.append(add_entry(disp(e), e["page"], 0, False))\n'
     '        last_anchor = insert_after_seq(g_anchor, g_made)\n'
     '        # v0.6.26: jeda halaman setelah entri Daftar Gambar terakhir\n'
     '        # bila tidak ada Daftar Tabel sesudahnya — tanpa ini isi bab\n'
     '        # mengalir ke halaman yang sama dengan daftar (cacat 22 Nov\n'
     '        # 2026). Dilewati bila paragraf berikutnya sudah mulai di\n'
     '        # halaman baru.\n'
     '        if not tabel and _needs_break_after(last_anchor):\n'
     '            last_anchor = _break_after(last_anchor)\n'),
    # 6. Jeda halaman setelah Daftar Tabel terakhir.
    (TOC,
     '        for e in tabel:\n'
     '            t_made.append(add_entry(disp(e), e["page"], 0, False))\n'
     '        insert_after_seq(t_anchor, t_made)\n',
     '        for e in tabel:\n'
     '            t_made.append(add_entry(disp(e), e["page"], 0, False))\n'
     '        last_t = insert_after_seq(t_anchor, t_made)\n'
     '        # v0.6.26: sama seperti Daftar Gambar di atas — jeda halaman\n'
     '        # setelah entri Daftar Tabel terakhir.\n'
     '        if _needs_break_after(last_t):\n'
     '            _break_after(last_t)\n'),
    # 7. Naikkan versi di server.py.
    (SRV,
     'VERSION = "0.6.25"  # v0.6.25:',
     'VERSION = "0.6.26"  # v0.6.26: daftar isi diletakkan sebelum judul '
     'BAB I + jeda halaman setelah Daftar Gambar/Tabel terakhir (dokumen '
     'tanpa heading "DAFTAR ISI"; cacat 22 Nov 2026). v0.6.25:'),
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
        bak = path + ".bak-0.6.25"
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
    print("SELESAI: engine v0.6.26 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
