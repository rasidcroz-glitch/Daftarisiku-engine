#!/usr/bin/env python3
"""Patch v0.6.33 untuk engine DaftarIsiKu di VPS Awan.

Posisi "Kanan atas" (page_pos: top-right) menggantikan "Kiri bawah":
  - toc.py: helper header baru (_ensure_header_number_on,
    _ensure_section_headers_top_right, _strip_section_headers,
    _strip_section_footers, _strip_page_fields_from_header_footer);
    add_footer_page_numbers() menaruh nomor halaman isi di HEADER
    kanan atas bila body_align="top-right" dan mengosongkan footer
    isi dari nomor; halaman depan tetap Romawi TENGAH bawah dan
    halaman pertama tiap BAB tetap TENGAH bawah (first-page footer).
    Section BAB continuous: footer default TIDAK lagi ditengahkan
    sebagai cadangan (akan jadi nomor ganda) — diganti titlePg +
    first-page footer tengah. Perbaiki juga pemanggilan
    _ensure_bab_first_page_footer 2-argumen (TypeError laten v0.6.32).
  - server.py: POST /api/toc terima page_pos "top-right";
    "left" tetap diterima mesin untuk job lama.

Jalankan SETELAH v0.6.32 ter-deploy (patch ini mengharapkan server.py
v0.6.32 + toc.py v0.6.32 di server).

Cara pakai (di konsol VPS, sebagai root):
  curl -sSL https://raw.githubusercontent.com/rasidcroz-glitch/Daftarisiku-engine/main/patch-0.6.33.py -o /tmp/patch-0.6.33.py && python3 /tmp/patch-0.6.33.py

Aman: tiap penggantian diverifikasi harus cocok TEPAT SATU kali; bila ada
yang tidak cocok, script BERHENTI tanpa mengubah apa pun. File asli
dicadangkan dulu ke *.bak-0.6.32. Setelah patch, syntax diperiksa lalu
service daftarisiku di-restart dan /api/health dicek.
"""
import os
import subprocess
import sys

TOC = "/opt/daftarisiku-engine/toc.py"
SRV = "/opt/daftarisiku-engine/server.py"

PAIRS = [
    (TOC,
     '    # 3. Tambah tepat satu field PAGE dengan posisi benar.\n'
     '    target = next((p for p in footer.paragraphs\n'
     '                   if not p.text.strip()), None)\n'
     '    if target is None:\n'
     '        target = footer.add_paragraph()\n'
     '    _add_page_field_para(target, align)\n'
     '\n'
     '\n'
     'def _ensure_section_footers(section, align, is_first, even_odd):\n',
     '    # 3. Tambah tepat satu field PAGE dengan posisi benar.\n'
     '    target = next((p for p in footer.paragraphs\n'
     '                   if not p.text.strip()), None)\n'
     '    if target is None:\n'
     '        target = footer.add_paragraph()\n'
     '    _add_page_field_para(target, align)\n'
     '\n'
     '\n'
     'def _hdrftr_effective_part(hf):\n'
     '    """Part header/footer yang berlaku untuk section ini (milik\n'
     '    sendiri atau warisan dari section sebelumnya), atau None bila\n'
     '    tidak ada definisi di rantai section. TIDAK membuat part baru\n'
     '    — mengakses hf.part / hf._element akan membuat part baru bila\n'
     '    section pertama belum punya definisi (kasus 3 python-docx)."""\n'
     '    cur = hf\n'
     '    while cur is not None:\n'
     '        if cur._has_definition:\n'
     '            return cur._definition\n'
     '        cur = cur._prior_headerfooter\n'
     '    return None\n'
     '\n'
     '\n'
     'def _strip_page_fields_from_header_footer(hf, ensure_explicit=False):\n'
     '    """Buang semua field PAGE dari satu header/footer (teks lain\n'
     '    dipertahankan). Dipakai v0.6.33 untuk posisi \'Kanan atas\':\n'
     '    nomor halaman isi pindah ke header, jadi footer isi (dan header\n'
     '    halaman depan / halaman pertama BAB) tidak boleh lagi membawa\n'
     '    field PAGE. Bila header/footer tertaut ke sebelumnya, tautan\n'
     '    diputus dulu supaya part bersama section lain tidak ikut\n'
     '    berubah — teks non-nomor disalin kembali seperti di\n'
     '    _ensure_footer_number_on.\n'
     '    v0.6.33-perbaikan: bila part yang berlaku TIDAK punya field\n'
     '    PAGE, jangan sentuh apa pun — memutuskan tautan (membuat part\n'
     '    + referensi baru) membuat LibreOffice mengabaikan pgNumType\n'
     '    lowerRoman section itu, sehingga nomor Romawi halaman depan\n'
     '    tampil sebagai Arab (nifas: vi/vii jadi 6/7).\n'
     '    ensure_explicit=True (hanya untuk first-page header section\n'
     '    isi): part HARUS eksplisit meski tidak ada field yang dibuang\n'
     '    — tanpa referensi \'first\', halaman pertama BAB memakai header\n'
     '    default (nomor kanan atas) sehingga nomor jadi ganda (atas +\n'
     '    tengah bawah)."""\n'
     '    part = _hdrftr_effective_part(hf)\n'
     '    if part is None:\n'
     '        if ensure_explicit and hf.is_linked_to_previous:\n'
     '            hf.is_linked_to_previous = False\n'
     '        return\n'
     '    has_field = any(_para_el_has_page_field(el)\n'
     '                    for el in part._element.iter(qn("w:p")))\n'
     '    if not has_field and not ensure_explicit:\n'
     '        return\n'
     '    if hf.is_linked_to_previous:\n'
     '        old_texts = [q.text for q in hf.paragraphs\n'
     '                     if q.text.strip()\n'
     '                     and not _para_has_page_field(q)]\n'
     '        hf.is_linked_to_previous = False\n'
     '        for t in old_texts:\n'
     '            hf.add_paragraph(t)\n'
     '        return\n'
     '    if has_field:\n'
     '        for el in list(hf._element.iter(qn("w:p"))):\n'
     '            if _para_el_has_page_field(el):\n'
     '                _strip_page_field_el(el)\n'
     '\n'
     '\n'
     'def _ensure_header_number_on(header, align):\n'
     '    """Cermin _ensure_footer_number_on untuk header: pastikan satu\n'
     '    header punya tepat satu field PAGE dengan posisi align.\n'
     '    v0.6.33: dipakai untuk posisi \'Kanan atas\' (header kanan atas)."""\n'
     '    want = {WD_ALIGN_PARAGRAPH.CENTER: "center",\n'
     '            WD_ALIGN_PARAGRAPH.RIGHT: "right",\n'
     '            WD_ALIGN_PARAGRAPH.LEFT: "left"}.get(align)\n'
     '    # 1. Cari field PAGE di mana pun di xml header ini.\n'
     '    field_paras = [el for el in header._element.iter(qn("w:p"))\n'
     '                   if _para_el_has_page_field(el)]\n'
     '    if field_paras and _para_el_alignment(field_paras[0]) == want:\n'
     '        # Posisi sudah benar — buang field tambahan bila ada, selesai.\n'
     '        for extra in field_paras[1:]:\n'
     '            _strip_page_field_el(extra)\n'
     '        return\n'
     '    # 2. Perlu diubah/tambah: bila tertaut ke sebelumnya, putuskan\n'
     '    # dulu supaya hanya header ini yang berubah.\n'
     '    if header.is_linked_to_previous:\n'
     '        old_texts = [q.text for q in header.paragraphs\n'
     '                     if q.text.strip()\n'
     '                     and not _para_has_page_field(q)]\n'
     '        header.is_linked_to_previous = False\n'
     '        for t in old_texts:\n'
     '            header.add_paragraph(t)\n'
     '    else:\n'
     '        # Bukan tertaut: buang semua field lama yang ada di header ini.\n'
     '        for el in field_paras:\n'
     '            _strip_page_field_el(el)\n'
     '    # 3. Tambah tepat satu field PAGE dengan posisi benar.\n'
     '    target = next((p for p in header.paragraphs\n'
     '                   if not p.text.strip()), None)\n'
     '    if target is None:\n'
     '        target = header.add_paragraph()\n'
     '    _add_page_field_para(target, align)\n'
     '\n'
     '\n'
     'def _ensure_section_headers_top_right(section, even_odd):\n'
     '    """Header section isi untuk posisi \'Kanan atas\': field PAGE\n'
     '    rata kanan di header default (+ genap bila dokumen memakai\n'
     '    odd/even). Header first-page dikosongkan dari field PAGE bila\n'
     '    sudah dipakai (halaman pertama BAB tidak boleh punya nomor\n'
     '    atas — nomornya tetap tengah bawah lewat first-page footer)."""\n'
     '    _ensure_header_number_on(section.header, WD_ALIGN_PARAGRAPH.RIGHT)\n'
     '    if even_odd:\n'
     '        try:\n'
     '            _ensure_header_number_on(section.even_page_header,\n'
     '                                     WD_ALIGN_PARAGRAPH.RIGHT)\n'
     '        except Exception:\n'
     '            pass\n'
     '    if section.different_first_page_header_footer:\n'
     '        try:\n'
     '            # Part first-header HARUS eksplisit (kosong): tanpa\n'
     '            # referensi \'first\', halaman pertama BAB memakai header\n'
     '            # default (nomor kanan atas) sehingga nomor jadi ganda\n'
     '            # (atas + tengah bawah).\n'
     '            _strip_page_fields_from_header_footer(\n'
     '                section.first_page_header, ensure_explicit=True)\n'
     '        except Exception:\n'
     '            pass\n'
     '\n'
     '\n'
     'def _strip_section_headers(section, even_odd):\n'
     '    """Buang field PAGE dari semua header section ini (default,\n'
     '    genap, first-page bila dipakai). Dipakai v0.6.33 untuk halaman\n'
     '    depan saat posisi \'Kanan atas\': halaman depan hanya boleh\n'
     '    menampilkan nomor Romawi tengah bawah."""\n'
     '    _strip_page_fields_from_header_footer(section.header)\n'
     '    if even_odd:\n'
     '        try:\n'
     '            _strip_page_fields_from_header_footer(\n'
     '                section.even_page_header)\n'
     '        except Exception:\n'
     '            pass\n'
     '    if section.different_first_page_header_footer:\n'
     '        try:\n'
     '            _strip_page_fields_from_header_footer(\n'
     '                section.first_page_header)\n'
     '        except Exception:\n'
     '            pass\n'
     '\n'
     '\n'
     'def _strip_section_footers(section, even_odd):\n'
     '    """Buang field PAGE dari footer default (+ genap) section ini.\n'
     '    Dipakai v0.6.33 untuk section isi saat posisi \'Kanan atas\':\n'
     '    nomor pindah ke header, footer isi harus kosong dari nomor.\n'
     '    First-page footer TIDAK disentuh di sini — halaman pertama BAB\n'
     '    tetap ditangani _ensure_bab_first_page_footer / penghapusan\n'
     '    titlePg liar."""\n'
     '    _strip_page_fields_from_header_footer(section.footer)\n'
     '    if even_odd:\n'
     '        try:\n'
     '            _strip_page_fields_from_header_footer(\n'
     '                section.even_page_footer)\n'
     '        except Exception:\n'
     '            pass\n'
     '\n'
     '\n'
     'def _ensure_section_footers(section, align, is_first, even_odd):\n'),
    (TOC,
     '    body_align (v0.6.32): posisi nomor halaman ISI selain halaman\n'
     '    pertama BAB — "right" (default, sama seperti dulu), "center",\n'
     '    atau "left". Halaman depan tetap Romawi TENGAH dan halaman\n'
     '    pertama tiap BAB tetap TENGAH (aturan final user) — hanya\n'
     '    nomor isi biasa yang mengikuti pilihan ini.\n',
     '    body_align (v0.6.32): posisi nomor halaman ISI selain halaman\n'
     '    pertama BAB — "right" (default, sama seperti dulu), "center",\n'
     '    atau "left". Halaman depan tetap Romawi TENGAH dan halaman\n'
     '    pertama tiap BAB tetap TENGAH (aturan final user) — hanya\n'
     '    nomor isi biasa yang mengikuti pilihan ini.\n'
     '    body_align (v0.6.33): "top-right" (pilihan "Kanan atas") —\n'
     '    nomor halaman isi pindah ke HEADER kanan atas; footer isi\n'
     '    dikosongkan dari nomor. Halaman depan tetap Romawi TENGAH\n'
     '    bawah dan halaman pertama tiap BAB tetap TENGAH bawah.\n'),
    (TOC,
     '    doc = Document(out_path)\n'
     '    # v0.6.32: posisi nomor halaman isi (bukan halaman pertama BAB).\n'
     '    body_align_enum = {"right": WD_ALIGN_PARAGRAPH.RIGHT,\n'
     '                       "center": WD_ALIGN_PARAGRAPH.CENTER,\n'
     '                       "left": WD_ALIGN_PARAGRAPH.LEFT}.get(\n'
     '                           (body_align or "right").lower(),\n'
     '                           WD_ALIGN_PARAGRAPH.RIGHT)\n',
     '    doc = Document(out_path)\n'
     '    # v0.6.32: posisi nomor halaman isi (bukan halaman pertama BAB).\n'
     '    # v0.6.33: "top-right" — nomor isi di header kanan atas, footer\n'
     '    # isi dikosongkan dari nomor (halaman depan + halaman pertama\n'
     '    # BAB tetap tengah bawah).\n'
     '    top_right = (body_align or "right").lower() == "top-right"\n'
     '    body_align_enum = {"right": WD_ALIGN_PARAGRAPH.RIGHT,\n'
     '                       "center": WD_ALIGN_PARAGRAPH.CENTER,\n'
     '                       "left": WD_ALIGN_PARAGRAPH.LEFT,\n'
     '                       "top-right": WD_ALIGN_PARAGRAPH.RIGHT}.get(\n'
     '                           (body_align or "right").lower(),\n'
     '                           WD_ALIGN_PARAGRAPH.RIGHT)\n'),
    (TOC,
     '    for i, s in enumerate(doc.sections):\n'
     '        align = (WD_ALIGN_PARAGRAPH.CENTER if i < body_idx\n'
     '                 else body_align_enum)\n'
     '        _ensure_section_footers(s, align, i == 0, even_odd)\n',
     '    for i, s in enumerate(doc.sections):\n'
     '        if i < body_idx:\n'
     '            _ensure_section_footers(s, WD_ALIGN_PARAGRAPH.CENTER, i == 0,\n'
     '                                    even_odd)\n'
     '            if top_right:\n'
     '                # Halaman depan: header tidak boleh membawa nomor\n'
     '                # (hanya Romawi tengah bawah di footer).\n'
     '                _strip_section_headers(s, even_odd)\n'
     '        elif top_right:\n'
     '            # v0.6.33: nomor isi di header kanan atas; footer isi\n'
     '            # dikosongkan dari nomor.\n'
     '            _ensure_section_headers_top_right(s, even_odd)\n'
     '            _strip_section_footers(s, even_odd)\n'
     '        else:\n'
     '            _ensure_section_footers(s, body_align_enum, i == 0, even_odd)\n'),
    (TOC,
     '        _cont_secs.add(i)\n'
     '        _ensure_footer_number_on(\n'
     '            doc.sections[i].footer, WD_ALIGN_PARAGRAPH.CENTER)\n',
     '        _cont_secs.add(i)\n'
     '        if top_right:\n'
     '            # v0.6.33: JANGAN tengahkan footer DEFAULT section ini\n'
     '            # (cadangan v0.6.14) — footer default dipakai halaman\n'
     '            # isi berikutnya, dan nomor tengah di sana akan tampil\n'
     '            # BERSAMAAN dengan nomor kanan atas (nomor ganda).\n'
     '            # Sebagai gantinya, nomor tengah halaman pertama BAB\n'
     '            # lewat first-page footer (titlePg): hanya berlaku di\n'
     '            # halaman pertama section ini. Tambahkan ke `centered`\n'
     '            # supaya titlePg tidak dihapus loop "titlePg liar".\n'
     '            centered.add(i)\n'
     '            _ensure_bab_first_page_footer(doc.sections[i])\n'
     '        else:\n'
     '            _ensure_footer_number_on(\n'
     '                doc.sections[i].footer, WD_ALIGN_PARAGRAPH.CENTER)\n'),
    (TOC,
     '                if _all_empty:\n'
     '                    _cont_secs.add(i - 1)\n'
     '                    _ensure_footer_number_on(\n'
     '                        doc.sections[i - 1].footer,\n'
     '                        WD_ALIGN_PARAGRAPH.CENTER)\n'
     '                    _sp_prev = doc.sections[i - 1]._sectPr\n'
     '                    if _sp_prev.find(qn("w:titlePg")) is not None:\n'
     '                        _ensure_bab_first_page_footer(\n'
     '                            doc.sections[i - 1],\n'
     '                            WD_ALIGN_PARAGRAPH.CENTER)\n'
     '    for i, s in enumerate(doc.sections):\n'
     '        if i in _cont_secs:\n'
     '            continue\n'
     '        _align = (WD_ALIGN_PARAGRAPH.CENTER if i < body_idx\n'
     '                  else body_align_enum)\n'
     '        _ensure_footer_number_on(s.footer, _align)\n',
     '                if _all_empty:\n'
     '                    _cont_secs.add(i - 1)\n'
     '                    _ensure_footer_number_on(\n'
     '                        doc.sections[i - 1].footer,\n'
     '                        WD_ALIGN_PARAGRAPH.CENTER)\n'
     '                    if top_right:\n'
     '                        # Halaman judul BAB memakai header section\n'
     '                        # sebelumnya (paragraf pertamanya di sana,\n'
     '                        # v0.6.14) — buang nomor atas supaya halaman\n'
     '                        # pertama BAB hanya menampilkan nomor tengah\n'
     '                        # bawah.\n'
     '                        _strip_page_fields_from_header_footer(\n'
     '                            doc.sections[i - 1].header)\n'
     '                    _sp_prev = doc.sections[i - 1]._sectPr\n'
     '                    if _sp_prev.find(qn("w:titlePg")) is not None:\n'
     '                        # v0.6.33: perbaiki pemanggilan v0.6.32 yang\n'
     '                        # mengoper 2 argumen ke fungsi 1 argumen\n'
     '                        # (TypeError laten bila cabang ini tercapai).\n'
     '                        _ensure_bab_first_page_footer(\n'
     '                            doc.sections[i - 1])\n'
     '    for i, s in enumerate(doc.sections):\n'
     '        if i in _cont_secs:\n'
     '            continue\n'
     '        if i < body_idx:\n'
     '            _ensure_footer_number_on(s.footer, WD_ALIGN_PARAGRAPH.CENTER)\n'
     '        elif top_right:\n'
     '            _ensure_section_headers_top_right(s, even_odd)\n'
     '            _strip_section_footers(s, even_odd)\n'
     '        else:\n'
     '            _ensure_footer_number_on(s.footer, body_align_enum)\n'),
    (TOC,
     '    body_align (v0.6.32): posisi nomor halaman isi — "right"\n'
     '    (default), "center", atau "left"; diteruskan ke semua\n'
     '    pemanggilan add_footer_page_numbers. Nilai di luar itu\n'
     '    dianggap "right"."""\n',
     '    body_align (v0.6.32): posisi nomor halaman isi — "right"\n'
     '    (default), "center", atau "left"; diteruskan ke semua\n'
     '    pemanggilan add_footer_page_numbers. Nilai di luar itu\n'
     '    dianggap "right".\n'
     '    body_align (v0.6.33): ditambah "top-right" (pilihan "Kanan\n'
     '    atas") — nomor isi di header kanan atas; "left" (Kiri bawah)\n'
     '    dihapus dari website tapi tetap diterima mesin (perilaku lama)\n'
     '    bila ada job lama yang mengirimnya."""\n'),
    (SRV,
     '                               "page_pos" ("right" | "center" |\n'
     '                               "left"; posisi nomor halaman isi,\n'
     '                               default "right"; halaman depan +\n'
     '                               halaman pertama BAB tetap tengah).\n',
     '                               "page_pos" ("right" | "center" |\n'
     '                               "top-right"; posisi nomor halaman isi,\n'
     '                               default "right"; halaman depan +\n'
     '                               halaman pertama BAB tetap tengah).\n'
     '                               "left" (Kiri bawah) dihapus dari\n'
     '                               website v0.6.33 tapi tetap diterima\n'
     '                               mesin bila ada job lama yang\n'
     '                               mengirimnya.\n'),
    (SRV,
     'VERSION = "0.6.32"  # v0.6.32: pilihan posisi nomor halaman isi (page_pos: right/center/left; halaman depan + halaman pertama BAB tetap tengah).',
     'VERSION = "0.6.33"  # v0.6.33: posisi "Kanan atas" (page_pos: top-right) — nomor halaman isi di header kanan atas, footer isi dikosongkan; "Kiri bawah" (left) dihapus dari website. Halaman depan + halaman pertama BAB tetap tengah. Perbaiki juga pemanggilan _ensure_bab_first_page_footer 2-argumen (TypeError laten v0.6.32). v0.6.32: pilihan posisi nomor halaman isi (page_pos: right/center/left; halaman depan + halaman pertama BAB tetap tengah).'),
    (SRV,
     '    # v0.6.32: posisi nomor halaman isi ("right"|"center"|"left").\n'
     '    page_pos = (request.form.get("page_pos", "right") or "right").strip().lower()\n'
     '    if page_pos not in ("right", "center", "left"):\n'
     '        page_pos = "right"\n',
     '    # v0.6.33: posisi nomor halaman isi ("right"|"center"|"top-right";\n'
     '    # "left" tetap diterima mesin untuk job lama).\n'
     '    page_pos = (request.form.get("page_pos", "right") or "right").strip().lower()\n'
     '    if page_pos not in ("right", "center", "left", "top-right"):\n'
     '        page_pos = "right"\n'),
    (TOC,
     '    for i, s in enumerate(doc.sections):\n'
     '        if i == 0 or i in centered:\n'
     '            continue\n'
     '        _sp = s._sectPr\n'
     '        _tp = _sp.find(qn("w:titlePg"))\n'
     '        if _tp is not None:\n'
     '            _sp.remove(_tp)\n'
     '    doc.save(out_path)\n'
     '    return True\n',
     '    for i, s in enumerate(doc.sections):\n'
     '        if i == 0 or i in centered:\n'
     '            continue\n'
     '        _sp = s._sectPr\n'
     '        _tp = _sp.find(qn("w:titlePg"))\n'
     '        if _tp is not None:\n'
     '            _sp.remove(_tp)\n'
     '    # v0.6.33: normalisasi urutan anak sectPr SETELAH semua perubahan\n'
     '    # header/footer \u2014 python-docx menyisipkan referensi header baru\n'
     '    # SETELAH referensi footer yang sudah ada (urutan tidak valid\n'
     '    # menurut schema), dan LibreOffice lalu mengabaikan footer\n'
     '    # first-page section itu (nifas: halaman pertama BAB II-IV\n'
     '    # tanpa nomor sama sekali).\n'
     '    for _s in doc.sections:\n'
     '        _normalize_sectPr_order(_s._sectPr)\n'
     '    doc.save(out_path)\n'
     '    return True\n'),
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
        bak = path + ".bak-0.6.32"
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
    return 0


def main():
    if apply_pairs() != 0:
        return 1
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
    print("SELESAI: engine v0.6.33 jalan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

