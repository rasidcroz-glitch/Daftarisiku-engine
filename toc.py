"""Mesin Daftar Isi Otomatis — DaftarIsiKu.

Alur:
  1. Ekstrak struktur heading dari .docx:
     - BAB (gaya apa pun, termasuk Normal; judul boleh di baris berikut),
     - sub-bab Heading 3 (tanpa huruf -> diberi A./B./C. di daftar isi),
     - sub-sub List Paragraph (pendek / istilah+definisi / tebal),
     - sub-sub-sub (a./b./c.), grup non-bold, halaman depan, pustaka,
       lampiran.
  2. Render .docx -> PDF dengan LibreOffice headless, lalu cari halaman
     tiap heading dari teks PDF (nomor halaman TIDAK ditebak).
  3. Penomoran: cover = i, halaman depan berurutan Romawi (ii, iii, iv,
     v, ...); angka Arab (1, 2, 3, ...) baru mulai dari BAB I.
  4. Keluaran: .docx baru berisi daftar isi yang disisipkan setelah
     heading "DAFTAR ISI" (+ DAFTAR GAMBAR / DAFTAR TABEL bila ada).
     Setelah penyisipan, total halaman diukur ulang dan label halaman
     depan yang bergeser (mis. DAFTAR TABEL/GAMBAR yang judulnya ada
     setelah DAFTAR ISI) dikoreksi dari PDF hasil, lalu verifikasi ulang.
  5. Fitur tambahan (v0.3.0): DAFTAR GAMBAR dan DAFTAR TABEL —
     caption pola "Gambar 1.1 ..."/"Tabel 2 ..." dideteksi dari
     paragraf, nomor halaman dihitung dari PDF yang sama, lalu
     disisipkan setelah DAFTAR ISI (judul bagian tebal, entri tidak).
     Bagian yang tidak punya entri tidak dibuat.
  6. Normalisasi tampilan (v0.4.0): penomoran caption yang salah ketik
     ("Gambar 2 1 ..." -> "Gambar 2.1 ...") hanya diperbaiki di daftar
     yang ditampilkan, isi dokumen tidak diubah.
  7. Hierarki List Paragraph (v0.5.0): ilvl Word didahulukan atas
     indentasi untuk butir bukan istilah (teori sejajar ilvl=1 tidak
     lagi tersarang di bawah dimensi ilvl=2 yang indentasinya nyaris
     sama); "Kesimpulan" yang ditulis tepat sebelum judul BAB
     dipindah ke bawah BAB itu sebagai sub-bab pertama bila BAB
     tersebut belum punya sub-bab "Kesimpulan".
  8. Nomor halaman footer (v0.6.0): field PAGE di footer — halaman
     depan Romawi kecil tengah bawah (cover tidak menampilkan nomor),
     mulai BAB I angka Arab kanan bawah dari 1.
  9. Daftar isi lama diabaikan (v0.6.1): dokumen yang sudah berisi
     daftar isi lama ("BAB I\t1", "1. Rumusan Masalah\t7" — baris
     Normal ber-akhir tab/titik + nomor halaman, atau tab saja)
     tidak lagi dibaca sebagai heading (tidak jadi entri duplikat),
     baris-baris lama itu dihapus dari dokumen hasil supaya hanya ada
     satu daftar isi yang baru, dan PDF input di-render dari salinan
     yang daftar lamanya sudah dibuang supaya pencarian halaman tidak
     ketemu teks daftar lama. Karena penyisipan daftar isi bisa
     menggeser alur halaman depan (BAB I tidak selalu mulai di
     halaman baru), SEMUA label dikoreksi ulang dari PDF hasil
     sebelum dokumen final disimpan. Judul yang sangat panjang
     dicari dengan 60 karakter pertamanya bila terbelah halaman.
     Berlaku juga untuk baris daftar gambar/tabel lama.
  10. Hemat memori (v0.6.2): sebelum tiap render LibreOffice, arena
      Python dikembalikan ke OS (malloc_trim) supaya RSS Python turun
      dari ~160 MB ke ~60 MB — LibreOffice sendiri ~325 MB, jadi total
      puncak ~390 MB dan muat di container 512 MB untuk dokumen besar
      (144 halaman). Gunicorn 1 thread supaya dua request tidak
      me-render bersamaan.
  11. Samakan format dengan file asli (v0.6.3): sebelum baris daftar
      lama dihapus, formatnya dibaca dulu — indent tiap level, posisi
      tab stop kanan, spasi antar-baris, dan font — lalu dipakai untuk
      daftar isi/gambar/tabel yang baru. Bila dokumen tidak punya
      daftar lama, dipakai nilai default (indent 0,6 cm/level, tab di
      ujung lebar teks, Times New Roman 12).
  12. Spasi antar-entri dilegakan (v0.6.4): daftar isi asli dokumen
      umumnya 1 spasi (space_after 2 pt) sehingga baris-barisnya
      terlihat dempet; entri baru memakai space_after minimal 6 pt
      supaya ada napas antar-baris (bila format asli sudah lebih lega,
      dipakai nilai aslinya). Baris lanjutan judul yang sama tetap
      rapat (1 spasi) supaya judul tidak terpecah.
  13. Lembar terpisah (v0.6.5): DAFTAR GAMBAR dan DAFTAR TABEL
      masing-masing mulai di halaman baru (pageBreakBefore pada
      judulnya, baik judul asli dokumen maupun judul yang dibuat
      mesin) — tidak digabung dengan daftar isi atau satu sama lain.
  14. Dokumen ber-Heading (v0.6.6): dokumen yang memakai Heading 1/2
      asli Word (bukan List Paragraph) kini dibaca dari stylenya —
      Heading 1 jadi sub-bab, Heading 2 jadi sub-sub — dan butir
      List Paragraph di dokumen seperti itu kena gerbang ketat: di
      bawah BAB hanya butir tebal yang diterima, di bawah sub-bab
      hanya butir level 2 yang induknya heading asli dan nomor
      listnya sama dengan induk. Dokumen skripsi biasa (tanpa
      Heading 1/2 isi) tidak berubah. "LEMBAR PERSETUJUAN" dikenal
      sebagai halaman depan. _strip_prefix kini membuang nomor
      bertitik penuh ("4.1 ..." -> "", bukan "1 ..."). Judul bagian
      yang tertulis sebelum judul BAB-nya ditampung lalu ditempel di
      bawah BAB itu, dan diberi page_break_before supaya tidak
      tertinggal di akhir halaman daftar isi (keepNext tidak mempan
      karena ada section break di judul BAB). Baris judul BAB boleh
      bergaya Heading 1/2, bukan hanya Normal.
  15. Nomor halaman awal BAB di tengah bawah (v0.6.10): halaman
      pertama tiap BAB memakai footer first-page (titlePg) dengan
      field PAGE rata tengah; halaman lain tetap Romawi tengah
      (halaman depan) atau Arab kanan (isi). TIDAK ada section break
      baru dan TIDAK ada perubahan tipe break: nextPage sebelum BAB
      terbukti merusak pagination di LibreOffice (memecah daftar
      isi karena sectPr sumber tanpa w:type, halaman kosong karena
      run break kolom jadi double break, judul terbelah dua kolom
      karena section body sumber ber-cols 2). Sebagai gantinya,
      titlePg hanya dipasang pada section YANG SUDAH ADA yang
      paragraf berisi pertamanya adalah judul BAB yang mulai di
      halaman baru — halaman pertama section itu memang halaman BAB,
      jadi nomornya jadi di tengah tanpa mengubah pagination. BAB
      yang bukan paragraf pertama section-nya (BAB I, BAB tengah
      halaman) dibiarkan seperti semula. Tiga hal pendukung: (a)
      cover dikosongkan SEBELUM footer tengah BAB dibuat — urutan
      sebaliknya membuat LibreOffice menampilkan '1' di cover;
      (b) titlePg LIAR warisan sumber dihapus dari section yang
      paragraf pertamanya bukan BAB baru (mencegah halaman yang
      salah jadi tengah); (c) sectPr continuous di BAB I hanya
      ditambah bila section BAB I tidak memuat judul halaman depan
      — tanpanya daftar isi nifas/Gusti pecah, dengannya halaman
      depan Gunak bergeser. Sekalian diperbaiki: deteksi BAB kini
      kenal judul BAB telanjang ("BAB I" + judul di baris berikut)
      dan judul pending sebelum BAB I, sehingga Gunak (BAB
      telanjang) dan nifas (pending "1.1 Latar Belakang") tidak
      lagi jatuh ke fallback semua-tengah; dan section BAB mendapat
      referensi footer 'first' SENDIRI (bukan tautan) karena
      LibreOffice tidak mewariskan referensi 'first' melewati
      section tanpa titlePg. Batasan yang tersisa: BAB tengah
      halaman tetap kanan; nilai nomor BAB I tengah halaman tetap
      Romawi (bukan Arab 1); dokumen yang halaman depannya satu
      section dengan isi (Gunak) halaman depannya Arab kanan dan
      nilai isi bergeser sejumlah halaman depan di section itu.
  16. Halaman pertama BAB I tampil "1" Arab (v0.6.11): sectPr
      continuous di judul BAB I dipindah ke paragraf sebelumnya,
      supaya section Arab (decimal start=1) mulai TEPAT di judul
      BAB I dan halaman pertamanya mendapat titlePg + footer
      tengah "1" — bukan "viii"/"xii" Romawi lagi. Pemindahan saja
      merusak rendering di LibreOffice, jadi tiga pendamping:
      (a) run break halaman/kolom di paragraf sebelumnya dibuang
      dan diganti pageBreakBefore di judul BAB I (sectPr + break
      run dalam satu paragraf = dua break: halaman kosong + BAB I
      geser); (b) w:cols di sectPr penutup section Arab baru
      dihapus (bila num=2, judul BAB I terbelah huruf per huruf
      ke dua kolom karena w:ind right-nya melebihi lebar kolom;
      mengubah num jadi 1 tidak mempan, menghapus elemennya yang
      mempan); (c) run break halaman (w:br type=page) disisipkan
      di awal paragraf isi pertama sesudah sectPr penutup itu,
      karena LibreOffice menggabung section break nextPage
      penutup itu dengan perubahan format Romawi->Arab di batas
      pindahan sehingga body naik ke halaman BAB I (total halaman
      berkurang 1) — break level-run tidak digabung, pagination
      baseline pulih. Tipe sectPr penutup dibiarkan nextPage
      supaya restart start=1 tetap jalan (continuous membuat
      footer BAB I tampil "8"). Tidak ada section break baru dan
      BAB I tidak dipaksa ke halaman baru. Batasan: BAB tengah
      halaman tetap kanan; Gunak (halaman depan satu section
      dengan isi) BAB I-nya tetap Romawi bila tidak bisa
      diperbaiki tanpa merusak pagination — dilaporkan apa adanya.

Format daftar isi (contoh dosen): titik-titik sampai kanan, nomor
halaman rata kanan; yang tebal hanya halaman depan + judul BAB.
"""

import copy
import ctypes
import gc
import os
import re
import shutil
import subprocess
import tempfile

import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Twips

def _trim_memory():
    """Kembalikan memori arena Python ke OS (v0.6.2): objek Document
    besar sudah dibebaskan (refcount) tapi arena-nya dipertahankan
    Python — tanpa trim, RSS Python bisa 150 MB+ saat LibreOffice
    (325 MB) jalan bersamaan dan total melewati 512 MB. Dipanggil
    tepat sebelum tiap render LibreOffice. Aman bila gagal (glibc
    tidak ada / bukan Linux): lanjut tanpa trim."""
    try:
        gc.collect()
        libc = ctypes.CDLL("libc.so.6")
        libc.malloc_trim(0)
    except Exception:
        pass


# ---------------------------------------------------------------- pola
RE_BAB = re.compile(r"^BAB\s+([IVXLC]+)\b", re.IGNORECASE)
RE_BAB_BARE = re.compile(r"^BAB\s+[IVXLC]+\s*$", re.IGNORECASE)
RE_LETTER = re.compile(r"^([A-Z])\.\s+\S")
RE_NUMBER = re.compile(r"^(\d+)\.\s+\S")
RE_LOWER = re.compile(r"^([a-z])\.\s+\S")
# Caption gambar/tabel pola standar: "Gambar 1.1 Peta Lokasi",
# "Tabel 2.1 Data Responden", nomor tunggal "Gambar 1 Denah", dan
# salah ketik pemisah spasi "Gambar 2 1 Kerangka Berpikir" (v0.4.0).
RE_CAPTION = re.compile(r"^(Gambar|Tabel)\s+(\d+(?:[.\s]\d+)*)\s+\S",
                         re.IGNORECASE)
RE_LAMPIRAN = re.compile(r"^LAMPIRAN\s+\d+\b", re.IGNORECASE)
# Baris daftar isi LAMA (v0.6.1): "BAB I\t1", "1. Rumusan Masalah\t7",
# "HALAMAN PERSETUJUAN PEMBIMBING\tii" — teks diakhiri tab / titik-titik /
# spasi ganda + nomor halaman (Arab atau Romawi), atau tab saja di akhir
# (nomornya terhapus). Judul asli tidak pernah diakhiri pola itu, jadi
# pola ini aman dipakai sebagai penanda baris daftar isi lama yang
# harus diabaikan (dan dihapus dari hasil).
RE_OLD_TOC_STRONG = re.compile(
    r"(?:\t\s*([0-9]+|[ivxlcdm]+)?\s*$"
    r"|\.{2,}\s*([0-9]+|[ivxlcdm]+)\s*$"
    r"|…+\s*([0-9]+|[ivxlcdm]+)\s*$)",
    re.IGNORECASE)
RE_OLD_TOC_SPACES = re.compile(
    r" {2,}\s*([0-9]+|[ivxlcdm]+)\s*$", re.IGNORECASE)


def _is_old_toc_line(p):
    """True bila paragraf ini baris daftar isi lama: teksnya diakhiri
    tab/titik-titik + nomor halaman.
    Dipakai raw text (bukan norm) karena norm() menghapus tab.
    v0.6.12: gaya tidak lagi disyaratkan Normal/TOC — daftar lama
    pada revisi nifas __1 bergaya 'Body Text' sehingga lolos dari
    deteksi dan ikut jadi entri (duplikat BAB, nomor '?' dan huruf
    salah). Sinyal kuat (tab / titik-titik / elipsis + nomor) berlaku
    untuk semua gaya KECUALI Heading (judul asli tidak pernah jadi
    baris daftar lama); sinyal lemah (spasi ganda + nomor) tetap
    hanya untuk Normal/TOC seperti sebelumnya."""
    style = p.style.name
    if style in ("Heading 1", "Heading 2", "Heading 3"):
        return False
    text = p.text or ""
    if RE_OLD_TOC_STRONG.search(text):
        return True
    if (style == "Normal" or style.startswith("TOC")) \
            and RE_OLD_TOC_SPACES.search(text):
        return True
    return False


def _para_has_page_break(p):
    """Paragraf mengandung page break eksplisit (w:br type=page)."""
    return 'w:type="page"' in p._p.xml


def _capture_old_toc_format(doc):
    """Baca format baris daftar isi LAMA sebelum dihapus (v0.6.3):
    indent tiap level, tab stop kanan, spasi, dan font — supaya daftar
    isi baru memakai format yang SAMA PERSIS dengan file aslinya,
    bukan nilai hardcode. Kembalikan dict, atau None bila dokumen
    tidak punya baris daftar lama."""
    from collections import Counter
    paras = doc.paragraphs
    indents, tabs, afters, befores, firsts, fonts = [], [], [], [], [], []
    for idx, p in enumerate(paras):
        if norm(p.text).upper() not in ("DAFTAR ISI", "DAFTAR GAMBAR",
                                        "DAFTAR TABEL"):
            continue
        for q in paras[idx + 1:]:
            if _para_has_page_break(q):
                break
            if not norm(q.text):
                continue
            if not _is_old_toc_line(q):
                break
            pf = q.paragraph_format
            indents.append(pf.left_indent.cm
                           if pf.left_indent is not None else 0.0)
            for t in pf.tab_stops:
                if str(t.alignment).startswith("RIGHT"):
                    tabs.append(round(t.position.cm, 2))
            if pf.space_after is not None:
                afters.append(round(pf.space_after.pt, 1))
            if pf.space_before is not None:
                befores.append(round(pf.space_before.pt, 1))
            firsts.append(pf.first_line_indent.cm
                          if pf.first_line_indent is not None else 0.0)
            if q.runs:
                r = q.runs[0]
                fonts.append((r.font.name,
                              r.font.size.pt if r.font.size else None))
    if not indents:
        return None
    uniq = sorted(set(round(v, 2) for v in indents))
    def _common(vals, default=None):
        vals = [v for v in vals if v is not None]
        return Counter(vals).most_common(1)[0][0] if vals else default
    return {
        "indents": uniq,
        "tab_cm": _common(tabs),
        "space_after_pt": _common(afters),
        "space_before_pt": _common(befores),
        "first_cm": _common(firsts, 0.0),
        "font_name": _common([f[0] for f in fonts]),
        "font_size_pt": _common([f[1] for f in fonts]),
    }


def _remove_old_toc_runs(doc):
    """Hapus baris-baris daftar lama dari dokumen (v0.6.1).

    Setelah heading 'DAFTAR ISI' / 'DAFTAR GAMBAR' / 'DAFTAR TABEL',
    hapus run paragraf kosong + baris daftar lama, sampai paragraf
    pertama yang bukan keduanya (atau paragraf berisi page break —
    itu pemisah bagian, jangan dihapus). Dipakai sebelum render PDF
    input (supaya pencarian halaman tidak ketemu teks daftar lama)
    dan saat membangun dokumen hasil (supaya hanya ada satu daftar
    yang baru). Kembalikan jumlah paragraf yang dihapus."""
    paras = doc.paragraphs
    doomed = []
    seen = set()
    for idx, p in enumerate(paras):
        if norm(p.text).upper() not in ("DAFTAR ISI", "DAFTAR GAMBAR",
                                        "DAFTAR TABEL"):
            continue
        for q in paras[idx + 1:]:
            if _para_has_page_break(q):
                break
            if not norm(q.text) or _is_old_toc_line(q):
                if id(q) not in seen:
                    seen.add(id(q))
                    doomed.append(q)
                continue
            break
    for p in doomed:
        p._element.getparent().remove(p._element)
    return len(doomed)
FRONT_TITLES = [
    "PERNYATAAN ORIGINALITAS",
    "HALAMAN PERSETUJUAN PEMBIMBING",
    "HALAMAN PENGESAHAN TIM PENGUJI",
    "KATA PENGANTAR",
    "DAFTAR ISI",
    "HALAMAN PERSETUJUAN",
    "LEMBAR PERSETUJUAN",
    "LEMBAR PENGESAHAN",
    "ABSTRAK",
    "DAFTAR TABEL",
    "DAFTAR GAMBAR",
    "DAFTAR LAMPIRAN",
]
# Pencocokan halaman depan tidak peka spasi: dokumen asli kadang
# menulis "PERNYATAANORIGINALITAS" (satu run tanpa spasi).
FRONT_BY_NOSPACE = {t.replace(" ", ""): t for t in FRONT_TITLES}


def to_roman(n):
    """1 -> 'i', 2 -> 'ii', ... (huruf kecil, sesuai contoh dosen)."""
    table = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
    out = ""
    for val, sym in table:
        while n >= val:
            out += sym
            n -= val
    return out


def norm(text):
    return re.sub(r"\s+", " ", text or "").strip()


def _nospace(text):
    return re.sub(r"\s+", "", text or "").lower()


def _strip_prefix(text, lvl=None):
    """Buang awalan penomoran ('A. ', '1. ', 'a. ', '4.1 ') dari judul.
    Nomor bertitik dibuang penuh ('4.1 Asuhan...' -> 'Asuhan...'),
    bukan hanya kelompok pertamanya."""
    return re.sub(r"^(?:[A-Za-z]\.|\d+(?:\.\d+)+\.?|\d+\.)\s*", "",
                  text or "")


# ---------------------------------------------------------------- info paragraf
def _ilvl(p):
    """Level list Word (numPr/ilvl) atau None bila tidak ada."""
    try:
        pPr = p._p.pPr
        if pPr is None:
            return None
        numPr = pPr.numPr
        if numPr is None:
            return None
        il = numPr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl")
        if il is None:
            return None
        v = il.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
        return int(v) if v is not None else None
    except Exception:
        return None


def _numId(p):
    """Nomor list Word (numPr/numId) atau None bila tidak ada.
    ilvl hanya boleh dibandingkan antar-paragraf yang numId-nya
    sama; satu dokumen bisa memakai banyak list dengan ilvl yang
    tidak sebanding."""
    try:
        pPr = p._p.pPr
        if pPr is None:
            return None
        numPr = pPr.numPr
        if numPr is None:
            return None
        nid = numPr.find(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId")
        if nid is None:
            return None
        v = nid.get(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
        return int(v) if v is not None else None
    except Exception:
        return None


def _indent_cm(p):
    li = p.paragraph_format.left_indent
    return round(li.cm, 2) if li is not None else None


def _is_bold(p):
    return any(r.bold for r in p.runs if r.text.strip())


def _known_terms(doc):
    """Istilah pendek dari List Paragraph (<=60 karakter) di seluruh
    dokumen — dipakai untuk mengenali butir 'istilah + definisi'
    ('Pengenalan Kebutuhan Konsumen menyadari ...' -> 'Pengenalan
    Kebutuhan'). Diurutkan terpanjang dulu agar cocok yang terpanjang."""
    terms = set()
    for p in doc.paragraphs:
        if p.style.name == "List Paragraph":
            t = norm(p.text)
            if 4 <= len(t) <= 60 and not t.endswith("."):
                terms.add(t)
    return sorted(terms, key=len, reverse=True)


def _extract_term(text, known_terms):
    """Kembalikan istilah saja dari butir 'istilah + definisi', atau
    None bila paragraf ini bukan butir semacam itu."""
    if ": " in text:
        term = text.split(": ", 1)[0].strip()
        if 2 <= len(term) <= 60 and not term.endswith("."):
            return term
    for term in known_terms:
        if len(term) >= 4 and text.startswith(term + " "):
            return term
    return None


# ---------------------------------------------------------------- ekstraksi
def extract_structure(docx_path):
    """Kembalikan list entri: dict(text, display, level, kind,
    para_index, ...).

    level: 0 = halaman depan / BAB / DAFTAR PUSTAKA / DAFTAR LAMPIRAN,
           1 = sub-bab (A./B.), 2 = sub-sub (1./2.), 3 = sub-sub-sub
           (a./b.).
    kind:  'front' | 'bab' | 'sub' | 'pustaka' | 'lampiran'
    'text' = teks apa adanya dari dokumen (untuk mencari halaman di
    PDF); 'display' = teks yang tampil di daftar isi (dengan huruf /
    angka yang ditetapkan mesin + normalisasi tampilan).
    """
    doc = Document(docx_path)
    paras = doc.paragraphs
    texts = [norm(p.text) for p in paras]
    known_terms = _known_terms(doc)
    entries = []
    consumed = set()
    seen_bab = False
    in_lampiran = False
    toc_anchor_idx = None
    # Penghitung untuk keputusan level pada awalan angka eksplisit
    # ("2. ..." -> level 2 bila melanjutkan urutan sub-sub, level 3
    # bila melanjutkan urutan di bawah sub-sub yang sedang berjalan).
    n_l2 = 0  # entri level 2 di bagian Heading 3 yang sedang berjalan
    n_l3 = 0  # entri level 3 di bawah entri level 2 yang sedang berjalan
    # Tumpukan heading terakhir (untuk menentukan level butir List
    # Paragraph): tiap item {"level", "ilvl", "indent"}.
    stack = []
    # v0.6.6: dokumen "ber-Heading" — penulis memakai Heading 1/2
    # asli Word untuk judul bagian (bukan List Paragraph). Dihitung
    # dari H1/H2 isi (bukan judul BAB, pustaka/lampiran, atau halaman
    # depan); >= 3 berarti dokumen ber-Heading (laporan nifas = 14;
    # skripsi Gunak = 1, Gusti Ayu = 0 -> bukan).
    n_head = 0
    for _p in paras:
        if _p.style.name in ("Heading 1", "Heading 2"):
            _t = norm(_p.text)
            if not _t:
                continue
            _u = _t.upper()
            if RE_BAB.match(_t):
                continue
            if _u in ("DAFTAR PUSTAKA", "DAFTAR LAMPIRAN"):
                continue
            if _u.replace(" ", "") in FRONT_BY_NOSPACE:
                continue
            n_head += 1
    heading_styled = n_head >= 3
    # v0.6.6: judul bagian yang penulis letakkan SEBELUM judul BAB-nya
    # (mis. "Latar Belakang" tebal tepat sebelum "BAB I PENDAHULUAN"
    # pada laporan nifas — satu list Word yang sama dengan isi BAB I).
    # Ditampung dulu, lalu ditempel di bawah BAB pertama sebagai
    # sub-sub. Hanya untuk dokumen ber-Heading dan hanya butir tebal
    # pendek setelah jangkar DAFTAR ISI yang bukan baris daftar lama.
    pending_pre_bab = []

    def push(e):
        while stack and stack[-1]["level"] >= e["level"]:
            stack.pop()
        stack.append({"level": e["level"],
                      "ilvl": e.get("ilvl"), "indent": e.get("indent"),
                      "numId": e.get("numId"),
                      "lp_bold": e.get("lp_bold", False),
                      "is_heading": e.get("is_heading", True)})

    def add(text, level, kind, idx, bold=False, display=None, search=None,
            post_anchor=False, term=None, ilvl=None, indent=None,
            numId=None, lp_bold=False, is_heading=True):
        nonlocal n_l2, n_l3
        if kind == "bab":
            n_l2 = 0
            n_l3 = 0
        elif level == 1:
            n_l2 = 0
            n_l3 = 0
        elif level == 2:
            n_l2 += 1
            n_l3 = 0
        elif level == 3:
            n_l3 += 1
        e = {"text": text, "level": level, "kind": kind,
             "para_index": idx, "bold": bold,
             "display": display if display is not None else text,
             "search": search if search is not None else [("text", text)],
             "post_anchor": post_anchor, "term": term,
             "ilvl": ilvl, "indent": indent, "numId": numId,
             "lp_bold": lp_bold, "is_heading": is_heading}
        entries.append(e)
        push(e)
        return e

    def level_for_list(ilvl, ind, is_term, top, numId=None):
        """Level (2 atau 3) untuk kandidat List Paragraph.
        top = heading terakhir di tumpukan (atau None)."""
        if top is None or top["level"] <= 1:
            return 2
        # Butir istilah+definisi di bawah sub-sub selalu anaknya.
        if is_term and top["level"] == 2:
            return 3
        # Indentasi render (posisi di halaman) paling tepercaya bila
        # berbeda jelas (>0,3 cm): butir yang menjorok lebih dalam
        # adalah anak, yang menjorok lebih dangkal adalah saudara
        # yang lebih tinggi.
        if ind is not None and top.get("indent") is not None:
            if ind > top["indent"] + 0.3:
                return 3
            if ind < top["indent"] - 0.3:
                return 2
        # Indentasi sama (atau tidak tersedia di salah satu sisi):
        # ilvl Word memutuskan — TAPI hanya bila nomor list-nya
        # sama (numId). Satu dokumen memakai banyak list Word yang
        # ilvl-nya tidak sebanding (di BAB III butir level 3
        # ber-ilvl 0, di Landasan Teori ber-ilvl 2; "Uji Asumsi
        # Klasik" BAB IV ber-ilvl 0 padahal setingkat dengan butir
        # ilvl 1). Dalam satu list yang sama, ilvl adalah sinyal
        # paling tepercaya: teori sejajar ber-ilvl 1 dan dimensinya
        # ber-ilvl 2 meski indentasinya nyaris sama (2.15 vs
        # 2.16 cm). Pengecualian: bila indentasi tidak tersedia di
        # salah satu sisi, ilvl satu-satunya sinyal yang ada, jadi
        # dipakai apa pun numId-nya (perilaku v0.4.0).
        if ilvl is not None and top.get("ilvl") is not None:
            same_list = (numId is not None
                         and top.get("numId") == numId)
            no_indent = (ind is None or top.get("indent") is None)
            if same_list or no_indent:
                if ilvl < top["ilvl"]:
                    return max(2, top["level"] - 1)
                if ilvl > top["ilvl"]:
                    return min(3, top["level"] + 1)
        return top["level"]

    def is_body(text):
        return len(text) > 120

    for i, p in enumerate(paras):
        if i in consumed:
            continue
        text = texts[i]
        if not text:
            continue
        if _is_old_toc_line(p):
            # v0.6.1: baris daftar isi LAMA ("BAB I\t1",
            # "1. Rumusan Masalah\t7") bukan heading asli — abaikan
            # supaya tidak jadi entri duplikat dengan nomor salah.
            continue
        upper = text.upper()
        style = p.style.name
        # --- halaman depan (sebelum BAB I) ---
        if not seen_bab:
            key = upper.replace(" ", "")
            if key in FRONT_BY_NOSPACE:
                canon = FRONT_BY_NOSPACE[key]
                if canon == "DAFTAR ISI":
                    toc_anchor_idx = i
                add(canon, 0, "front", i, bold=True, display=canon,
                    post_anchor=(toc_anchor_idx is not None
                                 and i > toc_anchor_idx))
                continue
        # --- BAB (gaya apa pun; judul boleh di baris berikut) ---
        m_bab = RE_BAB.match(text)
        if m_bab:
            seen_bab = True
            roman = m_bab.group(1).upper()
            disp = text.upper()
            search = [("text", text)]
            if RE_BAB_BARE.match(text):
                j = i + 1
                while j < len(paras) and not texts[j]:
                    j += 1
                if j < len(paras):
                    t2 = texts[j]
                    # v0.6.6: baris judul BAB boleh bergaya Normal ATAU
                    # Heading 1/2 ("BAB I" + "PENDAHULUAN" yang keduanya
                    # Heading 1 pada proposal Gunak).
                    if (paras[j].style.name in ("Normal", "Heading 1",
                                                "Heading 2")
                            and len(t2) <= 120
                            and not t2.endswith(".")
                            and not RE_CAPTION.match(t2)
                            and not RE_LETTER.match(t2)
                            and not RE_NUMBER.match(t2)
                            and not RE_LOWER.match(t2)
                            and t2.upper().replace(" ", "")
                            not in FRONT_BY_NOSPACE):
                        disp = f"{text.upper()} {t2}"
                        search = [("text", disp), ("bab", roman)]
                        consumed.add(j)
            add(disp, 0, "bab", i, bold=True, display=disp.upper(),
                search=search)
            for (t, idx, il, ind, nid) in pending_pre_bab:
                pe = add(t, 2, "sub", idx, bold=False,
                         ilvl=il, indent=ind, numId=nid, is_heading=False)
                # v0.6.6: mulai di halaman baru bersama BAB-nya
                # (break_before, diterapkan di build_output_docx)
                # supaya tidak tertinggal di akhir halaman daftar isi.
                pe["break_before"] = True
            pending_pre_bab.clear()
            continue
        if not seen_bab and heading_styled and toc_anchor_idx is not None:
            # v0.6.6: judul bagian sebelum BAB pertama (lihat di atas).
            if (style == "List Paragraph" and _is_bold(p)
                    and len(text) <= 70
                    and not _is_old_toc_line(p)):
                pending_pre_bab.append(
                    (text, i, _ilvl(p), _indent_cm(p), _numId(p)))
                continue
        if not seen_bab:
            continue  # abaikan isi halaman depan selain judul yang dikenal
        # --- penutup: pustaka & lampiran ---
        if upper == "DAFTAR PUSTAKA":
            add(upper, 0, "pustaka", i, bold=False)
            continue
        if upper == "DAFTAR LAMPIRAN":
            add(upper, 0, "pustaka", i, bold=False)
            in_lampiran = True
            continue
        if in_lampiran:
            if RE_LAMPIRAN.match(text):
                add(text, 1, "lampiran", i, bold=False)
            continue
        if RE_CAPTION.match(text):
            continue  # caption gambar/tabel ditangani extract_captions
        # --- sub-bab: Heading 3 (tanpa huruf -> diberi A./B./C.) ---
        if style == "Heading 3":
            add(text, 1, "sub", i, bold=False,
                ilvl=_ilvl(p), indent=_indent_cm(p), numId=_numId(p))
            continue
        # --- v0.6.6: Heading 1/2 asli (dokumen ber-Heading) ---
        # Judul yang berupa kalimat utuh (diakhiri titik) atau sangat
        # panjang bukan judul bagian — abaikan.
        if style == "Heading 1":
            if len(text) > 100 or text.endswith("."):
                continue
            add(text, 1, "sub", i, bold=False,
                ilvl=_ilvl(p), indent=_indent_cm(p), numId=_numId(p))
            continue
        if style == "Heading 2":
            if len(text) > 100 or text.endswith("."):
                continue
            add(text, 2, "sub", i, bold=False,
                ilvl=_ilvl(p), indent=_indent_cm(p), numId=_numId(p))
            continue
        # --- awalan eksplisit (A. / 1. / a.) ---
        if RE_LETTER.match(text):
            add(text, 1, "sub", i, bold=False,
                ilvl=_ilvl(p), indent=_indent_cm(p), numId=_numId(p))
            continue
        m_num = RE_NUMBER.match(text)
        if m_num:
            n = int(m_num.group(1))
            if n == n_l2 + 1:
                lvl = 2
            elif n == n_l3 + 1:
                lvl = 3
            else:
                lvl = 2
            e = add(text, lvl, "sub", i, bold=False,
                    ilvl=_ilvl(p), indent=_indent_cm(p), numId=_numId(p))
            if lvl == 2 and n > n_l2:
                n_l2 = n  # sinkronkan bila penomoran dokumen melompat
            continue
        if RE_LOWER.match(text):
            add(text, 3, "sub", i, bold=False,
                ilvl=_ilvl(p), indent=_indent_cm(p), numId=_numId(p))
            continue
        # --- sub-sub: List Paragraph ---
        if style == "List Paragraph":
            ilvl = _ilvl(p)
            ind = _indent_cm(p)
            nid = _numId(p)
            bold = _is_bold(p)
            term = _extract_term(text, known_terms) if len(text) > 70 else None
            is_term = term is not None
            short = len(text) <= 70
            top = stack[-1] if stack else None
            top_lvl = top["level"] if top else 0
            # Kandidat bila: pendek (frasa topik), atau istilah+definisi,
            # atau tebal (<=120 karakter) yang langsung di bawah
            # Heading 3/BAB, atau sejajar dengan butir tebal List
            # Paragraph level 2 sebelumnya (sub-bab yang berurutan,
            # mis. butir Pembahasan kedua). Kalimat isi yang panjang
            # dan tidak tebal (tinjauan pustaka, kriteria sampel,
            # penjelasan rumus) ditolak; begitu pula butir tebal yang
            # sangat panjang (awal paragraf isi, bukan judul).
            bold_heading = (bold and len(text) <= 120
                            and (top_lvl <= 1
                                 or (top_lvl == 2
                                     and top.get("lp_bold"))))
            if not (short or is_term or bold_heading):
                continue
            lvl = level_for_list(ilvl, ind, is_term, top, numId=nid)
            if lvl is None or lvl > 3:
                continue
            if heading_styled:
                # v0.6.6 gerbang ketat: induk = entri tumpukan
                # terdekat yang levelnya lebih dangkal dari kandidat.
                parent = None
                for s in reversed(stack):
                    if s["level"] < lvl:
                        parent = s
                        break
                if parent is None:
                    continue
                if parent["level"] == 0:
                    # Langsung di bawah BAB: hanya butir tebal
                    # (judul bagian) yang diterima — dinilai dari
                    # butirnya sendiri, bukan dari tumpukan (tumpukan
                    # bisa berisi sub-bab Heading 2 sehingga
                    # bold_heading salah False).
                    if not (bold and len(text) <= 120):
                        continue
                else:
                    # Di bawah sub-bab: induk harus heading asli,
                    # kandidat harus level 2, dan nomor list harus
                    # sama dengan induk (bila keduanya ada).
                    if not parent.get("is_heading"):
                        continue
                    if lvl != 2:
                        continue
                    if (parent.get("numId") is not None
                            and nid is not None
                            and parent.get("numId") != nid):
                        continue
            add(text, lvl, "sub", i, bold=False, term=term,
                ilvl=ilvl, indent=ind, numId=nid, lp_bold=bold,
                is_heading=False)
            continue
        # --- Heading grup NON-BOLD: paragraf pendek yang langsung
        # diikuti sub a./b./c. (mis. "Teori Pelestarian"). ---
        nxt = next((t for t in texts[i + 1:] if t), "")
        if (len(text) <= 60 and not text.endswith(".")
                and not is_body(text) and RE_LOWER.match(nxt)):
            add(text, 2, "sub", i, bold=False,
                ilvl=_ilvl(p), indent=_indent_cm(p), numId=_numId(p))
    # --- "Kesimpulan" yatim sebelum BAB (v0.5.0): di dokumen ini
    # sub-bab "Kesimpulan" ditulis tepat sebelum judul "BAB V
    # PENUTUP", padahal isinya kesimpulan BAB V — tiga butir temuan
    # ("Harga berpengaruh signifikan ...", dst.) ada SETELAH judul
    # BAB V, dan pengantar "Saran" merujuk "kesimpulan yang
    # didapatkan pada penelitian ini". Bila entri "Kesimpulan"
    # tepat sebelum sebuah BAB yang di bawahnya belum ada sub-bab
    # "Kesimpulan", pindahkan ke bawah BAB itu sebagai sub-bab
    # pertama (level 1); "Saran" yang sudah ada otomatis menjadi B.
    # assign_display menomori ulang dari level, jadi cukup pindah
    # dan ubah level.
    for idx, e in enumerate(entries):
        if (e["kind"] == "sub"
                and e["text"].strip().upper() == "KESIMPULAN"
                and idx + 1 < len(entries)
                and entries[idx + 1]["kind"] == "bab"):
            has_kes = False
            j = idx + 2
            while j < len(entries) and entries[j]["kind"] != "bab":
                if (entries[j]["level"] == 1
                        and entries[j]["text"].strip().upper()
                        == "KESIMPULAN"):
                    has_kes = True
                    break
                j += 1
            if not has_kes:
                entries.pop(idx)
                entries.insert(idx + 1, e)
                e["level"] = 1
            break
    return entries


def assign_display(entries):
    """Isi e['display']: huruf/angka penomoran sesuai format contoh
    dosen. Entri yang sudah berawalan dinomori ulang berurutan (hasil
    sama bila dokumen sudah rapi); grup tanpa awalan (mis. 'Teori
    Pelestarian') mendapat nomor; butir istilah+definisi hanya
    menampilkan istilahnya; lampiran tampil apa adanya."""
    l1 = l2 = l3 = 0
    for e in entries:
        kind = e["kind"]
        lvl = e["level"]
        if kind in ("front", "bab", "pustaka", "lampiran"):
            if kind == "bab":
                l1 = l2 = l3 = 0
            continue  # display sudah benar saat ekstraksi
        base = e.get("term") or _strip_prefix(e["text"])
        if lvl == 1:
            l1 += 1
            l2 = 0
            l3 = 0
            e["display"] = f"{chr(ord('A') + l1 - 1)}. {base}"
        elif lvl == 2:
            l2 += 1
            l3 = 0
            e["display"] = f"{l2}. {base}"
        elif lvl == 3:
            l3 += 1
            e["display"] = f"{chr(ord('a') + l3 - 1)}. {base}"
        else:
            e["display"] = e["text"]
    return entries


def extract_captions(docx_path):
    """Kembalikan (gambar, tabel): entri caption dari paragraf dokumen,
    dalam urutan kemunculan. Hanya pola standar yang terdeteksi
    ('Gambar 1.1 ...', 'Tabel 2 ...', 'Gambar 1 ...'); caption tanpa
    nomor (mis. 'Gambar: Peta') atau caption di dalam sel tabel
    tidak terdeteksi. kind 'caption' -> nomor halaman Arab relatif
    terhadap BAB I, sama seperti heading isi.

    Normalisasi tampilan (v0.4.0): pemisah spasi pada nomor
    ('Gambar 2 1 ...') ditampilkan sebagai titik ('Gambar 2.1 ...').
    Isi dokumen TIDAK diubah; 'text' tetap apa adanya (untuk mencari
    halaman di PDF), 'display' yang ternormalisasi untuk daftar."""
    doc = Document(docx_path)
    gambar, tabel = [], []
    for i, p in enumerate(doc.paragraphs):
        text = norm(p.text)
        if _is_old_toc_line(p):
            # v0.6.1: baris daftar gambar/tabel LAMA ("Gambar 1.1 ...\t5")
            # bukan caption asli — abaikan supaya tidak jadi entri duplikat.
            continue
        m = RE_CAPTION.match(text)
        if not m:
            continue
        kind = m.group(1)
        num = ".".join(re.split(r"[.\s]+", m.group(2)))
        rest = text[m.end(2):].strip()
        display = f"{kind} {num} {rest}"
        entry = {"text": text, "level": 0, "kind": "caption",
                 "list": kind.lower(), "para_index": i, "bold": False,
                 "display": display, "post_anchor": False,
                 "search": [("text", text)]}
        (gambar if kind.lower() == "gambar" else tabel).append(entry)
    return gambar, tabel


# ---------------------------------------------------------------- render PDF
def render_pdf(docx_path, out_dir):
    """Render .docx ke PDF via LibreOffice headless. Kembalikan path PDF."""
    os.makedirs(out_dir, exist_ok=True)
    _trim_memory()  # v0.6.2: kecilkan RSS Python sebelum soffice jalan
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf",
         "--outdir", out_dir, docx_path],
        check=True, capture_output=True, timeout=600,
    )
    base = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
    pdf_path = os.path.join(out_dir, base)
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"LibreOffice gagal membuat {pdf_path}")
    return pdf_path


def pdf_pages_text(pdf_path):
    doc = pymupdf.open(pdf_path)
    return [norm(page.get_text()) for page in doc]


def pdf_pages_lines(pdf_path):
    """Teks PDF per halaman, per baris (tanpa baris kosong). Dipakai
    untuk mencari heading yang berdiri sendiri (bukan baris daftar
    isi dan bukan kalimat isi)."""
    doc = pymupdf.open(pdf_path)
    out = []
    for page in doc:
        lines = [norm(ln) for ln in page.get_text().split("\n")]
        out.append([ln for ln in lines if ln])
    return out


# ---------------------------------------------------------------- cari halaman
def _page_contains(page_text, key):
    kl = key.lower()
    if kl in page_text.lower():
        return True
    return _nospace(key) in _nospace(page_text)


def _key_in_page(page_text, key):
    """Cek satu kunci ('text' = substring, 'bab' = regex BAB dengan
    batas kata agar 'BAB II' tidak cocok dengan 'BAB III').
    Judul yang sangat panjang (kalimat definisi) bisa terbelah
    halaman di PDF — bila teks penuh tidak ketemu, 60 karakter
    pertamanya cukup sebagai kunci (v0.6.1)."""
    kind, pat = key
    if kind == "bab":
        rx = re.compile(rf"BAB\s+{pat}\b(?![IVXLC])", re.IGNORECASE)
        return rx.search(page_text) is not None
    if _page_contains(page_text, pat):
        return True
    if len(pat) > 60:
        return _page_contains(page_text, pat[:60])
    return False


def _find_key(pages_text, keys, start=0):
    """Cari kunci pertama yang cocok, dari halaman start (0-based).
    Kembalikan nomor halaman (1-based) atau None."""
    for key in keys:
        for pi in range(start, len(pages_text)):
            if _key_in_page(pages_text[pi], key):
                return pi + 1
    return None


def find_pages(pages_text, entries):
    """Cari halaman PDF (1-based) tiap entri, berurutan dari depan.

    Pencarian berurutan dengan pointer agar judul yang berulang
    (mis. 'Teori Pelestarian') ketemu di urutan yang benar.
    """
    pointer = 0  # indeks halaman (0-based) tempat pencarian dimulai
    for e in entries:
        keys = e.get("search") or [("text", e["text"])]
        found = _find_key(pages_text, keys, pointer)
        e["pdf_page"] = found
        if found is not None:
            pointer = found - 1  # entri berikut boleh di halaman sama
    return entries


def compute_labels(entries, bab1=None):
    """Isi e['page']: Romawi untuk halaman depan, Arab mulai dari BAB I.
    bab1 = halaman PDF (1-based) tempat BAB I mulai; bila tidak
    diberikan, dihitung dari entri kind 'bab' di daftar ini. WAJIB
    diteruskan dari daftar isi saat melabeli daftar gambar/tabel,
    karena daftar itu tidak punya entri BAB sendiri."""
    if bab1 is None:
        bab_pages = [e["pdf_page"] for e in entries
                     if e["kind"] == "bab" and e.get("pdf_page")]
        bab1 = min(bab_pages) if bab_pages else None
    for e in entries:
        p = e.get("pdf_page")
        if p is None:
            e["page"] = "?"
        elif e["kind"] == "front":
            e["page"] = to_roman(p)          # cover=i, dst.
        elif bab1 is not None:
            e["page"] = str(p - bab1 + 1)    # BAB I = 1
        else:
            e["page"] = str(p)
    return entries


# ---------------------------------------------------------------- format teks
def format_toc_text(entries, width=78):
    lines = []
    for e in entries:
        label = e.get("display") or e["text"]
        page = e["page"]
        indent = "  " * e["level"]
        body = indent + label
        dots = max(3, width - len(body) - len(page))
        lines.append(f"{body}{'.' * dots}{page}")
    return "\n".join(lines)


# ---------------------------------------------------------------- sisip ke docx
def _insert_after(ref_para, new_para):
    ref_para._p.addnext(new_para._p)


def _page_break_before(p):
    """Pastikan paragraf mulai di halaman baru (v0.6.5): daftar
    gambar dan daftar tabel masing-masing di lembar sendiri, tidak
    digabung dengan daftar isi atau satu sama lain."""
    p.paragraph_format.page_break_before = True


def build_output_docx(src_path, entries, out_path, gambar=None, tabel=None):
    """Sisipkan daftar isi setelah heading 'DAFTAR ISI' (atau sebelum
    BAB I bila heading itu tidak ada), lalu DAFTAR GAMBAR dan
    DAFTAR TABEL setelahnya. Bila dokumen sudah punya heading
    'DAFTAR GAMBAR' / 'DAFTAR TABEL' di halaman depan, entri disisipkan
    setelah heading itu (mengikuti urutan heading dokumen); bila tidak
    ada, judul bagian dibuat baru setelah daftar isi. Bagian yang tidak
    punya entri tidak dibuat.
    Kembalikan jumlah halaman TOC yang disisipkan tidak diketahui di
    sini — diukur oleh pemanggil."""
    doc = Document(src_path)
    # v0.6.6: judul bagian yang ditempel sebelum BAB-nya (pending)
    # jangan sampai tertinggal di akhir halaman daftar isi — mulai
    # di halaman baru bersama BAB-nya (page_break_before). keepNext
    # tidak mempan di LibreOffice karena ada section break di judul
    # BAB. Hanya untuk entri bertanda break_before; dokumen lain
    # tidak tersentuh.
    for e in entries:
        if not e.get("break_before"):
            continue
        _pi = e.get("para_index")
        if _pi is None or not (0 <= _pi < len(doc.paragraphs)):
            continue
        if norm(doc.paragraphs[_pi].text) != e["text"]:
            continue
        doc.paragraphs[_pi].paragraph_format.page_break_before = True
    sec = doc.sections[0]
    usable_twips = int(
        (sec.page_width - sec.left_margin - sec.right_margin) / 635
    )  # 1 twip = 635 EMU

    # v0.6.3: baca dulu format daftar isi lama (indent per level,
    # tab stop, spasi, font) supaya daftar baru sama persis dengan
    # file aslinya; bila tidak ada daftar lama, pakai nilai default.
    old_fmt = _capture_old_toc_format(doc)

    # v0.6.1: buang dulu baris-baris daftar lama, baru cari anchor —
    # supaya paras tidak memuat paragraf yang sudah dilepas.
    _remove_old_toc_runs(doc)
    paras = doc.paragraphs

    def find_heading(title):
        return next((p for p in paras
                     if norm(p.text).upper() == title), None)

    toc_para = find_heading("DAFTAR ISI")
    if toc_para is not None:
        anchor = toc_para
        insert_before_bab = False
    else:
        bab_idx = next((i for i, p in enumerate(paras)
                        if RE_BAB.match(norm(p.text))), None)
        anchor = paras[bab_idx] if bab_idx is not None else paras[-1]
        insert_before_bab = True

    def _indent_for_level(level):
        if old_fmt is None:
            return Cm(0.6 * level)
        inds = old_fmt["indents"]
        if level < len(inds):
            return Cm(inds[level])
        step = (inds[-1] - inds[-2]) if len(inds) >= 2 else 0.6
        return Cm(inds[-1] + step * (level - len(inds) + 1))

    def add_entry(text, page, level, bold):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # v0.6.4: spasi antar-entri minimal 6 pt supaya daftar isi
        # tidak dempet — aslinya 2 pt (1 spasi) terlalu rapat di mata.
        # Bila format asli sudah lebih lega, pakai nilai aslinya.
        _after = (old_fmt["space_after_pt"]
                  if old_fmt and old_fmt["space_after_pt"] is not None
                  else 2)
        p.paragraph_format.space_after = Pt(max(_after, 6))
        p.paragraph_format.space_before = Pt(
            old_fmt["space_before_pt"]
            if old_fmt and old_fmt["space_before_pt"] is not None else 0)
        p.paragraph_format.left_indent = _indent_for_level(level)
        p.paragraph_format.first_line_indent = Cm(
            old_fmt["first_cm"] if old_fmt else 0)
        tab_cm = old_fmt["tab_cm"] if old_fmt and old_fmt["tab_cm"] else None
        p.paragraph_format.tab_stops.add_tab_stop(
            Cm(tab_cm) if tab_cm else Twips(usable_twips),
            WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        r = p.add_run(f"{text}\t{page}")
        r.font.name = (old_fmt["font_name"]
                       if old_fmt and old_fmt["font_name"]
                       else "Times New Roman")
        r.font.size = Pt(old_fmt["font_size_pt"]
                         if old_fmt and old_fmt["font_size_pt"] else 12)
        r.bold = bold
        return p

    def add_section_title(title):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.space_before = Pt(10)
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
        r.bold = True
        return p

    def insert_after_seq(anchor_para, new_paras):
        """Sisipkan new_paras berurutan setelah anchor_para.
        Kembalikan paragraf terakhir yang disisipkan."""
        cursor = anchor_para
        for p in new_paras:
            _insert_after(cursor, p)
            cursor = p
        return cursor

    def disp(e):
        return e.get("display") or e["text"]

    # --- 1. Daftar isi ---
    toc_made = []
    if insert_before_bab:
        toc_made.append(add_section_title("DAFTAR ISI"))
    for e in entries:
        toc_made.append(add_entry(disp(e), e["page"], e["level"],
                                  e["bold"]))
    last_toc = insert_after_seq(anchor, toc_made)
    if insert_before_bab:
        br = doc.add_paragraph()
        br.add_run().add_break(WD_BREAK.PAGE)
        _insert_after(last_toc, br)
        last_toc = br

    # --- 2. Daftar gambar (setelah daftar isi, atau setelah heading
    # 'DAFTAR GAMBAR' bila dokumen sudah punya) ---
    last_anchor = last_toc
    if gambar:
        g_head = find_heading("DAFTAR GAMBAR")
        g_made = []
        if g_head is None:
            g_made.append(add_section_title("DAFTAR GAMBAR"))
            g_anchor = last_anchor
        else:
            g_anchor = g_head
        # v0.6.5: daftar gambar mulai di lembar baru.
        _page_break_before(g_made[0] if g_head is None else g_head)
        for e in gambar:
            g_made.append(add_entry(disp(e), e["page"], 0, False))
        last_anchor = insert_after_seq(g_anchor, g_made)

    # --- 3. Daftar tabel (setelah daftar gambar, atau setelah heading
    # 'DAFTAR TABEL' bila dokumen sudah punya) ---
    if tabel:
        t_head = find_heading("DAFTAR TABEL")
        t_made = []
        if t_head is None:
            t_made.append(add_section_title("DAFTAR TABEL"))
            t_anchor = last_anchor
        else:
            t_anchor = t_head
        # v0.6.5: daftar tabel mulai di lembar baru.
        _page_break_before(t_made[0] if t_head is None else t_head)
        for e in tabel:
            t_made.append(add_entry(disp(e), e["page"], 0, False))
        insert_after_seq(t_anchor, t_made)

    doc.save(out_path)
    return out_path


# ---------------------------------------------------------------- nomor halaman footer (v0.6.0)
# Nomor halaman dicetak di footer dokumen hasil dengan field PAGE
# (dihitung Word/LibreOffice saat dibuka, jadi selalu akurat):
# halaman depan Romawi kecil tengah bawah (cover = i, tidak
# ditampilkan di cover), mulai BAB I angka Arab kanan bawah dari 1.
_LAYOUT_TAGS = ("w:pgSz", "w:pgMar", "w:paperSrc", "w:pgBorders",
                "w:lnNumType", "w:cols", "w:formProt", "w:vAlign",
                "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid")


_SECTPR_ORDER = ("w:headerReference", "w:footerReference", "w:footnotePr",
                  "w:endnotePr", "w:type", "w:pgSz", "w:pgMar",
                  "w:paperSrc", "w:pgBorders", "w:lnNumType",
                  "w:pgNumType", "w:cols", "w:formProt", "w:vAlign",
                  "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi",
                  "w:rtlGutter", "w:docGrid", "w:printerSettings")


def _normalize_sectPr_order(sectPr):
    """Urutkan anak sectPr sesuai urutan schema OOXML. Dokumen sumber
    kadang menyimpan referensi header/footer terbalik (footer dulu),
    dan LibreOffice mengabaikan pgNumType pada sectPr yang urutannya
    salah — akibatnya nomor Romawi tampil sebagai Arab."""
    order = {qn(t): i for i, t in enumerate(_SECTPR_ORDER)}
    kids = sorted(sectPr, key=lambda c: order.get(c.tag, 99))
    for child in list(sectPr):
        sectPr.remove(child)
    for child in kids:
        sectPr.append(child)


def _sectPr_pgnumtype(sectPr, fmt, start=None):
    """Setel w:pgNumType pada sectPr (hapus yang lama bila ada)."""
    for old in sectPr.findall(qn("w:pgNumType")):
        sectPr.remove(old)
    el = OxmlElement("w:pgNumType")
    el.set(qn("w:fmt"), fmt)
    if start is not None:
        el.set(qn("w:start"), str(start))
    after = ("w:cols", "w:formProt", "w:vAlign", "w:noEndnote", "w:titlePg",
             "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
             "w:printerSettings")
    for child in sectPr:
        if child.tag in [qn(t) for t in after]:
            child.addprevious(el)
            break
    else:
        sectPr.append(el)
    _normalize_sectPr_order(sectPr)


def _new_sectPr_from(body_sectPr, break_type="continuous"):
    """sectPr baru (section break) dengan layout disalin dari sectPr
    akhir dokumen."""
    sectPr = OxmlElement("w:sectPr")
    type_el = OxmlElement("w:type")
    type_el.set(qn("w:val"), break_type)
    sectPr.append(type_el)
    if body_sectPr is not None:
        for tag in _LAYOUT_TAGS:
            src = body_sectPr.find(qn(tag))
            if src is not None:
                sectPr.append(copy.deepcopy(src))
    return sectPr


def _bab_starts_new_page(doc, bab_para, bab_j):
    """Apakah paragraf BAB I mulai di halaman baru (sinyal .docx:
    pageBreakBefore, page break di paragraf sebelumnya, atau
    sectPr nextPage/oddPage/evenPage di paragraf ini)."""
    pPr = bab_para._p.find(qn("w:pPr"))
    if pPr is not None:
        if pPr.find(qn("w:pageBreakBefore")) is not None:
            return True
        sp = pPr.find(qn("w:sectPr"))
        if sp is not None:
            t = sp.find(qn("w:type"))
            if t is not None and t.get(qn("w:val")) in (
                    "nextPage", "oddPage", "evenPage"):
                return True
    if bab_j and bab_j > 0:
        for r in doc.paragraphs[bab_j - 1]._p.findall(qn("w:r")):
            br = r.find(qn("w:br"))
            if br is not None and br.get(qn("w:type")) == "page":
                return True
    return False


def _strip_break_runs(para):
    """Hapus run break halaman/kolom (w:br type page/column) dari
    paragraf — dipakai saat section break nextPage sudah menjamin
    ganti halaman (dua break = satu halaman kosong). Kembalikan
    True bila ada break yang dihapus."""
    removed = False
    for r in list(para._p.findall(qn("w:r"))):
        br = r.find(qn("w:br"))
        if br is not None and br.get(qn("w:type")) in ("page", "column"):
            r.remove(br)
            removed = True
    return removed


def _move_break_before(doc, bab_para, bab_j, body_sectPr):
    """Pindahkan batas section ke paragraf SEBELUM BAB dengan tipe
    nextPage, supaya halaman BAB sendiri sudah masuk section Arab.
    sectPr lama di BAB (bila ada) dipakai ulang agar layout halaman
    depan tidak berubah.
    CATATAN v0.6.10: fungsi ini TIDAK dipakai lagi — nextPage
    sebelum BAB terbukti merusak pagination di LibreOffice untuk
    dokumen ini (lihat docstring add_footer_page_numbers).
    Disimpan untuk pemakaian mendatang bila diperlukan.
    Break ditaruh di paragraf BERISI terakhir sebelum BAB (paragraf
    kosong di antara dilewati — paragraf kosong itu duduk di atas
    halaman BAB dan break di sana justru menggeser BAB ke halaman
    berikut). Run break halaman/kolom di paragraf target dan di
    paragraf kosong yang dilewati, serta pageBreakBefore di BAB,
    dibuang karena sudah digantikan section break nextPage (dua
    break = satu halaman kosong)."""
    if bab_j <= 0:
        return False
    target_j = bab_j - 1
    while target_j > 0 and not doc.paragraphs[target_j].text.strip():
        target_j -= 1
    target = doc.paragraphs[target_j]
    pPr = bab_para._p.find(qn("w:pPr"))
    old = pPr.find(qn("w:sectPr")) if pPr is not None else None
    if old is not None:
        pPr.remove(old)
        sectPr = copy.deepcopy(old)
    else:
        sectPr = _new_sectPr_from(body_sectPr)
    t = sectPr.find(qn("w:type"))
    if t is None:
        t = OxmlElement("w:type")
        sectPr.insert(0, t)
    t.set(qn("w:val"), "nextPage")
    target_pPr = target._p.get_or_add_pPr()
    existing = target_pPr.find(qn("w:sectPr"))
    if existing is not None:
        target_pPr.remove(existing)
    target_pPr.append(sectPr)
    for j in range(target_j, bab_j):
        _strip_break_runs(doc.paragraphs[j])
    if pPr is not None:
        pb = pPr.find(qn("w:pageBreakBefore"))
        if pb is not None:
            pPr.remove(pb)
    return True


def _para_has_page_field(p):
    if p._p.find(qn("w:fldSimple")) is not None:
        return True
    for r in p._p.findall(qn("w:r")):
        if (r.find(qn("w:fldChar")) is not None
                or r.find(qn("w:instrText")) is not None):
            return True
    return False


def _strip_leading_number(line):
    """Buang awalan penomoran ('1.1 ', 'A. ', 'IV. ') dari baris
    judul — nomor list di PDF tidak selalu sama dengan teks entri
    (nomor list dirender LibreOffice, display entri beda lagi)."""
    return re.sub(r"^(?:\d+(?:\.\d+)*\.?|[A-Za-z]\.|[IVXLCDM]+\.)\s+",
                  "", line)


def _bab_start_page(lines, entry):
    """Halaman (1-based) judul BAB di PDF hasil: kemunculan TERAKHIR
    baris yang persis sama dengan teks entri; bila tidak ketemu
    (judul BAB telanjang / terpotong baris di PDF), cari baris
    'BAB <romawi>' telanjang (nomor romawi diambil dari teks entri,
    bukan dari field search — field itu hanya berisi ('text', ...))."""
    p = _heading_page_out(lines, entry)
    if p is not None:
        return p
    m = RE_BAB.match(entry["text"])
    if m:
        return _heading_page_out(lines, {"text": "BAB " + m.group(1)})
    return None


def _bab_starts_new_page_pdf(lines, page, pending_texts=()):
    """True bila judul BAB adalah baris isi pertama halamannya.
    Baris judul bagian pending (v0.6.6) yang ditempel sebelum BAB
    ditoleransi — baris itu bagian dari BAB yang sama, bukan isi bab
    sebelumnya. pending_texts: teks-teks judul pending itu (tanpa
    nomor — nomor list di PDF dibuang sebelum dicocokkan)."""
    if page is None or page < 1 or page > len(lines):
        return False
    first_lines = [ln for ln in lines[page - 1] if ln.strip()]
    pend = {_nospace(t) for t in pending_texts}
    for k, ln in enumerate(first_lines[:3]):
        first = norm(ln)
        if (RE_BAB.match(first) and "..." not in first
                and not re.search(r"\d\s*$", first)):
            preceding = first_lines[:k]
            return all(_nospace(_strip_leading_number(x)) in pend
                       for x in preceding)
    return False


def _find_bab_paras(doc, entries):
    """Cari paragraf judul tiap BAB di dokumen, sesuai urutan entri.
    Cocok dengan teks entri (case-insensitive) ATAU judul BAB
    telanjang ('BAB I') bila judulnya di baris berikut — keduanya
    dicocokkan per nomor romawi agar tidak salah bab.
    Kembalikan list (entry, paragraph, index) yang ketemu saja."""
    paras = doc.paragraphs
    out = []
    start = 0
    for e in [x for x in entries if x.get("kind") == "bab"]:
        target = _nospace(e["text"])
        m = RE_BAB.match(e["text"])
        roman = m.group(1).lower() if m else None
        for j in range(start, len(paras)):
            t = _nospace(paras[j].text)
            if not t:
                continue
            if t == target:
                out.append((e, paras[j], j))
                start = j + 1
                break
            if roman and t == "bab" + roman:
                out.append((e, paras[j], j))
                start = j + 1
                break
    return out


def _ensure_bab_first_page_footer(section):
    """Aktifkan different-first-page pada section dan isi first-page
    footer-nya dengan field PAGE rata tengah (nomor halaman awal BAB
    di bawah tengah). Section mendapat referensi 'first' SENDIRI
    (bukan tautan ke section sebelumnya): LibreOffice tidak
    mewariskan referensi 'first' melewati section tanpa titlePg,
    jadi bila dibiarkan tertaut, nomor halaman pertama BAB tetap
    memakai footer default (kanan)."""
    section.different_first_page_header_footer = True
    try:
        fp = section.first_page_footer
        if fp.part is not None and fp.part is section.footer.part:
            for ref in list(section._sectPr.findall(qn("w:footerReference"))):
                if ref.get(qn("w:type")) == "first":
                    section._sectPr.remove(ref)
            fp = section.first_page_footer
        if fp.is_linked_to_previous:
            fp.is_linked_to_previous = False
            fp = section.first_page_footer
        _ensure_footer_number_on(fp, WD_ALIGN_PARAGRAPH.CENTER)
    except Exception:
        pass


def _footer_has_page_field(footer):
    return any(_para_has_page_field(p) for p in footer.paragraphs)


def _strip_page_field(p):
    """Hapus field PAGE dari paragraf (biarkan teks lain)."""
    for child in list(p._p):
        if child.tag == qn("w:fldSimple"):
            p._p.remove(child)
    for r in list(p._p.findall(qn("w:r"))):
        if (r.find(qn("w:fldChar")) is not None
                or r.find(qn("w:instrText")) is not None):
            p._p.remove(r)


def _add_page_field_para(paragraph, align):
    """Isi paragraf dengan field PAGE (Times New Roman 12)."""
    paragraph.alignment = align
    p = paragraph._p

    def run_with(child):
        r = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        fonts = OxmlElement("w:rFonts")
        fonts.set(qn("w:ascii"), "Times New Roman")
        fonts.set(qn("w:hAnsi"), "Times New Roman")
        rPr.append(fonts)
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "24")
        rPr.append(sz)
        r.append(rPr)
        r.append(child)
        return r

    fld = OxmlElement("w:fldChar")
    fld.set(qn("w:fldCharType"), "begin")
    p.append(run_with(fld))
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    p.append(run_with(instr))
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    p.append(run_with(sep))
    t = OxmlElement("w:t")
    t.text = "1"  # nilai sementara; field diperbarui saat dibuka/dicetak
    p.append(run_with(t))
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    p.append(run_with(end))


def _ensure_footer_number_on(footer, align):
    """Pastikan satu footer (default/genap/halaman pertama) punya
    nomor halaman field PAGE dengan posisi align. Bila footer sudah
    punya field PAGE, dibiarkan (kecuali posisinya beda — maka
    tautan diputus dan dibuat ulang). Teks footer lain dipertahankan;
    teks hasil field PAGE lama TIDAK disalin (itu yang menimbulkan
    angka '1' liar)."""
    if _footer_has_page_field(footer):
        for p in footer.paragraphs:
            if _para_has_page_field(p):
                if p.alignment == align:
                    return
                if footer.is_linked_to_previous:
                    old_texts = [q.text for q in footer.paragraphs
                                 if q.text.strip()
                                 and not _para_has_page_field(q)]
                    footer.is_linked_to_previous = False
                    for t in old_texts:
                        footer.add_paragraph(t)
                    target = next((q for q in footer.paragraphs
                                   if not q.text.strip()), None)
                    if target is None:
                        target = footer.add_paragraph()
                    _add_page_field_para(target, align)
                else:
                    _strip_page_field(p)
                    _add_page_field_para(p, align)
                return
    old_texts = [p.text for p in footer.paragraphs
                 if p.text.strip() and not _para_has_page_field(p)]
    if footer.is_linked_to_previous:
        footer.is_linked_to_previous = False
        for t in old_texts:
            footer.add_paragraph(t)
    target = next((p for p in footer.paragraphs if not p.text.strip()),
                  None)
    if target is None:
        target = footer.add_paragraph()
    _add_page_field_para(target, align)


def _ensure_section_footers(section, align, is_first, even_odd):
    """Nomor di semua jenis footer section ini: default, genap (bila
    dokumen memakai odd/even), dan halaman pertama (bila section
    memakai different-first-page) — kecuali cover (is_first)."""
    _ensure_footer_number_on(section.footer, align)
    if even_odd:
        try:
            _ensure_footer_number_on(section.even_page_footer, align)
        except Exception:
            pass
    if not is_first and section.different_first_page_header_footer:
        try:
            _ensure_footer_number_on(section.first_page_footer, align)
        except Exception:
            pass


def add_footer_page_numbers(out_path, entries, bab_new_page=None,
                            bab_new_pages=None):
    """Tambahkan nomor halaman di footer dokumen hasil.
    bab_new_page: True bila BAB I mulai di halaman baru (diisi dari
    PDF hasil oleh process(); bila None, ditebak dari sinyal .docx).
    bab_new_pages: list bool sejajar entri BAB — True bila BAB itu
    mulai di halaman baru (diisi dari PDF hasil oleh process()).
    Halaman pertama tiap BAB memakai footer first-page (titlePg)
    dengan field PAGE rata tengah di bawah; halaman lain tetap
    Romawi tengah (halaman depan) atau Arab kanan (isi).
    PENTING (v0.6.10): TIDAK ada section break baru dan TIDAK ada
    perubahan tipe break. Section break nextPage sebelum BAB
    terbukti merusak pagination di LibreOffice untuk dokumen ini
    (memecah daftar isi, halaman kosong, judul terbelah dua kolom)
    karena sectPr sumber tanpa w:type dan run break kolom. Sebagai
    gantinya, titlePg hanya dipasang pada section YANG SUDAH ADA
    yang paragraf berisi pertamanya ADALAH judul BAB yang mulai di
    halaman baru — section itu halaman pertamanya memang halaman
    BAB, jadi nomornya jadi di tengah tanpa mengubah pagination
    sama sekali. BAB I (v0.6.11): sectPr continuous di judulnya
    dipindah ke paragraf sebelumnya supaya section Arab mulai
    tepat di judul BAB I (footer tengah "1", bukan "viii"), plus
    hapus w:cols di sectPr penutup section itu dan run break
    halaman di awal paragraf isi pertama sesudahnya (lihat kode).
    BAB yang bukan paragraf pertama section-nya (BAB yang mulai
    di tengah halaman) dibiarkan seperti semula: BAB tengah
    halaman tetap kanan. DAFTAR PUSTAKA / lampiran bukan BAB:
    tidak diberi titlePg.
    Kembalikan True bila batas BAB I ditemukan (Romawi -> Arab)."""
    doc = Document(out_path)
    # Cari paragraf judul tiap BAB (kenal judul telanjang "BAB I" yang
    # judulnya di baris berikut; baris daftar isi mengandung tab+nomor
    # sehingga tidak cocok persis — itu yang diinginkan).
    bab_paras = _find_bab_paras(doc, entries)
    if not bab_paras:
        for s in doc.sections:
            _ensure_section_footers(s, WD_ALIGN_PARAGRAPH.CENTER, False,
                                    False)
        doc.save(out_path)
        return False

    bab_para, bab_j = bab_paras[0][1], bab_paras[0][2]
    body = doc.element.body
    body_sectPr = body.find(qn("w:sectPr"))

    # BAB I: sectPr continuous di paragrafnya SENDIRI, tapi HANYA
    # bila section tempat BAB I berada TIDAK mengandung judul
    # halaman depan (KATA PENGANTAR, HALAMAN PERSETUJUAN, dsb.)
    # sebelum BAB I. sectPr itu menutup section yang berakhir di
    # BAB I; bila section itu memuat halaman depan (Gunak: section
    # besar dari akhir cover sampai jauh ke isi), sectPr continuous
    # mengubah tipe break-nya dari nextPage menjadi continuous
    # sehingga halaman depan tidak lagi mulai di halaman baru (isi
    # hal. 1-16 bergeser). Bila section BAB I hanya memuat daftar
    # isi + BAB I (nifas, Gusti), sectPr continuous justru
    # DIPERLUKAN: tanpa itu LibreOffice memecah/menggeser daftar
    # isi (judul sendiri di satu halaman, entri di halaman berikut).
    if bab_new_page is None:
        bab_new_page = _bab_starts_new_page(doc, bab_para, bab_j)
    _front_markers = ("KATA PENGANTAR", "HALAMAN PERSETUJUAN",
                      "LEMBAR PENGESAHAN", "LEMBAR PERSETUJUAN",
                      "HALAMAN JUDUL", "ABSTRAK", "KATA SAMBUTAN",
                      "PERNYATAAN", "HALAMAN PERNYATAAN")
    _sec_start = 0
    for _i in range(bab_j - 1, -1, -1):
        _pPr = doc.paragraphs[_i]._p.find(qn("w:pPr"))
        if _pPr is not None and _pPr.find(qn("w:sectPr")) is not None:
            _sec_start = _i + 1
            break
    _has_front = any(
        norm(doc.paragraphs[_i].text) in _front_markers
        for _i in range(_sec_start, bab_j))
    if not _has_front:
        pPr = bab_para._p.get_or_add_pPr()
        if pPr.find(qn("w:sectPr")) is None:
            pPr.append(_new_sectPr_from(body_sectPr))

    # --- v0.6.11: pindahkan sectPr dari paragraf BAB I ke paragraf
    # SEBELUMNYA, supaya section Arab mulai TEPAT di judul BAB I.
    # Sebelumnya batas section ada di judul BAB I sendiri: halaman
    # pertama BAB I masih masuk section depan (Romawi) sehingga
    # nomornya tampil "viii"/"xii" tengah, bukan "1". Dengan batas
    # dipindah satu paragraf lebih awal, judul BAB I jadi paragraf
    # pertama section Arab (decimal start=1) dan mendapat titlePg
    # + footer tengah "1" — tanpa mengubah pagination, karena tipe
    # break dipertahankan. Hanya sectPr continuous yang dipindah:
    # sectPr nextPage/oddPage/evenPage — atau tanpa w:type (default
    # = nextPage) — TIDAK dipindah, karena memindah break halaman
    # akan menggeser BAB I ke halaman baru. Bila paragraf
    # sebelumnya sudah punya sectPr sendiri, atau BAB I paragraf
    # pertama dokumen, atau section BAB I memuat judul halaman
    # depan (Gunak), tidak dipindah. Langkah ini HARUS sebelum
    # pemetaan section di bawah, supaya bab_idx/body_idx dan
    # sec_first_text sudah memakai batas yang baru.
    # PERINGATAN (ditemukan saat uji): paragraf sebelumnya kadang
    # membawa run break halaman/kolom (nifas: w:br type=column di
    # paragraf kosong sesudah daftar isi). sectPr + break run
    # dalam SATU paragraf membuat LibreOffice menghitung dua
    # break (halaman kosong + BAB I geser satu halaman). Karena
    # itu break run itu dibuang dan diganti pageBreakBefore di
    # judul BAB I: BAB I tetap mulai di halaman baru, tanpa break
    # ganda.
    # PENTING v0.6.11: pemindahan hanya aman bila BAB I SUDAH mulai
    # di halaman baru (bab_new_page True). Bila BAB I mulai di
    # tengah halaman (Gusti: judulnya di bawah halaman DAFTAR
    # GAMBAR), perubahan format Romawi->Arab di batas pindahan
    # MEMAKSA halaman baru — judul BAB I loncat ke halaman
    # berikut dan total halaman bertambah 1 (Gusti 142 -> 143).
    # BAB tengah halaman dibiarkan seperti v0.6.10 (footer tetap
    # Romawi tengah dari section depan); memaksanya jadi "1"
    # butuh halaman baru, dan itu mengubah pagination.
    _moved = False
    if not _has_front and bab_j > 0 and bab_new_page:
        _bpPr = bab_para._p.find(qn("w:pPr"))
        _bsp = (_bpPr.find(qn("w:sectPr"))
                if _bpPr is not None else None)
        if _bsp is not None:
            _bt = _bsp.find(qn("w:type"))
            _bval = _bt.get(qn("w:val")) if _bt is not None else None
            if _bval == "continuous":
                _prev = doc.paragraphs[bab_j - 1]
                _ppPr = _prev._p.get_or_add_pPr()
                if _ppPr.find(qn("w:sectPr")) is None:
                    if _strip_break_runs(_prev):
                        bab_para.paragraph_format.page_break_before = True
                    _bpPr.remove(_bsp)
                    _ppPr.append(_bsp)
                    _moved = True

    # --- v0.6.11 (lanjutan): dua perbaikan yang membuat pemindahan
    # sectPr di atas tidak merusak rendering di LibreOffice.
    # (a) sectPr yang menutup section Arab baru BAB I (paragraf
    # pertama ber-sectPr sesudah judul BAB I) sering membawa
    # w:cols num=2 dari sumber (section isi sumber 2 kolom). Judul
    # BAB I yang kini duduk di section itu ikut ter-layout 2 kolom
    # — indent kanannya (w:ind right) melebihi lebar kolom
    # sehingga judul terbelah huruf per huruf ke dua kolom
    # (nifas: "B"/"A" kolom kiri, "B I"/"PENDAHULUAN" kolom
    # kanan). Hapus elemen w:cols-nya: section ini hanya berisi
    # judul BAB I + paragraf kosong, jadi 1 kolom sudah benar;
    # isi (section berikut) tidak tersentuh. Mengubah num=2
    # menjadi num=1 TIDAK cukup (LibreOffice mengabaikannya bila
    # equalWidth=0 tanpa w:col); menghapus elemennya yang
    # berhasil (eksperimen F3).
    # (b) Setelah batas dipindah, LibreOffice mengabaikan section
    # break nextPage di sectPr penutup itu: body ikut naik ke
    # halaman BAB I dan total halaman berkurang 1 (nifas 32 -> 31).
    # Penyebabnya: dua section break berurutan (perubahan format
    # Romawi->Arab di batas pindahan + nextPage penutup) digabung
    # LibreOffice menjadi satu break (dibuktikan: G1 — ubah
    # nextPage penutup jadi continuous di baseline saja sudah
    # menggeser body ke halaman BAB I). Gantinya, sisipkan run
    # break halaman (w:br type=page) di awal paragraf isi pertama
    # sesudah sectPr penutup: break level-run tidak digabung
    # dengan section break, jadi body kembali mulai di halaman
    # baru (pagination baseline pulih, nifas 32 halaman). Tipe
    # sectPr penutup DIBIARKAN nextPage supaya restart pgNumType
    # start=1 tetap jalan — mengubahnya jadi continuous membuat
    # footer BAB I tampil "8", bukan "1" (eksperimen I1 vs I2).
    # Run break ini BUKAN section break nextPage baru dan TIDAK
    # memaksa BAB I ke halaman baru (BAB I sudah di halamannya
    # sendiri); ia hanya mengembalikan pemisah halaman body yang
    # hilang akibat pemindahan batas.
    if _moved:
        _close_j = None
        for _j in range(bab_j, len(doc.paragraphs)):
            _p = doc.paragraphs[_j]._p.find(qn("w:pPr"))
            if _p is not None and _p.find(qn("w:sectPr")) is not None:
                _close_j = _j
                break
        if _close_j is not None:
            _cpPr = doc.paragraphs[_close_j]._p.find(qn("w:pPr"))
            _csp = _cpPr.find(qn("w:sectPr"))
            _cols = _csp.find(qn("w:cols"))
            if _cols is not None:
                _csp.remove(_cols)
            if _close_j + 1 < len(doc.paragraphs):
                _nxt = doc.paragraphs[_close_j + 1]
                _r = OxmlElement("w:r")
                _br = OxmlElement("w:br")
                _br.set(qn("w:type"), "page")
                _r.append(_br)
                _nxt._p.insert(0, _r)

    # --- BAB II+: TIDAK ada break baru (lihat docstring). Yang
    # dibutuhkan hanya peta "BAB ini mulai di halaman baru?" untuk
    # menentukan section mana yang diberi titlePg di bawah. ---
    bab_entries = [e for e in entries if e.get("kind") == "bab"]
    starts_by_entry = {}
    if bab_new_pages is not None:
        for e, starts in zip(bab_entries, bab_new_pages):
            starts_by_entry[id(e)] = bool(starts)
    bab_new_by_id = {id(bab_para._p): bool(bab_new_page)}
    for entry, para, j in bab_paras[1:]:
        starts = starts_by_entry.get(id(entry))
        if starts is None:
            starts = _bab_starts_new_page(doc, para, j)
        bab_new_by_id[id(para._p)] = bool(starts)

    # --- petakan ulang: section tempat paragraf BAB I berada, dan
    # paragraf pertama tiap section (untuk footer tengah halaman BAB).
    # sec_first_text = paragraf BERISI pertama tiap section (paragraf
    # kosong di atas judul BAB dilewati — judul BAB-lah yang
    # menentukan halaman pertama section itu). ---
    idx = 0
    bab_idx = None
    bab_has_break = False
    sec_first = [None]
    sec_first_text = [None]
    for child in body.iterchildren():
        if child.tag != qn("w:p"):
            continue
        if sec_first[idx] is None:
            sec_first[idx] = child
        if sec_first_text[idx] is None:
            txt = "".join(n.text or ""
                          for n in child.findall(".//" + qn("w:t")))
            if txt.strip():
                sec_first_text[idx] = child
        if child is bab_para._p:
            bab_idx = idx
            pPr = child.find(qn("w:pPr"))
            bab_has_break = (pPr is not None
                             and pPr.find(qn("w:sectPr")) is not None)
        pPr = child.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            idx += 1
            sec_first.append(None)
            sec_first_text.append(None)
    while sec_first and sec_first[-1] is None:
        sec_first.pop()
    while sec_first_text and sec_first_text[-1] is None:
        sec_first_text.pop()
    if bab_idx is None:
        bab_idx = len(doc.sections) - 1
    # Bila break ada di paragraf BAB I sendiri, BAB I masih di section
    # depan -> body mulai section berikut; bila tidak, BAB I sudah
    # paragraf pertama section body.
    body_idx = bab_idx + 1 if bab_has_break else bab_idx
    for i, s in enumerate(doc.sections):
        if i < body_idx:
            _sectPr_pgnumtype(s._sectPr, "lowerRoman")
        elif i == body_idx:
            _sectPr_pgnumtype(s._sectPr, "decimal", start=1)
        else:
            # Section setelah body: lanjutkan Arab. Dokumen sumber
            # kadang punya restart sendiri (pgNumType start=14/23
            # tanpa fmt) yang memecah urutan — dinormalisasi.
            _sectPr_pgnumtype(s._sectPr, "decimal")
    try:
        even_odd = bool(doc.settings.odd_and_even_pages_header_footer)
    except Exception:
        even_odd = False
    for i, s in enumerate(doc.sections):
        align = (WD_ALIGN_PARAGRAPH.CENTER if i < body_idx
                 else WD_ALIGN_PARAGRAPH.RIGHT)
        _ensure_section_footers(s, align, i == 0, even_odd)
    # --- cover dulu, SEBELUM footer tengah BAB ---
    # (v0.6.10) Urutan penting: first-page footer cover harus
    # dikosongkan SEBELUM _ensure_bab_first_page_footer membuat
    # part footer 'first' baru untuk section BAB. Bila cover
    # dikosongkan sesudahnya, LibreOffice menampilkan '1' di cover
    # (footer 'first' section BAB dipakai untuk halaman 1).
    s0 = doc.sections[0]
    if not any(p.text.strip() for p in s0.header.paragraphs):
        s0.different_first_page_header_footer = True
    # Cover (halaman pertama section 0) tidak menampilkan nomor:
    # pastikan first-page footer-nya ADA dan kosong. Bila dokumen
    # tidak punya referensi first footer, LibreOffice jatuh kembali
    # ke footer default sehingga nomor ikut tampil di cover.
    # (Bila section 0 sekaligus section BAB I — dokumen tanpa halaman
    # depan — cover tidak ada, jadi jangan dikosongkan: BAB I di
    # halaman 1 tetap menampilkan nomornya di tengah.)
    if body_idx != 0:
        try:
            fp = s0.first_page_footer
            if fp.part is not None and fp.part is s0.footer.part:
                # first & default berbagi part yang sama: pisahkan dulu
                for ref in list(s0._sectPr.findall(qn("w:footerReference"))):
                    if ref.get(qn("w:type")) == "first":
                        s0._sectPr.remove(ref)
                fp = s0.first_page_footer
            if fp.is_linked_to_previous:
                fp.is_linked_to_previous = False
            for p in list(fp.paragraphs):
                _strip_page_field(p)
        except Exception:
            pass
    # --- footer tengah untuk halaman pertama BAB ---
    # Section yang paragraf berisi pertamanya adalah judul BAB yang
    # mulai di halaman baru mendapat first-page footer dengan nomor
    # di tengah bawah. Section BAB yang mulai di tengah halaman
    # TIDAK diberi titlePg: halaman pertamanya di LibreOffice
    # adalah halaman BERIKUT (halaman yang pertama kali dimulai
    # paragraf section itu), dan nomornya akan salah tengah. BAB I
    # tidak masuk di sini: ia paragraf terakhir section depan,
    # halamannya sudah tengah (Romawi) lewat footer default.
    centered = set()
    for i, fp in enumerate(sec_first_text):
        if (i >= body_idx and fp is not None
                and bab_new_by_id.get(id(fp), False)):
            centered.add(i)
    for i in sorted(centered):
        if i < len(doc.sections):
            _ensure_bab_first_page_footer(doc.sections[i])
    # Hapus titlePg LIAR dari section lain (kecuali cover sec 0):
    # titlePg warisan sumber pada section yang paragraf pertamanya
    # BUKAN BAB yang mulai di halaman baru akan menengahkan halaman
    # pertama section itu — halaman yang salah (Gunak: sec 3 punya
    # titlePg warisan, paragraf pertamanya "Penjelasan Bagan :",
    # sehingga hal. 31 jadi tengah padahal BAB III di hal. 32).
    # Tanpa titlePg, section itu memakai footer default (kanan).
    for i, s in enumerate(doc.sections):
        if i == 0 or i in centered:
            continue
        _sp = s._sectPr
        _tp = _sp.find(qn("w:titlePg"))
        if _tp is not None:
            _sp.remove(_tp)
    doc.save(out_path)
    return True


# ---------------------------------------------------------------- verifikasi
def _heading_page_out(pages_lines, entry):
    """Halaman (1-based) heading entri di PDF HASIL: kemunculan
    TERAKHIR dari baris yang persis sama dengan teks entri (bukan
    baris daftar isi dan bukan kalimat isi). Dipakai untuk entri
    halaman depan yang judulnya ada SETELAH daftar isi."""
    key = _nospace(entry["text"])
    found = None
    for pi, lines in enumerate(pages_lines):
        for ln in lines:
            if _nospace(ln) == key:
                found = pi + 1
    return found


def verify_output(pdf_path, entries):
    """Verifikasi: hitung ulang label dari PDF HASIL dan bandingkan
    dengan label di daftar isi."""
    pages = pdf_pages_text(pdf_path)
    lines = pdf_pages_lines(pdf_path)
    # BAB I = halaman kemunculan terakhir "BAB I PENDAHULUAN" di PDF hasil.
    bab1_hits = [i + 1 for i, t in enumerate(pages)
                 if _page_contains(t, "BAB I PENDAHULUAN")]
    bab1 = bab1_hits[-1] if bab1_hits else None
    mismatches = []
    for e, (pdf_page, expect) in zip(
            entries, _expected_labels(pages, lines, entries, bab1)):
        if pdf_page is None:
            mismatches.append((e["text"], e["page"], "tidak ketemu"))
        elif expect != e["page"]:
            mismatches.append((e["text"], e["page"], expect))
    return mismatches



def _expected_labels(pages, lines, entries, bab1):
    """Untuk tiap entri, hitung (pdf_page, label) dari PDF — dipakai
    verifikasi dan koreksi label. Logika sama seperti verify_output:
    entri halaman depan dicari dari kemunculan pertama (post_anchor:
    terakhir), entri isi dicari berurutan mulai dari halaman BAB I
    agar tidak ketemu baris daftar isi duluan."""
    pointer = (bab1 - 1) if bab1 is not None else 0
    out = []
    for e in entries:
        keys = e.get("search") or [("text", e["text"])]
        if e["kind"] == "front":
            if e.get("post_anchor"):
                pdf_page = _heading_page_out(lines, e)
            else:
                pdf_page = _find_key(pages, keys, 0)
        else:
            pdf_page = _find_key(pages, keys, pointer)
            if pdf_page is not None:
                pointer = pdf_page - 1
        if pdf_page is None:
            out.append((None, None))
            continue
        if e["kind"] == "front":
            expect = to_roman(pdf_page)
        elif bab1 is not None:
            expect = str(pdf_page - bab1 + 1)
        else:
            expect = str(pdf_page)
        out.append((pdf_page, expect))
    return out


def _fix_labels_from_output(pdf_path, entries, gambar, tabel):
    """Koreksi SEMUA label (isi + depan + gambar/tabel) dari PDF hasil
    (v0.6.1): penyisipan daftar isi mengubah alur halaman depan
    (BAB I tidak selalu mulai di halaman baru), jadi label dari PDF
    masuk bisa bergeser 1 halaman. Hitung ulang dari PDF hasil;
    kembalikan True bila ada yang berubah. Mengubah angka di baris
    daftar isi tidak mengubah pagination, jadi satu rebuild cukup."""
    pages = pdf_pages_text(pdf_path)
    lines = pdf_pages_lines(pdf_path)
    bab1_hits = [i + 1 for i, t in enumerate(pages)
                 if _page_contains(t, "BAB I PENDAHULUAN")]
    bab1 = bab1_hits[-1] if bab1_hits else None
    changed = False
    for lst in (entries, gambar, tabel):
        for e, (_pg, expect) in zip(
                lst, _expected_labels(pages, lines, lst, bab1)):
            if expect is None:
                continue
            if expect != e["page"]:
                e["page"] = expect
                changed = True
    return changed


# ---------------------------------------------------------------- pipeline
def process(docx_path, work_dir, include_lists=True):
    """Pipeline lengkap. Kembalikan dict laporan + path hasil.
    include_lists=False: DAFTAR GAMBAR / DAFTAR TABEL tidak dibuat
    (daftar isi tetap jalan seperti biasa)."""
    os.makedirs(work_dir, exist_ok=True)
    entries = extract_structure(docx_path)
    assign_display(entries)
    if not entries:
        raise ValueError("Tidak ada heading yang terdeteksi di dokumen.")
    gambar, tabel = extract_captions(docx_path) if include_lists else ([], [])

    # v0.6.1: render PDF input dari salinan yang baris daftar lamanya
    # sudah dibuang — kalau tidak, pencarian halaman (substring)
    # ketemu teks daftar lama di halaman depan dan semua label isi
    # jadi salah (mis. 'Rumusan Masalah' ketemu di halaman daftar
    # isi, bukan di BAB I).
    clean_path = os.path.join(work_dir, "_bersih.docx")
    _clean = Document(docx_path)
    _remove_old_toc_runs(_clean)
    _clean.save(clean_path)
    pdf_in = render_pdf(clean_path, work_dir)
    pages_in = pdf_pages_text(pdf_in)
    total_in = len(pages_in)
    find_pages(pages_in, entries)
    compute_labels(entries)
    bab1_pages = [e["pdf_page"] for e in entries
                  if e["kind"] == "bab" and e.get("pdf_page")]
    bab1 = min(bab1_pages) if bab1_pages else None
    if gambar:
        find_pages(pages_in, gambar)
        compute_labels(gambar, bab1=bab1)
    if tabel:
        find_pages(pages_in, tabel)
        compute_labels(tabel, bab1=bab1)

    out_docx = os.path.join(work_dir, "hasil-daftar-isi.docx")
    build_output_docx(docx_path, entries, out_docx,
                      gambar=gambar, tabel=tabel)

    # --- verifikasi: nomor Arab relatif terhadap BAB I, jadi penyisipan
    # TOC sebelum BAB I TIDAK mengubah nomor (BAB I dan semua heading
    # bergeser sama besar) — kecuali bila BAB I tidak mulai di halaman
    # baru: alur halaman depannya berubah dan label isi bisa bergeser
    # 1 halaman (v0.6.1). Karena itu SEMUA label dikoreksi ulang dari
    # PDF hasil; mengubah angka di baris daftar isi tidak mengubah
    # pagination, jadi satu pass rebuild cukup.
    pdf_out = render_pdf(out_docx, work_dir)
    total_out = len(pdf_pages_text(pdf_out))
    if _fix_labels_from_output(pdf_out, entries, gambar, tabel):
        build_output_docx(docx_path, entries, out_docx,
                          gambar=gambar, tabel=tabel)
        pdf_out = render_pdf(out_docx, work_dir)
        total_out = len(pdf_pages_text(pdf_out))
    mismatches = verify_output(pdf_out, entries)
    mismatches_gambar = verify_output(pdf_out, gambar) if gambar else []
    mismatches_tabel = verify_output(pdf_out, tabel) if tabel else []

    # --- nomor halaman footer (v0.6.0): field PAGE dinamis, jadi
    # tidak mengubah pagination dan tidak perlu verifikasi ulang.
    # Tiap BAB mulai di halaman baru? Dicek dari PDF HASIL (bukan PDF
    # sumber — penyisipan daftar isi bisa menggeser posisi BAB dalam
    # halaman): cari halaman judul BAB (kenal judul telanjang), lalu
    # lihat apakah judulnya baris pertama halaman itu. Judul bagian
    # pending (v0.6.6) yang ditempel sebelum BAB I ditoleransi sebagai
    # baris sebelum judul BAB I. Hasilnya dipakai add_footer_page_
    # numbers: BAB yang mulai di halaman baru dapat nextPage break +
    # first-page footer tengah; yang mulai di tengah halaman tidak.
    bab_new_page = False
    bab_new_pages = []
    bab_entries = [e for e in entries if e.get("kind") == "bab"]
    if bab_entries:
        try:
            lines_out = pdf_pages_lines(pdf_out)
            # posisi tiap entri BAB di entries, untuk mencari judul
            # pending-nya (sub break_before tepat setelah BAB itu)
            bab_pos = [i for i, e in enumerate(entries)
                       if e.get("kind") == "bab"]
            for bi, be in enumerate(bab_entries):
                pg = _bab_start_page(lines_out, be)
                pend = []
                for e in entries[bab_pos[bi] + 1:]:
                    if e.get("kind") == "bab":
                        break
                    if e.get("break_before"):
                        pend.append(e["text"])
                bab_new_pages.append(
                    _bab_starts_new_page_pdf(lines_out, pg, pend))
        except Exception:
            bab_new_pages = [False] * len(bab_entries)
        bab_new_page = bab_new_pages[0] if bab_new_pages else False
    footer_ok = add_footer_page_numbers(out_docx, entries,
                                        bab_new_page=bab_new_page,
                                        bab_new_pages=bab_new_pages)

    return {
        "entries": entries,
        "daftar_gambar": gambar,
        "daftar_tabel": tabel,
        "toc_text": format_toc_text(entries),
        "output_docx": out_docx,
        "pages_in": total_in,
        "pages_out": total_out,
        "shift": total_out - total_in,
        "corrected": False,
        "footer_ok": footer_ok,
        "bab_new_pages": bab_new_pages,
        "mismatches": mismatches,
        "mismatches_gambar": mismatches_gambar,
        "mismatches_tabel": mismatches_tabel,
    }


if __name__ == "__main__":
    import sys
    import json
    src = sys.argv[1]
    work = sys.argv[2] if len(sys.argv) > 2 else tempfile.mkdtemp()
    rep = process(src, work)
    print(rep["toc_text"])
    if rep["daftar_gambar"]:
        print("\nDAFTAR GAMBAR")
        print(format_toc_text(rep["daftar_gambar"]))
    if rep["daftar_tabel"]:
        print("\nDAFTAR TABEL")
        print(format_toc_text(rep["daftar_tabel"]))
    print(f"\n[halaman masuk: {rep['pages_in']}, keluar: {rep['pages_out']}, "
          f"geser: {rep['shift']}]")
    print(f"[verifikasi TOC: {'LOLOS' if not rep['mismatches'] else rep['mismatches']}]")
    print(f"[verifikasi gambar: {'LOLOS' if not rep['mismatches_gambar'] else rep['mismatches_gambar']}]")
    print(f"[verifikasi tabel: {'LOLOS' if not rep['mismatches_tabel'] else rep['mismatches_tabel']}]")
    print(f"[hasil: {rep['output_docx']}]")
