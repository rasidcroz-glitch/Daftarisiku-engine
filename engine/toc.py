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

Format daftar isi (contoh dosen): titik-titik sampai kanan, nomor
halaman rata kanan; yang tebal hanya halaman depan + judul BAB.
"""

import os
import re
import shutil
import subprocess
import tempfile

import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.shared import Cm, Pt, Twips

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
FRONT_TITLES = [
    "PERNYATAAN ORIGINALITAS",
    "HALAMAN PERSETUJUAN PEMBIMBING",
    "HALAMAN PENGESAHAN TIM PENGUJI",
    "KATA PENGANTAR",
    "DAFTAR ISI",
    "HALAMAN PERSETUJUAN",
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
    """Buang awalan penomoran ('A. ', '1. ', 'a. ') dari judul."""
    return re.sub(r"^([A-Za-z]\.|\d+\.)\s*", "", text or "")


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

    def push(e):
        while stack and stack[-1]["level"] >= e["level"]:
            stack.pop()
        stack.append({"level": e["level"],
                      "ilvl": e.get("ilvl"), "indent": e.get("indent"),
                      "numId": e.get("numId"),
                      "lp_bold": e.get("lp_bold", False)})

    def add(text, level, kind, idx, bold=False, display=None, search=None,
            post_anchor=False, term=None, ilvl=None, indent=None,
            numId=None, lp_bold=False):
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
             "lp_bold": lp_bold}
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
                    if (paras[j].style.name == "Normal"
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
            lvl = level_for_list(ilvl, ind, is_term, top,
                                 numId=_numId(p))
            if lvl is None or lvl > 3:
                continue
            add(text, lvl, "sub", i, bold=False, term=term,
                ilvl=ilvl, indent=ind, numId=_numId(p), lp_bold=bold)
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
    batas kata agar 'BAB II' tidak cocok dengan 'BAB III')."""
    kind, pat = key
    if kind == "bab":
        rx = re.compile(rf"BAB\s+{pat}\b(?![IVXLC])", re.IGNORECASE)
        return rx.search(page_text) is not None
    return _page_contains(page_text, pat)


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
    sec = doc.sections[0]
    usable_twips = int(
        (sec.page_width - sec.left_margin - sec.right_margin) / 635
    )  # 1 twip = 635 EMU

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

    def add_entry(text, page, level, bold):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.left_indent = Cm(0.6 * level)
        p.paragraph_format.tab_stops.add_tab_stop(
            Twips(usable_twips), WD_TAB_ALIGNMENT.RIGHT,
            WD_TAB_LEADER.DOTS)
        r = p.add_run(f"{text}\t{page}")
        r.font.name = "Times New Roman"
        r.font.size = Pt(12)
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
        for e in tabel:
            t_made.append(add_entry(disp(e), e["page"], 0, False))
        insert_after_seq(t_anchor, t_made)

    doc.save(out_path)
    return out_path


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
    dengan label di daftar isi. Entri halaman depan dicari dari
    kemunculan PERTAMA (judulnya ada sebelum daftar isi; yang
    post_anchor memakai kemunculan terakhir baris judulnya).
    Entri isi dicari berurutan MULAI DARI HALAMAN BAB I — baris
    daftar isi/gambar/tabel (yang memuat teks heading juga) ada
    sebelum BAB I, jadi pointer harus mulai di sana agar tidak
    ketemu baris daftar duluan."""
    pages = pdf_pages_text(pdf_path)
    lines = pdf_pages_lines(pdf_path)
    # BAB I = halaman kemunculan terakhir "BAB I PENDAHULUAN" di PDF hasil.
    bab1_hits = [i + 1 for i, t in enumerate(pages)
                 if _page_contains(t, "BAB I PENDAHULUAN")]
    bab1 = bab1_hits[-1] if bab1_hits else None
    mismatches = []
    pointer = (bab1 - 1) if bab1 is not None else 0
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
            mismatches.append((e["text"], e["page"], "tidak ketemu"))
            continue
        if e["kind"] == "front":
            expect = to_roman(pdf_page)
        elif bab1 is not None:
            expect = str(pdf_page - bab1 + 1)
        else:
            expect = str(pdf_page)
        if expect != e["page"]:
            mismatches.append((e["text"], e["page"], expect))
    return mismatches



def _fix_post_anchor_labels(pdf_path, entries):
    """Koreksi label Romawi entri halaman depan yang judulnya ada
    SETELAH daftar isi (mis. DAFTAR TABEL/GAMBAR): halaman judul itu
    bergeser karena daftar isi yang disisipkan, jadi label dari PDF
    masuk sudah tidak berlaku. Hitung ulang dari PDF hasil.
    Kembalikan True bila ada label yang berubah."""
    lines = pdf_pages_lines(pdf_path)
    changed = False
    for e in entries:
        if e["kind"] != "front" or not e.get("post_anchor"):
            continue
        pdf_page = _heading_page_out(lines, e)
        if pdf_page is None:
            continue
        label = to_roman(pdf_page)
        if label != e["page"]:
            e["page"] = label
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

    pdf_in = render_pdf(docx_path, work_dir)
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
    # bergeser sama besar). Render ulang dan hitung ulang labelnya.
    # Label Romawi entri halaman depan yang judulnya ada setelah
    # daftar isi dikoreksi dari PDF hasil (satu pass rebuild bila
    # berubah — mengubah angka di baris daftar isi tidak mengubah
    # pagination, jadi satu pass cukup).
    pdf_out = render_pdf(out_docx, work_dir)
    total_out = len(pdf_pages_text(pdf_out))
    if _fix_post_anchor_labels(pdf_out, entries):
        build_output_docx(docx_path, entries, out_docx,
                          gambar=gambar, tabel=tabel)
        pdf_out = render_pdf(out_docx, work_dir)
        total_out = len(pdf_pages_text(pdf_out))
    mismatches = verify_output(pdf_out, entries)
    mismatches_gambar = verify_output(pdf_out, gambar) if gambar else []
    mismatches_tabel = verify_output(pdf_out, tabel) if tabel else []

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
