"""
attendance_register_filler.py - MULTI-COMPANY / MULTI-LAYOUT VERSION
=============================================================
Fills the PDF attendance register starting from the Platform Excel reports.

Supports multiple companies and two register layouts:
  - LAYOUT A: Cliente A (3 pages per session: A=students 1-25,
              B=26-57, C=instructors+topics). ELENCO ALLIEVI with
              interleaved rows.
  - LAYOUT B: single-row courses (clean single row in the ELENCO
              ALLIEVI, time in DALLE/ALLE format).

ANTI-BUG PRINCIPLE: the source of the names is ALWAYS the ELENCO ALLIEVI
page of the PDF we are filling in. A fixed cache is never used, as it
could contain names from another company. The per-company storage
(docs/aziende/) is only a verification archive, never the active source.

USAGE:
    py attendance_register_filler.py               -> input/*.pdf + input/*.xlsx -> output/
    py attendance_register_filler.py --no-open --no-preview
    py attendance_register_filler.py --archive     -> archives input into elaborati/<ts>/

LOGIC:
1. Detects the layout (A or B) from the ELENCO ALLIEVI page.
2. Extracts the students from the ELENCO ALLIEVI of the current PDF
   -> {numero: {nome, cf, data, azienda}}.
3. Detects the company from the AZIENDA column (fallback: PDF title).
4. If the company is new -> creates docs/aziende/<slug>.json (self-learning).
5. For each session: date + time -> finds the Platform Excel with the same
   date -> writes NAME + PRESENTE/ASSENTE for each student.

MATCHING: normalized surname+name + fuzzy similarity (>=80%).
"""
import sys
import os
import re
import json
import glob
import unicodedata
import subprocess
from collections import Counter
from pathlib import Path

import openpyxl
import pdfplumber
import fitz


# ============================ utility =====================================
def norm(s: str) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.upper()
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def slugify(s: str) -> str:
    """Slug for the file name: 'CLIENTE_A S.P.A.' -> 'datamanagement_italia'."""
    s = norm(s)
    s = re.sub(r"\b(SPA|S P A|SRL|S R L|SNC|S N C)\b", "", s)
    s = re.sub(r"[^A-Z0-9]+", "_", s)
    return s.strip("_").lower()


def nome_chiave(cognome, nome=""):
    return norm(f"{cognome} {nome}")


def similarity(a: str, b: str) -> float:
    import difflib
    return difflib.SequenceMatcher(None, norm(a), norm(b)).ratio()


# ============================ path ========================================
BASE = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE / "input"
OUTPUT_DIR = BASE / "output"
DOCS_DIR = BASE / "docs"
AZIENDE_DIR = DOCS_DIR / "aziende"
REF_CACHE_LEGACY = DOCS_DIR / "reference_list_final.json"  # no longer used as a source

CF_RE = re.compile(r"^[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]$")


# ============================ ELENCO ALLIEVI: layout ======================
def _trova_pagine_elenco(pdf):
    """Returns the 0-based indices of the ELENCO ALLIEVI pages.

    The first page has the 'ELENCO ALLIEVI' header; the list may continue
    on the following pages WITHOUT a repeated header. Continuation pages
    are recognized because they contain tax codes (16-char CF), which
    never appear on the attendance pages. We stop when a page has no
    CF (typically the first attendance/topics page)."""
    out = []
    start_idx = None
    for idx, page in enumerate(pdf.pages):
        text = (page.extract_text() or "").upper()
        if "ELENCO ALLIEVI" in text or "ELENCO  ALLIEVI" in text:
            start_idx = idx
            out.append(idx)
            break
    if start_idx is None:
        return out
    # extend to the following pages as long as we find CF codes
    for idx in range(start_idx + 1, len(pdf.pages)):
        page = pdf.pages[idx]
        words = page.extract_words()
        ha_cf = any(CF_RE.match(w['text']) for w in words)
        if ha_cf:
            out.append(idx)
        else:
            break
    return out


def rileva_layout(pdf, pagine_elenco_idx):
    """Returns 'A' or 'B' based on the structure of the ELENCO ALLIEVI page.

    LAYOUT A (Cliente A): interleaved rows - the number (N) is on a
      separate row (y) with respect to surname/name (y+/-6). Moreover the
      CF/DATA/AZIENDA fields are vertically very close to the surname.
    LAYOUT B (single-row courses): clean single row -
      '1 ASCOLI ASCOLI CF DATA AZIENDA' all on the same y.

    Robust heuristic: we group the words by row (2px tolerance) and
    count how many rows start with a number (1-100) followed by alphabetic
    text ON THE SAME row. Layout B -> many rows like that; Layout A ->
      few (because the number and the name are on different y values).
    """
    if not pagine_elenco_idx:
        return "A"  # default
    page = pdf.pages[pagine_elenco_idx[0]]
    words = page.extract_words()
    rows = {}
    for w in words:
        key = round(w['top'] / 2) * 2
        rows.setdefault(key, []).append(w)
    righe_complete = 0
    for top in sorted(rows):
        line = sorted(rows[top], key=lambda w: w['x0'])
        txts = [w['text'] for w in line]
        if not txts:
            continue
        # row starting with a number and immediately followed by a surname (alpha word)
        if re.fullmatch(r"\d{1,3}", txts[0]) and len(txts) >= 2 and \
                re.match(r"^[A-Z]+$", txts[1]):
            righe_complete += 1
    # Layout B: many complete rows (>5). Layout A: few.
    return "B" if righe_complete >= 5 else "A"


# Typical company-name words that, as soon as they are found in an
# ELENCO ALLIEVI row (layout B), close the student's "nome" field and
# start the "azienda" field. Prevents names from being polluted when the
# CF or the date is not recognized by the parser (e.g. "ROSSI MARIA
# INFORMATION AND" instead of "ROSSI MARIA", or "BIANCHI LUCA I.CON."
# instead of "BIANCHI LUCA").
AZIENDA_WORDS = {
    "INFORMATION", "AND", "STARTUP", "START", "UP", "COSTITUITA",
    "COSTITUZIONE", "SOCIETA", "SOC", "SRL", "SPA", "SNC", "SAS",
    "TECNOLOGIA", "TECNOLOGIE", "DIGITAL", "SERVIZI", "SISTEMI",
    "SOLUTIONS", "SOLUTION", "GROUP", "CONSULENZE", "INFORMATICA",
    "INGEGNERIA", "IMPRESA", "AZIENDA", "HOLDING", "COOP", "COOPERATIVA",
    "CORSO_A", "CORSO_B", "CORSO_C", "CORSO_D",
}


def _is_azienda_word(t):
    """True if the token is a typical company-name word.

    IMPORTANT: periods must be removed WITHOUT replacing them with a space,
    otherwise 'X.Y.Z.' would become 'X Y Z' (more words) and would not match
    'XYZ'. So a "tight" normalization is done here (periods/apostrophes
    removed), different from norm() which turns them into spaces.
    """
    s = t.upper()
    # remove periods/apostrophes/hyphens without leaving a space: X.Y.Z. -> XYZ
    s = re.sub(r"[.\'\-]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    # first try the whole token (e.g. S.P.A.->SPA)
    if s in AZIENDA_WORDS:
        return True
    # then try the single words (e.g. START-UP -> 'START UP' -> match 'START')
    for w in s.split():
        if w in AZIENDA_WORDS:
            return True
    return False


def _estrai_riga_elenco_B(page):
    """LAYOUT B parser: each row = 'N COGNOME NOME CF DATA AZIENDA' on one y.
    Returns {numero: {nome, cf, data, azienda}}."""
    words = page.extract_words()
    rows = {}
    for w in words:
        key = round(w['top'] / 2) * 2
        rows.setdefault(key, []).append(w)
    out = {}
    for top in sorted(rows):
        line = sorted(rows[top], key=lambda w: w['x0'])
        txts = [w['text'] for w in line]
        if not txts or not re.fullmatch(r"\d{1,3}", txts[0]):
            continue
        num = int(txts[0])
        if not (1 <= num <= 100):
            continue
        # collect the words after the number. The phase advances like this:
        #   "nome" -> (CF detected) -> "after_cf" -> (date) -> "after_data"
        #   "nome" -> (company-word) -> "after_azienda"  [anti-pollution filter]
        # In any phase after "nome", the leftover tokens (unless they are a
        # CF or a date) go into the company: this way "INFORMATION AND ..."
        # does not leave the word "AND" in the name.
        parts = []
        cf = data = azienda = ""
        phase = "nome"
        for t in txts[1:]:
            if CF_RE.match(t):
                cf = t
                phase = "after_cf"
                continue
            if phase == "after_cf":
                if re.match(r"^\d{4}-\d{2}-\d{2}$", t) or \
                        re.match(r"^\d{2}/\d{2}/\d{4}$", t):
                    data = t
                    phase = "after_data"
                    continue
                # after the CF but before the date (rare): ignore
                continue
            # Anti-pollution filter: a company-word closes the name.
            if phase == "nome" and _is_azienda_word(t):
                phase = "after_azienda"
                azienda = (azienda + " " + t).strip()
                continue
            # In any "after" phase (after_data / after_azienda): everything
            # goes into the company, nothing ends up in the name anymore.
            if phase in ("after_data", "after_azienda"):
                azienda = (azienda + " " + t).strip()
                continue
            parts.append(t)
        nome = " ".join(parts).strip()
        # Cap backstop: max 3 words (2 spaces) in the name. Safety net for
        # company names not covered by AZIENDA_WORDS. Keeps compound
        # surnames (DI PIETRO ALESSANDRO = 3 words) and double first names.
        nome_words = nome.split()
        if len(nome_words) > 3:
            nome = " ".join(nome_words[:3])
        if nome and any(c.isalpha() for c in nome):
            out[num] = {"nome": nome, "cf": cf, "data": data, "azienda": azienda}
    return out


NOME_WORD_RE = re.compile(r"^[A-Za-z][A-Za-z'\u00C0-\u024F]+$")

# ELENCO ALLIEVI header words to exclude from the name (case-insensitive)
HEADER_WORDS = {
    "COGNOME", "NOME", "E", "DELL'ALLIEVO", "DELLALLIEVO",
    "CF", "DATA", "AZIENDA", "N",
    "VIDIMATO", "DA", "PAGINA", "DI",
}


def _is_nome_valido(t):
    """True if the word is a plausible component of a name (not a header)."""
    if t.upper() in HEADER_WORDS:
        return False
    # noble particles/phrases in Italian surnames
    if t in ("De", "Di", "Da", "La", "Lo", "Del", "Della", "Delle", "Dello"):
        return True
    return bool(NOME_WORD_RE.match(t))


def _estrai_riga_elenco_A(page):
    """LAYOUT A parser (Cliente A): interleaved rows.
    The CF, the surname+name and the number sit on NEARBY but not always
    identical y values (delta up to +/-12px). Example:
        y124: CF@x236 CLIENTE_A@x413
        y126: Battoli@x96 Sigfrido@x127          <- name on a row different from the CF
        y130: 1@x77                               <- number on yet another row

    "Logical band" strategy: we group the words into wide horizontal
    bands (~30px tolerance) so that the CF, name and number of the same
    person fall into the same band. From each band we extract:
      - numero  = integer 1-100 with x < 95
      - CF      = token matching CF_RE
      - nome    = alphabetic words (not header) with 90 < x < 200
      - azienda = alphabetic words with x > x_CF
    """
    words = page.extract_words()
    bande = {}
    for w in words:
        key = round(w['top'] / 30) * 30
        bande.setdefault(key, []).append(w)
    X_NUM_MAX = 95
    X_NOME_MAX = 200   # name sits at x ~93-153; CF starts ~233
    out = {}
    sorted_keys = sorted(bande)
    for bi, bkey in enumerate(sorted_keys):
        gruppo = sorted(bande[bkey], key=lambda w: w['x0'])
        num = None
        cf = ""
        nome_parts = []
        azienda_parts = []
        for w in gruppo:
            t = w['text']
            if re.fullmatch(r"\d{1,3}", t):
                if w['x0'] < X_NUM_MAX and 1 <= int(t) <= 100:
                    num = int(t)
            elif CF_RE.match(t):
                cf = t
            elif _is_nome_valido(t):
                if w['x0'] < X_NOME_MAX:
                    nome_parts.append((w['x0'], t))
                elif cf:
                    azienda_parts.append((w['x0'], t))
        # missing number: look in an adjacent band
        if num is None and cf and nome_parts:
            for dbkey in (sorted_keys[bi - 1] if bi > 0 else None,
                          sorted_keys[bi + 1] if bi + 1 < len(sorted_keys) else None):
                if dbkey is None:
                    continue
                for w in bande[dbkey]:
                    if re.fullmatch(r"\d{1,3}", w['text']) and w['x0'] < X_NUM_MAX:
                        n = int(w['text'])
                        if 1 <= n <= 100 and n not in out:
                            num = n
                            break
                if num is not None:
                    break
        if num is None or not nome_parts:
            continue
        nome_parts.sort()
        nome = " ".join(p[1] for p in nome_parts)
        azienda_parts.sort()
        azienda = " ".join(p[1] for p in azienda_parts)
        if num not in out:
            out[num] = {"nome": nome, "cf": cf, "data": "", "azienda": azienda}
    return out


def estrai_allievi_elenco(pdf, layout, pagine_elenco_idx):
    """Extracts {numero: {nome, cf, data, azienda}} from all the ELENCO pages.
    Uses majority voting across multiple pages when present."""
    if not pagine_elenco_idx:
        return {}
    parser = _estrai_riga_elenco_A if layout == "A" else _estrai_riga_elenco_B
    nomi = {}  # numero -> Counter() over dict-str
    for idx in pagine_elenco_idx:
        estratti = parser(pdf.pages[idx])
        for num, info in estratti.items():
            key = json.dumps(info, ensure_ascii=False)
            nomi.setdefault(num, Counter())[key] += 1
    out = {}
    for num in sorted(nomi):
        info_str, _ = nomi[num].most_common(1)[0]
        out[num] = json.loads(info_str)
    return out


# ============================ company detection ===========================
# known keywords in the title -> company slug (fallback if the AZIENDA column is empty)
AZIENDE_NOTE = {
    "CLIENTE_A": "cliente_a",
    "CORSO_X": "corso_x",
    "CORSO_Y": "corso_y",
    "CORSO_A": "corso_a",
    "CORSO_B": "corso_b",
}


def rileva_azienda(allievi_estratti, pdf):
    """Detects the company from the AZIENDA column of the list (majority).
    Fallback: keyword in the title (first 3 pages)."""
    # 1) majority vote on the company column
    c = Counter()
    for info in allievi_estratti.values():
        az = info.get("azienda", "").strip()
        if az:
            c[slugify(az)] += 1
    if c:
        return c.most_common(1)[0][0]
    # 2) title fallback
    try:
        for page in pdf.pages[:3]:
            text = (page.extract_text() or "").upper()
            for kw, slug in AZIENDE_NOTE.items():
                if kw in text:
                    return slug
    except Exception:
        pass
    return "sconosciuta"


def gestisci_storage_azienda(slug, allievi_estratti):
    """If the company is new, creates docs/aziende/<slug>.json.
    Returns (azienda_conosciuta: bool, path_storage)."""
    AZIENDE_DIR.mkdir(parents=True, exist_ok=True)
    path = AZIENDE_DIR / f"{slug}.json"
    conosciuta = path.exists()
    if not conosciuta:
        # saves the names extracted from the PDF (self-learning)
        data = {
            "azienda_slug": slug,
            "allievi": {str(n): info for n, info in sorted(allievi_estratti.items())},
        }
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return conosciuta, path


# ============================ parsing Excel ===============================
def leggi_allievi_excel(path_excel: Path) -> dict:
    """{normalized name key: last disconnection str} for students
    (no instructors/tutors)."""
    wb = openpyxl.load_workbook(path_excel, data_only=True)
    ws = wb[wb.sheetnames[0]]
    righe = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
             for r in range(1, ws.max_row + 1)]
    header_idx = None
    for i, row in enumerate(righe):
        upper = [str(v).strip().upper() if v else "" for v in row]
        if "COGNOME" in upper and "NOME" in upper:
            header_idx = i
            break
    if header_idx is None:
        return {}
    header = [str(v).strip().upper() if v else "" for v in righe[header_idx]]
    ic = header.index("COGNOME")
    in_ = header.index("NOME")
    iruolo = header.index("RUOLO") if "RUOLO" in header else None
    iud = None
    for ci, h in enumerate(header):
        if "ULTIMA" in h and "DISC" in h:
            iud = ci
            break
    out = {}
    for row in righe[header_idx + 1:]:
        cog = row[ic] if ic < len(row) else None
        nom = row[in_] if in_ < len(row) else None
        if not cog or not str(cog).strip():
            continue
        ruolo = ""
        if iruolo is not None and iruolo < len(row) and row[iruolo]:
            ruolo = str(row[iruolo]).upper()
        if ruolo in ("DOCENTE", "TUTOR", "ESPERTO"):
            continue
        ud = row[iud] if (iud is not None and iud < len(row)) else None
        ud_str = str(ud) if ud is not None else ""
        out[nome_chiave(str(cog), str(nom or ""))] = ud_str
    return out


def parse_orario(s):
    if s is None or s == "":
        return None
    s = str(s)
    if " " in s:
        s = s.split(" ")[-1]
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?$", s.strip())
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def leggi_data_excel(path_excel: Path):
    wb = openpyxl.load_workbook(path_excel, data_only=True)
    ws = wb[wb.sheetnames[0]]
    data = ora_in = ora_out = None
    for r in range(1, min(ws.max_row, 15) + 1):
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v:
                s = str(v)
                if not data and re.search(r"\d{2}/\d{2}/\d{4}", s):
                    data = re.search(r"(\d{2}/\d{2}/\d{4})", s).group(1)
                if s.lower().startswith("ora inizio") and c + 1 <= ws.max_column:
                    ora_in = ws.cell(r, c + 1).value
                if s.lower().startswith("ora fine") and c + 1 <= ws.max_column:
                    ora_out = ws.cell(r, c + 1).value
    return data, str(ora_in or ""), str(ora_out or "")


def leggi_multi_excel(paths):
    out = []
    for p in paths:
        try:
            data, oi, of = leggi_data_excel(p)
            allievi = leggi_allievi_excel(p)
            out.append({"data": data, "ora_in": oi, "ora_out": of,
                        "allievi": allievi, "fonte": p.name})
            print(f"  {p.name}: date={data} time={oi}-{of}  "
                  f"students={len(allievi)}")
        except Exception as e:
            print(f"  ERROR {p.name}: {e}")
    return out


# ============================ PDF parsing: sessions =======================
def info_pagina(page):
    text = page.extract_text() or ""
    ha_header = "REGISTRO PRESENZE ALLIEVI" in text
    m = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    data = m.group(1) if m else None
    # time: layout A (8:30-13:30) or layout B (DALLE:.. ALLE:..)
    orar = None
    m_orar = re.search(r"(\d{1,2}:\d{2}-\d{1,2}:\d{2})", text)
    if m_orar:
        orar = m_orar.group(1)
    else:
        m_b = re.search(r"DALLE[:\s]*(\d{1,2}:\d{2}(?::\d{2})?)"
                        r"[\s\S]{0,20}?ALLE[:\s]*(\d{1,2}:\d{2}(?::\d{2})?)",
                        text, re.IGNORECASE)
        if m_b:
            inizio = m_b.group(1).split(":")
            fine = m_b.group(2).split(":")
            orar = f"{inizio[0]}:{inizio[1]}-{fine[0]}:{fine[1]}"
    return {"ha_header": ha_header, "data": data, "orar": orar,
            "is_docente_o_argomenti": ("DOCENTI" in text and "TUTOR" in text) or
                                       ("ARGOMENTI" in text.upper() and "VIDIMATO" in text.upper()),
            "ha_tabella_allievi": "COGNOME" in text.upper() and "ENTRATA" in text.upper()}


def raggruppa_sessioni(pdf):
    sessioni = []
    current = None
    for idx, page in enumerate(pdf.pages):
        pg = idx + 1
        info = info_pagina(page)
        if info["ha_header"]:
            if current:
                sessioni.append(current)
            current = {"A": pg, "B": [], "C": [],
                       "data": info["data"], "orar": info["orar"]}
        elif current is None:
            continue
        else:
            is_B = info["ha_tabella_allievi"]
            if not is_B:
                righe_b, _, _ = estrai_righe_pagina_B_nuda(page)
                if len(righe_b) >= 2:
                    is_B = True
            if is_B:
                current["B"].append(pg)
            if info["is_docente_o_argomenti"]:
                current["C"].append(pg)
    if current:
        sessioni.append(current)
    return sessioni


def estrai_righe_allievi_da_pagina(page):
    """[(numero, top_y, nome_esistente, x_entrata, x_uscita)] for an
    attendance page. Valid for both layouts (both have COGNOME/ENTRATA/USCITA)."""
    words = page.extract_words()
    entrata_x = uscita_x = None
    for w in words:
        if w['text'].upper() == "ENTRATA":
            entrata_x = (w['x0'] + w['x1']) / 2
        elif w['text'].upper() == "USCITA":
            uscita_x = (w['x0'] + w['x1']) / 2
    rows = {}
    for w in words:
        key = round(w['top'] / 2) * 2
        rows.setdefault(key, []).append(w)
    header_top = None
    sorted_tops_list = sorted(rows)
    for i, top in enumerate(sorted_tops_list):
        txt = " ".join(w['text'] for w in sorted(rows[top], key=lambda w: w['x0'])).upper()
        if "COGNOME" in txt and "ENTRATA" in txt:
            header_top = top
            break
        if "ENTRATA" in txt and "USCITA" in txt and i + 1 < len(sorted_tops_list):
            next_top = sorted_tops_list[i + 1]
            next_txt = " ".join(w['text'] for w in sorted(rows[next_top], key=lambda w: w['x0'])).upper()
            if "COGNOME" in next_txt:
                header_top = next_top
                break
        if "ENTRATA" in txt and "USCITA" in txt and i > 0:
            prev_top = sorted_tops_list[i - 1]
            prev_txt = " ".join(w['text'] for w in sorted(rows[prev_top], key=lambda w: w['x0'])).upper()
            if "COGNOME" in prev_txt:
                header_top = top
                break
    righe = []
    if header_top is None:
        return righe, entrata_x, uscita_x
    docenti_top = None
    for top in sorted(rows):
        txt = " ".join(w['text'] for w in sorted(rows[top], key=lambda w: w['x0'])).upper()
        if "DOCENTI" in txt and "TUTOR" in txt:
            docenti_top = top
            break
        if "ARGOMENTI" in txt and "ATTIVIT" in txt:
            docenti_top = top
            break
    for top in sorted(rows):
        if top <= header_top:
            continue
        if docenti_top is not None and top >= docenti_top:
            break
        line = sorted(rows[top], key=lambda w: w['x0'])
        txts = [w['text'] for w in line]
        if not txts or not re.fullmatch(r"\d{1,3}", txts[0]):
            continue
        num = int(txts[0])
        if not (1 <= num <= 200):
            continue
        nome_esistente_parts = []
        nome_x0 = None
        for w in line[1:]:
            if w['text'] in ("DOCENTI", "|", "TUTOR", "RUOLO", "FIRMA"):
                break
            if w['text'] in ("Tutor", "Docente", "Esperto"):
                break
            if nome_x0 is None and w['text'].strip():
                nome_x0 = w['x0']
            nome_esistente_parts.append(w['text'])
        nome_esistente = " ".join(nome_esistente_parts).strip()
        if nome_x0 is None:
            nome_x0 = 90
        righe.append({"numero": num, "top": top,
                      "nome_esistente": nome_esistente, "nome_x0": nome_x0})
    return righe, entrata_x, uscita_x


def estrai_righe_pagina_B_nuda(page):
    """Parser for B pages without header (numbers+names only)."""
    words = page.extract_words(use_text_flow=True)
    rows = {}
    for w in words:
        key = round(w['top'] / 2) * 2
        rows.setdefault(key, []).append(w)
    docenti_top = None
    for top in sorted(rows):
        txt = " ".join(w['text'] for w in sorted(rows[top], key=lambda w: w['x0'])).upper()
        if "DOCENTI" in txt and "TUTOR" in txt:
            docenti_top = top
            break
        if "ARGOMENTI" in txt and "ATTIVIT" in txt:
            docenti_top = top
            break
    righe = []
    for top in sorted(rows):
        if docenti_top is not None and top >= docenti_top:
            break
        line = sorted(rows[top], key=lambda w: w['x0'])
        txts = [w['text'] for w in line]
        if not txts or not re.fullmatch(r"\d{1,3}", txts[0]):
            continue
        num = int(txts[0])
        if not (1 <= num <= 200):
            continue
        nome_parts = []
        nome_x0 = None
        for w in line[1:]:
            if w['text'] in ("DOCENTI", "|", "TUTOR", "RUOLO", "FIRMA"):
                break
            if w['text'] in ("Tutor", "Docente", "Esperto"):
                break
            if nome_x0 is None and w['text'].strip():
                nome_x0 = w['x0']
            nome_parts.append(w['text'])
        nome = " ".join(nome_parts).strip()
        if nome_x0 is None:
            nome_x0 = 90
        righe.append({"numero": num, "top": top,
                      "nome_esistente": nome, "nome_x0": nome_x0})
    return righe, None, None


# ============================ PDF writing =================================
def decidi_presenza(nome_ref, excel_giornata, is_pomeriggio, inizio_pom_min):
    """'PRESENTE', 'ASSENTE' or None."""
    if not excel_giornata:
        return None
    allievi = excel_giornata["allievi"]
    nome_match = None
    nt = norm(nome_ref)
    for k in allievi:
        if norm(k) == nt:
            nome_match = k
            break
    if nome_match is None:
        best_k = None; best_s = 0
        for k in allievi:
            s = similarity(nt, norm(k))
            if s > best_s:
                best_s = s; best_k = k
            if best_s >= 0.95:
                break
        if best_s >= 0.80:
            nome_match = best_k
    if nome_match is None:
        return "ASSENTE"
    # Simplified rule: if the student is in the report -> always PRESENTE,
    # otherwise ASSENTE (handled above). No heuristics on times/exits.
    # (is_pomeriggio / inizio_pom_min kept in the signature for backward
    #  compatibility, but no longer used for the decision.)
    return "PRESENTE"


def compila_sessione(doc, sessione, righe_per_pagina, ref_list, excel_giornata):
    azioni = []
    pagine_allievi = [sessione["A"]] + sessione["B"]
    is_pomeriggio = False
    inizio_pom_min = None
    if sessione.get("orar"):
        m = re.match(r"^(\d{1,2}):(\d{2})-", sessione["orar"])
        if m:
            inizio_pom_min = int(m.group(1)) * 60 + int(m.group(2))
            is_pomeriggio = inizio_pom_min >= 12 * 60
    ref_entrata_x = ref_uscita_x = None
    if sessione["A"] in righe_per_pagina:
        _, ref_entrata_x, ref_uscita_x = righe_per_pagina[sessione["A"]]
    for pg in pagine_allievi:
        if pg not in righe_per_pagina:
            continue
        righe, entrata_x, uscita_x = righe_per_pagina[pg]
        if entrata_x is None:
            entrata_x = ref_entrata_x
        if uscita_x is None:
            uscita_x = ref_uscita_x
        page = doc[pg - 1]
        for riga in righe:
            num = riga["numero"]
            nome_ref = ref_list.get(num, "")
            if not nome_ref:
                continue
            if not riga["nome_esistente"].strip():
                page.insert_text((riga["nome_x0"], riga["top"]),
                                 nome_ref, fontsize=9, color=(0, 0, 0))
            esito = decidi_presenza(nome_ref, excel_giornata,
                                    is_pomeriggio, inizio_pom_min)
            if esito:
                if entrata_x:
                    page.insert_text((entrata_x, riga["top"]),
                                     esito, fontsize=9, color=(0, 0, 0))
                if uscita_x:
                    page.insert_text((uscita_x, riga["top"]),
                                     esito, fontsize=9, color=(0, 0, 0))
            azioni.append({"pagina": pg, "numero": num, "nome": nome_ref,
                           "esito": esito or "VUOTO"})
    return azioni


def compila_pdf(registro_pdf, sessioni, ref_list, excel_giornate, output_pdf):
    righe_per_pagina = {}
    with pdfplumber.open(registro_pdf) as pdf:
        for idx, page in enumerate(pdf.pages):
            righe, ex, ux = estrai_righe_allievi_da_pagina(page)
            if righe:
                righe_per_pagina[idx + 1] = (righe, ex, ux)
            else:
                righe_b, _, _ = estrai_righe_pagina_B_nuda(page)
                if righe_b:
                    righe_per_pagina[idx + 1] = (righe_b, None, None)
    doc = fitz.open(str(registro_pdf))
    report = {"sessioni": [], "present_totali": 0, "assente_totali": 0,
              "sessioni_senza_excel": []}
    by_data = {}
    for g in excel_giornate:
        if g["data"]:
            by_data.setdefault(g["data"], []).append(g)
    for s in sessioni:
        gx_list = by_data.get(s["data"], [])
        gx = gx_list[0] if gx_list else None
        azioni = compila_sessione(doc, s, righe_per_pagina, ref_list, gx)
        n_pres = sum(1 for a in azioni if a["esito"] == "PRESENTE")
        n_ass = sum(1 for a in azioni if a["esito"] == "ASSENTE")
        n_vuoto = sum(1 for a in azioni if a["esito"] == "VUOTO")
        report["present_totali"] += n_pres
        report["assente_totali"] += n_ass
        report["sessioni"].append({
            "data": s["data"], "orar": s["orar"],
            "pag_A": s["A"], "pag_B": s["B"], "pag_C": s["C"],
            "excel_usato": gx["fonte"] if gx else None,
            "present": n_pres, "assente": n_ass, "vuoto": n_vuoto,
        })
        if not gx:
            report["sessioni_senza_excel"].append(
                {"data": s["data"], "orar": s["orar"]})
    doc.save(str(output_pdf))
    doc.close()
    return report


# ============================ preview/opening ============================
def renderizza_anteprima(pdf_path, pagina_idx, png_path):
    try:
        doc = fitz.open(str(pdf_path))
        pix = doc[pagina_idx].get_pixmap(dpi=130)
        pix.save(str(png_path))
        doc.close()
        return True
    except Exception as e:
        print(f"  preview failed: {e}")
        return False


def apri_pdf(p):
    if sys.platform.startswith("win"):
        os.startfile(str(p))
    elif sys.platform == "darwin":
        subprocess.run(["open", str(p)])
    else:
        subprocess.run(["xdg-open", str(p)])


# ============================ main ========================================
def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    no_open = "--no-open" in flags
    no_preview = "--no-preview" in flags
    archive = "--archive" in flags

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    pdfs = sorted(INPUT_DIR.glob("*.pdf"))
    if not pdfs:
        print("ERROR: no PDF in input/"); sys.exit(2)
    registro_pdf = pdfs[0]
    output_pdf = OUTPUT_DIR / f"{registro_pdf.stem}_compilato.pdf"

    excel_paths = sorted(p for p in INPUT_DIR.glob("*.xlsx")
                         if "TEST" not in p.name.upper())
    if not excel_paths:
        print("ERROR: no Excel in input/"); sys.exit(2)

    print(f"[1/5] PDF analysis: {registro_pdf.name}")
    with pdfplumber.open(registro_pdf) as pdf:
        pagine_elenco = _trova_pagine_elenco(pdf)
        print(f"      ELENCO ALLIEVI pages: {[i+1 for i in pagine_elenco] or '(none!)'}")
        layout = rileva_layout(pdf, pagine_elenco)
        print(f"      detected layout: {layout}")
        print(f"      extracting students from the ELENCO ALLIEVI (primary source)...")
        allievi_pdf = estrai_allievi_elenco(pdf, layout, pagine_elenco)
        azienda_slug = rileva_azienda(allievi_pdf, pdf)

    ref_list = {num: info["nome"] for num, info in allievi_pdf.items()}
    print(f"      {len(ref_list)} students extracted from the PDF")
    print(f"      detected company: {azienda_slug}")

    if not ref_list:
        print("ERROR: no students extracted from the ELENCO ALLIEVI!")
        print("        (the PDF may not have an ELENCO ALLIEVI page,")
        print("         or the layout is not recognized). See the README.")
        sys.exit(3)

    print(f"[2/5] Company storage...")
    conosciuta, path_az = gestisci_storage_azienda(azienda_slug, allievi_pdf)
    if conosciuta:
        print(f"      company already known: {path_az.name}")
    else:
        print(f"      NEW company! storage created: {path_az.name}")

    print(f"[3/5] Reading Platform Excel files ({len(excel_paths)} files):")
    giornate = leggi_multi_excel(excel_paths)

    print(f"[4/5] PDF session analysis...")
    with pdfplumber.open(registro_pdf) as pdf:
        sessioni = raggruppa_sessioni(pdf)
    print(f"      {len(sessioni)} sessions found")

    print(f"[5/5] Filling PDF: {output_pdf.name}")
    rep = compila_pdf(registro_pdf, sessioni, ref_list, giornate, output_pdf)
    rep["layout"] = layout
    rep["azienda_rilevata"] = azienda_slug
    rep["allievi_estratti"] = {str(n): ref_list[n] for n in sorted(ref_list)}
    report_json = output_pdf.with_suffix(".report.json")
    report_json.write_text(json.dumps(rep, indent=2, ensure_ascii=False),
                           encoding="utf-8")

    if not no_preview and sessioni:
        png = output_pdf.with_suffix(".preview.png")
        renderizza_anteprima(output_pdf, sessioni[0]["A"] - 1, png)

    print()
    print("=" * 60)
    print("COMPLETED")
    print(f"  Layout:         {layout}")
    print(f"  Company:        {azienda_slug}"
          f"{' (NEW)' if not conosciuta else ''}")
    print(f"  Students:       {len(ref_list)}")
    print(f"  PDF:            {output_pdf}")
    print(f"  Report JSON:    {report_json}")
    print(f"  PRESENTE totals: {rep['present_totali']}")
    print(f"  ASSENTE totals:  {rep['assente_totali']}")
    if rep['sessioni_senza_excel']:
        print(f"  Sessions without Excel ({len(rep['sessioni_senza_excel'])}):")
        for s in rep['sessioni_senza_excel']:
            print(f"    - date={s['data']} time={s['orar']}")
    print("=" * 60)

    if archive and not args:
        archivia_input(BASE, registro_pdf, excel_paths)

    if not no_open:
        try:
            apri_pdf(output_pdf)
        except Exception as e:
            print(f"  failed to open: {e}")


def archivia_input(base, registro_pdf, excel_paths):
    from datetime import datetime
    import shutil
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    elab_dir = base / "elaborati" / ts
    elab_dir.mkdir(parents=True, exist_ok=True)
    spostati = []
    try:
        shutil.move(str(registro_pdf), str(elab_dir / registro_pdf.name))
        spostati.append(registro_pdf.name)
    except Exception as e:
        print(f"  (cannot move {registro_pdf.name}: {e})")
    for xp in excel_paths:
        try:
            shutil.move(str(xp), str(elab_dir / xp.name))
            spostati.append(xp.name)
        except Exception as e:
            print(f"  (cannot move {xp.name}: {e})")
    print(f"  Files archived in: elaborati/{ts}/ ({len(spostati)} files)")


if __name__ == "__main__":
    main()
