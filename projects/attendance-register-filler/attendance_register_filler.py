"""
attendance_register_filler.py - VERSIONE MULTI-AZIENDA / MULTI-LAYOUT
=============================================================
Compila il registro PDF delle presenze a partire dai report Excel Piattaforma.

Supporta piu' aziende e due layout di registro:
  - LAYOUT A: Cliente A (3 pagine per sessione: A=allievi 1-25,
              B=26-57, C=docenti+argomenti). ELENCO ALLIEVI con righe
              interlacciate.
  - LAYOUT B: corsi a riga unica (riga unica pulita nell'ELENCO
              ALLIEVI, orario in formato DALLE/ALLE).

PRINCIPIO ANTI-BUG: la fonte dei nomi e' SEMPRE la pagina ELENCO ALLIEVI
del PDF che stiamo compilando. Non si usa una cache fissa che potrebbe
contenere nomi di altra azienda. Lo storage per-azienda (docs/aziende/)
e' solo un archivio di verifica, mai la fonte attiva.

USO:
    py attendance_register_filler.py               -> input/*.pdf + input/*.xlsx -> output/
    py attendance_register_filler.py --no-open --no-preview
    py attendance_register_filler.py --archive     -> archivia input in elaborati/<ts>/

LOGICA:
1. Rileva il layout (A o B) dalla pagina ELENCO ALLIEVI.
2. Estrae gli allievi dall'ELENCO ALLIEVI del PDF corrente
   -> {numero: {nome, cf, data, azienda}}.
3. Rileva l'azienda dalla colonna AZIENDA (fallback: titolo PDF).
4. Se azienda nuova -> crea docs/aziende/<slug>.json (auto-apprendimento).
5. Per ogni sessione: data + orario -> trova Excel Piattaforma con stessa data ->
   per ogni allievo scrive NOME + PRESENTE/ASSENTE.

MATCHING: cognome+nome normalizzato + similarita' fuzzy (>=80%).
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
    """Slug per nome file: 'CLIENTE_A S.P.A.' -> 'datamanagement_italia'."""
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
REF_CACHE_LEGACY = DOCS_DIR / "reference_list_final.json"  # non piu' usata come fonte

CF_RE = re.compile(r"^[A-Z]{6}\d{2}[A-Z]\d{2}[A-Z]\d{3}[A-Z]$")


# ============================ ELENCO ALLIEVI: layout ======================
def _trova_pagine_elenco(pdf):
    """Restituisce gli indici (0-based) delle pagine ELENCO ALLIEVI.

    La pagina iniziale ha l'header 'ELENCO ALLIEVI'; l'elenco puo' continuare
    sulle pagine successive SENZA header ripetuto. Le pagine di continuazione
    si riconoscono perche' contengono codici fiscali (CF a 16 char), che non
    appaiono mai nelle pagine presenze. Ci fermiamo quando una pagina non ha
    CF (tipicamente la prima pagina presenze/argomenti)."""
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
    # estendi alle pagine successive finche' troviamo CF
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
    """Restituisce 'A' o 'B' in base alla struttura della pagina ELENCO ALLIEVI.

    LAYOUT A (Cliente A): righe interlacciate - il numero (N) e' su una
      riga separata (y) rispetto a cognome/nome (y+/-6). Inoltre i campi
      CF/DATA/AZIENDA sono molto vicini verticalmente al cognome.
    LAYOUT B (corsi a riga unica): riga unica pulita -
      '1 ASCOLI ASCOLI CF DATA AZIENDA' tutto sulla stessa y.

    Euristica robusta: raggruppiamo le parole per riga (toleranza 2px) e
    contiamo quante righe iniziano con un numero (1-100) seguito da testo
    alfabetico SULLA STESSA riga. Layout B -> tante righe cosi'; Layout A ->
      poche (perche' numero e nome sono su y diverse).
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
        # riga che inizia con numero e ha subito dopo un cognome (parola alpha)
        if re.fullmatch(r"\d{1,3}", txts[0]) and len(txts) >= 2 and \
                re.match(r"^[A-Z]+$", txts[1]):
            righe_complete += 1
    # Layout B: molte righe complete (>5). Layout A: poche.
    return "B" if righe_complete >= 5 else "A"


# Parole tipiche di ragione sociale che, appena incontrate nella riga
# ELENCO ALLIEVI (layout B), chiudono il campo "nome" dell'allievo e fanno
# iniziare il campo "azienda". Evita che nomi vengano inquinati quando il CF
# o la data non vengono riconosciuti dal parser (es. "ROSSI MARIA
# INFORMATION AND" invece di "ROSSI MARIA", o "BIANCHI LUCA I.CON."
# invece di "BIANCHI LUCA").
AZIENDA_WORDS = {
    "INFORMATION", "AND", "STARTUP", "START", "UP", "COSTITUITA",
    "COSTITUZIONE", "SOCIETA", "SOC", "SRL", "SPA", "SNC", "SAS",
    "TECNOLOGIA", "TECNOLOGIE", "DIGITAL", "SERVIZI", "SISTEMI",
    "SOLUTIONS", "SOLUTION", "GROUP", "CONSULENZE", "INFORMATICA",
    "INGEGNERIA", "IMPRESA", "AZIENDA", "HOLDING", "COOP", "COOPERATIVA",
    "CORSO_A", "CORSO_B", "CORSO_C", "CORSO_D",
}


def _is_azienda_word(t):
    """True se il token e' una parola tipica di ragione sociale.

    IMPORTANTE: i punti vanno rimossi SENZA sostituirli con spazio, altrimenti
    'X.Y.Z.' diventerebbe 'X Y Z' (piu' parole) e non matcherebbe 'XYZ'.
    Quindi qui si fa una normalizzazione "tight" (punti/apostrofi rimossi),
    diversa da norm() che li trasforma in spazio.
    """
    s = t.upper()
    # togli punti/apostrofi/trattini senza lasciare spazio: X.Y.Z. -> XYZ
    s = re.sub(r"[.\'\-]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    # prima prova l'intero token (es. S.P.A.->SPA)
    if s in AZIENDA_WORDS:
        return True
    # poi prova le singole parole (es. START-UP -> 'START UP' -> match 'START')
    for w in s.split():
        if w in AZIENDA_WORDS:
            return True
    return False


def _estrai_riga_elenco_B(page):
    """Parser LAYOUT B: ogni riga = 'N COGNOME NOME CF DATA AZIENDA' su una y.
    Restituisce {numero: {nome, cf, data, azienda}}."""
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
        # raccogli parole dopo il numero. La fase avanza cosi':
        #   "nome" -> (CF rilevato) -> "after_cf" -> (data) -> "after_data"
        #   "nome" -> (parola-azienda) -> "after_azienda"  [filtro anti-inquinamento]
        # In ogni fase successiva a "nome", i token residui (che non siano CF
        # o data) vanno nell'azienda: cosi' "INFORMATION AND ..." non lascia
        # la parola "AND" nel nome.
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
                # dopo il CF ma prima della data (raro): ignora
                continue
            # Filtro anti-inquinamento: una parola-azienda chiude il nome.
            if phase == "nome" and _is_azienda_word(t):
                phase = "after_azienda"
                azienda = (azienda + " " + t).strip()
                continue
            # In qualsiasi fase "dopo" (after_data / after_azienda): tutto va
            # in azienda, niente finisce piu' nel nome.
            if phase in ("after_data", "after_azienda"):
                azienda = (azienda + " " + t).strip()
                continue
            parts.append(t)
        nome = " ".join(parts).strip()
        # Cap backstop: massimo 3 parole (2 spazi) nel nome. Rete di sicurezza
        # per ragioni sociali non coperte da AZIENDA_WORDS. Mantiene cognomi
        # composti (DI PIETRO ALESSANDRO = 3 parole) e nomi doppi.
        nome_words = nome.split()
        if len(nome_words) > 3:
            nome = " ".join(nome_words[:3])
        if nome and any(c.isalpha() for c in nome):
            out[num] = {"nome": nome, "cf": cf, "data": data, "azienda": azienda}
    return out


NOME_WORD_RE = re.compile(r"^[A-Za-z][A-Za-z'\u00C0-\u024F]+$")

# parole dell'header ELENCO ALLIEVI da escludere dal nome (case-insensitive)
HEADER_WORDS = {
    "COGNOME", "NOME", "E", "DELL'ALLIEVO", "DELLALLIEVO",
    "CF", "DATA", "AZIENDA", "N",
    "VIDIMATO", "DA", "PAGINA", "DI",
}


def _is_nome_valido(t):
    """True se la parola e' un componente plausibile di un nome (non header)."""
    if t.upper() in HEADER_WORDS:
        return False
    # particelle nobiliari/locuzioni nei cognomi italiani
    if t in ("De", "Di", "Da", "La", "Lo", "Del", "Della", "Delle", "Dello"):
        return True
    return bool(NOME_WORD_RE.match(t))


def _estrai_riga_elenco_A(page):
    """Parser LAYOUT A (Cliente A): righe interlacciate.
    Il CF, il cognome+nome e il numero sono su y VICINE ma non sempre
    identiche (delta fino a +/-12px). Esempio:
        y124: CF@x236 CLIENTE_A@x413
        y126: Battoli@x96 Sigfrido@x127          <- nome su riga diversa dal CF
        y130: 1@x77                               <- numero su un'altra riga

    Strategia "banda logica": raggruppiamo le parole per bande orizzontali
    larghe (toleranza ~30px) in modo che CF, nome e numero della stessa
    persona cadano nella stessa banda. Da ogni banda estraiamo:
      - numero  = intero 1-100 con x < 95
      - CF      = token che matcha CF_RE
      - nome    = parole alfabetiche (non header) con 90 < x < 200
      - azienda = parole alfabetiche con x > x_CF
    """
    words = page.extract_words()
    bande = {}
    for w in words:
        key = round(w['top'] / 30) * 30
        bande.setdefault(key, []).append(w)
    X_NUM_MAX = 95
    X_NOME_MAX = 200   # nome sta a x ~93-153; CF inizia ~233
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
        # numero mancante: cerca in banda adiacente
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
    """Estrae {numero: {nome, cf, data, azienda}} da tutte le pagine ELENCO.
    Usa voto a maggioranza tra pagine multiple se presenti."""
    if not pagine_elenco_idx:
        return {}
    parser = _estrai_riga_elenco_A if layout == "A" else _estrai_riga_elenco_B
    nomi = {}  # numero -> Counter() su dict-str
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


# ============================ rilevazione azienda =========================
# keyword note nel titolo -> slug azienda (fallback se colonna AZIENDA vuota)
AZIENDE_NOTE = {
    "CLIENTE_A": "cliente_a",
    "CORSO_X": "corso_x",
    "CORSO_Y": "corso_y",
    "CORSO_A": "corso_a",
    "CORSO_B": "corso_b",
}


def rileva_azienda(allievi_estratti, pdf):
    """Rileva l'azienda dalla colonna AZIENDA dell'elenco (maggioranza).
    Fallback: keyword nel titolo (prime 3 pagine)."""
    # 1) maggioranza sulla colonna azienda
    c = Counter()
    for info in allievi_estratti.values():
        az = info.get("azienda", "").strip()
        if az:
            c[slugify(az)] += 1
    if c:
        return c.most_common(1)[0][0]
    # 2) fallback titolo
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
    """Se l'azienda e' nuova, crea docs/aziende/<slug>.json.
    Restituisce (azienda_conosciuta: bool, path_storage)."""
    AZIENDE_DIR.mkdir(parents=True, exist_ok=True)
    path = AZIENDE_DIR / f"{slug}.json"
    conosciuta = path.exists()
    if not conosciuta:
        # salva i nomi estratti dal PDF (auto-apprendimento)
        data = {
            "azienda_slug": slug,
            "allievi": {str(n): info for n, info in sorted(allievi_estratti.items())},
        }
        path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    return conosciuta, path


# ============================ parsing Excel ===============================
def leggi_allievi_excel(path_excel: Path) -> dict:
    """{nome_chiave_normalizzato: ultima_disconnessione_str} per allievi
    (no docenti/tutor)."""
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
            print(f"  {p.name}: data={data} orario={oi}-{of}  "
                  f"allievi={len(allievi)}")
        except Exception as e:
            print(f"  ERRORE {p.name}: {e}")
    return out


# ============================ parsing PDF: sessioni =======================
def info_pagina(page):
    text = page.extract_text() or ""
    ha_header = "REGISTRO PRESENZE ALLIEVI" in text
    m = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    data = m.group(1) if m else None
    # orario: layout A (8:30-13:30) o layout B (DALLE:.. ALLE:..)
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
    """[(numero, top_y, nome_esistente, x_entrata, x_uscita)] per una pagina
    presenze. Valido per entrambi i layout (entrambi hanno COGNOME/ENTRATA/USCITA)."""
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
    """Parser per pagine B senza header (solo numeri+nomi)."""
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


# ============================ scrittura PDF ===============================
def decidi_presenza(nome_ref, excel_giornata, is_pomeriggio, inizio_pom_min):
    """'PRESENTE', 'ASSENTE' o None."""
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
    # Regola semplificata: se l'allievo e' nel report -> sempre PRESENTE,
    # altrimenti ASSENTE (gestito sopra). Nessuna euristica su orari/uscite.
    # (is_pomeriggio / inizio_pom_min mantenuti nella firma per retrocompat,
    #  ma non piu' usati per la decisione.)
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


# ============================ anteprima/apertura ==========================
def renderizza_anteprima(pdf_path, pagina_idx, png_path):
    try:
        doc = fitz.open(str(pdf_path))
        pix = doc[pagina_idx].get_pixmap(dpi=130)
        pix.save(str(png_path))
        doc.close()
        return True
    except Exception as e:
        print(f"  anteprima fallita: {e}")
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
        print("ERRORE: nessun PDF in input/"); sys.exit(2)
    registro_pdf = pdfs[0]
    output_pdf = OUTPUT_DIR / f"{registro_pdf.stem}_compilato.pdf"

    excel_paths = sorted(p for p in INPUT_DIR.glob("*.xlsx")
                         if "TEST" not in p.name.upper())
    if not excel_paths:
        print("ERRORE: nessun Excel in input/"); sys.exit(2)

    print(f"[1/5] Analisi PDF: {registro_pdf.name}")
    with pdfplumber.open(registro_pdf) as pdf:
        pagine_elenco = _trova_pagine_elenco(pdf)
        print(f"      pagine ELENCO ALLIEVI: {[i+1 for i in pagine_elenco] or '(nessuna!)'}")
        layout = rileva_layout(pdf, pagine_elenco)
        print(f"      layout rilevato: {layout}")
        print(f"      estrazione allievi dall'ELENCO ALLIEVI (fonte primaria)...")
        allievi_pdf = estrai_allievi_elenco(pdf, layout, pagine_elenco)
        azienda_slug = rileva_azienda(allievi_pdf, pdf)

    ref_list = {num: info["nome"] for num, info in allievi_pdf.items()}
    print(f"      {len(ref_list)} allievi estratti dal PDF")
    print(f"      azienda rilevata: {azienda_slug}")

    if not ref_list:
        print("ERRORE: nessun allievo estratto dall'ELENCO ALLIEVI!")
        print("        (il PDF potrebbe non avere la pagina ELENCO ALLIEVI,")
        print("         o il layout non e' riconosciuto). Vedi il README.")
        sys.exit(3)

    print(f"[2/5] Storage aziende...")
    conosciuta, path_az = gestisci_storage_azienda(azienda_slug, allievi_pdf)
    if conosciuta:
        print(f"      azienda gia' nota: {path_az.name}")
    else:
        print(f"      NUOVA azienda! creato storage: {path_az.name}")

    print(f"[3/5] Lettura Excel Piattaforma ({len(excel_paths)} file):")
    giornate = leggi_multi_excel(excel_paths)

    print(f"[4/5] Analisi sessioni PDF...")
    with pdfplumber.open(registro_pdf) as pdf:
        sessioni = raggruppa_sessioni(pdf)
    print(f"      {len(sessioni)} sessioni trovate")

    print(f"[5/5] Compilazione PDF: {output_pdf.name}")
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
    print("COMPLETATO")
    print(f"  Layout:         {layout}")
    print(f"  Azienda:        {azienda_slug}"
          f"{' (NUOVA)' if not conosciuta else ''}")
    print(f"  Allievi:        {len(ref_list)}")
    print(f"  PDF:            {output_pdf}")
    print(f"  Report JSON:    {report_json}")
    print(f"  PRESENTE totali: {rep['present_totali']}")
    print(f"  ASSENTE totali:  {rep['assente_totali']}")
    if rep['sessioni_senza_excel']:
        print(f"  Sessioni senza Excel ({len(rep['sessioni_senza_excel'])}):")
        for s in rep['sessioni_senza_excel']:
            print(f"    - data={s['data']} orar={s['orar']}")
    print("=" * 60)

    if archive and not args:
        archivia_input(BASE, registro_pdf, excel_paths)

    if not no_open:
        try:
            apri_pdf(output_pdf)
        except Exception as e:
            print(f"  apertura fallita: {e}")


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
        print(f"  (impossibile spostare {registro_pdf.name}: {e})")
    for xp in excel_paths:
        try:
            shutil.move(str(xp), str(elab_dir / xp.name))
            spostati.append(xp.name)
        except Exception as e:
            print(f"  (impossibile spostare {xp.name}: {e})")
    print(f"  File archiviati in: elaborati/{ts}/ ({len(spostati)} file)")


if __name__ == "__main__":
    main()
