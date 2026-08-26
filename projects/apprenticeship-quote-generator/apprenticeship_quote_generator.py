# -*- coding: utf-8 -*-
"""
apprenticeship_quote_generator.py
-------------------------------
Automatic generator of ECONOMIC OFFERS ("OFFERTE ECONOMICHE") for private
apprenticeship courses.

How it works:
  - Uses as its base a faithful copy of the sample document (template_offerta_apprendistato.docx),
    preserving logo, header, footer, Arial font and original formatting.
  - Asks interactively (keyboard input, VS Code / terminal):
      1. whether the calendar has been agreed upon (if NO, the calendar Excel file is required);
      2. the number of participants (basis of the calculation);
      3. the apprenticeship year ("annualita'", first / second);
      4. the recipient company's details.
  - Calculates:
      * Total amount      = 480.00 €  x  n. participants
      * Final price ("PREZZO MIGLIOR FAVORE") = total x 0.90 rounded to the nearest hundred
  - Fills in the OFFER NUMBER ("NUMERO OFFERTA") with an automatic progressive
    counter, separate for each apprenticeship year (stored in contatore_offerte.json):
      * first year  starts at 20260044 and increments by 1 with each offer;
      * second year starts at 20260066 and increments by 1 with each offer.
  - Sets the date to the day the script is run.
  - If the calendar is NOT agreed upon, reads the calendar Excel file:
      * extracts the edition code (CORSO 07 / Edizione 07) from the header and propagates it into the document;
      * rebuilds a 6-column Word table (Data | Orario | Ore |
        Modulo | Docente | Modalità) under the "Calendario" heading, as in the
        template "template_offerta_apprendistato con calendario.docx".
  - The apprenticeship year (first/second) is updated both in the course subtitle
    and in the description row of the cost table.
  - Saves the offer as output/Offerta_Apprendistato_<RagioneSociale>_<YYYYMMDD>.docx.

The original "template_offerta_apprendistato.docx" is never touched: work always
happens on a copy of the template.
"""

import os
import sys
import re
from datetime import date, datetime

import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
from openpyxl import load_workbook


# =========================================================================
# CONSTANTS
# =========================================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(SCRIPT_DIR, "template_offerta_apprendistato.docx")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")

UNITARIO = 480.0          # unit amount per participant (fixed)
SCONTO = 0.10             # 10% discount applied to get the final price

# Calendar columns: 6 full columns (as in the "con calendario" template).
HEADERS_CALENDARIO = ["Data", "Orario", "Ore", "Modulo", "Docente", "Modalità"]

SEGNAPOSTO_VUOTO = "______________"   # for company fields left empty

# In the original offer the subtitle uses "CORSO5" (number without leading zero)
# and the cost description "Ed. 05" (2-digit zero-padded number).
EDIZIONE_DEFAULT = 5      # fallback if not readable from the calendar

# --- Progressive offer numbering -------------------------------------------
# Local file storing the last offer number used for each apprenticeship year.
# The first year starts at 20260044, the second at 20260066;
# each new offer produced increments the counter of its year by 1.
CONTATORE_FILE = os.path.join(SCRIPT_DIR, "contatore_offerte.json")
CONTATORI_BASE = {1: 20260046, 2: 20260067}   # starting point for year 1 / 2


# =========================================================================
# OFFER NUMBER COUNTER (persistent on a local file)
# =========================================================================
def leggi_contatori():
    """
    Loads the JSON file with the LAST ASSIGNED offer number per apprenticeship year.
    Returns a dict {1: <int>, 2: <int>}. If the file does not exist or is invalid,
    initializes each counter to (BASE - 1), so that the next number to assign
    is exactly the BASE (first year 20260044, second year 20260066).
    """
    out = {k: v - 1 for k, v in CONTATORI_BASE.items()}   # starting at base-1
    if os.path.isfile(CONTATORE_FILE):
        try:
            import json
            with open(CONTATORE_FILE, "r", encoding="utf-8") as f:
                dati = json.load(f)
            for k in (1, 2):
                v = dati.get(str(k))
                # accept only valid values (>= base-1)
                if isinstance(v, int) and v >= CONTATORI_BASE[k] - 1:
                    out[k] = v
        except Exception:
            # corrupt file: ignore and use base-1
            pass
    return out


def prossimo_numero_offerta(annualita_num, contatori=None):
    """
    Returns the next offer number to assign for the given apprenticeship
    year (1 or 2): always (last_assigned + 1).
    On the first run (no offer produced yet) the counter equals BASE-1,
    so the first assigned number will be the BASE
    (20260044 for the first year, 20260066 for the second).
    """
    if contatori is None:
        contatori = leggi_contatori()
    return contatori.get(annualita_num, CONTATORI_BASE[annualita_num] - 1) + 1


def aggiorna_contatore(annualita_num, numero_assegnato):
    """
    Records the offer number just assigned for the given apprenticeship year,
    if it is greater than the last stored one (this way the maximum is kept
    even after re-runs with a manually modified number). Saves to a local JSON file.
    """
    import json
    contatori = leggi_contatori()
    if numero_assegnato > contatori.get(annualita_num, CONTATORI_BASE[annualita_num] - 1):
        contatori[annualita_num] = numero_assegnato
        try:
            with open(CONTATORE_FILE, "w", encoding="utf-8") as f:
                # save with string keys (JSON standard)
                json.dump({str(k): v for k, v in contatori.items()}, f,
                          indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"  WARNING: could not save the offer counter ({e}).")


# =========================================================================
# FORMATTING UTILITIES
# =========================================================================
def euro(value):
    """Formats a number as an Italian-style euro amount: 5760 -> '5.760,00 €'."""
    s = f"{value:,.2f}"          # 5,760.00
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")  # 5.760,00
    return s + " €"


def calc_totali(n_partecipanti):
    """Returns (total, final) given the number of participants."""
    totale = UNITARIO * n_partecipanti
    scontato = totale * (1 - SCONTO)
    # round to the nearest hundred
    finale = round(scontato / 100.0) * 100.0
    return totale, finale


def codice_corso(edizione):
    """Returns the CORSO code with no leading zero: 7 -> 'CORSO7'."""
    return f"CORSO{int(edizione)}"


def codice_edizione_padded(edizione):
    """Returns the zero-padded 2-digit edition: 7 -> 'Ed. 07'."""
    return f"Ed. {int(edizione):02d}"


def _parse_data(valore):
    """
    Converts a date value (datetime, date, or a string in various formats) to 'dd/mm/yyyy'.
    Handles datetimes produced by Excel (e.g. datetime(2026,9,14)).
    """
    if valore is None or (isinstance(valore, str) and not valore.strip()):
        return ""
    if isinstance(valore, datetime):
        return valore.strftime("%d/%m/%Y")
    if isinstance(valore, date):
        return valore.strftime("%d/%m/%Y")
    s = str(valore).strip()
    # try common formats
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s  # fallback: keep the original text


# =========================================================================
# INTERACTIVE INPUT
# =========================================================================
def _chiedi(prompt, default=None):
    """Wrapper around input() with an optional default shown in square brackets."""
    if default is not None:
        full = f"{prompt} [{default}]: "
    else:
        full = f"{prompt}: "
    val = input(full).strip()
    if val == "" and default is not None:
        return default
    return val


def _si_no(prompt, default="s"):
    """Asks a yes/no question. Returns True for yes, False for no."""
    while True:
        v = _chiedi(prompt + " (y/n)", default=default).lower()
        if v in ("s", "si", "y", "yes"):
            return True
        if v in ("n", "no"):
            return False
        print("  Answer with 'y' or 'n'.")


def chiedi_calendario():
    """
    Handles the calendar block.
    Returns a dict with:
      - "concordato": True/False
      - "righe":      list[dict] (Data/Orario/Ore/Modulo/Docente/Modalità) or None
      - "edizione":   int (edition number read from the calendar, or EDIZIONE_DEFAULT)
    """
    print("\n--- CALENDAR ---")
    concordato = _si_no("Has the calendar been agreed upon?", default="s")
    if concordato:
        print("  >> Calendar agreed upon: the document will keep the original text.")
        return {"concordato": True, "righe": None, "edizione": EDIZIONE_DEFAULT}

    print("  >> Calendar NOT agreed upon: the calendar Excel file must be loaded.")
    while True:
        path = _chiedi("Path of the calendar Excel file").strip().strip('"')
        if not path:
            print("  Invalid path, try again.")
            continue
        risultato = leggi_calendario_excel(path)
        if risultato is not None:
            righe, edizione = risultato
            print(f"  >> Read {len(righe)} session days from the calendar. Detected edition: CORSO {edizione}.")
            return {"concordato": False, "righe": righe, "edizione": edizione}
        riprova = _si_no("Do you want to try another file?", default="s")
        if not riprova:
            print("  >> No calendar available. The original text will be left in place.")
            return {"concordato": True, "righe": None, "edizione": EDIZIONE_DEFAULT}


def leggi_calendario_excel(path):
    """
    Reads the calendar Excel file (real CORSO07 format):
      - header row with the edition title ("... Edizione 07 - CORSO 07 ...")
      - column header row: Data | Orario | Ore | Docente | Modulo | Modalita'
      - data rows (dates as datetime)
      - possible total-hours row to exclude
    Searches for the columns by name in a flexible way.
    Returns (righe, edizione) or (None, None) in case of error.
      - righe: list[dict] with keys Data, Orario, Ore, Modulo, Docente, Modalità
      - edizione: int
    """
    if not os.path.isfile(path):
        print(f"  ERROR: the file '{path}' does not exist.")
        return None, None
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        print(f"  ERROR: could not open the Excel file ({e}).")
        return None, None

    # expected columns (all lower-case for matching)
    target = [h.lower() for h in HEADERS_CALENDARIO]   # data, orario, ore, modulo, docente, modalità
    ESSENZIALI = ["data"]   # at least Data is needed to recognize the table

    def cerca_in_foglio(ws):
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None
        # extract the edition from all the text in the sheet
        edizione = _estrai_edizione(rows)
        # find the header row containing the 'data' column
        for r_idx, row in enumerate(rows):
            cells = [("" if c is None else str(c)).strip().lower() for c in row]
            # map each target to the corresponding column (if present)
            col_map = {}
            for t in target:
                trovato = None
                for ci, c in enumerate(cells):
                    if c and (c == t or c.startswith(t) or t.startswith(c)):
                        trovato = ci
                        break
                if trovato is not None:
                    col_map[t] = trovato
            # the essential 'data' column is required
            if not all(t in col_map for t in ESSENZIALI):
                continue
            # collect the data rows
            out = []
            for data_row in rows[r_idx + 1:]:
                def get(t):
                    idx = col_map.get(t)
                    if idx is None:
                        return None
                    return data_row[idx] if idx < len(data_row) else None
                data_val = get("data")
                # skip rows without a valid date (e.g. total-hours row)
                if not _ha_data_valida(data_val):
                    continue
                riga = {
                    "Data": _parse_data(data_val),
                    "Orario": _str(get("orario")),
                    "Ore": _str(get("ore")),
                    "Modulo": _str(get("modulo")),
                    "Docente": _str(get("docente")),
                    "Modalità": _str(get("modalità")),
                }
                # skip completely empty rows
                if any(v for v in riga.values()):
                    out.append(riga)
            return out, edizione
        return None

    # try all sheets
    for ws in wb.worksheets:
        res = cerca_in_foglio(ws)
        if res:
            wb.close()
            return res
    wb.close()
    print("  ERROR: could not find the 'Data' column in the calendar file.")
    return None, None


def _str(v):
    """Converts an Excel value into a clean string (None -> '')."""
    if v is None:
        return ""
    s = str(v).strip()
    return s


def _ha_data_valida(v):
    """True if the value represents a valid date (datetime/date or date string)."""
    if v is None:
        return False
    if isinstance(v, (datetime, date)):
        return True
    if isinstance(v, (int, float)):
        # numbers are not dates here
        return False
    s = str(v).strip()
    if not s:
        return False
    # try to parse
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            datetime.strptime(s, fmt)
            return True
        except ValueError:
            continue
    return False


def _estrai_edizione(rows):
    """
    Searches for the edition/CORSO number in all the text of the sheet.
    Typical patterns: 'Edizione 07', 'CORSO 07', 'CORSO7', 'Ed. 7'.
    Returns int (or EDIZIONE_DEFAULT if not found).
    """
    testo = " ".join(
        "" if c is None else str(c) for row in rows for c in row
    )
    # look for "edizione XX" or "CORSO XX" (with optional spaces/zeros)
    for pat in (r"edizione\s*0*(\d{1,3})", r"corso\s*0*(\d{1,3})", r"ed\.?\s*0*(\d{1,3})"):
        m = re.search(pat, testo, re.IGNORECASE)
        if m:
            try:
                n = int(m.group(1))
                if 1 <= n <= 999:
                    return n
            except ValueError:
                pass
    return EDIZIONE_DEFAULT


def chiedi_dati():
    """Asks for the main data (participants, apprenticeship year) and the company details."""
    print("\n--- COURSE DATA ---")
    # Number of participants
    while True:
        v = _chiedi("Number of participants (employees involved)")
        try:
            n = int(v)
            if n > 0:
                break
        except ValueError:
            pass
        print("  Enter a positive integer.")

    # Apprenticeship year (ann_testo values "prima"/"seconda")
    print("  Year: 1 = first year (default), 2 = second year")
    while True:
        a = _chiedi("Year (1/2)", default="1")
        if a in ("1", "2"):
            ann_num = int(a)                            # 1 or 2 (for the offer counter)
            ann_circ = "I°" if a == "1" else "II°"      # for the subtitle
            ann_desc = "I" if a == "1" else "II"        # for the cost description
            ann_testo = "prima" if a == "1" else "seconda"
            break
        print("  Answer with 1 or 2.")

    print("\n--- RECIPIENT COMPANY DATA ---")
    print("  (leave empty to keep the placeholder ______________)")
    ragione = _chiedi("Company name / Ragione sociale (e.g. Azienda Esempio srl)") or SEGNAPOSTO_VUOTO
    via = _chiedi("Street and number (e.g. Via Emilio Ghione, 12)") or SEGNAPOSTO_VUOTO
    cap = _chiedi("Postal code / CAP (e.g. 00128)") or SEGNAPOSTO_VUOTO
    citta = _chiedi("City (e.g. Roma)") or SEGNAPOSTO_VUOTO
    prov = _chiedi("Province code (e.g. RM)") or SEGNAPOSTO_VUOTO
    telefono = _chiedi("Phone/Office (e.g. +39 0669331247)") or SEGNAPOSTO_VUOTO
    piva = _chiedi("VAT number / Partita IVA (e.g. 16823051004)") or SEGNAPOSTO_VUOTO
    cf = _chiedi("Tax code / Codice Fiscale (e.g. 16823051004)") or SEGNAPOSTO_VUOTO
    pec = _chiedi("PEC (certified e-mail)") or SEGNAPOSTO_VUOTO
    sdi = _chiedi("Electronic invoicing identifier code (SDI)") or SEGNAPOSTO_VUOTO

    return {
        "n_partecipanti": n,
        "ann_num": ann_num,
        "ann_circ": ann_circ,
        "ann_desc": ann_desc,
        "ann_testo": ann_testo,
        "ragione": ragione,
        "via": via,
        "cap": cap,
        "citta": citta,
        "prov": prov,
        "telefono": telefono,
        "piva": piva,
        "cf": cf,
        "pec": pec,
        "sdi": sdi,
    }


# =========================================================================
# DOCUMENT MANIPULATION
# =========================================================================
def imposta_testo_paragrafo(paragrafo, nuovo_testo):
    """
    Replaces the text of a paragraph keeping the formatting of the FIRST run.
    All runs are emptied except the first, which receives the new text.
    """
    if not paragrafo.runs:
        paragrafo.add_run(nuovo_testo)
        return
    paragrafo.runs[0].text = nuovo_testo
    for r in paragrafo.runs[1:]:
        r.text = ""


def imposta_testo_cella(cell, nuovo_testo):
    """
    Replaces the text of a cell keeping the formatting of the first run
    of the first non-empty paragraph. The other paragraphs/runs are emptied.
    """
    paragrafi = cell.paragraphs
    if not paragrafi:
        cell.text = nuovo_testo
        return
    principale = None
    for p in paragrafi:
        if p.text.strip():
            principale = p
            break
    if principale is None:
        principale = paragrafi[0]
    if principale.runs:
        principale.runs[0].text = nuovo_testo
        for r in principale.runs[1:]:
            r.text = ""
    else:
        principale.add_run(nuovo_testo)
    # empty the other paragraphs
    for p in paragrafi:
        if p is principale:
            continue
        for r in p.runs:
            r.text = ""


def compila_destinatario(doc, dati):
    """Rewrites the recipient block (paragraphs 0-7, right-aligned)."""
    p = doc.paragraphs
    imposta_testo_paragrafo(p[0], f"Spett.le {dati['ragione']}")
    imposta_testo_paragrafo(p[1], f"{dati['via']}")
    imposta_testo_paragrafo(p[2], f"CAP {dati['cap']} - {dati['citta']} ({dati['prov']})")
    imposta_testo_paragrafo(p[3], f"Uff. {dati['telefono']}")
    imposta_testo_paragrafo(p[4], f"\xa0Partita IVA\xa0 {dati['piva']}\xa0")
    imposta_testo_paragrafo(p[5], f"Codice Fiscale {dati['cf']}\xa0")
    imposta_testo_paragrafo(p[6], f"PEC: {dati['pec']}")
    imposta_testo_paragrafo(p[7], f"Codice identificativo per fatturazione elettronica: {dati['sdi']}")


def compila_titolo_data_annualita_edizione(doc, dati, edizione, numero_offerta):
    """
    - Offer number: filled with the progressive counter (e.g. 20260044 / 20260066 ...)
    - Today's date (del <gg/mm/aaaa>)
    - Subtitle: apprenticeship year (I°/II°) + edition code (CORSOX)
      E.g.: 'Apprendistato I° annualita' CORSO7'
    """
    p = doc.paragraphs
    # PARA10: offer number
    imposta_testo_paragrafo(p[10], f"OFFERTA ECONOMICA n. {numero_offerta}")
    # PARA11: today's date
    oggi = date.today().strftime("%d/%m/%Y")
    imposta_testo_paragrafo(p[11], f"\xa0del {oggi}\xa0")
    # PARA13: subtitle -> 'Apprendistato <I°/II°> annualita' CORSO<edizione>'
    nuovo_sottotitolo = f"Apprendistato {dati['ann_circ']} annualità {codice_corso(edizione)}"
    imposta_testo_paragrafo(p[13], nuovo_sottotitolo)


def compila_tabella_costi(doc, dati, totale, finale, edizione):
    """
    Fills in the Costi table (3x4):
      - description (cell 1,0): updates the year (I/II) and the edition (Ed. 0X)
      - unit amount (1,1), n.participants (1,2), total (1,3), final (2,3)
    """
    t = doc.tables[0]
    # cell[1,0] description: "Competenze di base e trasversale <I/II> Annualita' Ed. 0X"
    nuova_desc = f"Competenze di base e trasversale {dati['ann_desc']} Annualità {codice_edizione_padded(edizione)}"
    imposta_testo_cella(t.cell(1, 0), nuova_desc)
    # cell[1,1] = 480,00 € (unit amount, fixed)
    imposta_testo_cella(t.cell(1, 1), f"\xa0\n{euro(UNITARIO)}\xa0")
    # cell[1,2] = n. participants
    imposta_testo_cella(t.cell(1, 2), f"\xa0\n{dati['n_partecipanti']}\n\n\xa0")
    # cell[1,3] = calculated total
    imposta_testo_cella(t.cell(1, 3), f"\xa0\n{euro(totale).replace(' €', '€')}\xa0")
    # cell[2,3] = final (PREZZO MIGLIOR FAVORE) - bold as in the original
    imposta_testo_cella(t.cell(2, 3), euro(finale).replace(" €", "€"))


# ---------------------------------------------------------------------------
# Calendar table (inserted as a new table in the document body)
# ---------------------------------------------------------------------------
def _set_cella_tabella(cell, testo, bold=False, font_name="Arial", size_pt=11):
    """Sets the text of a new table cell, centered Arial font."""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = 1  # CENTER
    run = p.add_run(testo)
    run.bold = bold
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), font_name)
    rfonts.set(qn('w:hAnsi'), font_name)
    rfonts.set(qn('w:cs'), font_name)


def _bordi_tabella(tabella):
    """Adds borders to all cells of a table."""
    tbl = tabella._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{border_name}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), '000000')
        tblBorders.append(b)
    tblPr.append(tblBorders)


def inserisci_tabella_calendario(doc, righe_calendario):
    """
    Removes the paragraph 'Come gia' concordato.' and inserts right after
    the 'Calendario' heading a table with the session days, 6 columns:
    Data | Orario | Ore | Modulo | Docente | Modalità (as in the template
    'template_offerta_apprendistato con calendario.docx').
    """
    corpo = doc.element.body

    par_calendario = None
    par_concordato = None
    for par in doc.paragraphs:
        t = par.text.strip().lower()
        if t.startswith("calendario"):
            par_calendario = par
        elif t.startswith("come già concordato") or t.startswith("come gia concordato"):
            par_concordato = par

    n_righe = len(righe_calendario) + 1  # +1 header row
    n_colonne = len(HEADERS_CALENDARIO)  # 6 columns
    tabella = doc.add_table(rows=n_righe, cols=n_colonne)
    tabella.alignment = 1  # CENTER
    tabella.autofit = True

    # header row
    for ci, header in enumerate(HEADERS_CALENDARIO):
        _set_cella_tabella(tabella.cell(0, ci), header, bold=True)
    # data rows: each key of the row dict goes into the column corresponding
    # to the header (Data, Orario, Ore, Modulo, Docente, Modalità)
    for ri, riga in enumerate(righe_calendario, start=1):
        for ci, header in enumerate(HEADERS_CALENDARIO):
            _set_cella_tabella(tabella.cell(ri, ci), riga.get(header, ""))

    _bordi_tabella(tabella)

    # move the table right after the 'Calendario' heading
    elemento_tabella = tabella._tbl
    corpo.remove(elemento_tabella)
    if par_calendario is not None:
        par_calendario._element.addnext(elemento_tabella)

    # remove 'Come gia' concordato.'
    if par_concordato is not None:
        par_concordato._element.getparent().remove(par_concordato._element)


# =========================================================================
# SAVING
# =========================================================================
def sanifica_nome(s):
    """Makes a string safe to use as a file name."""
    s = s.strip()
    s = re.sub(r"[\\/:*?\"<>|]", "", s)
    s = s.replace(" ", "_")
    if not s:
        s = "offerta"
    return s


def salva_offerta(doc, dati, edizione):
    """Saves the compiled document to output/ and returns the path.
    Includes the edition code in the file name."""
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    corso = codice_corso(edizione).replace(".", "")  # CORSO 7 -> CORSO7
    nome = (f"Offerta_Apprendistato_{sanifica_nome(dati['ragione'])}_"
            f"{corso}_{date.today().strftime('%Y%m%d')}.docx")
    percorso = os.path.join(OUTPUT_DIR, nome)
    base, ext = os.path.splitext(percorso)
    i = 1
    while os.path.exists(percorso):
        percorso = f"{base}_{i}{ext}"
        i += 1
    doc.save(percorso)
    return percorso


# =========================================================================
# MAIN
# =========================================================================
def main():
    print("=" * 70)
    print("  OFFER GENERATOR - PRIVATE APPRENTICESHIP")
    print("=" * 70)

    if not os.path.isfile(TEMPLATE):
        print(f"\nERROR: template not found at:\n  {TEMPLATE}")
        print("Make sure template_offerta_apprendistato.docx is present in the folder.")
        sys.exit(1)

    # --- 1. calendar (FIRST of all) ---
    info_cal = chiedi_calendario()
    righe_calendario = info_cal["righe"]
    edizione = info_cal["edizione"]
    calendario_concordato = info_cal["concordato"]

    # --- 2. data (participants, year, company) ---
    dati = chiedi_dati()

    # --- 3. calculations ---
    totale, finale = calc_totali(dati["n_partecipanti"])
    # progressive offer number (1st year from 20260044, 2nd from 20260066)
    numero_offerta = prossimo_numero_offerta(dati["ann_num"])
    print("\n--- CALCULATIONS ---")
    print(f"  Unit amount:           {euro(UNITARIO)}")
    print(f"  N. participants:       {dati['n_partecipanti']}")
    print(f"  Total amount:          {euro(totale)}")
    print(f"  Discount {int(SCONTO*100)}%:             {euro(totale*(1-SCONTO))}")
    print(f"  Final price (rounded to the nearest hundred): {euro(finale)}")
    print(f"  Edition:               {codice_corso(edizione)} / {codice_edizione_padded(edizione)}")
    print(f"  Year:                  {dati['ann_testo']}")
    print(f"  Offer number:          {numero_offerta}")

    # --- 4. confirm before generating ---
    if not _si_no("\nProceed with generating the offer?", default="s"):
        print("Generation cancelled.")
        return

    # --- 5. load the template and fill it in ---
    print("\nGenerating the document...")
    doc = docx.Document(TEMPLATE)

    compila_destinatario(doc, dati)
    compila_titolo_data_annualita_edizione(doc, dati, edizione, numero_offerta)
    compila_tabella_costi(doc, dati, totale, finale, edizione)
    if righe_calendario:  # only if not agreed upon
        inserisci_tabella_calendario(doc, righe_calendario)

    # --- 5b. record the used offer number in the persistent counter ---
    aggiorna_contatore(dati["ann_num"], numero_offerta)

    # --- 6. save ---
    percorso = salva_offerta(doc, dati, edizione)
    print(f"\n  >> Offer generated:\n     {percorso}")

    # --- 7. open option ---
    if _si_no("Do you want to open the document now?", default="s"):
        try:
            os.startfile(percorso)  # Windows
        except Exception as e:
            print(f"  Could not open it automatically ({e}). Open the file manually.")

    print("\nOperation completed.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrupted by the user.")
        sys.exit(0)
