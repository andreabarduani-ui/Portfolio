# -*- coding: utf-8 -*-
"""
transparency_certificates.py
===================
Duplicates the "Attestato Apprendimenti Acquisiti_.docx" file for each
participant listed in "File madre.xlsx", filling the Word tables with the
Excel data and saving each certificate with the participant's surname.

How to use it (VS Code):
  1. Open this folder in VS Code.
  2. Open the transparency_certificates.py file.
  3. Press F5 (or the "Run" button) or, from the terminal:
        py transparency_certificates.py

At the top (CONFIGURATION section) you can change the paths and other
options without touching the rest of the code.
"""

import os
import re
import datetime
import shutil
import tempfile

import openpyxl
import docx


# ============================================================================
# CONFIGURATION  (edit here if needed)
# ============================================================================
BASE_DIR   = "."
FILE_WORD  = os.path.join(BASE_DIR, "Attestato Apprendimenti Acquisiti_.docx")
FILE_EXCEL = os.path.join(BASE_DIR, "File madre.xlsx")

# Folder where the generated certificates will be saved
OUT_DIR    = os.path.join(BASE_DIR, "Attestati_Generati")

# Excel rows/sheets
NOME_FOGLIO   = "Foglio1"
RIGA_INTESTAZIONE = 2   # row containing the column headers
RIGA_PRIMO_DATO   = 3   # first row with participant data

# Output file name prefix (result: <prefix>_<Cognome>.docx)
PREFISSO_FILE = "Attestato Apprendimenti Acquisiti"

# If True, overwrites any already existing certificates with the same name
SOVRASCRIVI = True
# ============================================================================


# ----------------------------------------------------------------------------
# MAPPING  Word <-> Excel
# ----------------------------------------------------------------------------
# Key    = EXACT label that appears in the left cell of the Word
#          table (column 0).
# Value  = Excel column index (1-based: B=2, C=3, D=4, ...).
#
# IMPORTANT NOTE about Excel columns D and E:
#   - the headers are SWAPPED relative to the content;
#   - column D actually contains the birth DATES;
#   - column E actually contains the birth CITIES/STATES.
#   Here we rely on the actual CONTENT, not on the header.
# ----------------------------------------------------------------------------
MAPPING = {
    # --- Table 1: Allievo ---
    "Cognome":                              2,   # col B
    "Nome":                                 3,   # col C
    "Codice fiscale":                       8,   # col H
    "Data nascita":                         4,   # col D (contains the dates)
    "Comune/stato straniero nascita":       5,   # col E (contains the cities)
    "Cittadinanza":                         6,   # col F

    # --- Table 3: Esperienza formativa (only the empty fields get filled) ---
    "Titolo del percorso":                                  7,   # col G
    "N\u00b0 di ore frequentate su numero di ore totali del percorso": 14,  # col N

    # --- Table 4: Competenze ---
    "Risultato di apprendimento":                              9,   # col I
    "Risultato Atteso (RA)":                                   10,  # col J
    "Competenza/descrittore di cui ai quadri europei":        11,  # col K
    "Eventuali ulteriori evidenze a supporto riferite alla personalizzazione del percorso": 12,  # col L
    "Prove di valutazione \u2013 data \u2013 esito":          13,  # col M
}
# NB: Excel fields stored as timedelta (column N) are converted to hours.
#     Word cells that already contain a value are NOT touched (they remain
#     unchanged), as requested.


# ----------------------------------------------------------------------------
# Utility functions
# ----------------------------------------------------------------------------

def normalizza_etichetta(s):
    """Makes a label comparable: lowercase, collapsed spaces, no punctuation."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def valore_come_testo(val):
    """Converts an Excel value (datetime, timedelta, None, ...) to a string."""
    if val is None:
        return ""
    # Birth date -> Italian format gg/mm/aaaa (dd/mm/yyyy)
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d/%m/%Y")
    # Duration (timedelta) -> total hours (integer or with one decimal)
    if isinstance(val, datetime.timedelta):
        ore = val.total_seconds() / 3600.0
        # if practically an integer, show without decimals
        if abs(ore - round(ore)) < 1e-6:
            return f"{int(round(ore))}"
        return f"{ore:.1f}".rstrip("0").rstrip(".")
    # Numbers
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def scrivi_in_cella(cell, testo):
    """
    Writes 'testo' into the cell while KEEPING the document style:
    - uses the first existing run (copies its format);
    - if the cell is empty, creates a run inheriting the paragraph style.
    """
    par = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    # Remove the existing runs keeping the first one (for the format)
    if par.runs:
        primo_run = par.runs[0]
        primo_run.text = testo
        # empty the other runs
        for r in par.runs[1:]:
            r.text = ""
    else:
        run = par.add_run(testo)
    # if there were any other paragraphs in the cell, empty them (rare case)
    for extra in cell.paragraphs[1:]:
        for r in extra.runs:
            r.text = ""


def sanitize_filename(name):
    """Removes characters that are invalid in Windows file names."""
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name


# ----------------------------------------------------------------------------
# Main logic
# ----------------------------------------------------------------------------

def leggi_partecipanti():
    """Reads the Excel file and returns a list of dicts {col_index: value}."""
    wb = openpyxl.load_workbook(FILE_EXCEL, data_only=True)
    ws = wb[NOME_FOGLIO]

    partecipanti = []
    for riga in range(RIGA_PRIMO_DATO, ws.max_row + 1):
        cognome = ws.cell(row=riga, column=2).value  # col B
        if cognome is None or str(cognome).strip() == "":
            continue  # empty row: skip
        dati = {}
        for col in range(1, ws.max_column + 1):
            dati[col] = ws.cell(row=riga, column=col).value
        dati["_cognome"] = str(cognome).strip()
        dati["_nome"] = str(ws.cell(row=riga, column=3).value or "").strip()
        partecipanti.append(dati)
    return partecipanti


def compila_documento(template_doc, dati):
    """
    Creates a NEW copy of the template document (from file, to avoid
    problems sharing the internal XML structure across copies) and
    fills the tables with the participant's data.
    The original template file is NEVER modified.
    """
    # Copies the template to a temporary file, then opens it: this way each
    # certificate starts from a clean, independent XML structure.
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    try:
        shutil.copyfile(template_doc, tmp_path)
        doc = docx.Document(tmp_path)
    finally:
        # remove the temporary file right away; the 'doc' stays in memory
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    # Map: normalized label -> Excel column
    mappa_norm = {normalizza_etichetta(k): v for k, v in MAPPING.items()}

    for tabella in doc.tables:
        ncol = len(tabella.columns)
        if ncol < 2:
            continue  # at least 2 columns needed (label | value)
        for riga in tabella.rows:
            celle = riga.cells
            etichetta = celle[0].text
            etich_norm = normalizza_etichetta(etichetta)

            # FUNDAMENTAL RULE (requested by the user):
            #   fill in ONLY the value cells that are currently EMPTY.
            #   All cells that already contain a value remain UNCHANGED
            #   in every certificate (e.g. Ente/organisation data, manager,
            #   instance code, fund, etc.).
            if etich_norm in mappa_norm and celle[1].text.strip() == "":
                col_excel = mappa_norm[etich_norm]
                valore = dati.get(col_excel)
                testo = valore_come_testo(valore)
                if testo != "":
                    scrivi_in_cella(celle[1], testo)
    return doc


def main():
    # --- Existence checks ---
    if not os.path.isfile(FILE_WORD):
        print(f"ERROR: Word file not found:\n  {FILE_WORD}")
        return
    if not os.path.isfile(FILE_EXCEL):
        print(f"ERROR: Excel file not found:\n  {FILE_EXCEL}")
        return

    os.makedirs(OUT_DIR, exist_ok=True)

    print("Reading participants from the Excel file...")
    partecipanti = leggi_partecipanti()
    print(f"Found {len(partecipanti)} participants.\n")

    print(f"Word template:\n  {FILE_WORD}")

    generati = 0
    saltati = 0
    for i, dati in enumerate(partecipanti, start=1):
        cognome = sanitize_filename(dati["_cognome"])
        nome = dati["_nome"]
        out_name = f"{PREFISSO_FILE}_{cognome}.docx"
        out_path = os.path.join(OUT_DIR, out_name)

        if os.path.exists(out_path) and not SOVRASCRIVI:
            print(f"[{i:02d}] SKIP (already exists): {out_name}")
            saltati += 1
            continue

        doc = compila_documento(FILE_WORD, dati)
        doc.save(out_path)
        print(f"[{i:02d}] Generated: {out_name}   ({cognome} {nome})")
        generati += 1

    print("\n" + "=" * 60)
    print(f"Done. Generated: {generati} | Skipped: {saltati}")
    print(f"Output folder:\n  {OUT_DIR}")
    print("=" * 60)


if __name__ == "__main__":
    main()
