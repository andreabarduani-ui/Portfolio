"""
certificate_manager.py
=======================

Single script for the complete management of CORSO certificates, with three
subcommands corresponding to the three steps of the workflow:

  1. dividi       Splits a mass-print PDF (alternating front+back)
                   into one file per person, joining front and back.

  2. organizza     Sorts the PDFs of ONE edition (e.g. folder "CORSO10") into
                   subfolders by company, cross-referencing the names with an
                   Excel rubrica (columns: Cognome, Nome, NomeCorso, Azienda).

  3. riorganizza    Rebuilds MULTIPLE editions already organized by company
                   (output of step 2) into an inverted view:
                   first level = company, second level = edition.
                   With --rubrica it also generates, in each company folder,
                   a Word file "Registro attestati <corso> - <azienda>.docx"
                   with the list of trainees taken from the rubrica.

  (bonus) tutto     Runs "dividi" + "organizza" in sequence for a single
                   edition, handy when you already have both the mass-print
                   PDF and the Excel rubrica of that edition.

USAGE FROM VS CODE / TERMINAL
----------------------------
    py certificate_manager.py dividi ["input.pdf"]
    py certificate_manager.py organizza ["CORSO10"] ["rubrica.xls"] [--copia]
    py certificate_manager.py riorganizza ["cartelle"...] [--output DIR] [--sposta] [--rubrica "rubrica.xls"]
    py certificate_manager.py tutto ["input.pdf"] ["rubrica.xls"] [--edizione CORSO10] [--copia]

Every positional argument is OPTIONAL: if omitted, the script tries to find
it automatically in the folder where it lives (useful to launch with F5 in
VS Code without configuring anything). Launch without a subcommand to see
this summary.

REQUIREMENTS
---------
    py -m pip install -r requirements.txt
    (pypdf to read the PDFs; pandas + xlrd/openpyxl to read the Excel
    rubrica; python-docx to generate the Word registers — needed only by
    the 'organizza', 'riorganizza' and 'tutto' subcommands, but installing
    them all together avoids surprises when you switch from one
    subcommand to another)
"""

from __future__ import annotations

import argparse
import logging
import re
import shutil
import sys
from difflib import get_close_matches
from pathlib import Path

import pandas as pd
from docx import Document as DocxDocument
from docx.shared import Cm, Pt
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from pypdf import PdfReader, PdfWriter

logging.basicConfig(level=logging.INFO, format="%(message)s")
log = logging.getLogger(__name__)


# ========================================================================
# SHARED FUNCTIONS
# ========================================================================
_ACCENTI = {
    "à": "a", "á": "a", "À": "A", "Á": "A",
    "è": "e", "é": "e", "È": "E", "É": "E",
    "ì": "i", "í": "i", "Ì": "I", "Í": "I",
    "ò": "o", "ó": "o", "Ò": "O", "Ó": "O",
    "ù": "u", "ú": "u", "Ù": "U", "Ú": "U",
}


def estrai_nome_da_testo(testo: str) -> str | None:
    """Extracts the participant name from the text of a front page,
    looking for the pattern 'Si attesta che' followed by the name on the
    next line. Returns None if not found."""
    match = re.search(r"Si attesta che\s*\n?\s*(.+?)\s*\n", testo)
    return match.group(1).strip() if match else None


def normalizza(nome: str) -> str:
    """Normalizes a name for comparison: uppercase, without accents,
    without punctuation, without double spaces."""
    for accentata, semplice in _ACCENTI.items():
        nome = nome.replace(accentata, semplice)
    nome = re.sub(r"[^A-Za-z ]", " ", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome.upper()


def slug_nome_persona(nome: str) -> str:
    """Converts a name into a format suitable for a filename (Nome_Cognome)."""
    nome_norm = normalizza(nome)
    parti = [p.capitalize() for p in nome_norm.split()]
    return "_".join(parti) if parti else "Sconosciuto"


def sanitizza_nome_cartella(nome: str) -> str:
    """Removes characters not allowed in Windows folder names."""
    nome = re.sub(r'[<>:"/\\|?*]', "", nome).strip()
    nome = re.sub(r"\s+", " ", nome)
    return nome or "AZIENDA_SCONOSCIUTA"


# ========================================================================
# SECTION 1 — DIVIDI: mass-print PDF -> one file per person
# ========================================================================
def trova_pdf_input(cartella: Path) -> Path | None:
    """Automatically looks for a PDF in the folder (for launching with F5
    with no arguments). If there is more than one, it asks you to specify it."""
    candidati = sorted(cartella.glob("*.pdf"))
    if len(candidati) == 1:
        return candidati[0]
    if len(candidati) > 1:
        log.warning("Multiple PDFs found in the folder. Specify which one to use:")
        for c in candidati:
            log.warning(f"  - {c.name}")
    return None


def dividi_pdf(input_path: Path) -> Path:
    """Splits the input PDF into one file per person (front+back joined).
    Returns the path of the output folder."""
    reader = PdfReader(str(input_path))
    n_pagine = len(reader.pages)

    if n_pagine == 0:
        raise ValueError("The PDF contains no pages.")
    if n_pagine % 2 != 0:
        raise ValueError(
            f"The PDF has {n_pagine} pages (odd number). "
            "I expect front+back pairs: check the input file."
        )

    n_persone = n_pagine // 2
    out_dir = input_path.parent / "attestati_output"
    out_dir.mkdir(exist_ok=True)

    log.info(f"Found {n_pagine} pages -> {n_persone} participants.\n")

    for i in range(n_persone):
        fronte = reader.pages[i * 2]
        retro = reader.pages[i * 2 + 1]

        testo = fronte.extract_text() or ""
        nome = estrai_nome_da_testo(testo)
        nome_file = slug_nome_persona(nome) if nome else f"persona_{i + 1}"

        writer = PdfWriter()
        writer.add_page(fronte)
        writer.add_page(retro)

        out_path = out_dir / f"{i + 1:02d}_{nome_file}.pdf"
        with open(out_path, "wb") as f:
            writer.write(f)

        etichetta = nome if nome else "(name not recognized)"
        log.info(f"  [{i + 1:02d}] {etichetta} -> {out_path.name}")

    log.info(f"\nCompleted. Files saved in: {out_dir}")
    return out_dir


def azione_dividi(args: argparse.Namespace, script_dir: Path) -> Path:
    if args.input_pdf:
        input_path = Path(args.input_pdf)
    else:
        input_path = trova_pdf_input(script_dir)
        if input_path is None:
            log.error(
                "No PDF specified and no unique PDF found in the folder.\n"
                "Drag the PDF into the project folder or run:\n"
                '  py certificate_manager.py dividi "filename.pdf"'
            )
            sys.exit(1)
        log.info(f"No argument provided: using the automatically found PDF -> {input_path.name}\n")

    if not input_path.exists():
        log.error(f"File not found: {input_path}")
        sys.exit(1)

    try:
        return dividi_pdf(input_path)
    except ValueError as e:
        log.error(f"Error: {e}")
        sys.exit(1)


# ========================================================================
# SECTION 2 — ORGANIZZA: one edition -> subfolders by company
# ========================================================================
def carica_rubrica(excel_path: Path) -> pd.DataFrame:
    """Loads the Excel file and checks the required columns."""
    df = pd.read_excel(excel_path)

    colonne_richieste = {"Cognome", "Nome", "NomeCorso", "Azienda"}
    mancanti = colonne_richieste - set(df.columns)
    if mancanti:
        raise ValueError(
            f"The Excel file is missing the columns: {', '.join(mancanti)}. "
            f"Columns found: {', '.join(df.columns)}"
        )

    df = df.dropna(subset=["Cognome", "Nome", "Azienda"]).copy()
    df["nome_completo_norm"] = (df["Nome"].astype(str) + " " + df["Cognome"].astype(str)).apply(normalizza)
    df["edizione"] = df["NomeCorso"].astype(str).apply(estrai_numero_edizione_da_corso)
    return df


def estrai_numero_edizione_da_corso(testo: str) -> str | None:
    """Extracts the CORSO edition number from a string like
    'CORSO EDIZIONE 10 - ...'."""
    match = re.search(r"EDIZIONE\s+(\d+)", testo, re.IGNORECASE)
    return match.group(1) if match else None


def estrai_numero_edizione_da_cartella(nome_cartella: str) -> str | None:
    """Extracts the edition number from the folder name, e.g. 'CORSO10' -> '10'."""
    match = re.search(r"CORSO\s*(\d+)", nome_cartella, re.IGNORECASE)
    return match.group(1) if match else None


def estrai_nome_da_pdf(pdf_path: Path) -> str | None:
    """Extracts the participant name from the text of the first page (front),
    with a fallback on the file name if the text cannot be extracted."""
    try:
        reader = PdfReader(str(pdf_path))
        testo = reader.pages[0].extract_text() or ""
    except Exception as e:
        log.warning(f"  Cannot read the text of {pdf_path.name}: {e}")
        testo = ""

    nome = estrai_nome_da_testo(testo)
    if nome:
        return nome

    base = re.sub(r"^\d+_", "", pdf_path.stem)
    return base.replace("_", " ") if base else None


def trova_azienda(
    nome_pdf: str,
    rubrica: pd.DataFrame,
    edizione: str | None,
) -> tuple[str | None, str]:
    """Finds the company matching the name extracted from the PDF.
    Returns (company_or_None, method)."""
    nome_norm = normalizza(nome_pdf)

    subset = rubrica
    if edizione is not None:
        filtrata = rubrica[rubrica["edizione"] == edizione]
        if not filtrata.empty:
            subset = filtrata

    match_esatto = subset[subset["nome_completo_norm"] == nome_norm]
    if not match_esatto.empty:
        return match_esatto.iloc[0]["Azienda"], "esatto"

    candidati = subset["nome_completo_norm"].tolist()
    vicini = get_close_matches(nome_norm, candidati, n=1, cutoff=0.85)
    if vicini:
        riga = subset[subset["nome_completo_norm"] == vicini[0]].iloc[0]
        return riga["Azienda"], "edizione+fuzzy" if edizione else "fuzzy"

    if edizione is not None:
        candidati_globali = rubrica["nome_completo_norm"].tolist()
        vicini_globali = get_close_matches(nome_norm, candidati_globali, n=1, cutoff=0.85)
        if vicini_globali:
            riga = rubrica[rubrica["nome_completo_norm"] == vicini_globali[0]].iloc[0]
            return riga["Azienda"], "fuzzy_globale"

    return None, "nessuno"


def organizza_cartella_edizione(cartella: Path, rubrica: pd.DataFrame, copia: bool = False) -> None:
    edizione = estrai_numero_edizione_da_cartella(cartella.name)
    if edizione:
        log.info(f"\n=== {cartella.name} (CORSO edition {edizione}) ===")
    else:
        log.info(f"\n=== {cartella.name} (edition not recognized from folder name) ===")

    pdf_files = sorted(cartella.glob("*.pdf"))
    if not pdf_files:
        log.info("  No PDFs found in this folder (maybe already organized?).")
        return

    n_trovati = 0
    n_non_trovati = 0

    for pdf_path in pdf_files:
        nome = estrai_nome_da_pdf(pdf_path)
        if not nome:
            log.warning(f"  [!] Cannot determine the name for {pdf_path.name}")
            azienda, metodo = None, "nessuno"
        else:
            azienda, metodo = trova_azienda(nome, rubrica, edizione)

        if azienda:
            cartella_azienda = cartella / sanitizza_nome_cartella(str(azienda))
            simbolo = "~" if metodo != "esatto" else " "
            log.info(f"  [{simbolo}] {nome or pdf_path.stem} -> {azienda}  ({metodo})")
            n_trovati += 1
        else:
            cartella_azienda = cartella / "_DA_VERIFICARE"
            log.warning(f"  [X] {nome or pdf_path.stem} -> NO COMPANY FOUND")
            n_non_trovati += 1

        cartella_azienda.mkdir(exist_ok=True)
        destinazione = cartella_azienda / pdf_path.name

        if copia:
            shutil.copy2(pdf_path, destinazione)
        else:
            shutil.move(str(pdf_path), str(destinazione))

    log.info(f"  --> {n_trovati} assigned, {n_non_trovati} to verify manually.")


def trova_cartella_singola_edizione(base: Path) -> Path | None:
    """Looks for a folder starting with 'CORSO' in the base folder,
    or returns base itself if it already contains PDFs."""
    if any(base.glob("*.pdf")):
        return base
    candidati = sorted(p for p in base.iterdir() if p.is_dir() and p.name.upper().startswith("CORSO"))
    if len(candidati) == 1:
        return candidati[0]
    return None


def trova_file_excel(base: Path) -> Path | None:
    candidati = sorted(base.glob("*.xls")) + sorted(base.glob("*.xlsx"))
    if len(candidati) == 1:
        return candidati[0]
    return None


def azione_organizza(args: argparse.Namespace, script_dir: Path) -> None:
    if args.cartella_edizione:
        input_path = Path(args.cartella_edizione)
    else:
        input_path = trova_cartella_singola_edizione(script_dir)
        if input_path is None:
            log.error(
                "Could not automatically find a CORSO* folder with PDFs.\n"
                "Specify the path, for example:\n"
                '  py certificate_manager.py organizza "CORSO10" "rubrica.xls"'
            )
            sys.exit(1)
        log.info(f"Edition folder found automatically: {input_path.name}")

    if not input_path.exists():
        log.error(f"Folder not found: {input_path}")
        sys.exit(1)

    if args.file_excel:
        excel_path = Path(args.file_excel)
    else:
        excel_path = trova_file_excel(script_dir)
        if excel_path is None:
            log.error(
                "Could not automatically find an Excel file (.xls/.xlsx) in the script folder.\n"
                "Specify the path, for example:\n"
                '  py certificate_manager.py organizza "CORSO10" "rubrica.xls"'
            )
            sys.exit(1)
        log.info(f"Excel file found automatically: {excel_path.name}")

    if not excel_path.exists():
        log.error(f"Excel file not found: {excel_path}")
        sys.exit(1)

    try:
        rubrica = carica_rubrica(excel_path)
    except ValueError as e:
        log.error(f"Error reading the rubrica: {e}")
        sys.exit(1)

    log.info(f"Rubrica loaded: {len(rubrica)} entries.")

    if any(input_path.glob("*.pdf")):
        cartelle_da_processare = [input_path]
    else:
        cartelle_da_processare = sorted(
            p for p in input_path.iterdir() if p.is_dir() and p.name.upper().startswith("CORSO")
        )
        if not cartelle_da_processare:
            log.error(f"In folder '{input_path}' neither PDFs nor CORSO* subfolders were found.")
            sys.exit(1)

    for cartella in cartelle_da_processare:
        organizza_cartella_edizione(cartella, rubrica, copia=args.copia)

    log.info("\nCompleted.")


# ========================================================================
# SECTION 3 — RIORGANIZZA: multiple editions -> company/edition view
# ========================================================================
def trova_tutte_cartelle_edizione(base: Path) -> list[Path]:
    """Finds all subfolders whose name starts with 'CORSO' inside 'base'.
    If 'base' itself starts with CORSO, returns it directly."""
    if base.name.upper().startswith("CORSO") and base.is_dir():
        return [base]
    return sorted(p for p in base.iterdir() if p.is_dir() and p.name.upper().startswith("CORSO"))


def riorganizza(cartelle_edizione: list[Path], output_root: Path, sposta: bool = False) -> None:
    output_root.mkdir(parents=True, exist_ok=True)

    totale_spostati = 0
    totale_edizioni = 0
    riepilogo: dict[str, set[str]] = {}

    for cartella_edizione in cartelle_edizione:
        edizione = cartella_edizione.name
        log.info(f"\n=== Edition: {edizione} (from {cartella_edizione}) ===")

        sottocartelle_azienda = sorted(p for p in cartella_edizione.iterdir() if p.is_dir())

        if not sottocartelle_azienda:
            log.warning(
                f"  No company subfolder found in '{edizione}'. "
                f"Have you already run 'organizza' on this folder?"
            )
            continue

        for cartella_azienda in sottocartelle_azienda:
            azienda = cartella_azienda.name
            pdf_files = sorted(cartella_azienda.glob("*.pdf"))
            if not pdf_files:
                continue

            destinazione_azienda_edizione = output_root / azienda / edizione
            destinazione_azienda_edizione.mkdir(parents=True, exist_ok=True)

            for pdf_path in pdf_files:
                destinazione = destinazione_azienda_edizione / pdf_path.name
                if sposta:
                    shutil.move(str(pdf_path), str(destinazione))
                else:
                    shutil.copy2(pdf_path, destinazione)
                totale_spostati += 1

            log.info(f"  {azienda}: {len(pdf_files)} file -> {destinazione_azienda_edizione}")
            riepilogo.setdefault(azienda, set()).add(edizione)

        totale_edizioni += 1

    log.info("\n" + "=" * 60)
    log.info("SUMMARY")
    log.info("=" * 60)
    for azienda in sorted(riepilogo):
        edizioni = ", ".join(sorted(riepilogo[azienda]))
        log.info(f"  {azienda}: {edizioni}")

    azione = "moved" if sposta else "copied"
    log.info(
        f"\n{totale_spostati} files {azione} in total, "
        f"from {totale_edizioni} editions, into {len(riepilogo)} different companies."
    )
    log.info(f"Output folder: {output_root}")


def azione_riorganizza(args: argparse.Namespace, script_dir: Path) -> None:
    if args.cartelle:
        percorsi = [Path(c) for c in args.cartelle]
        if len(percorsi) == 1 and percorsi[0].is_dir() and not percorsi[0].name.upper().startswith("CORSO"):
            cartelle_edizione = trova_tutte_cartelle_edizione(percorsi[0])
        else:
            cartelle_edizione = percorsi
    else:
        cartelle_edizione = trova_tutte_cartelle_edizione(script_dir)

    cartelle_edizione = [c for c in cartelle_edizione if c.exists() and c.is_dir()]

    if not cartelle_edizione:
        log.error(
            "No edition folder (CORSO*) found.\n"
            "Specify the folders, for example:\n"
            '  py certificate_manager.py riorganizza "CORSO10" "CORSO11" "CORSO12" "CORSO13" "CORSO14"\n'
            "or a container folder:\n"
            '  py certificate_manager.py riorganizza "C:\\path\\Attestati"'
        )
        sys.exit(1)

    log.info(f"Edition folders found ({len(cartelle_edizione)}):")
    for c in cartelle_edizione:
        log.info(f"  - {c}")

    output_root = Path(args.output) if args.output else cartelle_edizione[0].parent / "ATTESTATI_PER_AZIENDA"

    log.info(f"\nOutput folder: {output_root}")
    log.info(f"Mode: {'MOVE' if args.sposta else 'COPY (originals intact)'}")

    riorganizza(cartelle_edizione, output_root, sposta=args.sposta)

    # ---- EXTRA STEP: Word register generation per company ----
    # Runs at the end of the reorganization process: for each company folder
    # present in output_root it generates a .docx file with the list of
    # trainees taken from the Excel rubrica.
    rubrica_excel = getattr(args, "rubrica", None)
    if rubrica_excel:
        excel_path = Path(rubrica_excel)
        if not excel_path.exists():
            log.error(f"\nRubrica Excel file not found: {excel_path}")
        else:
            try:
                rubrica = carica_rubrica(excel_path)
                log.info(f"\nRubrica loaded for the registers: {len(rubrica)} entries.")
                log.info("\n### Generating Word registers per company ###")
                genera_registri_word(output_root, rubrica)
            except ValueError as e:
                log.error(f"\nError reading the rubrica: {e}")

    log.info("\nCompleted.")


def azione_registri(args: argparse.Namespace, script_dir: Path) -> None:
    """Generates the Word registers from template + rubrica ONLY, with no need
    of CORSO* folders or PDFs. Creates a folder for each company present in
    the rubrica and saves the corresponding register .docx in it."""
    if args.rubrica:
        excel_path = Path(args.rubrica)
    else:
        excel_path = trova_file_excel(script_dir)
        if excel_path is None:
            log.error(
                "No Excel file (.xls/.xlsx) found in the script folder.\n"
                "Specify the path, for example:\n"
                '  py certificate_manager.py registri "rubrica.xls"'
            )
            sys.exit(1)
        log.info(f"Excel file found automatically: {excel_path.name}")

    if not excel_path.exists():
        log.error(f"Excel file not found: {excel_path}")
        sys.exit(1)

    try:
        rubrica = carica_rubrica(excel_path)
    except ValueError as e:
        log.error(f"Error reading the rubrica: {e}")
        sys.exit(1)

    output_dir = Path(args.output) if args.output else script_dir / "REGISTRI"
    output_dir.mkdir(parents=True, exist_ok=True)

    log.info(f"Rubrica loaded: {len(rubrica)} entries.")
    log.info(f"Output folder: {output_dir}")
    log.info("\n### Generating Word registers from rubrica ###")
    genera_registri_da_rubrica(rubrica, output_dir)
    log.info("\nCompleted.")


# ========================================================================
# SECTION 4 — WORD REGISTERS: generates one register per company
# ========================================================================
def genera_registri_word(output_root: Path, rubrica: pd.DataFrame) -> None:
    """Generates a 'Registro attestati' .docx file for each company
    present in the output structure (output_root/<azienda>/...).

    For each company:
      - retrieves from the rubrica all the trainees of that company
        (normalized name match, with fuzzy fallback);
      - creates a Word document with:
          * Header table:  Descrizione | Attestati di Frequenza azienda <NOME>
          * List table:    N° | Cognome e Nome | Data | Firma per Ricezione
      - saves the file in the company folder as:
          'Registro attestati <NomeCorso> - <Azienda>.docx'

    The format follows that of the reference file
    'Registro attestati Cliente B.doc'.
    """
    # Prepares a normalized column of the company for the matching
    rubrica = rubrica.copy()
    rubrica["azienda_norm"] = rubrica["Azienda"].astype(str).apply(normalizza)

    aziende_dirs = sorted(p for p in output_root.iterdir() if p.is_dir())

    if not aziende_dirs:
        log.warning("  No company folder found in output: register not generated.")
        return

    n_registri = 0
    for cartella_azienda in aziende_dirs:
        azienda_cartella = cartella_azienda.name
        azienda_norm = normalizza(azienda_cartella)

        # Looks for the trainees of this company (normalized match, then fuzzy)
        corsisti = rubrica[rubrica["azienda_norm"] == azienda_norm]
        if corsisti.empty:
            candidati = rubrica["azienda_norm"].unique().tolist()
            vicini = get_close_matches(azienda_norm, candidati, n=1, cutoff=0.85)
            if vicini:
                corsisti = rubrica[rubrica["azienda_norm"] == vicini[0]]
            else:
                log.warning(
                    f"  [!] No trainee in the rubrica for '{azienda_cartella}': "
                    f"register not generated."
                )
                continue

        azienda_rubrica = str(corsisti.iloc[0]["Azienda"]).strip()

        # If a company has trainees in multiple courses, groups them by
        # course: generates a separate register for each course.
        for nome_corso, gruppo in corsisti.groupby("NomeCorso"):
            nome_corso = str(nome_corso).strip()
            _genera_registro_word(
                cartella_azienda=cartella_azienda,
                azienda=azienda_rubrica,
                nome_corso=nome_corso,
                corsisti=gruppo,
            )
            n_registri += 1

    log.info(f"\nGenerated {n_registri} Word register(s).")


def genera_registri_da_rubrica(rubrica: pd.DataFrame, output_dir: Path) -> None:
    """Generates a 'Registro attestati' .docx file for each company present
    IN THE RUBRICA (the CORSO* folder with the PDFs is not needed).

    For each company:
      - retrieves from the rubrica all the trainees of that company;
      - creates a <output_dir>/<azienda>/ folder;
      - generates the Word register (on modello_registro.docx) replacing
        company, course, CORSO code and trainee list.
    """
    n_registri = 0
    for azienda, gruppo in rubrica.groupby("Azienda"):
        azienda = str(azienda).strip()
        cartella_azienda = output_dir / sanitizza_nome_cartella(azienda)
        cartella_azienda.mkdir(parents=True, exist_ok=True)

        # If a company has trainees in multiple courses, one register per course.
        for nome_corso, sottogruppo in gruppo.groupby("NomeCorso"):
            _genera_registro_word(
                cartella_azienda=cartella_azienda,
                azienda=azienda,
                nome_corso=str(nome_corso).strip(),
                corsisti=sottogruppo,
            )
            n_registri += 1

    log.info(f"\nGenerated {n_registri} Word register(s).")


def tronca_per_percorso(testo: str, cartella: Path, suffix_base: str = "") -> str:
    """Truncates 'testo' so that the full path of the .docx file generated
    in 'cartella' respects the Windows MAX_PATH limit (260 chars).

    Leaves a safety margin for the company suffix, separators and '.docx'.
    """
    # Characters reserved for the suffix: separator " - " + company
    # (estimated equal to the folder name) + ".docx". We use a wide margin
    # for safety.
    azienda_len = len(cartella.name)
    margine = len(suffix_base) + len(" - ") + azienda_len + len(".docx")
    limite = 259 - len(str(cartella)) - 1 - margine  # -1 for the '\' separator
    if limite < 10:
        limite = 10
    if len(testo) > limite:
        testo = testo[:limite].rstrip()
    return testo


def _grassetto_riga(riga) -> None:
    """Sets all the text of a table row to bold."""
    for cella in riga.cells:
        for paragrafo in cella.paragraphs:
            for run in paragrafo.runs:
                run.bold = True


def _imposta_font_cella(cell, font_name: str, size_pt: float, bold: bool = False) -> None:
    """Sets font name, size and bold on all the runs of the
    paragraph (and of the first paragraph) of the cell. Also sets the
    font for complex scripts (rPr/rFonts cs) for consistency with Word."""
    for paragrafo in cell.paragraphs:
        for run in paragrafo.runs:
            run.font.name = font_name
            run.font.size = Pt(size_pt)
            run.bold = bold
            # Also sets the font for ASCII/hAnsi/cs so that Word really
            # applies it (python-docx only sets w:ascii via
            # font.name, which is not always sufficient).
            rpr = run._element.get_or_add_rPr()
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = rpr.makeelement(qn("w:rFonts"), {})
                rpr.append(rfonts)
            for attr in ("w:ascii", "w:hAnsi", "w:cs"):
                rfonts.set(qn(attr), font_name)


def _centra_paragrafo(paragrafo) -> None:
    """Horizontally centers a paragraph."""
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _estrai_codice_corso(nome_corso: str) -> str:
    """Returns the edition code like 'CORSO08' from the NomeCorso.
    Tries first with 'EDIZIONE N', then with the '[CORSO N]' prefix.
    Returns 'CORSO?' if it cannot extract the number."""
    corso_breve = estrai_numero_edizione_da_corso(nome_corso)
    if not corso_breve:
        corso_breve = estrai_numero_edizione_da_cartella(nome_corso)
    return f"CORSO{corso_breve}" if corso_breve else "CORSO?"


def _sostituisci_testo_header(part, vecchio: str, nuovo: str) -> None:
    """Replaces all occurrences of 'vecchio' with 'nuovo' in the runs of
    a Word part (header/footer/body) preserving the formatting.

    Works at the XML level to handle the (frequent) case where the text
    is split across multiple w:r/w:t runs: it recomputes the distribution
    keeping the same run elements.
    """
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xml = part._element

    for paragrafo in xml.iter(f"{{{W_NS}}}p"):
        run_t = []
        # Collects all the w:t nodes of the paragraph and the mapping to the runs
        run_nodes = paragrafo.findall(f"{{{W_NS}}}r")
        for r in run_nodes:
            for t in r.findall(f"{{{W_NS}}}t"):
                run_t.append((r, t, t.text or ""))

        if not run_t:
            continue

        testo_unito = "".join(txt for _, _, txt in run_t)
        if vecchio not in testo_unito:
            continue

        nuovo_unito = testo_unito.replace(vecchio, nuovo)

        # Redistributes the new text over the existing runs:
        # puts everything into the first w:t run and empties the others, so
        # the first run's formatting is applied to the whole text.
        primo = True
        for r, t, _ in run_t:
            if primo:
                t.text = nuovo_unito
                primo = False
            else:
                t.text = ""


def _genera_registro_word(
    cartella_azienda: Path,
    azienda: str,
    nome_corso: str,
    corsisti: pd.DataFrame,
) -> None:
    """Creates and saves a single register .docx file.

    Uses the 'modello_registro.docx' file as base (a .docx copy of the
    reference file 'Registro attestati Cliente B.doc') and merely
    replaces the variable contents, preserving the original format
    intact (logo, page header, tables, sizes, fonts):
      - header: course name and CORSO edition code (project row)
      - body: company name (holder table) and trainee list
        (register table), with correct 1..N numbering.
    """
    # ---- Opening the template ----
    modello_path = Path(__file__).parent / "modello_registro.docx"
    if not modello_path.exists():
        raise FileNotFoundError(
            f"Template not found: {modello_path}. "
            "Place 'modello_registro.docx' in the script folder."
        )
    doc = DocxDocument(str(modello_path))

    # Edition code for the project row (e.g. "CORSO08")
    codice_corso = _estrai_codice_corso(nome_corso)

    # ---- Replacements in the page header ----
    # The template contains the texts of the reference document: we
    # replace them with those of the current course/edition.
    header = doc.sections[0].header
    # Course name row (in the template: 'PROBLEM SOLVING ... II ANNUALITA')
    _sostituisci_testo_header(header, "PROBLEM SOLVING E PROCESSI DECISIONALI  II ANNUALITÀ", nome_corso)
    # Project row: CORSO code (in the template: 'CORSO 20' -> 'CORSOxx')
    _sostituisci_testo_header(header, "CORSO 20", codice_corso)

    # ---- Table 1: holder (Descrizione | company value) ----
    tab_intro = doc.tables[0]
    # Replaces the template company name with the real one
    cella_val = tab_intro.cell(0, 1)
    _sostituisci_testo_header(cella_val, "CLIENTE_B S.N.C.", azienda)

    # ---- Table 2: trainee list ----
    tab_elenco = doc.tables[1]
    # The template has 1 header row + N data rows (one with the example
    # 'CORSISTA ESEMPIO' and the others empty). We remove all the existing
    # data rows and rebuild them with the real trainees.
    n_da_rimuovere = len(tab_elenco.rows) - 1  # keep only the header
    for _ in range(n_da_rimuovere):
        # Removes the last data row (the XML tr element)
        tbl = tab_elenco._tbl
        tbl.remove(tab_elenco.rows[-1]._tr)

    # Adds a row for each real trainee, numbering 1..N
    for idx, (_, riga) in enumerate(corsisti.iterrows(), start=1):
        cells = tab_elenco.add_row().cells
        cells[0].text = str(idx)
        _imposta_font_cella(cells[0], "Arial", 10, bold=True)
        _centra_paragrafo(cells[0].paragraphs[0])
        cells[1].text = f"{riga['Cognome']} {riga['Nome']}".strip()
        _imposta_font_cella(cells[1], "Times New Roman", 12, bold=False)
        # Data and Firma columns left empty (to be filled in by hand)

    # ---- Saving ----
    # The file name uses a short course identifier (edition number if
    # available, otherwise the truncated NomeCorso) to respect the
    # Windows MAX_PATH limit (260 characters): the path is
    # already very long. The full NomeCorso still remains in the
    # document header.
    corso_breve = estrai_numero_edizione_da_corso(nome_corso)
    if corso_breve:
        corso_slug = f"CORSO{corso_breve}"
    else:
        corso_slug = sanitizza_nome_cartella(nome_corso)
    corso_slug = tronca_per_percorso(corso_slug, cartella_azienda, suffix_base="Registro attestati ")
    azienda_slug = sanitizza_nome_cartella(azienda)
    nome_file = f"Registro attestati {corso_slug} - {azienda_slug}.docx"
    out_path = cartella_azienda / nome_file

    doc.save(str(out_path))
    log.info(
        f"  {cartella_azienda.name}: register generated -> {out_path.name} "
        f"({len(corsisti)} trainees)"
    )


# ========================================================================
# BONUS — TUTTO: dividi + organizza in one go (single edition)
# ========================================================================
def azione_tutto(args: argparse.Namespace, script_dir: Path) -> None:
    log.info("### Step 1/2: front+back split ###\n")

    dividi_args = argparse.Namespace(input_pdf=args.input_pdf)
    out_dir = azione_dividi(dividi_args, script_dir)

    # If an edition name was given, renames the output folder
    # (e.g. attestati_output -> CORSO10) before organizing it, so the
    # result is already ready for a later 'riorganizza'.
    if args.edizione:
        nuova_cartella = out_dir.parent / args.edizione
        if nuova_cartella.exists():
            log.error(
                f"The folder '{nuova_cartella}' already exists: remove it or rename it "
                f"manually before relaunching, to avoid mixing different editions."
            )
            sys.exit(1)
        out_dir.rename(nuova_cartella)
        out_dir = nuova_cartella
        log.info(f"\nFolder renamed to: {out_dir.name}")
    else:
        log.info(
            "\nTip: use --edizione CORSO10 (with the right number) to rename "
            "the output folder right away and get better per-edition matching "
            "and an easier future use of 'riorganizza'."
        )

    log.info("\n### Step 2/2: organization by company ###")

    organizza_args = argparse.Namespace(
        cartella_edizione=str(out_dir),
        file_excel=args.file_excel,
        copia=args.copia,
    )
    azione_organizza(organizza_args, script_dir)


# ========================================================================
# ENTRY POINT
# ========================================================================
def main() -> None:
    parser = argparse.ArgumentParser(
        prog="certificate_manager.py",
        description="Complete management of CORSO certificates: dividi, organizza, riorganizza.",
    )
    sub = parser.add_subparsers(dest="comando")

    sp_dividi = sub.add_parser(
        "dividi",
        help="Splits a mass-print PDF into one file per person (front+back joined).",
    )
    sp_dividi.add_argument(
        "input_pdf", nargs="?", default=None,
        help="Path of the PDF to split. If omitted, looks for a PDF in the script folder automatically.",
    )

    sp_organizza = sub.add_parser(
        "organizza",
        help="Sorts the PDFs of an edition into company subfolders using an Excel rubrica.",
    )
    sp_organizza.add_argument(
        "cartella_edizione", nargs="?", default=None,
        help="Edition folder (e.g. 'CORSO10') or folder with multiple CORSO* subfolders. If omitted, searched automatically.",
    )
    sp_organizza.add_argument(
        "file_excel", nargs="?", default=None,
        help="Path of the rubrica Excel file (.xls or .xlsx). If omitted, searched automatically.",
    )
    sp_organizza.add_argument(
        "--copia", action="store_true",
        help="Copies the PDFs into the subfolders instead of moving them.",
    )

    sp_riorganizza = sub.add_parser(
        "riorganizza",
        help="Rebuilds multiple editions already organized by company into a company/edition view.",
    )
    sp_riorganizza.add_argument(
        "cartelle", nargs="*",
        help="Container folder with multiple CORSO* subfolders, or a list of edition folders. If omitted, searched automatically.",
    )
    sp_riorganizza.add_argument("--output", default=None, help="Destination folder.")
    sp_riorganizza.add_argument(
        "--sposta", action="store_true",
        help="Moves the files instead of copying them.",
    )
    sp_riorganizza.add_argument(
        "--rubrica", default=None,
        help="Rubrica Excel file (.xls/.xlsx): if specified, generates a Word register "
             "for each company in the output folder.",
    )

    sp_registri = sub.add_parser(
        "registri",
        help="Generates the Word registers from template + rubrica ONLY (no CORSO* folders/PDFs).",
    )
    sp_registri.add_argument(
        "rubrica", nargs="?", default=None,
        help="Rubrica Excel file (.xls/.xlsx). If omitted, searched automatically "
             "in the script folder.",
    )
    sp_registri.add_argument(
        "--output", default=None,
        help="Destination folder (default: <script folder>/REGISTRI).",
    )

    sp_tutto = sub.add_parser(
        "tutto",
        help="Full pipeline: splits the PDF and immediately organizes it by company (a single edition).",
    )
    sp_tutto.add_argument("input_pdf", nargs="?", default=None, help="Mass-print PDF to split.")
    sp_tutto.add_argument("file_excel", nargs="?", default=None, help="Excel rubrica of that edition.")
    sp_tutto.add_argument(
        "--edizione", default=None,
        help="Edition name, e.g. CORSO10: renames the output folder before organizing it.",
    )
    sp_tutto.add_argument("--copia", action="store_true", help="Copies instead of moving the PDFs into the company subfolders.")

    args = parser.parse_args()

    if args.comando is None:
        parser.print_help()
        sys.exit(0)

    script_dir = Path(__file__).parent

    if args.comando == "dividi":
        azione_dividi(args, script_dir)
    elif args.comando == "organizza":
        azione_organizza(args, script_dir)
    elif args.comando == "riorganizza":
        azione_riorganizza(args, script_dir)
    elif args.comando == "registri":
        azione_registri(args, script_dir)
    elif args.comando == "tutto":
        azione_tutto(args, script_dir)


if __name__ == "__main__":
    main()
