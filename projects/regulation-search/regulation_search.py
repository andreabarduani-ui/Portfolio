#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INFORMATION EXTRACTOR - FUNDED TRAINING DOCUMENTS
==================================================

Interactive tool to extract information from manuals, calls and guidelines
related to funded training (from the main interprofessional
funds, etc.).

The script:
- Indexes the PDF page by page
- Allows searches by predefined categories (variations, rendicontazione,
  eligible costs, training modes, state aid, etc.)
- Supports free keyword searches
- ALWAYS returns the page number and the context of the match

Dependencies: pdfplumber (preferred) or pypdf
    pip install pdfplumber

Usage:
    python regulation_search.py [path_to_pdf]
    # or without arguments, the script asks for the path
"""

import os
import re
import sys
import textwrap
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# -------- PDF library import with fallback --------
try:
    import pdfplumber
    PDF_LIB = "pdfplumber"
except ImportError:
    try:
        from pypdf import PdfReader
        PDF_LIB = "pypdf"
    except ImportError:
        print("ERROR: at least one PDF library is required.")
        print("Install one with: pip install pdfplumber")
        sys.exit(1)


# =============================================================================
# PREDEFINED CATEGORIES FOR FUNDED TRAINING
# =============================================================================
# Each category is a list of keywords/patterns (case-insensitive).
# Add to or modify these categories to adapt the script to other
# Interprofessional Funds or document types.

CATEGORIE: Dict[str, List[str]] = {
    "variazioni_progetto": [
        "variazion", "modific", "richiesta di variazione", "modello 4",
        "subordinat", "autorizzazion", "in itinere",
        "rinuncia", "subentro", "sostituzione", "rimodulazione",
        "proroga", "differimento",
    ],
    "modalita_formative": [
        "modalità formativ", "aula", "fad", "e-learning", "teleformazione",
        "affiancamento", "training on the job", "coaching", "asincron",
        "sincron", "webinar", "videoconferenza", "lms",
    ],
    "rendicontazione": [
        "rendiconto", "rendicontazione", "consuntivo", "preventivo",
        "scostament", "variant", "saldo", "documenti giustificativ",
        "tracciabilità", "estratto conto", "quietanza",
    ],
    "costi_ammissibili": [
        "costi ammissibili", "costo ammissibile", "costi non ammissibili",
        "spesa ammissibile", "massimal", "parametro ora", "costo orario",
        "iva", "indiretti", "diretti", "forfait", "forfetar",
    ],
    "aiuti_di_stato": [
        "aiuti di stato", "aiuto di stato", "de minimis", "regolamento ue",
        "regolamento (ue)", "651/2014", "1407/2013", "intensità",
        "rna", "registro nazionale", "cor",
    ],
    "personale_docenza": [
        "docenza", "docente", "tutor", "tutoraggio", "personale interno",
        "personale esterno", "fascia a", "fascia b", "fascia c",
        "consulent", "coordinament", "curriculum", "incarico professional",
    ],
    "presentazione_progetto": [
        "presentazione", "formulario", "allegato", "ammissibilità formal",
        "valutazione tecnica", "punteggio", "graduatoria", "sportello",
        "parti sociali", "condivisione", "ppss",
    ],
    "controlli_verifiche": [
        "verifica", "controllo", "ispezion", "verifich", "i livello",
        "ii livello", "ex post", "in itinere", "verbale", "controdeduzion",
        "campionament", "sanzion", "revoca",
    ],
    "voucher_individuale": [
        "voucher", "formazione individuale", "ente erogatore",
        "alta formazione", "formazione specialistic",
    ],
    "tempistiche_scadenze": [
        "giorni", "termine", "scadenza", "tempistica", "data di fine",
        "data di inizio", "365 giorni", "545 giorni", "180 giorni",
        "30 giorni", "60 giorni", "90 giorni",
    ],
    "ati_ats_partenariato": [
        "a.t.i.", "a.t.s.", "ati", "ats", "associazione temporanea",
        "partner", "partenariato", "capofila", "mandatari",
        "delega", "delegato", "soggetti terzi",
    ],
    "imprese_beneficiarie": [
        "impresa beneficiaria", "imprese beneficiarie", "destinatari",
        "lavoratori", "dipendenti", "dirigenti", "apprendist",
        "cassintegrat", "stagional",
    ],
    "fideiussione_anticipo": [
        "fideiussione", "fidejussione", "fidejussori", "polizza",
        "anticipazion", "anticipo", "svincolo",
    ],
}

# User-friendly labels for the menu
ETICHETTE_CATEGORIE: Dict[str, str] = {
    "variazioni_progetto": "Variations / deadline extensions / withdrawals / substitutions",
    "modalita_formative": "Training delivery modes (classroom, FAD, e-learning...)",
    "rendicontazione": "Rendicontazione (reporting) and supporting documents",
    "costi_ammissibili": "Eligible costs, parameters and maximum amounts",
    "aiuti_di_stato": "State aid (de minimis, Reg. 651/2014, RNA)",
    "personale_docenza": "Staff: trainers, tutors, consultants",
    "presentazione_progetto": "Project submission (application form, attachments, evaluation)",
    "controlli_verifiche": "Controls and checks (first level, second level, inspections)",
    "voucher_individuale": "Vouchers / individual training",
    "tempistiche_scadenze": "Project timelines and deadlines",
    "ati_ats_partenariato": "ATI / ATS / partnership / delegation",
    "imprese_beneficiarie": "Beneficiary companies and recipients",
    "fideiussione_anticipo": "Bank guarantee and advance payment requests",
}


# =============================================================================
# PDF TEXT EXTRACTION
# =============================================================================

def estrai_testo_pdf(path: str) -> List[Tuple[int, str]]:
    """
    Extracts the text page by page.
    Returns a list of (page_number_1based, page_text) tuples.
    """
    pagine: List[Tuple[int, str]] = []
    if PDF_LIB == "pdfplumber":
        with pdfplumber.open(path) as pdf:
            for idx, page in enumerate(pdf.pages, start=1):
                testo = page.extract_text() or ""
                pagine.append((idx, testo))
    else:
        reader = PdfReader(path)
        for idx, page in enumerate(reader.pages, start=1):
            testo = page.extract_text() or ""
            pagine.append((idx, testo))
    return pagine


# =============================================================================
# SEARCH
# =============================================================================

def trova_corrispondenze(
    pagine: List[Tuple[int, str]],
    parole_chiave: List[str],
    contesto_caratteri: int = 250,
    max_risultati_per_pagina: int = 3,
) -> List[Dict]:
    """
    Searches for the keywords (case-insensitive) in the pages.
    Returns a list of dicts with: pagina (page), parola_trovata (found keyword), contesto (context).
    """
    risultati: List[Dict] = []
    # Compile patterns: \b does not work well with Italian accents -> use lookaround
    patterns = [
        (kw, re.compile(re.escape(kw), re.IGNORECASE))
        for kw in parole_chiave
    ]

    for num_pag, testo in pagine:
        if not testo:
            continue
        # Counts matches per page and extracts context
        trovati_pag = 0
        gia_visti_offset: List[int] = []  # avoids nearby duplicates
        for kw, pat in patterns:
            for match in pat.finditer(testo):
                if trovati_pag >= max_risultati_per_pagina:
                    break
                start = match.start()
                # Skip if too close to an already taken match (within 80 chars)
                if any(abs(start - x) < 80 for x in gia_visti_offset):
                    continue
                gia_visti_offset.append(start)

                inizio = max(0, start - contesto_caratteri // 2)
                fine = min(len(testo), match.end() + contesto_caratteri // 2)
                contesto = testo[inizio:fine]
                # Clean up multiple line breaks
                contesto = re.sub(r"\s+", " ", contesto).strip()

                risultati.append({
                    "pagina": num_pag,
                    "parola_trovata": kw,
                    "contesto": contesto,
                    "posizione": start,
                })
                trovati_pag += 1
            if trovati_pag >= max_risultati_per_pagina:
                break
    return risultati


def stampa_risultati(risultati: List[Dict], titolo: str) -> None:
    """Formatted printing of the results."""
    print()
    print("=" * 78)
    print(f"  RESULTS: {titolo}")
    print(f"  Matches found: {len(risultati)}")
    print("=" * 78)
    if not risultati:
        print("\n  No matches found.\n")
        return

    # Group by page
    per_pagina: Dict[int, List[Dict]] = {}
    for r in risultati:
        per_pagina.setdefault(r["pagina"], []).append(r)

    for pag in sorted(per_pagina.keys()):
        items = per_pagina[pag]
        print(f"\n  📄 PAGE {pag}  ({len(items)} match{'' if len(items)==1 else 'es'})")
        print("  " + "-" * 74)
        for it in items:
            kw = it["parola_trovata"]
            ctx = it["contesto"]
            # Highlights (simply — uppercase) the found keyword in the context
            ctx_evid = re.sub(
                re.escape(kw),
                lambda m: f">>>{m.group(0).upper()}<<<",
                ctx,
                flags=re.IGNORECASE,
                count=1,
            )
            wrapped = textwrap.fill(
                ctx_evid,
                width=72,
                initial_indent="    ",
                subsequent_indent="    ",
            )
            print(f"\n    🔍 «{kw}»")
            print(wrapped)
    print()


# =============================================================================
# INTERACTIVE MENU
# =============================================================================

def chiedi_path_pdf() -> str:
    """Asks for the path if not passed as an argument."""
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = input("Path of the PDF to analyze: ").strip().strip('"').strip("'")
    if not path:
        print("No path provided.")
        sys.exit(1)
    if not Path(path).is_file():
        print(f"File not found: {path}")
        sys.exit(1)
    return path


def menu_principale() -> str:
    """Shows the menu and returns the choice."""
    print()
    print("╔" + "═" * 76 + "╗")
    print("║" + "  WHAT DO YOU WANT TO EXTRACT FROM THE DOCUMENT?".ljust(76) + "║")
    print("╚" + "═" * 76 + "╝")
    print()
    chiavi = list(ETICHETTE_CATEGORIE.keys())
    for i, k in enumerate(chiavi, start=1):
        print(f"  {i:>2}. {ETICHETTE_CATEGORIE[k]}")
    print(f"  {len(chiavi)+1:>2}. 🔎 Free search (custom keywords)")
    print(f"  {len(chiavi)+2:>2}. 📑 Document index (first lines of each page)")
    print(f"  {len(chiavi)+3:>2}. 📊 Document statistics")
    print(f"   0. Exit")
    print()
    scelta = input("Choice: ").strip()
    return scelta


def gestisci_scelta(
    scelta: str,
    pagine: List[Tuple[int, str]],
) -> bool:
    """Executes the chosen action. Returns True to continue, False to exit."""
    chiavi = list(ETICHETTE_CATEGORIE.keys())

    if scelta == "0":
        print("\nBye!\n")
        return False

    # Search by category
    if scelta.isdigit() and 1 <= int(scelta) <= len(chiavi):
        cat = chiavi[int(scelta) - 1]
        parole = CATEGORIE[cat]
        risultati = trova_corrispondenze(pagine, parole)
        stampa_risultati(risultati, ETICHETTE_CATEGORIE[cat])
        return True

    # Free search
    if scelta == str(len(chiavi) + 1):
        kw_input = input(
            "\nEnter keywords separated by commas "
            "(e.g.: proroga, rinuncia, modello 4): "
        ).strip()
        if not kw_input:
            print("No keyword entered.")
            return True
        parole = [p.strip() for p in kw_input.split(",") if p.strip()]
        risultati = trova_corrispondenze(pagine, parole)
        stampa_risultati(risultati, f"Free search: {', '.join(parole)}")
        return True

    # Quick index
    if scelta == str(len(chiavi) + 2):
        print("\n" + "=" * 78)
        print("  QUICK INDEX (first lines of each page)")
        print("=" * 78)
        for num_pag, testo in pagine:
            prima_riga = ""
            for riga in (testo or "").splitlines():
                riga_pulita = riga.strip()
                if riga_pulita and len(riga_pulita) > 5:
                    prima_riga = riga_pulita[:90]
                    break
            print(f"  Pg. {num_pag:>3} | {prima_riga}")
        print()
        return True

    # Statistics
    if scelta == str(len(chiavi) + 3):
        totale_pagine = len(pagine)
        totale_caratteri = sum(len(t) for _, t in pagine)
        totale_parole = sum(len((t or "").split()) for _, t in pagine)
        print("\n" + "=" * 78)
        print("  DOCUMENT STATISTICS")
        print("=" * 78)
        print(f"  Total pages       : {totale_pagine}")
        print(f"  Total characters  : {totale_caratteri:,}".replace(",", "."))
        print(f"  Total words       : {totale_parole:,}".replace(",", "."))
        print(f"  Avg words/page    : {totale_parole // max(totale_pagine,1)}")
        print()
        return True

    print(f"Invalid choice: {scelta!r}")
    return True


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    print()
    print("╔" + "═" * 76 + "╗")
    titolo = "EXTRACTOR - FUNDED TRAINING DOCUMENTS"
    print("║" + titolo.center(76) + "║")
    sotto = f"(PDF library in use: {PDF_LIB})"
    print("║" + sotto.center(76) + "║")
    print("╚" + "═" * 76 + "╝")

    path = chiedi_path_pdf()
    print(f"\n📥 Extracting text from: {path}")
    pagine = estrai_testo_pdf(path)
    print(f"✅ {len(pagine)} pages indexed.")

    while True:
        scelta = menu_principale()
        try:
            continua = gestisci_scelta(scelta, pagine)
        except KeyboardInterrupt:
            print("\nInterrupted by user.\n")
            break
        if not continua:
            break


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrupted.\n")
