"""
certificate_manager.py
=======================

Script unico per la gestione completa degli attestati CORSO, con tre
sottocomandi che corrispondono ai tre passaggi del flusso di lavoro:

  1. dividi       Divide un PDF di stampa massiva (fronte+retro alternati)
                   in un file per persona, unendo fronte e retro.

  2. organizza     Smista i PDF di UNA edizione (es. cartella "CORSO10") in
                   sottocartelle per azienda, incrociando i nomi con una
                   rubrica Excel (colonne: Cognome, Nome, NomeCorso, Azienda).

  3. riorganizza    Ricompone PIÙ edizioni già organizzate per azienda
                   (output del punto 2) in una vista invertita:
                   primo livello = azienda, secondo livello = edizione.
                   Con --rubrica genera anche, in ogni cartella azienda,
                   un file Word "Registro attestati <corso> - <azienda>.docx"
                   con l'elenco dei corsisti tratto dalla rubrica.

  (bonus) tutto     Esegue in sequenza "dividi" + "organizza" per una singola
                   edizione, comodo quando hai già pronti sia il PDF di stampa
                   massiva sia la rubrica Excel di quell'edizione.

USO DA VS CODE / TERMINALE
----------------------------
    py certificate_manager.py dividi ["input.pdf"]
    py certificate_manager.py organizza ["CORSO10"] ["rubrica.xls"] [--copia]
    py certificate_manager.py riorganizza ["cartelle"...] [--output DIR] [--sposta] [--rubrica "rubrica.xls"]
    py certificate_manager.py tutto ["input.pdf"] ["rubrica.xls"] [--edizione CORSO10] [--copia]

Ogni argomento posizionale è OPZIONALE: se omesso, lo script prova a trovarlo
automaticamente nella cartella in cui si trova (utile per lanciare con F5 in
VS Code senza configurare nulla). Lancia senza sottocomando per vedere questo
riepilogo.

REQUISITI
---------
    py -m pip install -r requirements.txt
    (pypdf per leggere i PDF; pandas + xlrd/openpyxl per leggere la rubrica
    Excel; python-docx per generare i registri Word — servono solo ai
    sottocomandi 'organizza', 'riorganizza' e 'tutto', ma installarli tutti
    insieme evita sorprese quando passi da un sottocomando all'altro)
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
# FUNZIONI CONDIVISE
# ========================================================================
_ACCENTI = {
    "à": "a", "á": "a", "À": "A", "Á": "A",
    "è": "e", "é": "e", "È": "E", "É": "E",
    "ì": "i", "í": "i", "Ì": "I", "Í": "I",
    "ò": "o", "ó": "o", "Ò": "O", "Ó": "O",
    "ù": "u", "ú": "u", "Ù": "U", "Ú": "U",
}


def estrai_nome_da_testo(testo: str) -> str | None:
    """Estrae il nome del partecipante dal testo di una pagina fronte,
    cercando il pattern 'Si attesta che' seguito dal nome sulla riga
    successiva. Restituisce None se non trovato."""
    match = re.search(r"Si attesta che\s*\n?\s*(.+?)\s*\n", testo)
    return match.group(1).strip() if match else None


def normalizza(nome: str) -> str:
    """Uniforma un nome per il confronto: maiuscolo, senza accenti,
    senza punteggiatura, senza spazi doppi."""
    for accentata, semplice in _ACCENTI.items():
        nome = nome.replace(accentata, semplice)
    nome = re.sub(r"[^A-Za-z ]", " ", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome.upper()


def slug_nome_persona(nome: str) -> str:
    """Converte un nome in un formato adatto a un filename (Nome_Cognome)."""
    nome_norm = normalizza(nome)
    parti = [p.capitalize() for p in nome_norm.split()]
    return "_".join(parti) if parti else "Sconosciuto"


def sanitizza_nome_cartella(nome: str) -> str:
    """Rimuove caratteri non ammessi nei nomi di cartella Windows."""
    nome = re.sub(r'[<>:"/\\|?*]', "", nome).strip()
    nome = re.sub(r"\s+", " ", nome)
    return nome or "AZIENDA_SCONOSCIUTA"


# ========================================================================
# SEZIONE 1 — DIVIDI: PDF di stampa massiva -> un file per persona
# ========================================================================
def trova_pdf_input(cartella: Path) -> Path | None:
    """Cerca automaticamente un PDF nella cartella (per l'avvio con F5
    senza argomenti). Se ce n'è più di uno, chiede di specificarlo."""
    candidati = sorted(cartella.glob("*.pdf"))
    if len(candidati) == 1:
        return candidati[0]
    if len(candidati) > 1:
        log.warning("Trovati più PDF nella cartella. Specifica quale usare:")
        for c in candidati:
            log.warning(f"  - {c.name}")
    return None


def dividi_pdf(input_path: Path) -> Path:
    """Divide il PDF di input in un file per persona (fronte+retro uniti).
    Restituisce il percorso della cartella di output."""
    reader = PdfReader(str(input_path))
    n_pagine = len(reader.pages)

    if n_pagine == 0:
        raise ValueError("Il PDF non contiene pagine.")
    if n_pagine % 2 != 0:
        raise ValueError(
            f"Il PDF ha {n_pagine} pagine (numero dispari). "
            "Mi aspetto coppie fronte+retro: controlla il file di input."
        )

    n_persone = n_pagine // 2
    out_dir = input_path.parent / "attestati_output"
    out_dir.mkdir(exist_ok=True)

    log.info(f"Trovate {n_pagine} pagine -> {n_persone} partecipanti.\n")

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

        etichetta = nome if nome else "(nome non riconosciuto)"
        log.info(f"  [{i + 1:02d}] {etichetta} -> {out_path.name}")

    log.info(f"\nCompletato. File salvati in: {out_dir}")
    return out_dir


def azione_dividi(args: argparse.Namespace, script_dir: Path) -> Path:
    if args.input_pdf:
        input_path = Path(args.input_pdf)
    else:
        input_path = trova_pdf_input(script_dir)
        if input_path is None:
            log.error(
                "Nessun PDF specificato e nessun PDF univoco trovato nella cartella.\n"
                "Trascina il PDF nella cartella del progetto oppure lancia:\n"
                '  py certificate_manager.py dividi "nomefile.pdf"'
            )
            sys.exit(1)
        log.info(f"Nessun argomento fornito: uso il PDF trovato automaticamente -> {input_path.name}\n")

    if not input_path.exists():
        log.error(f"File non trovato: {input_path}")
        sys.exit(1)

    try:
        return dividi_pdf(input_path)
    except ValueError as e:
        log.error(f"Errore: {e}")
        sys.exit(1)


# ========================================================================
# SEZIONE 2 — ORGANIZZA: un'edizione -> sottocartelle per azienda
# ========================================================================
def carica_rubrica(excel_path: Path) -> pd.DataFrame:
    """Carica il file Excel e verifica le colonne necessarie."""
    df = pd.read_excel(excel_path)

    colonne_richieste = {"Cognome", "Nome", "NomeCorso", "Azienda"}
    mancanti = colonne_richieste - set(df.columns)
    if mancanti:
        raise ValueError(
            f"Nel file Excel mancano le colonne: {', '.join(mancanti)}. "
            f"Colonne trovate: {', '.join(df.columns)}"
        )

    df = df.dropna(subset=["Cognome", "Nome", "Azienda"]).copy()
    df["nome_completo_norm"] = (df["Nome"].astype(str) + " " + df["Cognome"].astype(str)).apply(normalizza)
    df["edizione"] = df["NomeCorso"].astype(str).apply(estrai_numero_edizione_da_corso)
    return df


def estrai_numero_edizione_da_corso(testo: str) -> str | None:
    """Estrae il numero di edizione CORSO da una stringa tipo
    'CORSO EDIZIONE 10 - ...'."""
    match = re.search(r"EDIZIONE\s+(\d+)", testo, re.IGNORECASE)
    return match.group(1) if match else None


def estrai_numero_edizione_da_cartella(nome_cartella: str) -> str | None:
    """Estrae il numero edizione dal nome della cartella, es. 'CORSO10' -> '10'."""
    match = re.search(r"CORSO\s*(\d+)", nome_cartella, re.IGNORECASE)
    return match.group(1) if match else None


def estrai_nome_da_pdf(pdf_path: Path) -> str | None:
    """Estrae il nome del partecipante dal testo della prima pagina (fronte),
    con fallback sul nome del file se il testo non è estraibile."""
    try:
        reader = PdfReader(str(pdf_path))
        testo = reader.pages[0].extract_text() or ""
    except Exception as e:
        log.warning(f"  Impossibile leggere il testo di {pdf_path.name}: {e}")
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
    """Cerca l'azienda corrispondente al nome estratto dal PDF.
    Ritorna (azienda_o_None, metodo)."""
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
        log.info(f"\n=== {cartella.name} (edizione CORSO {edizione}) ===")
    else:
        log.info(f"\n=== {cartella.name} (edizione non riconosciuta dal nome cartella) ===")

    pdf_files = sorted(cartella.glob("*.pdf"))
    if not pdf_files:
        log.info("  Nessun PDF trovato in questa cartella (forse già organizzati?).")
        return

    n_trovati = 0
    n_non_trovati = 0

    for pdf_path in pdf_files:
        nome = estrai_nome_da_pdf(pdf_path)
        if not nome:
            log.warning(f"  [!] Impossibile determinare il nome per {pdf_path.name}")
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
            log.warning(f"  [X] {nome or pdf_path.stem} -> NESSUNA AZIENDA TROVATA")
            n_non_trovati += 1

        cartella_azienda.mkdir(exist_ok=True)
        destinazione = cartella_azienda / pdf_path.name

        if copia:
            shutil.copy2(pdf_path, destinazione)
        else:
            shutil.move(str(pdf_path), str(destinazione))

    log.info(f"  --> {n_trovati} assegnati, {n_non_trovati} da verificare manualmente.")


def trova_cartella_singola_edizione(base: Path) -> Path | None:
    """Cerca una cartella che inizia con 'CORSO' nella cartella base,
    oppure restituisce base stessa se contiene già dei PDF."""
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
                "Non ho trovato automaticamente una cartella CORSO* con i PDF.\n"
                "Specifica il percorso, ad esempio:\n"
                '  py certificate_manager.py organizza "CORSO10" "rubrica.xls"'
            )
            sys.exit(1)
        log.info(f"Cartella edizione trovata automaticamente: {input_path.name}")

    if not input_path.exists():
        log.error(f"Cartella non trovata: {input_path}")
        sys.exit(1)

    if args.file_excel:
        excel_path = Path(args.file_excel)
    else:
        excel_path = trova_file_excel(script_dir)
        if excel_path is None:
            log.error(
                "Non ho trovato automaticamente un file Excel (.xls/.xlsx) nella cartella dello script.\n"
                "Specifica il percorso, ad esempio:\n"
                '  py certificate_manager.py organizza "CORSO10" "rubrica.xls"'
            )
            sys.exit(1)
        log.info(f"File Excel trovato automaticamente: {excel_path.name}")

    if not excel_path.exists():
        log.error(f"File Excel non trovato: {excel_path}")
        sys.exit(1)

    try:
        rubrica = carica_rubrica(excel_path)
    except ValueError as e:
        log.error(f"Errore nella lettura della rubrica: {e}")
        sys.exit(1)

    log.info(f"Rubrica caricata: {len(rubrica)} nominativi.")

    if any(input_path.glob("*.pdf")):
        cartelle_da_processare = [input_path]
    else:
        cartelle_da_processare = sorted(
            p for p in input_path.iterdir() if p.is_dir() and p.name.upper().startswith("CORSO")
        )
        if not cartelle_da_processare:
            log.error(f"Nella cartella '{input_path}' non ho trovato né PDF né sottocartelle CORSO*.")
            sys.exit(1)

    for cartella in cartelle_da_processare:
        organizza_cartella_edizione(cartella, rubrica, copia=args.copia)

    log.info("\nCompletato.")


# ========================================================================
# SEZIONE 3 — RIORGANIZZA: più edizioni -> vista azienda/edizione
# ========================================================================
def trova_tutte_cartelle_edizione(base: Path) -> list[Path]:
    """Trova tutte le sottocartelle il cui nome inizia con 'CORSO' dentro 'base'.
    Se 'base' stessa inizia con CORSO, la restituisce direttamente."""
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
        log.info(f"\n=== Edizione: {edizione} (da {cartella_edizione}) ===")

        sottocartelle_azienda = sorted(p for p in cartella_edizione.iterdir() if p.is_dir())

        if not sottocartelle_azienda:
            log.warning(
                f"  Nessuna sottocartella azienda trovata in '{edizione}'. "
                f"Hai già eseguito 'organizza' su questa cartella?"
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
    log.info("RIEPILOGO")
    log.info("=" * 60)
    for azienda in sorted(riepilogo):
        edizioni = ", ".join(sorted(riepilogo[azienda]))
        log.info(f"  {azienda}: {edizioni}")

    azione = "spostati" if sposta else "copiati"
    log.info(
        f"\n{totale_spostati} file {azione} in totale, "
        f"da {totale_edizioni} edizioni, in {len(riepilogo)} aziende diverse."
    )
    log.info(f"Cartella di output: {output_root}")


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
            "Nessuna cartella edizione (CORSO*) trovata.\n"
            "Specifica le cartelle, ad esempio:\n"
            '  py certificate_manager.py riorganizza "CORSO10" "CORSO11" "CORSO12" "CORSO13" "CORSO14"\n'
            "oppure una cartella contenitore:\n"
            '  py certificate_manager.py riorganizza "C:\\percorso\\Attestati"'
        )
        sys.exit(1)

    log.info(f"Cartelle edizione trovate ({len(cartelle_edizione)}):")
    for c in cartelle_edizione:
        log.info(f"  - {c}")

    output_root = Path(args.output) if args.output else cartelle_edizione[0].parent / "ATTESTATI_PER_AZIENDA"

    log.info(f"\nCartella di output: {output_root}")
    log.info(f"Modalità: {'SPOSTAMENTO' if args.sposta else 'COPIA (originali intatti)'}")

    riorganizza(cartelle_edizione, output_root, sposta=args.sposta)

    # ---- FASE AGGIUNTIVA: generazione registri Word per azienda ----
    # Avviene alla fine del processo di riorganizzazione: per ogni cartella
    # azienda presente in output_root genera un file .docx con l'elenco dei
    # corsisti tratto dalla rubrica Excel.
    rubrica_excel = getattr(args, "rubrica", None)
    if rubrica_excel:
        excel_path = Path(rubrica_excel)
        if not excel_path.exists():
            log.error(f"\nFile Excel rubrica non trovato: {excel_path}")
        else:
            try:
                rubrica = carica_rubrica(excel_path)
                log.info(f"\nRubrica caricata per i registri: {len(rubrica)} nominativi.")
                log.info("\n### Generazione registri Word per azienda ###")
                genera_registri_word(output_root, rubrica)
            except ValueError as e:
                log.error(f"\nErrore nella lettura della rubrica: {e}")

    log.info("\nCompletato.")


def azione_registri(args: argparse.Namespace, script_dir: Path) -> None:
    """Genera i registri Word SOLO da modello + rubrica, senza bisogno di
    cartelle CORSO* o PDF. Crea una cartella per ogni azienda presente nella
    rubrica e vi salva il relativo .docx di registro."""
    if args.rubrica:
        excel_path = Path(args.rubrica)
    else:
        excel_path = trova_file_excel(script_dir)
        if excel_path is None:
            log.error(
                "Nessun file Excel (.xls/.xlsx) trovato nella cartella dello script.\n"
                "Specifica il percorso, ad esempio:\n"
                '  py certificate_manager.py registri "rubrica.xls"'
            )
            sys.exit(1)
        log.info(f"File Excel trovato automaticamente: {excel_path.name}")

    if not excel_path.exists():
        log.error(f"File Excel non trovato: {excel_path}")
        sys.exit(1)

    try:
        rubrica = carica_rubrica(excel_path)
    except ValueError as e:
        log.error(f"Errore nella lettura della rubrica: {e}")
        sys.exit(1)

    output_dir = Path(args.output) if args.output else script_dir / "REGISTRI"
    output_dir.mkdir(parents=True, exist_ok=True)

    log.info(f"Rubrica caricata: {len(rubrica)} nominativi.")
    log.info(f"Cartella di output: {output_dir}")
    log.info("\n### Generazione registri Word da rubrica ###")
    genera_registri_da_rubrica(rubrica, output_dir)
    log.info("\nCompletato.")


# ========================================================================
# SEZIONE 4 — REGISTRI WORD: genera un registro per ogni azienda
# ========================================================================
def genera_registri_word(output_root: Path, rubrica: pd.DataFrame) -> None:
    """Genera un file .docx di 'Registro attestati' per ogni azienda
    presente nella struttura di output (output_root/<azienda>/...).

    Per ogni azienda:
      - recupera dalla rubrica tutti i corsisti di quell'azienda (match
        normalizzato sul nome, con fallback fuzzy);
      - crea un documento Word con:
          * Tabella intestazione:  Descrizione | Attestati di Frequenza azienda <NOME>
          * Tabella elenco:        N° | Cognome e Nome | Data | Firma per Ricezione
      - salva il file nella cartella dell'azienda come:
          'Registro attestati <NomeCorso> - <Azienda>.docx'

    Il formato riprende quello del file di riferimento
    'Registro attestati Cliente B.doc'.
    """
    # Prepara una colonna normalizzata dell'azienda per il matching
    rubrica = rubrica.copy()
    rubrica["azienda_norm"] = rubrica["Azienda"].astype(str).apply(normalizza)

    aziende_dirs = sorted(p for p in output_root.iterdir() if p.is_dir())

    if not aziende_dirs:
        log.warning("  Nessuna cartella azienda trovata in output: registro non generato.")
        return

    n_registri = 0
    for cartella_azienda in aziende_dirs:
        azienda_cartella = cartella_azienda.name
        azienda_norm = normalizza(azienda_cartella)

        # Cerca i corsisti di questa azienda (match normalizzato, poi fuzzy)
        corsisti = rubrica[rubrica["azienda_norm"] == azienda_norm]
        if corsisti.empty:
            candidati = rubrica["azienda_norm"].unique().tolist()
            vicini = get_close_matches(azienda_norm, candidati, n=1, cutoff=0.85)
            if vicini:
                corsisti = rubrica[rubrica["azienda_norm"] == vicini[0]]
            else:
                log.warning(
                    f"  [!] Nessun corsista in rubrica per '{azienda_cartella}': "
                    f"registro non generato."
                )
                continue

        azienda_rubrica = str(corsisti.iloc[0]["Azienda"]).strip()

        # Se un'azienda ha corsisti in più corsi, li raggruppa per corso:
        # genera un registro separato per ogni corso.
        for nome_corso, gruppo in corsisti.groupby("NomeCorso"):
            nome_corso = str(nome_corso).strip()
            _genera_registro_word(
                cartella_azienda=cartella_azienda,
                azienda=azienda_rubrica,
                nome_corso=nome_corso,
                corsisti=gruppo,
            )
            n_registri += 1

    log.info(f"\nGenerati {n_registri} registro/i Word.")


def genera_registri_da_rubrica(rubrica: pd.DataFrame, output_dir: Path) -> None:
    """Genera un file .docx di 'Registro attestati' per ogni azienda presente
    NELLA RUBRICA (non serve la cartella CORSO* con i PDF).

    Per ogni azienda:
      - recupera dalla rubrica tutti i corsisti di quell'azienda;
      - crea una cartella <output_dir>/<azienda>/ ;
      - genera il registro Word (su modello_registro.docx) sostituendo
        azienda, corso, codice CORSO ed elenco corsisti.
    """
    n_registri = 0
    for azienda, gruppo in rubrica.groupby("Azienda"):
        azienda = str(azienda).strip()
        cartella_azienda = output_dir / sanitizza_nome_cartella(azienda)
        cartella_azienda.mkdir(parents=True, exist_ok=True)

        # Se un'azienda ha corsisti in piu' corsi, un registro per ogni corso.
        for nome_corso, sottogruppo in gruppo.groupby("NomeCorso"):
            _genera_registro_word(
                cartella_azienda=cartella_azienda,
                azienda=azienda,
                nome_corso=str(nome_corso).strip(),
                corsisti=sottogruppo,
            )
            n_registri += 1

    log.info(f"\nGenerati {n_registri} registro/i Word.")


def tronca_per_percorso(testo: str, cartella: Path, suffix_base: str = "") -> str:
    """Tronca 'testo' in modo che il percorso completo del file .docx generato
    nella 'cartella' rispetti il limite MAX_PATH (260 char) di Windows.

    Lascia un margine di sicurezza per suffisso azienda, separatori e '.docx'.
    """
    # Caratteri riservati al suffisso: separatore " - " + azienda (stimata
    # uguale al nome cartella) + ".docx". Usiamo un margine ampio per sicurezza.
    azienda_len = len(cartella.name)
    margine = len(suffix_base) + len(" - ") + azienda_len + len(".docx")
    limite = 259 - len(str(cartella)) - 1 - margine  # -1 per il separatore '\'
    if limite < 10:
        limite = 10
    if len(testo) > limite:
        testo = testo[:limite].rstrip()
    return testo


def _grassetto_riga(riga) -> None:
    """Imposta in grassetto tutto il testo di una riga di tabella."""
    for cella in riga.cells:
        for paragrafo in cella.paragraphs:
            for run in paragrafo.runs:
                run.bold = True


def _imposta_font_cella(cell, font_name: str, size_pt: float, bold: bool = False) -> None:
    """Imposta nome font, dimensione e grassetto su tutti i run del
    paragrafo (e del primo paragrafo) della cella. Imposta anche il font
    per i caratteri complessi (rPr/rFonts cs) per coerenza con Word."""
    for paragrafo in cell.paragraphs:
        for run in paragrafo.runs:
            run.font.name = font_name
            run.font.size = Pt(size_pt)
            run.bold = bold
            # Imposta anche il font per ASCII/hAnsi/cs in modo che Word lo
            # applichi davvero (python-docx imposta solo w:ascii tramite
            # font.name, non sempre sufficiente).
            rpr = run._element.get_or_add_rPr()
            rfonts = rpr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = rpr.makeelement(qn("w:rFonts"), {})
                rpr.append(rfonts)
            for attr in ("w:ascii", "w:hAnsi", "w:cs"):
                rfonts.set(qn(attr), font_name)


def _centra_paragrafo(paragrafo) -> None:
    """Centra orizzontalmente un paragrafo."""
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _estrai_codice_corso(nome_corso: str) -> str:
    """Restituisce il codice edizione tipo 'CORSO08' dal NomeCorso.
    Prova prima con 'EDIZIONE N', poi col prefisso '[CORSO N]'.
    Ritorna 'CORSO?' se non riesce a estrarre il numero."""
    corso_breve = estrai_numero_edizione_da_corso(nome_corso)
    if not corso_breve:
        corso_breve = estrai_numero_edizione_da_cartella(nome_corso)
    return f"CORSO{corso_breve}" if corso_breve else "CORSO?"


def _sostituisci_testo_header(part, vecchio: str, nuovo: str) -> None:
    """Sostituisce tutte le occorrenze di 'vecchio' con 'nuovo' nei run di
    una parte Word (header/footer/corpo) preservando la formattazione.

    Lavora a livello di XML per gestire il caso (frequente) in cui il testo
    sia spezzato su piu' run w:r/w:t: ricalcola la ripartizione mantenendo
    gli stessi elementi run.
    """
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xml = part._element

    for paragrafo in xml.iter(f"{{{W_NS}}}p"):
        run_t = []
        # Raccoglie tutti i nodi w:t del paragrafo e la mappatura ai run
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

        # Ridistribuisce il nuovo testo sui run esistenti:
        # mette tutto nel primo run w:t e svuota gli altri, cosi' la
        # formattazione del primo run viene applicata all'intero testo.
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
    """Crea e salva un singolo file .docx di registro attestati.

    Usa come base il file 'modello_registro.docx' (copia .docx del file di
    riferimento 'Registro attestati Cliente B.doc') e si limita
    a sostituire i contenuti variabili, preservando integro il formato
    originale (logo, header di pagina, tabelle, misure, font):
      - header: nome corso e codice edizione CORSO (riga di progetto)
      - corpo: nome azienda (tabella intestatario) ed elenco corsisti
        (tabella registro), con numerazione corretta 1..N.
    """
    # ---- Apertura del modello ----
    modello_path = Path(__file__).parent / "modello_registro.docx"
    if not modello_path.exists():
        raise FileNotFoundError(
            f"Modello non trovato: {modello_path}. "
            "Inserisci 'modello_registro.docx' nella cartella dello script."
        )
    doc = DocxDocument(str(modello_path))

    # Codice edizione per la riga di progetto (es. "CORSO08")
    codice_corso = _estrai_codice_corso(nome_corso)

    # ---- Sostituzioni nell'header di pagina ----
    # Il modello contiene i testi del documento di riferimento: li
    # sostituiamo con quelli del corso/edizione corrente.
    header = doc.sections[0].header
    # Riga denominazione corso (nel modello: 'PROBLEM SOLVING ... II ANNUALITA')
    _sostituisci_testo_header(header, "PROBLEM SOLVING E PROCESSI DECISIONALI  II ANNUALITÀ", nome_corso)
    # Riga progetto: codice CORSO (nel modello: 'CORSO 20' -> 'CORSOxx')
    _sostituisci_testo_header(header, "CORSO 20", codice_corso)

    # ---- Tabella 1: intestatario (Descrizione | valore azienda) ----
    tab_intro = doc.tables[0]
    # Sostituisce il nome azienda del modello con quello reale
    cella_val = tab_intro.cell(0, 1)
    _sostituisci_testo_header(cella_val, "CLIENTE_B S.N.C.", azienda)

    # ---- Tabella 2: elenco corsisti ----
    tab_elenco = doc.tables[1]
    # Il modello ha 1 riga di intestazione + N righe dati (una con esempio
    # 'CORSISTA ESEMPIO' e le altre vuote). Rimuoviamo tutte le righe
    # dati esistenti e le ricostruiamo con i corsisti reali.
    n_da_rimuovere = len(tab_elenco.rows) - 1  # mantengo solo l'intestazione
    for _ in range(n_da_rimuovere):
        # Rimuove l'ultima riga dati (l'elemento tr XML)
        tbl = tab_elenco._tbl
        tbl.remove(tab_elenco.rows[-1]._tr)

    # Aggiunge una riga per ogni corsista reale, numerazione 1..N
    for idx, (_, riga) in enumerate(corsisti.iterrows(), start=1):
        cells = tab_elenco.add_row().cells
        cells[0].text = str(idx)
        _imposta_font_cella(cells[0], "Arial", 10, bold=True)
        _centra_paragrafo(cells[0].paragraphs[0])
        cells[1].text = f"{riga['Cognome']} {riga['Nome']}".strip()
        _imposta_font_cella(cells[1], "Times New Roman", 12, bold=False)
        # colonne Data e Firma lasciate vuote (da compilare a mano)

    # ---- Salvataggio ----
    # Il nome del file usa un identificativo breve del corso (numero
    # edizione se reperibile, altrimenti il NomeCorso troncato) per
    # rispettare il limite MAX_PATH (260 caratteri) di Windows: il percorso
    # Il nome file e' gia' molto lungo. Il NomeCorso completo resta comunque
    # nell'intestazione del documento.
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
        f"  {cartella_azienda.name}: registro generato -> {out_path.name} "
        f"({len(corsisti)} corsisti)"
    )


# ========================================================================
# BONUS — TUTTO: dividi + organizza in un colpo solo (singola edizione)
# ========================================================================
def azione_tutto(args: argparse.Namespace, script_dir: Path) -> None:
    log.info("### Passo 1/2: divisione fronte+retro ###\n")

    dividi_args = argparse.Namespace(input_pdf=args.input_pdf)
    out_dir = azione_dividi(dividi_args, script_dir)

    # Se è stato indicato un nome edizione, rinomina la cartella di output
    # (es. attestati_output -> CORSO10) prima di organizzarla, così il
    # risultato è già pronto per un successivo 'riorganizza'.
    if args.edizione:
        nuova_cartella = out_dir.parent / args.edizione
        if nuova_cartella.exists():
            log.error(
                f"La cartella '{nuova_cartella}' esiste già: rimuovila o rinominala "
                f"a mano prima di rilanciare, per evitare di mischiare edizioni diverse."
            )
            sys.exit(1)
        out_dir.rename(nuova_cartella)
        out_dir = nuova_cartella
        log.info(f"\nCartella rinominata in: {out_dir.name}")
    else:
        log.info(
            "\nSuggerimento: usa --edizione CORSO10 (col numero giusto) per rinominare "
            "subito la cartella di output e ottenere risultati migliori nel matching "
            "per edizione e per un futuro uso di 'riorganizza'."
        )

    log.info("\n### Passo 2/2: organizzazione per azienda ###")

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
        description="Gestione completa degli attestati CORSO: dividi, organizza, riorganizza.",
    )
    sub = parser.add_subparsers(dest="comando")

    sp_dividi = sub.add_parser(
        "dividi",
        help="Divide un PDF di stampa massiva in un file per persona (fronte+retro uniti).",
    )
    sp_dividi.add_argument(
        "input_pdf", nargs="?", default=None,
        help="Percorso del PDF da dividere. Se omesso, cerca automaticamente un PDF nella cartella dello script.",
    )

    sp_organizza = sub.add_parser(
        "organizza",
        help="Smista i PDF di un'edizione in sottocartelle per azienda usando una rubrica Excel.",
    )
    sp_organizza.add_argument(
        "cartella_edizione", nargs="?", default=None,
        help="Cartella edizione (es. 'CORSO10') o cartella con più sottocartelle CORSO*. Se omessa, cercata automaticamente.",
    )
    sp_organizza.add_argument(
        "file_excel", nargs="?", default=None,
        help="Percorso del file Excel della rubrica (.xls o .xlsx). Se omesso, cercato automaticamente.",
    )
    sp_organizza.add_argument(
        "--copia", action="store_true",
        help="Copia i PDF nelle sottocartelle invece di spostarli.",
    )

    sp_riorganizza = sub.add_parser(
        "riorganizza",
        help="Ricompone più edizioni già organizzate per azienda in una vista azienda/edizione.",
    )
    sp_riorganizza.add_argument(
        "cartelle", nargs="*",
        help="Cartella contenitore con più sottocartelle CORSO*, oppure lista di cartelle edizione. Se omessa, cercata automaticamente.",
    )
    sp_riorganizza.add_argument("--output", default=None, help="Cartella di destinazione.")
    sp_riorganizza.add_argument(
        "--sposta", action="store_true",
        help="Sposta i file invece di copiarli.",
    )
    sp_riorganizza.add_argument(
        "--rubrica", default=None,
        help="File Excel della rubrica (.xls/.xlsx): se specificato, genera un registro Word "
             "per ogni azienda nella cartella di output.",
    )

    sp_registri = sub.add_parser(
        "registri",
        help="Genera i registri Word SOLO da modello + rubrica (senza cartelle CORSO*/PDF).",
    )
    sp_registri.add_argument(
        "rubrica", nargs="?", default=None,
        help="File Excel della rubrica (.xls/.xlsx). Se omesso, cercato automaticamente "
             "nella cartella dello script.",
    )
    sp_registri.add_argument(
        "--output", default=None,
        help="Cartella di destinazione (default: <cartella script>/REGISTRI).",
    )

    sp_tutto = sub.add_parser(
        "tutto",
        help="Pipeline completa: divide il PDF e lo organizza subito per azienda (una singola edizione).",
    )
    sp_tutto.add_argument("input_pdf", nargs="?", default=None, help="PDF di stampa massiva da dividere.")
    sp_tutto.add_argument("file_excel", nargs="?", default=None, help="Rubrica Excel di quell'edizione.")
    sp_tutto.add_argument(
        "--edizione", default=None,
        help="Nome edizione, es. CORSO10: rinomina la cartella di output prima di organizzarla.",
    )
    sp_tutto.add_argument("--copia", action="store_true", help="Copia invece di spostare i PDF nelle sottocartelle azienda.")

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
