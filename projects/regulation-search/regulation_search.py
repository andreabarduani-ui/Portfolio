#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ESTRATTORE INFORMAZIONI - DOCUMENTI FORMAZIONE FINANZIATA
==========================================================

Strumento interattivo per estrarre informazioni da manuali, avvisi e linee
guida relativi alla formazione finanziata (dei principali fondi
interprofessionali, ecc.).

Lo script:
- Indicizza il PDF pagina per pagina
- Permette ricerche per categorie predefinite (variazioni, rendicontazione,
  costi ammissibili, modalità formative, aiuti di stato, ecc.)
- Supporta ricerche libere per parola chiave
- Restituisce SEMPRE il numero di pagina e il contesto della corrispondenza

Dipendenze: pdfplumber (preferito) oppure pypdf
    pip install pdfplumber

Uso:
    python regulation_search.py [path_al_pdf]
    # oppure senza argomenti, lo script chiede il path
"""

import os
import re
import sys
import textwrap
from pathlib import Path
from typing import Dict, List, Optional, Tuple

# -------- Import PDF library con fallback --------
try:
    import pdfplumber
    PDF_LIB = "pdfplumber"
except ImportError:
    try:
        from pypdf import PdfReader
        PDF_LIB = "pypdf"
    except ImportError:
        print("ERRORE: serve almeno una libreria PDF.")
        print("Installa con: pip install pdfplumber")
        sys.exit(1)


# =============================================================================
# CATEGORIE PREDEFINITE PER FORMAZIONE FINANZIATA
# =============================================================================
# Ogni categoria è una lista di parole chiave/pattern (case-insensitive).
# Aggiungere o modificare queste categorie per adattare lo script ad altri
# Fondi Interprofessionali o tipologie di documento.

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

# Etichette user-friendly per il menu
ETICHETTE_CATEGORIE: Dict[str, str] = {
    "variazioni_progetto": "Variazioni / proroghe / rinunce / sostituzioni",
    "modalita_formative": "Modalità formative (aula, FAD, teleformazione...)",
    "rendicontazione": "Rendicontazione e documenti giustificativi",
    "costi_ammissibili": "Costi ammissibili, parametri e massimali",
    "aiuti_di_stato": "Aiuti di Stato (de minimis, Reg. 651/2014, RNA)",
    "personale_docenza": "Personale: docenti, tutor, consulenti",
    "presentazione_progetto": "Presentazione progetto (formulario, allegati, valutazione)",
    "controlli_verifiche": "Controlli e verifiche (I livello, II livello, ispezioni)",
    "voucher_individuale": "Voucher / formazione individuale",
    "tempistiche_scadenze": "Tempistiche e scadenze del progetto",
    "ati_ats_partenariato": "ATI / ATS / partenariato / delega",
    "imprese_beneficiarie": "Imprese beneficiarie e destinatari",
    "fideiussione_anticipo": "Fideiussione e richiesta di anticipo",
}


# =============================================================================
# ESTRAZIONE TESTO PDF
# =============================================================================

def estrai_testo_pdf(path: str) -> List[Tuple[int, str]]:
    """
    Estrae testo pagina per pagina.
    Ritorna lista di tuple (numero_pagina_1based, testo_pagina).
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
# RICERCA
# =============================================================================

def trova_corrispondenze(
    pagine: List[Tuple[int, str]],
    parole_chiave: List[str],
    contesto_caratteri: int = 250,
    max_risultati_per_pagina: int = 3,
) -> List[Dict]:
    """
    Cerca le parole chiave (case-insensitive) nelle pagine.
    Ritorna lista di dict con: pagina, parola_trovata, contesto.
    """
    risultati: List[Dict] = []
    # Compila pattern: \b non funziona bene con accenti italiani -> usa lookaround
    patterns = [
        (kw, re.compile(re.escape(kw), re.IGNORECASE))
        for kw in parole_chiave
    ]

    for num_pag, testo in pagine:
        if not testo:
            continue
        # Conta corrispondenze per pagina ed estrae contesto
        trovati_pag = 0
        gia_visti_offset: List[int] = []  # evita duplicati ravvicinati
        for kw, pat in patterns:
            for match in pat.finditer(testo):
                if trovati_pag >= max_risultati_per_pagina:
                    break
                start = match.start()
                # Salta se troppo vicino a un match già preso (entro 80 char)
                if any(abs(start - x) < 80 for x in gia_visti_offset):
                    continue
                gia_visti_offset.append(start)

                inizio = max(0, start - contesto_caratteri // 2)
                fine = min(len(testo), match.end() + contesto_caratteri // 2)
                contesto = testo[inizio:fine]
                # Pulisci a capo multipli
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
    """Stampa formattata dei risultati."""
    print()
    print("=" * 78)
    print(f"  RISULTATI: {titolo}")
    print(f"  Corrispondenze trovate: {len(risultati)}")
    print("=" * 78)
    if not risultati:
        print("\n  Nessuna corrispondenza trovata.\n")
        return

    # Raggruppa per pagina
    per_pagina: Dict[int, List[Dict]] = {}
    for r in risultati:
        per_pagina.setdefault(r["pagina"], []).append(r)

    for pag in sorted(per_pagina.keys()):
        items = per_pagina[pag]
        print(f"\n  📄 PAGINA {pag}  ({len(items)} corrispondenz{'a' if len(items)==1 else 'e'})")
        print("  " + "-" * 74)
        for it in items:
            kw = it["parola_trovata"]
            ctx = it["contesto"]
            # Evidenzia (semplice — uppercase) la parola trovata nel contesto
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
# MENU INTERATTIVO
# =============================================================================

def chiedi_path_pdf() -> str:
    """Richiede il path se non passato come argomento."""
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = input("Path del PDF da analizzare: ").strip().strip('"').strip("'")
    if not path:
        print("Path non fornito.")
        sys.exit(1)
    if not Path(path).is_file():
        print(f"File non trovato: {path}")
        sys.exit(1)
    return path


def menu_principale() -> str:
    """Mostra il menu e ritorna la scelta."""
    print()
    print("╔" + "═" * 76 + "╗")
    print("║" + "  COSA VUOI ESTRARRE DAL DOCUMENTO?".ljust(76) + "║")
    print("╚" + "═" * 76 + "╝")
    print()
    chiavi = list(ETICHETTE_CATEGORIE.keys())
    for i, k in enumerate(chiavi, start=1):
        print(f"  {i:>2}. {ETICHETTE_CATEGORIE[k]}")
    print(f"  {len(chiavi)+1:>2}. 🔎 Ricerca libera (parole chiave personalizzate)")
    print(f"  {len(chiavi)+2:>2}. 📑 Indice del documento (prime righe di ogni pagina)")
    print(f"  {len(chiavi)+3:>2}. 📊 Statistiche documento")
    print(f"   0. Esci")
    print()
    scelta = input("Scelta: ").strip()
    return scelta


def gestisci_scelta(
    scelta: str,
    pagine: List[Tuple[int, str]],
) -> bool:
    """Esegue l'azione scelta. Ritorna True se continuare, False per uscire."""
    chiavi = list(ETICHETTE_CATEGORIE.keys())

    if scelta == "0":
        print("\nCiao!\n")
        return False

    # Ricerca per categoria
    if scelta.isdigit() and 1 <= int(scelta) <= len(chiavi):
        cat = chiavi[int(scelta) - 1]
        parole = CATEGORIE[cat]
        risultati = trova_corrispondenze(pagine, parole)
        stampa_risultati(risultati, ETICHETTE_CATEGORIE[cat])
        return True

    # Ricerca libera
    if scelta == str(len(chiavi) + 1):
        kw_input = input(
            "\nInserisci parole chiave separate da virgola "
            "(es: proroga, rinuncia, modello 4): "
        ).strip()
        if not kw_input:
            print("Nessuna parola inserita.")
            return True
        parole = [p.strip() for p in kw_input.split(",") if p.strip()]
        risultati = trova_corrispondenze(pagine, parole)
        stampa_risultati(risultati, f"Ricerca libera: {', '.join(parole)}")
        return True

    # Indice rapido
    if scelta == str(len(chiavi) + 2):
        print("\n" + "=" * 78)
        print("  INDICE RAPIDO (prime righe di ogni pagina)")
        print("=" * 78)
        for num_pag, testo in pagine:
            prima_riga = ""
            for riga in (testo or "").splitlines():
                riga_pulita = riga.strip()
                if riga_pulita and len(riga_pulita) > 5:
                    prima_riga = riga_pulita[:90]
                    break
            print(f"  Pag. {num_pag:>3} | {prima_riga}")
        print()
        return True

    # Statistiche
    if scelta == str(len(chiavi) + 3):
        totale_pagine = len(pagine)
        totale_caratteri = sum(len(t) for _, t in pagine)
        totale_parole = sum(len((t or "").split()) for _, t in pagine)
        print("\n" + "=" * 78)
        print("  STATISTICHE DOCUMENTO")
        print("=" * 78)
        print(f"  Pagine totali     : {totale_pagine}")
        print(f"  Caratteri totali  : {totale_caratteri:,}".replace(",", "."))
        print(f"  Parole totali     : {totale_parole:,}".replace(",", "."))
        print(f"  Media parole/pag. : {totale_parole // max(totale_pagine,1)}")
        print()
        return True

    print(f"Scelta non valida: {scelta!r}")
    return True


# =============================================================================
# MAIN
# =============================================================================

def main() -> None:
    print()
    print("╔" + "═" * 76 + "╗")
    titolo = "ESTRATTORE - DOCUMENTI FORMAZIONE FINANZIATA"
    print("║" + titolo.center(76) + "║")
    sotto = f"(libreria PDF in uso: {PDF_LIB})"
    print("║" + sotto.center(76) + "║")
    print("╚" + "═" * 76 + "╝")

    path = chiedi_path_pdf()
    print(f"\n📥 Estrazione testo da: {path}")
    pagine = estrai_testo_pdf(path)
    print(f"✅ {len(pagine)} pagine indicizzate.")

    while True:
        scelta = menu_principale()
        try:
            continua = gestisci_scelta(scelta, pagine)
        except KeyboardInterrupt:
            print("\nInterrotto dall'utente.\n")
            break
        if not continua:
            break


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nInterrotto.\n")
