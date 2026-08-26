# -*- coding: utf-8 -*-
"""
transparency_certificates.py
===================
Duplica il file "Attestato Apprendimenti Acquisiti_.docx" per ogni partecipante
presente nel "File madre.xlsx", riempiendo le tabelle del Word con i dati
dell'Excel e salvando ogni attestato con il cognome del partecipante.

Come usarlo (VS Code):
  1. Apri questa cartella in VS Code.
  2. Apri il file transparency_certificates.py.
  3. Premi F5 (o il pulsante "Run") oppure, da terminale:
        py transparency_certificates.py

In alto (sezione CONFIGURAZIONE) puoi modificare i percorsi e altre opzioni
senza toccare il resto del codice.
"""

import os
import re
import datetime
import shutil
import tempfile

import openpyxl
import docx


# ============================================================================
# CONFIGURAZIONE  (modifica qui se serve)
# ============================================================================
BASE_DIR   = "."
FILE_WORD  = os.path.join(BASE_DIR, "Attestato Apprendimenti Acquisiti_.docx")
FILE_EXCEL = os.path.join(BASE_DIR, "File madre.xlsx")

# Cartella dove verranno salvati gli attestati generati
OUT_DIR    = os.path.join(BASE_DIR, "Attestati_Generati")

# Righe/fogli dell'Excel
NOME_FOGLIO   = "Foglio1"
RIGA_INTESTAZIONE = 2   # riga che contiene le intestazioni di colonna
RIGA_PRIMO_DATO   = 3   # prima riga con i dati dei partecipanti

# Prefisso del nome file di output (verrà: <prefisso>_<Cognome>.docx)
PREFISSO_FILE = "Attestato Apprendimenti Acquisiti"

# Se True, sovrascrive eventuali attestati gia' esistenti con stesso nome
SOVRASCRIVI = True
# ============================================================================


# ----------------------------------------------------------------------------
# MAPPATURA  Word <-> Excel
# ----------------------------------------------------------------------------
# Chiave  = etichetta ESATTA che compare nella cella di sinistra della tabella
#           Word (colonna 0).
# Valore  = indice di colonna Excel (1-based: B=2, C=3, D=4, ...).
#
# NOTA IMPORTANTE sulle colonne D ed E dell'Excel:
#   - le intestazioni sono SCAMBIATE rispetto al contenuto;
#   - la colonna D contiene in realta' le DATE di nascita;
#   - la colonna E contiene in realta' i COMUNI/STATI di nascita.
#   Qui ci basiamo sul CONTENUTO effettivo, non sull'intestazione.
# ----------------------------------------------------------------------------
MAPPING = {
    # --- Tabella 1: Allievo ---
    "Cognome":                              2,   # col B
    "Nome":                                 3,   # col C
    "Codice fiscale":                       8,   # col H
    "Data nascita":                         4,   # col D (contiene le date)
    "Comune/stato straniero nascita":       5,   # col E (contiene le citta')
    "Cittadinanza":                         6,   # col F

    # --- Tabella 3: Esperienza formativa (solo i campi vuoti da riempire) ---
    "Titolo del percorso":                                  7,   # col G
    "N\u00b0 di ore frequentate su numero di ore totali del percorso": 14,  # col N

    # --- Tabella 4: Competenze ---
    "Risultato di apprendimento":                              9,   # col I
    "Risultato Atteso (RA)":                                   10,  # col J
    "Competenza/descrittore di cui ai quadri europei":        11,  # col K
    "Eventuali ulteriori evidenze a supporto riferite alla personalizzazione del percorso": 12,  # col L
    "Prove di valutazione \u2013 data \u2013 esito":          13,  # col M
}
# NB: i campi dell'Excel scritti come timedelta (colonna N) vengono convertiti
#     in ore. Le celle del Word che contengono gia' un valore NON vengono
#     toccate (restano invariate), come richiesto.


# ----------------------------------------------------------------------------
# Funzioni di utilita'
# ----------------------------------------------------------------------------

def normalizza_etichetta(s):
    """Rende un'etichetta confrontabile: lower, collapse spazi, senza punteg.grega."""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def valore_come_testo(val):
    """Converte un valore Excel (datetime, timedelta, None, ...) in stringa."""
    if val is None:
        return ""
    # Data di nascita -> formato italiano gg/mm/aaaa
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime("%d/%m/%Y")
    # Durata (timedelta) -> ore totali (intero o con un decimale)
    if isinstance(val, datetime.timedelta):
        ore = val.total_seconds() / 3600.0
        # se e' praticamente intero, mostra senza decimali
        if abs(ore - round(ore)) < 1e-6:
            return f"{int(round(ore))}"
        return f"{ore:.1f}".rstrip("0").rstrip(".")
    # Numeri
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    return str(val).strip()


def scrivi_in_cella(cell, testo):
    """
    Scrive 'testo' nella cella MANTENENDO lo stile del documento:
    - usa il primo run esistente (copia il suo formato);
    - se la cella e' vuota crea un run ereditando lo stile del paragrafo.
    """
    par = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    # Rimuove i run esistenti mantenendo il primo (per il formato)
    if par.runs:
        primo_run = par.runs[0]
        primo_run.text = testo
        # svuota gli altri run
        for r in par.runs[1:]:
            r.text = ""
    else:
        run = par.add_run(testo)
    # se ci fossero altri paragrafi nella cella li svuotiamo (caso raro)
    for extra in cell.paragraphs[1:]:
        for r in extra.runs:
            r.text = ""


def sanitize_filename(name):
    """Rimuove i caratteri non validi per i nomi file Windows."""
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = re.sub(r"\s+", " ", name)
    return name


# ----------------------------------------------------------------------------
# Logica principale
# ----------------------------------------------------------------------------

def leggi_partecipanti():
    """Legge l'Excel e restituisce una lista di dizionari {col_index: valore}."""
    wb = openpyxl.load_workbook(FILE_EXCEL, data_only=True)
    ws = wb[NOME_FOGLIO]

    partecipanti = []
    for riga in range(RIGA_PRIMO_DATO, ws.max_row + 1):
        cognome = ws.cell(row=riga, column=2).value  # col B
        if cognome is None or str(cognome).strip() == "":
            continue  # riga vuota: salta
        dati = {}
        for col in range(1, ws.max_column + 1):
            dati[col] = ws.cell(row=riga, column=col).value
        dati["_cognome"] = str(cognome).strip()
        dati["_nome"] = str(ws.cell(row=riga, column=3).value or "").strip()
        partecipanti.append(dati)
    return partecipanti


def compila_documento(template_doc, dati):
    """
    Crea una NUOVA copia del documento template (da file, per evitare
    problemi di condivisione della struttura XML interna tra copie) e
    compila le tabelle con i dati del partecipante.
    Il file template originale non viene MAI modificato.
    """
    # Copia il template in un file temporaneo, poi lo apriamo: cosi' ogni
    # attestato parte da una struttura XML pulita e indipendente.
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".docx")
    os.close(tmp_fd)
    try:
        shutil.copyfile(template_doc, tmp_path)
        doc = docx.Document(tmp_path)
    finally:
        # rimuoviamo subito il temporaneo; il 'doc' resta in memoria
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    # Mappa etichetta normalizzata -> colonna Excel
    mappa_norm = {normalizza_etichetta(k): v for k, v in MAPPING.items()}

    for tabella in doc.tables:
        ncol = len(tabella.columns)
        if ncol < 2:
            continue  # servono almeno 2 colonne (etichetta | valore)
        for riga in tabella.rows:
            celle = riga.cells
            etichetta = celle[0].text
            etich_norm = normalizza_etichetta(etichetta)

            # REGOLA FONDAMENTALE (richiesta dall'utente):
            #   compila SOLO le celle-valore che al momento sono VUOTE.
            #   Tutte le celle che contengono gia' un valore restano INVARIATE
            #   su ogni attestato (es. dati dell'Ente, del responsabile,
            #   codice istanza, fondo, ecc.).
            if etich_norm in mappa_norm and celle[1].text.strip() == "":
                col_excel = mappa_norm[etich_norm]
                valore = dati.get(col_excel)
                testo = valore_come_testo(valore)
                if testo != "":
                    scrivi_in_cella(celle[1], testo)
    return doc


def main():
    # --- Controlli di esistenza ---
    if not os.path.isfile(FILE_WORD):
        print(f"ERRORE: file Word non trovato:\n  {FILE_WORD}")
        return
    if not os.path.isfile(FILE_EXCEL):
        print(f"ERRORE: file Excel non trovato:\n  {FILE_EXCEL}")
        return

    os.makedirs(OUT_DIR, exist_ok=True)

    print("Lettura partecipanti dall'Excel...")
    partecipanti = leggi_partecipanti()
    print(f"Trovati {len(partecipanti)} partecipanti.\n")

    print(f"Template Word:\n  {FILE_WORD}")

    generati = 0
    saltati = 0
    for i, dati in enumerate(partecipanti, start=1):
        cognome = sanitize_filename(dati["_cognome"])
        nome = dati["_nome"]
        out_name = f"{PREFISSO_FILE}_{cognome}.docx"
        out_path = os.path.join(OUT_DIR, out_name)

        if os.path.exists(out_path) and not SOVRASCRIVI:
            print(f"[{i:02d}] SALTO (gia' esistente): {out_name}")
            saltati += 1
            continue

        doc = compila_documento(FILE_WORD, dati)
        doc.save(out_path)
        print(f"[{i:02d}] Generato: {out_name}   ({cognome} {nome})")
        generati += 1

    print("\n" + "=" * 60)
    print(f"Completato. Generati: {generati} | Saltati: {saltati}")
    print(f"Cartella di output:\n  {OUT_DIR}")
    print("=" * 60)


if __name__ == "__main__":
    main()
