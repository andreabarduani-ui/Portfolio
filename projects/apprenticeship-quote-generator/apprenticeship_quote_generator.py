# -*- coding: utf-8 -*-
"""
apprenticeship_quote_generator.py
-------------------------------
Generatore automatico di OFFERTE ECONOMICHE per corsi di apprendistato privato.

Funzionamento:
  - Usa come base una copia fedele del documento di esempio (template_offerta_apprendistato.docx),
    preservando logo, intestazione, piè di pagina, font Arial e formattazione originale.
  - Chiede in modo interattivo (input da tastiera, VS Code / terminale):
      1. se il calendario e' stato concordato (se NO richiede il file Excel del calendario);
      2. il numero di partecipanti (base del calcolo);
      3. l'annualita' (prima / seconda);
      4. i dati dell'azienda destinataria.
  - Calcola:
      * Importo totale      = 480,00 €  x  n.partecipanti
      * Prezzo finale ("PREZZO MIGLIOR FAVORE") = totale x 0,90 arrotondato alla centinaio piu' vicina
  - Compila il NUMERO OFFERTA con un progressivo automatico, separato per
    annualita' (memorizzato in contatore_offerte.json):
      * prima annualita'  parte da 20260044 e si incrementa di 1 ad ogni offerta;
      * seconda annualita' parte da 20260066 e si incrementa di 1 ad ogni offerta.
  - Imposta la data del giorno di esecuzione dello script.
  - Se il calendario NON e' concordato, legge il file Excel del calendario:
      * estrae codice edizione (CORSO 07 / Edizione 07) dall'intestazione e lo propaga nel documento;
      * ricostruisce una tabella Word a 6 colonne (Data | Orario | Ore |
        Modulo | Docente | Modalità) sotto la voce "Calendario", come nel
        template "template_offerta_apprendistato con calendario.docx".
  - L'annualita' (prima/seconda) viene aggiornata sia nel sottotitolo del corso
    sia nella riga descrizione della tabella costi.
  - Salva l'offerta in output/Offerta_Apprendistato_<RagioneSociale>_<AAAAMMGG>.docx.

L'originale "template_offerta_apprendistato.docx" non viene mai toccato: si lavora sempre
su una copia del template.
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
# COSTANTI
# =========================================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(SCRIPT_DIR, "template_offerta_apprendistato.docx")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "output")

UNITARIO = 480.0          # importo unitario per partecipante (fisso)
SCONTO = 0.10             # sconto del 10% per il prezzo finale

# Colonne del calendario: 6 colonne complete (come template "con calendario").
HEADERS_CALENDARIO = ["Data", "Orario", "Ore", "Modulo", "Docente", "Modalità"]

SEGNAPOSTO_VUOTO = "______________"   # per campi azienda lasciati vuoti

# Nell'offerta originale il sottotitolo usa "CORSO5" (numero senza zero)
# e la descrizione costi "Ed. 05" (numero a 2 cifre con zero).
EDIZIONE_DEFAULT = 5      # fallback se non leggibile dal calendario

# --- Numerazione progressiva offerte --------------------------------------
# File locale che memorizza l'ultimo numero offerta usato per ciascuna
# annualita'. La prima annualita' parte da 20260044, la seconda da 20260066;
# ogni nuova offerta prodotta incrementa di 1 il contatore della sua annualita'.
CONTATORE_FILE = os.path.join(SCRIPT_DIR, "contatore_offerte.json")
CONTATORI_BASE = {1: 20260046, 2: 20260067}   # partenza per annualita' 1 / 2


# =========================================================================
# CONTATORE NUMERO OFFERTA (persistente su file locale)
# =========================================================================
def leggi_contatori():
    """
    Carica il file JSON con l'ULTIMO numero offerta ASSEGNATO per annualita'.
    Ritorna un dict {1: <int>, 2: <int>}. Se il file non esiste o e' invalido
    inizializza ciascun contatore a (BASE - 1), in modo che il prossimo numero
    da assegnare sia proprio la BASE (prima annualita' 20260044, seconda 20260066).
    """
    out = {k: v - 1 for k, v in CONTATORI_BASE.items()}   # base-1 di partenza
    if os.path.isfile(CONTATORE_FILE):
        try:
            import json
            with open(CONTATORE_FILE, "r", encoding="utf-8") as f:
                dati = json.load(f)
            for k in (1, 2):
                v = dati.get(str(k))
                # accetta solo valori validi (>= base-1)
                if isinstance(v, int) and v >= CONTATORI_BASE[k] - 1:
                    out[k] = v
        except Exception:
            # file corrotto: ignora e usa base-1
            pass
    return out


def prossimo_numero_offerta(annualita_num, contatori=None):
    """
    Restituisce il prossimo numero offerta da assegnare per la annualita'
    indicata (1 o 2): sempre (ultimo_assegnato + 1).
    Alla prima esecuzione (nessuna offerta ancora prodotta) il contatore vale
    BASE-1, quindi il primo numero assegnato sara' la BASE
    (20260044 per la prima annualita', 20260066 per la seconda).
    """
    if contatori is None:
        contatori = leggi_contatori()
    return contatori.get(annualita_num, CONTATORI_BASE[annualita_num] - 1) + 1


def aggiorna_contatore(annualita_num, numero_assegnato):
    """
    Registra il numero offerta appena assegnato per la annualita' indicata,
    se e' maggiore dell'ultimo memorizzato (cosi' mantiene il max anche in caso
    di riesecuzioni con numero modificato a mano). Salva su file JSON locale.
    """
    import json
    contatori = leggi_contatori()
    if numero_assegnato > contatori.get(annualita_num, CONTATORI_BASE[annualita_num] - 1):
        contatori[annualita_num] = numero_assegnato
        try:
            with open(CONTATORE_FILE, "w", encoding="utf-8") as f:
                # salva con chiavi stringa (standard JSON)
                json.dump({str(k): v for k, v in contatori.items()}, f,
                          indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"  AVVISO: impossibile salvare il contatore offerte ({e}).")


# =========================================================================
# UTILITA' DI FORMATTAZIONE
# =========================================================================
def euro(value):
    """Formatta un numero come importo in euro italiano: 5760 -> '5.760,00 €'."""
    s = f"{value:,.2f}"          # 5,760.00
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")  # 5.760,00
    return s + " €"


def calc_totali(n_partecipanti):
    """Restituisce (totale, finale) dato il numero di partecipanti."""
    totale = UNITARIO * n_partecipanti
    scontato = totale * (1 - SCONTO)
    # arrotonda alla centinaio piu' vicina
    finale = round(scontato / 100.0) * 100.0
    return totale, finale


def codice_corso(edizione):
    """Restituisce il codice CORSO con numero senza zero: 7 -> 'CORSO7'."""
    return f"CORSO{int(edizione)}"


def codice_edizione_padded(edizione):
    """Restituisce l'edizione a 2 cifre con zero: 7 -> 'Ed. 07'."""
    return f"Ed. {int(edizione):02d}"


def _parse_data(valore):
    """
    Converte un valore di data (datetime, date, o stringa vari formati) in 'gg/mm/aaaa'.
    Gestisce i datetime prodotti da Excel (es. datetime(2026,9,14)).
    """
    if valore is None or (isinstance(valore, str) and not valore.strip()):
        return ""
    if isinstance(valore, datetime):
        return valore.strftime("%d/%m/%Y")
    if isinstance(valore, date):
        return valore.strftime("%d/%m/%Y")
    s = str(valore).strip()
    # prova formati comuni
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return s  # fallback: lascia il testo originale


# =========================================================================
# INPUT INTERATTIVO
# =========================================================================
def _chiedi(prompt, default=None):
    """Wrapper su input() con eventuale default tra parentesi quadre."""
    if default is not None:
        full = f"{prompt} [{default}]: "
    else:
        full = f"{prompt}: "
    val = input(full).strip()
    if val == "" and default is not None:
        return default
    return val


def _si_no(prompt, default="s"):
    """Chiede una domanda si/no. Ritorna True per si', False per no."""
    while True:
        v = _chiedi(prompt + " (s/n)", default=default).lower()
        if v in ("s", "si", "y", "yes"):
            return True
        if v in ("n", "no"):
            return False
        print("  Rispondere con 's' o 'n'.")


def chiedi_calendario():
    """
    Gestisce il blocco calendario.
    Ritorna un dict con:
      - "concordato": True/False
      - "righe":     list[dict] (Data/Orario/Ore/Modulo/Docente/Modalità) oppure None
      - "edizione":  int (numero edizione letta dal calendario, o EDIZIONE_DEFAULT)
    """
    print("\n--- CALENDARIO ---")
    concordato = _si_no("Il calendario e' stato concordato?", default="s")
    if concordato:
        print("  >> Calendario concordato: il documento manterra' il testo originale.")
        return {"concordato": True, "righe": None, "edizione": EDIZIONE_DEFAULT}

    print("  >> Calendario NON concordato: e' necessario caricare il file Excel del calendario.")
    while True:
        path = _chiedi("Percorso del file Excel del calendario").strip().strip('"')
        if not path:
            print("  Percorso non valido, riprovare.")
            continue
        risultato = leggi_calendario_excel(path)
        if risultato is not None:
            righe, edizione = risultato
            print(f"  >> Lette {len(righe)} giornate dal calendario. Edizione rilevata: CORSO {edizione}.")
            return {"concordato": False, "righe": righe, "edizione": edizione}
        riprova = _si_no("Vuoi provare con un altro file?", default="s")
        if not riprova:
            print("  >> Nessun calendario disponibile. Verra' lasciato il testo originale.")
            return {"concordato": True, "righe": None, "edizione": EDIZIONE_DEFAULT}


def leggi_calendario_excel(path):
    """
    Legge il file Excel del calendario (formato reale CORSO07):
      - riga di intestazione con titolo edizione ("... Edizione 07 - CORSO 07 ...")
      - riga header colonne: Data | Orario | Ore | Docente | Modulo | Modalita'
      - righe dati (date come datetime)
      - eventuale riga di totale ore da escludere
    Ricerca le colonne per nome in modo flessibile.
    Ritorna (righe, edizione) oppure (None, None) in caso di errore.
      - righe: list[dict] con chiavi Data, Orario, Ore, Modulo, Docente, Modalità
      - edizione: int
    """
    if not os.path.isfile(path):
        print(f"  ERRORE: il file '{path}' non esiste.")
        return None, None
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        print(f"  ERRORE: impossibile aprire il file Excel ({e}).")
        return None, None

    # colonne attese (tutte lower-case per il matching)
    target = [h.lower() for h in HEADERS_CALENDARIO]   # data, orario, ore, modulo, docente, modalità
    ESSENZIALI = ["data"]   # serve almeno Data per riconoscere la tabella

    def cerca_in_foglio(ws):
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return None
        # estrai edizione da tutto il testo del foglio
        edizione = _estrai_edizione(rows)
        # trova la riga header che contiene la colonna 'data'
        for r_idx, row in enumerate(rows):
            cells = [("" if c is None else str(c)).strip().lower() for c in row]
            # mappa ogni target alla colonna corrispondente (se presente)
            col_map = {}
            for t in target:
                trovato = None
                for ci, c in enumerate(cells):
                    if c and (c == t or c.startswith(t) or t.startswith(c)):
                        trovato = ci
                        break
                if trovato is not None:
                    col_map[t] = trovato
            # serve almeno la colonna essenziale 'data'
            if not all(t in col_map for t in ESSENZIALI):
                continue
            # raccogli le righe dati
            out = []
            for data_row in rows[r_idx + 1:]:
                def get(t):
                    idx = col_map.get(t)
                    if idx is None:
                        return None
                    return data_row[idx] if idx < len(data_row) else None
                data_val = get("data")
                # salta righe senza data valida (es. riga totali ore)
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
                # salta righe completamente vuote
                if any(v for v in riga.values()):
                    out.append(riga)
            return out, edizione
        return None

    # prova tutti i fogli
    for ws in wb.worksheets:
        res = cerca_in_foglio(ws)
        if res:
            wb.close()
            return res
    wb.close()
    print("  ERRORE: non ho trovato la colonna 'Data' nel file del calendario.")
    return None, None


def _str(v):
    """Converte un valore Excel in stringa pulita (None -> '')."""
    if v is None:
        return ""
    s = str(v).strip()
    return s


def _ha_data_valida(v):
    """True se il valore rappresenta una data valida (datetime/date o stringa di data)."""
    if v is None:
        return False
    if isinstance(v, (datetime, date)):
        return True
    if isinstance(v, (int, float)):
        # numeri non sono date qui
        return False
    s = str(v).strip()
    if not s:
        return False
    # prova a parsare
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
        try:
            datetime.strptime(s, fmt)
            return True
        except ValueError:
            continue
    return False


def _estrai_edizione(rows):
    """
    Cerca il numero di edizione/CORSO in tutto il testo del foglio.
    Pattern tipici: 'Edizione 07', 'CORSO 07', 'CORSO7', 'Ed. 7'.
    Ritorna int (o EDIZIONE_DEFAULT se non trovato).
    """
    testo = " ".join(
        "" if c is None else str(c) for row in rows for c in row
    )
    # cerca "edizione XX" o "CORSO XX" (con spazi/zero opzionali)
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
    """Chiede i dati principali (partecipanti, annualita') e i dati azienda."""
    print("\n--- DATI CORSO ---")
    # Numero partecipanti
    while True:
        v = _chiedi("Numero di partecipanti (dipendenti coinvolti)")
        try:
            n = int(v)
            if n > 0:
                break
        except ValueError:
            pass
        print("  Inserire un numero intero positivo.")

    # Annualita' (testo "prima"/"seconda")
    print("  Annualita': 1 = prima annualita' (default), 2 = seconda annualita'")
    while True:
        a = _chiedi("Annualita' (1/2)", default="1")
        if a in ("1", "2"):
            ann_num = int(a)                            # 1 o 2 (per contatore offerta)
            ann_circ = "I°" if a == "1" else "II°"      # per il sottotitolo
            ann_desc = "I" if a == "1" else "II"        # per la descrizione costi
            ann_testo = "prima" if a == "1" else "seconda"
            break
        print("  Rispondere con 1 o 2.")

    print("\n--- DATI AZIENDA DESTINATARIA ---")
    print("  (lascia vuoto per mantenere il segnaposto ______________)")
    ragione = _chiedi("Ragione sociale (es. Azienda Esempio srl)") or SEGNAPOSTO_VUOTO
    via = _chiedi("Via e civico (es. Via Emilio Ghione, 12)") or SEGNAPOSTO_VUOTO
    cap = _chiedi("CAP (es. 00128)") or SEGNAPOSTO_VUOTO
    citta = _chiedi("Citta' (es. Roma)") or SEGNAPOSTO_VUOTO
    prov = _chiedi("Provincia sigla (es. RM)") or SEGNAPOSTO_VUOTO
    telefono = _chiedi("Telefono/Uff. (es. +39 0669331247)") or SEGNAPOSTO_VUOTO
    piva = _chiedi("Partita IVA (es. 16823051004)") or SEGNAPOSTO_VUOTO
    cf = _chiedi("Codice Fiscale (es. 16823051004)") or SEGNAPOSTO_VUOTO
    pec = _chiedi("PEC") or SEGNAPOSTO_VUOTO
    sdi = _chiedi("Codice identificativo fatturazione elettronica (SDI)") or SEGNAPOSTO_VUOTO

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
# MANIPOLAZIONE DOCUMENTO
# =========================================================================
def imposta_testo_paragrafo(paragrafo, nuovo_testo):
    """
    Sostituisce il testo di un paragrafo mantenendo la formattazione del PRIMO run.
    Tutti i run vengono svuotati tranne il primo, che riceve il nuovo testo.
    """
    if not paragrafo.runs:
        paragrafo.add_run(nuovo_testo)
        return
    paragrafo.runs[0].text = nuovo_testo
    for r in paragrafo.runs[1:]:
        r.text = ""


def imposta_testo_cella(cell, nuovo_testo):
    """
    Sostituisce il testo di una cella mantenendo la formattazione del primo run
    del primo paragrafo non vuoto. Gli altri paragrafi/run vengono svuotati.
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
    # svuota gli altri paragrafi
    for p in paragrafi:
        if p is principale:
            continue
        for r in p.runs:
            r.text = ""


def compila_destinatario(doc, dati):
    """Riscrive il blocco destinatario (paragrafi 0-7, allineati a destra)."""
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
    - Numero offerta: compilato con il progressivo (es. 20260044 / 20260066 ...)
    - Data odierna (del <gg/mm/aaaa>)
    - Sottotitolo: annualita' (I°/II°) + codice edizione (CORSOX)
      Es: 'Apprendistato I° annualita' CORSO7'
    """
    p = doc.paragraphs
    # PARA10: numero offerta
    imposta_testo_paragrafo(p[10], f"OFFERTA ECONOMICA n. {numero_offerta}")
    # PARA11: data odierna
    oggi = date.today().strftime("%d/%m/%Y")
    imposta_testo_paragrafo(p[11], f"\xa0del {oggi}\xa0")
    # PARA13: sottotitolo -> 'Apprendistato <I°/II°> annualita' CORSO<edizione>'
    nuovo_sottotitolo = f"Apprendistato {dati['ann_circ']} annualità {codice_corso(edizione)}"
    imposta_testo_paragrafo(p[13], nuovo_sottotitolo)


def compila_tabella_costi(doc, dati, totale, finale, edizione):
    """
    Compila la tabella Costi (3x4):
      - descrizione (cella 1,0): aggiorna annualita' (I/II) ed edizione (Ed. 0X)
      - importo unitario (1,1), n.partecipanti (1,2), totale (1,3), finale (2,3)
    """
    t = doc.tables[0]
    # cella[1,0] descrizione: "Competenze di base e trasversale <I/II> Annualita' Ed. 0X"
    nuova_desc = f"Competenze di base e trasversale {dati['ann_desc']} Annualità {codice_edizione_padded(edizione)}"
    imposta_testo_cella(t.cell(1, 0), nuova_desc)
    # cella[1,1] = 480,00 € (unitario, fisso)
    imposta_testo_cella(t.cell(1, 1), f"\xa0\n{euro(UNITARIO)}\xa0")
    # cella[1,2] = n.partecipanti
    imposta_testo_cella(t.cell(1, 2), f"\xa0\n{dati['n_partecipanti']}\n\n\xa0")
    # cella[1,3] = totale calcolato
    imposta_testo_cella(t.cell(1, 3), f"\xa0\n{euro(totale).replace(' €', '€')}\xa0")
    # cella[2,3] = finale (PREZZO MIGLIOR FAVORE) - bold come nell'originale
    imposta_testo_cella(t.cell(2, 3), euro(finale).replace(" €", "€"))


# ---------------------------------------------------------------------------
# Tabella calendario (inserimento come nuova tabella nel corpo del documento)
# ---------------------------------------------------------------------------
def _set_cella_tabella(cell, testo, bold=False, font_name="Arial", size_pt=11):
    """Imposta testo di una cella di tabella nuova, font Arial centrato."""
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
    """Aggiunge bordi a tutte le celle di una tabella."""
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
    Rimuove il paragrafo 'Come gia' concordato.' e inserisce subito dopo
    il titolo 'Calendario' una tabella con le giornate, a 6 colonne:
    Data | Orario | Ore | Modulo | Docente | Modalità (come nel template
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

    n_righe = len(righe_calendario) + 1  # +1 intestazione
    n_colonne = len(HEADERS_CALENDARIO)  # 6 colonne
    tabella = doc.add_table(rows=n_righe, cols=n_colonne)
    tabella.alignment = 1  # CENTER
    tabella.autofit = True

    # intestazione
    for ci, header in enumerate(HEADERS_CALENDARIO):
        _set_cella_tabella(tabella.cell(0, ci), header, bold=True)
    # righe dati: ogni chiave del dict riga va nella colonna corrispondente
    # all'header (Data, Orario, Ore, Modulo, Docente, Modalità)
    for ri, riga in enumerate(righe_calendario, start=1):
        for ci, header in enumerate(HEADERS_CALENDARIO):
            _set_cella_tabella(tabella.cell(ri, ci), riga.get(header, ""))

    _bordi_tabella(tabella)

    # sposta la tabella subito dopo il titolo 'Calendario'
    elemento_tabella = tabella._tbl
    corpo.remove(elemento_tabella)
    if par_calendario is not None:
        par_calendario._element.addnext(elemento_tabella)

    # rimuove 'Come gia' concordato.'
    if par_concordato is not None:
        par_concordato._element.getparent().remove(par_concordato._element)


# =========================================================================
# SALVATAGGIO
# =========================================================================
def sanifica_nome(s):
    """Rende una stringa sicura come nome file."""
    s = s.strip()
    s = re.sub(r"[\\/:*?\"<>|]", "", s)
    s = s.replace(" ", "_")
    if not s:
        s = "offerta"
    return s


def salva_offerta(doc, dati, edizione):
    """Salva il documento compilato in output/ e ritorna il percorso.
    Include il codice edizione nel nome file."""
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
    print("  GENERATORE OFFERTE - APPRENDISTATO PRIVATO")
    print("=" * 70)

    if not os.path.isfile(TEMPLATE):
        print(f"\nERRORE: template non trovato in:\n  {TEMPLATE}")
        print("Assicurati che template_offerta_apprendistato.docx sia presente nella cartella.")
        sys.exit(1)

    # --- 1. calendario (PRIMA di tutto) ---
    info_cal = chiedi_calendario()
    righe_calendario = info_cal["righe"]
    edizione = info_cal["edizione"]
    calendario_concordato = info_cal["concordato"]

    # --- 2. dati (partecipanti, annualita', azienda) ---
    dati = chiedi_dati()

    # --- 3. calcoli ---
    totale, finale = calc_totali(dati["n_partecipanti"])
    # numero offerta progressivo (1a annualita' da 20260044, 2a da 20260066)
    numero_offerta = prossimo_numero_offerta(dati["ann_num"])
    print("\n--- CALCOLI ---")
    print(f"  Importo unitario:      {euro(UNITARIO)}")
    print(f"  N. partecipanti:       {dati['n_partecipanti']}")
    print(f"  Importo totale:        {euro(totale)}")
    print(f"  Sconto {int(SCONTO*100)}%:              {euro(totale*(1-SCONTO))}")
    print(f"  Prezzo finale (arrotondato alla centinaio): {euro(finale)}")
    print(f"  Edizione:              {codice_corso(edizione)} / {codice_edizione_padded(edizione)}")
    print(f"  Annualita':            {dati['ann_testo']}")
    print(f"  Numero offerta:        {numero_offerta}")

    # --- 4. conferma prima di generare ---
    if not _si_no("\nProcedere con la generazione dell'offerta?", default="s"):
        print("Generazione annullata.")
        return

    # --- 5. carica il template e compila ---
    print("\nGenerazione del documento in corso...")
    doc = docx.Document(TEMPLATE)

    compila_destinatario(doc, dati)
    compila_titolo_data_annualita_edizione(doc, dati, edizione, numero_offerta)
    compila_tabella_costi(doc, dati, totale, finale, edizione)
    if righe_calendario:  # solo se non concordato
        inserisci_tabella_calendario(doc, righe_calendario)

    # --- 5b. registra il numero offerta usato nel contatore persistente ---
    aggiorna_contatore(dati["ann_num"], numero_offerta)

    # --- 6. salva ---
    percorso = salva_offerta(doc, dati, edizione)
    print(f"\n  >> Offerta generata:\n     {percorso}")

    # --- 7. opzione apertura ---
    if _si_no("Vuoi aprire il documento ora?", default="s"):
        try:
            os.startfile(percorso)  # Windows
        except Exception as e:
            print(f"  Impossibile aprire automaticamente ({e}). Apri manualmente il file.")

    print("\nOperazione completata.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrotto dall'utente.")
        sys.exit(0)
