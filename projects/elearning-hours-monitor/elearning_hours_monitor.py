"""
=============================================================================
MONITORAGGIO ORE DI FORMAZIONE E-LEARNING
=============================================================================
Lo script rileva automaticamente i file di input presenti nella cartella
e produce un UNICO file Excel con un foglio per ogni mese elaborato.

  Celle giornaliere -> ore EFFETTIVE (= ore totali - ore in eccesso)
  Note sulle celle  -> "Ore totali" + "Eccesso" (solo se c'e' eccesso)

  Colonne riepilogative (sempre nello stesso ordine, MODALITA A e MODALITA B):
      1) Totale Ore Effettive  (HH:MM:SS)
      2) Ore Totali            (HH:MM:SS)
      3) Totale Eccesso        (HH:MM:SS)  = Ore Totali - Totale Ore Effettive
      4) Eccesso Weekend       (HH:MM:SS)
      5) Eccesso Dopo 18:00    (HH:MM:SS)
      6) Eccesso Mattutino     (HH:MM:SS)  (06:00-07:40)

  MODALITA A - Solo CSV
  MODALITA B - CSV + LUL (Presenze XLSX): la regola dell'eccesso ore LAV
      e del cap 8h/giorno viene comunque applicata internamente e confluisce
      nel Totale Eccesso, senza una colonna dedicata.

  MODALITA B - parsing LUL (V17):
      - leggi_mappa_colonne gestisce: "1 L", "1", date openpyxl
      - parse_blocco_dipendente cerca LAV anche in col 1, gestisce assenze
        con codice (F, M, ROL...) anche senza ore >= 8
      - carica_lul cerca la riga 'ORE' in tutti i fogli del workbook
      - calcola_giorno distingue: lul=None (no abbinamento -> come Mod.A),
        lav=None (giorno fuori LUL -> come Mod.A), lav=0 (giorno non lavorato
        -> assente), lav>0 (giorno lavorato -> cap a lav_ore)

SELEZIONE DEL PERIODO:
  - Imposta MESE = numero del mese (es. 3 per Marzo) per un mese specifico
  - Imposta MESE = None per elaborare TUTTI i mesi presenti nel CSV
  In entrambi i casi viene prodotto un unico file Excel con un foglio per mese.

DIPENDENZE:
    pip install pandas openpyxl

UTILIZZO:
    python monitoraggio_formazione.py
=============================================================================
"""

import os
import re
import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment


# =============================================================================
# CONFIGURAZIONE
# =============================================================================

CSV_FILE    = "Report_Accessi.csv"
# Il LUL puo' essere un PDF (stampa piattaforma orari) oppure un XLSX.
# Lo script rileva automaticamente il primo file disponibile tra questi nomi.
# IMPORTANTE: il PDF e' la fonte piu' affidabile. La conversione del PDF in
# Excel spesso PERDE i nomi dei dipendenti (tranne il primo), quindi quando
# possibile usare direttamente il PDF.
LUL_FILE_PDF  = "Presenze.pdf"
LUL_FILE_XLSX = "Presenze.xlsx"
LUL_FILE      = LUL_FILE_XLSX   # compatibilita': impostato dinamicamente in main()
AZIENDA_BREVE = "aicomply"   # nome breve dell'azienda per il file di output
# Il nome del file di output include automaticamente la data di oggi e l'azienda
# Es: Monitoraggio_15-04-2026.xlsx
OUTPUT_FILE = f"Monitoraggio_{datetime.date.today().strftime('%d-%m-%Y')}.xlsx"

# Mese da elaborare:
#   MESE = None    -> elabora TUTTI i mesi e anni presenti nel CSV (consigliato)
#   MESE = 32026   -> elabora solo Marzo 2026     (3  + 2026)
#   MESE = 112025  -> elabora solo Novembre 2025  (11 + 2025)
#   MESE = 12026   -> elabora solo Gennaio 2026   (1  + 2026)
# Formula: scrivi il numero del mese seguito dall anno a 4 cifre
# Nota: con MESE = None la variabile ANNO qui sotto viene ignorata
MESE = None

CSV_SEPARATOR      = ";"
SOGLIA_ORA         = 18.0
SOGLIA_MATTINO_INI = 6.0
SOGLIA_MATTINO_FIN = 7 + 40/60
TOLLERANZA_MINUTI  = 20
CAP_ORE_GIORNALIERO = 8.0    # ore massime riconoscibili per discente in un singolo giorno
ORE_LIMITE_FINANZIATE = 150  # ore massime riconoscibili per discente (per la colonna MIN nel Riepilogo)


# =============================================================================
# COSTANTI
# =============================================================================

NOMI_MESI = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
              "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

COLORI = {
    # Stati cella (tonalita' molto chiare per ridurre il "rumore" visivo)
    "verde":   "F0F6EC",   # fruizione normale
    "arancio": "FBEAD8",   # eccesso ore LAV
    "viola":   "F2E8EE",   # eccesso dopo soglia
    "azzurro": "E8EFF5",   # eccesso mattino (06:00-07:40)
    "rosa":    "FAE6E6",   # fruizione nel weekend
    "giallo":  "FBF4D9",   # giorno assente con fruizione
    "grigio":  "F2F2F2",   # weekend senza fruizione
    # UI
    "header":  "F4F5F7",   # intestazioni - grigio chiarissimo
    "titolo":  "ECF0F4",   # sfondo titolo - azzurro tenuissimo
    "titolo_testo": "2C3E50",  # colore testo titolo
    "totale":  "F1F3F5",   # riga totale - grigio chiarissimo
    "bordo":   "BDC3C7",   # bordi sottili
}


# Formato Excel per durate (anche superiori a 24 ore)
DURATION_FORMAT = "[h]:mm:ss"


# =============================================================================
# UTILITA GENERALI
# =============================================================================

def giorni_nel_mese(anno, mese):
    if mese == 12:
        return (datetime.date(anno + 1, 1, 1) - datetime.date(anno, mese, 1)).days
    return (datetime.date(anno, mese + 1, 1) - datetime.date(anno, mese, 1)).days


def abbreviazione_giorno(anno, mese, giorno):
    return ["L", "M", "M", "G", "V", "S", "D"][
        datetime.date(anno, mese, giorno).weekday()
    ]


def giorni_weekend(anno, mese):
    totale = giorni_nel_mese(anno, mese)
    return {d for d in range(1, totale + 1)
            if abbreviazione_giorno(anno, mese, d) in ("S", "D")}


def normalizza_nome(nome):
    """Normalizza un nome per l'ABBINAMENTO LUL<->CSV: ordina alfabeticamente
    i token cosi' "MARIO ROSSI" e "ROSSI MARIO" coincidono. NON usare questa
    funzione per ordinare i discenti nei fogli (vedi chiave_ordinamento)."""
    token = re.sub(r"[^A-Za-z ]", "", nome.upper()).split()
    return " ".join(sorted(token))


def chiave_ordinamento(nome):
    """Chiave per l'ORDINAMENTO ALFABETICO dei discenti nei fogli, per COGNOME.

    Nei file il nome e' scritto "COGNOME NOME" (es. "ROSSI FABIO"), quindi
    la stringa ripulita e maiuscola gia' ordina per cognome. A differenza di
    normalizza_nome, qui NON si riordinano i token: l'ordine delle parole
    (cognome prima) viene preservato."""
    return re.sub(r"[^A-Za-z ]", "", str(nome).upper()).strip()


def normalizza_ordine_nomi(discenti_globali, dipendenti_lul):
    """Riordina i nomi dei discenti nel formato "COGNOME NOME" usando il LUL
    come riferimento.

    Nel CSV i nomi possono essere scritti "Nome Cognome" (es. "Mario Rossi"),
    mentre nel LUL sono "Cognome Nome" (es. "ROSSI MARIO"). Per ordinare e
    mostrare i discenti per COGNOME servono i token nell'ordine giusto.

    Strategia:
      - per ogni discente si cercano nel LUL i token corrispondenti (match per
        prefisso, robusto ai troncamenti del PDF e all'ordine invertito);
      - se si trova il dipendente LUL, si adotta l'ordine dei token del LUL
        (cognome prima), completando pero' i token troncati con la versione
        intera presa dal CSV (es. LUL "GIOVA" -> si tiene "GIOVANNI" dal CSV);
      - se il discente NON e' nel LUL, si applica il fallback: si assume che
        l'ULTIMA parola sia il cognome e lo si porta in testa.
    Il campo "nome" di ogni discente viene aggiornato in "COGNOME NOME".
    """
    def _tok(n):
        return [t for t in re.sub(r"[^A-Za-z ]", " ", str(n).upper()).split() if t]

    def _match(a, b):
        if a == b:
            return True
        corto, lungo = (a, b) if len(a) <= len(b) else (b, a)
        return len(corto) >= 3 and lungo.startswith(corto)

    lul_tokens = [(_tok(d["name"]), d) for d in (dipendenti_lul or [])]

    for cf, info in discenti_globali.items():
        tok_csv = _tok(info["nome"])
        if not tok_csv:
            continue

        # Cerca il dipendente LUL compatibile
        ordine_lul = None
        for tlul, d in lul_tokens:
            piccolo, grande = (tok_csv, tlul) if len(tok_csv) <= len(tlul) else (tlul, tok_csv)
            disp = list(grande)
            ok = True
            for t in piccolo:
                trovato = next((u for u in disp if _match(t, u)), None)
                if trovato is None:
                    ok = False
                    break
                disp.remove(trovato)
            if ok and len(piccolo) >= 2:
                ordine_lul = tlul
                break

        if ordine_lul:
            # Adotta l'ordine del LUL, ma per ogni token del LUL preferisci la
            # versione PIU' LUNGA tra LUL e CSV (cosi' i troncamenti del PDF
            # vengono completati con il nome intero del CSV).
            csv_disp = list(tok_csv)
            nuovi = []
            for tl in ordine_lul:
                scelto = tl
                for tc in list(csv_disp):
                    if _match(tl, tc):
                        scelto = tc if len(tc) >= len(tl) else tl
                        csv_disp.remove(tc)
                        break
                nuovi.append(scelto)
            # eventuali token CSV non abbinati (rari) si accodano
            nuovi.extend(csv_disp)
            info["nome"] = " ".join(nuovi).upper()
        else:
            # Fallback (discente NON nel LUL): si assume "Nome ... Cognome", quindi
            # il cognome e' in coda. Si gestiscono i cognomi con particella
            # (DE, DEL, DELLA, DI, DA, LO, LA, LE, VAN, VON, MC, ...): la
            # particella che precede l'ultima parola fa parte del cognome.
            PARTICELLE = {"DE","DEL","DELLA","DELLE","DELLO","DEI","DEGLI","DI","DA",
                          "DAL","DALLA","LO","LA","LE","LI","VAN","VON","MC","MAC",
                          "SAN","SANTA","SANT","D"}
            if len(tok_csv) >= 2:
                # quante parole finali compongono il cognome?
                n_cog = 1
                # includi particelle immediatamente precedenti il cognome
                while len(tok_csv) - n_cog - 1 >= 1 and tok_csv[-(n_cog + 1)] in PARTICELLE:
                    n_cog += 1
                cognome = tok_csv[-n_cog:]
                nome    = tok_csv[:-n_cog]
                info["nome"] = (" ".join(cognome) + " " + " ".join(nome)).upper()
            else:
                info["nome"] = tok_csv[0].upper()


def fill_cell(colore):
    if not colore:
        return None
    return PatternFill("solid", start_color=colore, fgColor=colore)


LOGO_AZIENDA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "company-logo.png")
LOGO_FNC       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fnc-logo.png")


def inserisci_loghi(ws, riga_logo=2, altezza_riga=40):
    """Inserisce i due loghi nella riga indicata (default: riga 2).
    Logo aziendale a sinistra (cella A), FNC a destra (ultima colonna visibile).
    Gestisce gracefully l'assenza dei file immagine."""
    from openpyxl.drawing.image import Image as XLImage
    ws.row_dimensions[riga_logo].height = altezza_riga
    for path, anchor in [(LOGO_AZIENDA, "A2"), (LOGO_FNC, None)]:
        if not os.path.isfile(path):
            continue
        try:
            img = XLImage(path)
            # Scala mantenendo aspect ratio all'altezza target
            h_target = altezza_riga * 1.33   # punti -> pixel approssimativi
            scale    = h_target / img.height
            img.width  = int(img.width  * scale)
            img.height = int(img.height * scale)
            if anchor is None:
                # Calcola l'ultima colonna del merge del titolo per posizionare FNC
                # (usa colonna N abbastanza a destra — la funzione chiamante può
                # passare la lettera corretta via parametro se serve)
                anchor = "B2"   # placeholder; verrà sovrascritto dal chiamante
            img.anchor = anchor
            ws.add_image(img)
        except Exception:
            pass   # Se Pillow/openpyxl non reggono il file, ignora silenziosamente


def ore_decimali_a_hhmmss(ore_dec):
    if not ore_dec or ore_dec <= 0:
        return None
    totale_secondi = int(round(ore_dec * 3600))
    h = totale_secondi // 3600
    m = (totale_secondi % 3600) // 60
    s = totale_secondi % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


# =============================================================================
# PARSING CSV
# =============================================================================

def parse_durata(testo):
    m = re.match(r"(\d+)\s*h\s*(\d+)\s*min\s*(\d+)\s*s", str(testo))
    if m:
        return (int(m.group(1)) * 3600 + int(m.group(2)) * 60 + int(m.group(3))) / 3600
    return 0.0


def ora_in_decimale(testo):
    m = re.match(r"(\d+):(\d+):(\d+)", str(testo))
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60 + int(m.group(3)) / 3600
    m = re.match(r"(\d+):(\d+)", str(testo))
    if m:
        return int(m.group(1)) + int(m.group(2)) / 60
    return 0.0


def splitta_ore_per_soglia(riga, soglia):
    durata   = parse_durata(riga["Totale Ore"])
    t_inizio = ora_in_decimale(riga["Primo Accesso"])
    t_fine   = ora_in_decimale(riga["Ultimo Accesso"])
    if t_fine <= soglia:
        return durata, 0.0
    if t_inizio >= soglia:
        return 0.0, durata
    # Usa l'intervallo reale della sessione come denominatore (non la durata fruita),
    # per calcolare correttamente la frazione di studio proporzionale al tempo dopo soglia.
    elapsed = t_fine - t_inizio
    frazione_dopo = (t_fine - soglia) / elapsed if elapsed > 0 else 0
    # Clamp per sicurezza numerica
    frazione_dopo = max(0.0, min(1.0, frazione_dopo))
    ore_dopo = durata * frazione_dopo
    return durata - ore_dopo, ore_dopo


def ore_in_finestra_mattino(riga):
    """
    Calcola le ore di studio (proporzionali al tempo trascorso) che cadono
    nella finestra mattutina [SOGLIA_MATTINO_INI, SOGLIA_MATTINO_FIN].
    Usa la stessa logica proporzionale di splitta_ore_per_soglia: la frazione
    di studio in finestra e' proporzionale alla frazione di tempo elapsed in
    finestra, non alla durata fruita.
    """
    durata   = parse_durata(riga["Totale Ore"])
    t_inizio = ora_in_decimale(riga["Primo Accesso"])
    t_fine   = ora_in_decimale(riga["Ultimo Accesso"])
    # Sessione interamente fuori dalla finestra mattutina
    if t_fine <= SOGLIA_MATTINO_INI or t_inizio >= SOGLIA_MATTINO_FIN:
        return 0.0
    elapsed = t_fine - t_inizio
    if elapsed <= 0:
        return 0.0
    overlap_ini = max(t_inizio, SOGLIA_MATTINO_INI)
    overlap_fin = min(t_fine,   SOGLIA_MATTINO_FIN)
    overlap = max(0.0, overlap_fin - overlap_ini)
    if overlap == 0:
        return 0.0
    frazione = overlap / elapsed
    frazione = max(0.0, min(1.0, frazione))
    return durata * frazione


def fmt_durata_breve(ore_dec):
    """Formatta una durata decimale in modo compatto (per le note delle celle).
    Es. 1.5h -> '1 h 30 min', 0.42h -> '25 min', 0.005h -> '18 s'."""
    sec_tot = int(round(ore_dec * 3600))
    h = sec_tot // 3600
    m = (sec_tot % 3600) // 60
    s = sec_tot % 60
    if h > 0:
        return f"{h} h {m:02d} min"
    if m > 0:
        return f"{m} min {s:02d} s" if s > 0 else f"{m} min"
    return f"{s} s"


def carica_csv_completo(percorso_file):
    print(f"Caricamento CSV: {percorso_file}")
    df = pd.read_csv(percorso_file, sep=CSV_SEPARATOR, encoding="utf-8-sig")
    df["Giorno"] = pd.to_datetime(df["Giorno"], format="%d/%m/%Y")
    return df


def mesi_disponibili_nel_csv(df):
    periodi = df[["Giorno"]].copy()
    periodi["anno"] = periodi["Giorno"].dt.year
    periodi["mese"] = periodi["Giorno"].dt.month
    unici = periodi[["anno", "mese"]].drop_duplicates().sort_values(["anno", "mese"])
    return list(unici.itertuples(index=False, name=None))


def elabora_mese_dal_csv(df_completo, mese, anno, soglia):
    df_mese = df_completo[
        (df_completo["Giorno"].dt.month == mese) &
        (df_completo["Giorno"].dt.year  == anno)
    ].copy()

    if df_mese.empty:
        return pd.DataFrame(), pd.DataFrame(), "N/D"

    df_mese[["ore_prima_raw", "ore_dopo"]] = df_mese.apply(
        lambda r: pd.Series(splitta_ore_per_soglia(r, soglia)), axis=1)
    df_mese["ore_mattino"] = df_mese.apply(ore_in_finestra_mattino, axis=1)
    df_mese["ore_tot"]     = df_mese.apply(lambda r: parse_durata(r["Totale Ore"]), axis=1)
    # ore_prima = ore "valide" (orario normale = tra fine finestra mattutina e soglia)
    # = ore_prima_raw (prima della soglia) meno la parte che cade nella finestra mattino
    df_mese["ore_prima"] = (df_mese["ore_prima_raw"] - df_mese["ore_mattino"]).clip(lower=0)
    df_mese["day"] = df_mese["Giorno"].dt.day

    # Raggruppamento per (Codice Fiscale, giorno): SOMMA le ore di TUTTE le
    # sessioni di quel giorno, INCLUSI percorsi formativi diversi della stessa
    # persona. Se un discente segue due o piu' percorsi e nello stesso giorno
    # ha sessioni su piu' di essi (anche con orari sovrapposti), le ore vengono
    # sommate (somma semplice: eventuali sovrapposizioni temporali sono conteggiate).
    df_giornaliero = (
        df_mese.groupby(["Codice Fiscale", "day"])
        .agg(ore_prima  =("ore_prima",   "sum"),
             ore_dopo   =("ore_dopo",    "sum"),
             ore_mattino=("ore_mattino", "sum"),
             ore_tot    =("ore_tot",     "sum"))
        .reset_index()
    )

    # Elenco discenti: un discente = un Codice Fiscale. Se ha piu' percorsi,
    # vengono elencati tutti, concatenati con " | ".
    discenti_info = (
        df_mese.groupby("Codice Fiscale")
        .agg(nome    =("Nome Cognome", "first"),
             percorso=("Percorso",     lambda x: " | ".join(sorted(set(x.dropna().astype(str))))))
        .reset_index()
    )
    discenti_info["sort_key"] = discenti_info["nome"].apply(chiave_ordinamento)
    discenti_info = discenti_info.sort_values("sort_key").reset_index(drop=True)

    azienda = df_mese["Azienda"].iloc[0]
    print(f"  {NOMI_MESI[mese]} {anno}: {len(discenti_info)} discenti | Azienda: {azienda}")
    return df_giornaliero, discenti_info, azienda


# =============================================================================
# PARSING LUL
# =============================================================================

def leggi_mappa_colonne(riga_ore):
    """Costruisce la mappa {indice_colonna: numero_giorno} dalla riga intestazione del LUL.

    Gestisce i vari formati con cui le aziende numerano i giorni nella riga 'ORE':
      - "1 L", "2 M", ... (formato classico con spazio + lettera giorno)
      - "1", "2", ...     (solo numero intero, senza lettera)
      - datetime.date     (alcuni LUL restituiscono date openpyxl)
      - Salta "ORE", "TOT" e celle vuote.
    """
    col_giorno = {}
    for ci, valore in enumerate(riga_ore):
        if valore is None:
            continue
        # Oggetto data openpyxl -> usa il giorno del mese
        if isinstance(valore, (datetime.date, datetime.datetime)):
            col_giorno[ci] = valore.day
            continue
        s = str(valore).strip()
        if not s or s in ("ORE", "TOT"):
            continue
        # "1 L", "15 M", ecc.
        m = re.match(r"^(\d{1,2})\s+[A-Za-z]", s)
        if m:
            col_giorno[ci] = int(m.group(1))
            continue
        # Solo numero intero (es. "1", "15")
        if re.match(r"^\d{1,2}$", s):
            col_giorno[ci] = int(s)
    return col_giorno


def estrai_nome_dipendente(testo):
    m = re.search(
        r"Dipendente:\s*\d+\s*-\s*([A-Z][A-Z ]+?)(?:\s{2,}|\n|$|Data)", testo)
    return m.group(1).strip() if m else "SCONOSCIUTO"


def parse_blocco_dipendente(righe, indice_inizio, col_giorno):
    """Estrae le ore LAV e i giorni di assenza dal blocco di un dipendente nel LUL.

    Miglioramenti rispetto alla versione precedente:
    - Cerca la riga LAV anche quando il label è in una cella non prima (es. col 1).
    - Gestisce assenze con ore < 8 ma con codice assenza significativo (es. "F", "M").
    - Gestisce il formato alternativo ASSENZE con più righe di codici.
    - Distingue tra "giorno non presente nel LUL" (dipendente non in quel mese)
      e "giorno con LAV = 0" (dipendente presente ma senza ore lavorate dichiarate).
    - Non conta come "assente" un giorno con lav_ore > 0 anche se ha una voce assenza.
    """
    nome    = estrai_nome_dipendente(str(righe[indice_inizio][0]))
    lav     = {}
    assenti = set()

    # Codici assenza che indicano giorno non lavorato (in aggiunta al criterio ore >= 8)
    CODICI_ASSENZA = {"F", "M", "MR", "P", "ROL", "EX", "AL", "ASP", "INF",
                      "MAL", "CIG", "CIGS", "0", "ART", "SOS", "PERM"}

    for j in range(indice_inizio + 1, min(indice_inizio + 120, len(righe))):
        riga = righe[j]
        if not riga:
            continue
        # Cerca il label in colonna 0 o colonna 1 (alcuni LUL hanno una colonna
        # descrittiva in col 0 e il label vero in col 1)
        v0 = riga[0]
        v1 = riga[1] if len(riga) > 1 else None
        label = None
        if isinstance(v0, str):
            label = v0.strip()
        elif isinstance(v1, str):
            label = v1.strip()

        # ── Riga LAV ─────────────────────────────────────────────────────────
        if label == "LAV":
            for ci, val in enumerate(riga):
                if ci in col_giorno:
                    if isinstance(val, (int, float)):
                        lav[col_giorno[ci]] = float(val)
                    elif val is None or val == "":
                        # Giorno presente nella mappa ma senza valore ->
                        # marca esplicitamente come 0 (dipendente in quel mese,
                        # giorno senza ore lavorate dichiarate)
                        lav.setdefault(col_giorno[ci], 0.0)

        # ── Sezione ASSENZE ──────────────────────────────────────────────────
        elif isinstance(label, str) and "A S S E N Z E" in label:
            riga_codici = riga
            riga_ore    = righe[j + 1] if j + 1 < len(righe) else [None] * 50
            for ci in col_giorno:
                giorno_num = col_giorno[ci]
                codice  = riga_codici[ci] if ci < len(riga_codici) else None
                ore_ass = riga_ore[ci]    if ci < len(riga_ore)    else None
                if not codice:
                    continue
                codice_s = str(codice).strip().upper()
                # Assenza certa: ore >= 8 intere giornata
                if isinstance(ore_ass, (int, float)) and ore_ass >= 8:
                    assenti.add(giorno_num)
                # Assenza per codice riconosciuto (anche mezza giornata)
                elif codice_s in CODICI_ASSENZA:
                    # Solo se non ha già ore LAV > 0 registrate
                    if lav.get(giorno_num, 0.0) == 0.0:
                        assenti.add(giorno_num)

        # ── Fine blocco ───────────────────────────────────────────────────────
        elif j > indice_inizio + 2 and isinstance(label, str):
            if "PIATTAFORMA" in label or label.startswith("Azienda") or "Dipendente" in label:
                break

    return {"name": nome, "lav": lav, "assente": assenti}


def carica_lul(percorso_file):
    """Carica il LUL (Presenze.xlsx) e restituisce la lista dei dipendenti con
    ore LAV e giorni di assenza per ogni giorno del mese.

    Miglioramenti rispetto alla versione precedente:
    - Cerca la riga 'ORE' in tutte le righe (non solo le prime), utile per LUL
      con intestazioni aziendali di lunghezza variabile.
    - Prova tutti i fogli del workbook se nel foglio attivo non trova la mappa.
    - Log di debug sui dipendenti trovati e non trovati.
    """
    print(f"Caricamento LUL: {percorso_file}")
    wb = openpyxl.load_workbook(percorso_file, data_only=True)

    # Prova prima il foglio attivo, poi tutti gli altri
    fogli_da_provare = [wb.active] + [wb[s] for s in wb.sheetnames if wb[s] != wb.active]

    righe      = None
    col_giorno = {}

    for ws in fogli_da_provare:
        righe_candidate = list(ws.iter_rows(values_only=True))
        # Cerca la riga 'ORE' (intestazione dei giorni)
        for riga in righe_candidate:
            if riga and riga[0] == "ORE":
                col_giorno = leggi_mappa_colonne(riga)
                if col_giorno:
                    righe = righe_candidate
                    print(f"  Mappa colonne trovata nel foglio '{ws.title}' "
                          f"({len(col_giorno)} giorni)")
                    break
        if col_giorno:
            break

    if not col_giorno or righe is None:
        print("  ATTENZIONE: impossibile trovare la riga 'ORE' con la mappa giorni nel LUL.")
        print("  Verifica che il LUL abbia una riga con 'ORE' in colonna A e i giorni "
              "nel formato '1 L', '2 M', ... (o solo numero) nelle colonne successive.")
        return []

    # Individua gli inizi dei blocchi dipendente.
    # Una riga inizia un blocco se contiene "Dipendente:" in colonna A.
    # NON si vincola la posizione di "PIATTAFORMA" (che nei LUL reali può
    # comparire in coda alla riga, oltre i primi caratteri): basta che la
    # parola sia presente nella riga, oppure che la riga inizi con "Azienda".
    inizi_blocchi = [
        i for i, riga in enumerate(righe)
        if riga and isinstance(riga[0], str)
        and re.search(r"Dipendente\s*:", riga[0])
    ]

    if not inizi_blocchi:
        print("  ATTENZIONE: nessun blocco dipendente trovato nel LUL. "
              "Verifica il formato (atteso: riga con 'Dipendente:' in colonna A).")
        return []

    dipendenti = [parse_blocco_dipendente(righe, i, col_giorno) for i in inizi_blocchi]
    # Filtra eventuali blocchi vuoti / senza nome
    dipendenti = [d for d in dipendenti if d["name"] != "SCONOSCIUTO" or d["lav"]]
    print(f"  Trovati {len(dipendenti)} dipendenti nel LUL")
    return dipendenti


def carica_lul_da_pdf(percorso_file):
    """Carica il LUL direttamente dal PDF di piattaforma orari (stampa dettagliata
    controllo presenze) e restituisce la lista dei dipendenti con ore LAV e
    giorni di assenza.

    Perche' leggere il PDF e non l'Excel convertito:
    nel PDF ogni dipendente ha la propria intestazione di pagina con nome e
    "Ore lavorate". La conversione PDF->Excel tipicamente conserva solo la
    prima intestazione e rende anonimi gli altri blocchi: leggere il PDF
    risolve il problema alla radice.

    Tecnica di parsing:
    - una pagina = un dipendente;
    - dall'intestazione si estrae nome e "Ore lavorate" (per validazione);
    - dalla riga 'ORE' si mappa ogni numero-giorno alla sua coordinata X;
    - la riga 'LAV' contiene le ore lavorate, allineate per X alla colonna
      giorno piu' vicina (le colonne nel PDF non sono equispaziate);
    - un giorno con LAV > 0 e' lavorato; i giorni feriali senza valore LAV
      (festivi, ferie, malattia, ROL, permessi) vengono marcati a 0.0 e
      trattati a valle come non lavorati.
    """
    try:
        import pdfplumber
    except ImportError:
        print("  ERRORE: per leggere il LUL in PDF serve la libreria 'pdfplumber'.")
        print("          Installala dal terminale con:  pip install pdfplumber")
        print("          (oppure:  python -m pip install pdfplumber)")
        print("          Senza LUL i fogli saranno generati in Modalita' A (solo CSV).")
        return [], None
    from collections import defaultdict

    def _num(s):
        return float(str(s).replace(",", "."))

    print(f"Caricamento LUL (PDF): {percorso_file}")
    pdf = pdfplumber.open(percorso_file)
    dipendenti = []
    senza_ore  = 0
    mese_lul   = None   # (num_mese, anno) rilevato dall'intestazione del PDF

    MESI_IT = {"gennaio":1,"febbraio":2,"marzo":3,"aprile":4,"maggio":5,
               "giugno":6,"luglio":7,"agosto":8,"settembre":9,
               "ottobre":10,"novembre":11,"dicembre":12}

    for pagina in pdf.pages:
        words = pagina.extract_words(x_tolerance=1, y_tolerance=3)
        if not words:
            continue
        testo = pagina.extract_text() or ""
        mnome = re.search(r"Dipendente:\s*\d+\s*-\s*(.+?)\s+Data Ass", testo)
        if not mnome:
            continue
        nome = mnome.group(1).strip()
        mlav = re.search(r"Ore lavorate:\s*([\d.,]+)", testo)
        ore_dichiarate = _num(mlav.group(1)) if mlav else None

        # Raggruppa i token per riga (bin di 2px sull'asse verticale)
        righe_y = defaultdict(list)
        for w in words:
            righe_y[round(w['top'] / 2) * 2].append(w)

        # Mappa giorno -> centro X dalla riga 'ORE'
        giorni_x = {}
        ore_ykey = None
        for ykey in sorted(righe_y):
            rw = sorted(righe_y[ykey], key=lambda w: w['x0'])
            if rw and rw[0]['text'] == 'ORE':
                ore_ykey = ykey
                for w in rw:
                    t = w['text'].replace(',', '')
                    if t.isdigit() and 1 <= int(t) <= 31:
                        giorni_x[int(t)] = (w['x0'] + w['x1']) / 2
                break
        if not giorni_x:
            continue

        def _col_giorno(x):
            best, bestd = None, 999
            for g, gx in giorni_x.items():
                d = abs(x - gx)
                if d < bestd:
                    bestd, best = d, g
            return best if bestd < 13 else None

        # Riga 'LAV' -> ore lavorate per giorno (per posizione X)
        lav = {}
        for ykey in sorted(righe_y):
            if ykey <= ore_ykey:
                continue
            rw = sorted(righe_y[ykey], key=lambda w: w['x0'])
            if rw and rw[0]['text'] == 'LAV':
                for w in rw:
                    if w['text'] == 'LAV':
                        continue
                    if re.match(r'^[\d,]+$', w['text']):
                        v = _num(w['text'])
                        if v > 24:      # colonna TOT a fine riga
                            continue
                        g = _col_giorno((w['x0'] + w['x1']) / 2)
                        if g:
                            lav[g] = v
                break   # solo la prima riga 'LAV'

        if not lav:
            senza_ore += 1
        # I giorni feriali del mese SENZA valore LAV (ferie, malattia, ROL,
        # permessi, festivita') vengono marcati esplicitamente a 0.0: cosi'
        # calcola_giorno li tratta come "non lavorati" (eccesso LAV) e non come
        # "giorno fuori dal LUL". I weekend non si toccano: ci pensa la logica
        # weekend di calcola_giorno. Il mese/anno si ricavano dall'intestazione.
        mmese = re.search(r"Mese:\s*([A-Za-z]+)\s+(\d{4})", testo)
        if mmese:
            nome_mese = mmese.group(1).lower()
            anno_pdf  = int(mmese.group(2))
            num_mese = MESI_IT.get(nome_mese)
            if num_mese:
                if mese_lul is None:
                    mese_lul = (num_mese, anno_pdf)
                import calendar as _cal
                ndays = _cal.monthrange(anno_pdf, num_mese)[1]
                for d in range(1, ndays + 1):
                    wd = datetime.date(anno_pdf, num_mese, d).weekday()  # 0=lun..6=dom
                    if wd < 5 and d not in lav:   # feriale e non gia' lavorato
                        lav[d] = 0.0
        dipendenti.append({"name": nome, "lav": lav, "assente": set(),
                           "ore_dichiarate": ore_dichiarate})

    # Validazione: somma LAV == "Ore lavorate" dichiarate
    incongruenti = []
    for d in dipendenti:
        tot = sum(d["lav"].values())
        if d["ore_dichiarate"] is not None and abs(tot - d["ore_dichiarate"]) > 0.5:
            incongruenti.append(f"{d['name']} (LAV={tot:.0f}h vs dich={d['ore_dichiarate']:.0f}h)")

    print(f"  Trovati {len(dipendenti)} dipendenti nel LUL (PDF)")
    if mese_lul:
        print(f"  Mese del LUL: {mese_lul[0]:02d}/{mese_lul[1]} "
              f"(la Modalita' B sara' applicata SOLO a questo mese)")
    else:
        print("  ATTENZIONE: impossibile rilevare il mese dal LUL "
              "(intestazione 'Mese: ...' non trovata).")
    if senza_ore:
        print(f"  Di cui {senza_ore} senza ore lavorate nel mese (es. cessati/assenti).")
    if incongruenti:
        print(f"  ATTENZIONE: {len(incongruenti)} dipendenti con somma LAV diversa dalle ore dichiarate:")
        for s in incongruenti[:10]:
            print(f"     - {s}")
    return dipendenti, mese_lul


def carica_lul_auto(percorso_file):
    """Sceglie automaticamente il parser in base all'estensione del file LUL.

    Restituisce sempre una tupla (dipendenti, mese_lul), dove mese_lul e'
    (num_mese, anno) se rilevato, altrimenti None. Per il formato XLSX il mese
    non e' rilevabile dall'intestazione, quindi mese_lul = None e la Modalita' B
    viene applicata a tutti i mesi (comportamento legacy)."""
    ext = os.path.splitext(percorso_file)[1].lower()
    if ext == ".pdf":
        return carica_lul_da_pdf(percorso_file)
    return carica_lul(percorso_file), None


def abbina_discenti_lul(discenti_info, dipendenti_lul):
    """Abbina i discenti del CSV ai dipendenti del LUL tramite il nome.

    Problema dei PDF piattaforma orari: il nome del dipendente viene TRONCATO a
    circa 19 caratteri (es. "GIORDANETTI ALESSAN" invece di
    "GIORDANETTI ALESSANDRO", "LOMBARDINI BEATRIC" invece di
    "LOMBARDINI BEATRICE"). Un confronto esatto fallirebbe per tutti i nomi
    lunghi, lasciandoli erroneamente in Modalita' A (nessun controllo ore LAV).

    Strategia di abbinamento (in ordine):
      1) match esatto sul nome normalizzato (token ordinati);
      2) match per PREFISSO COGNOME+NOME: il nome del LUL (eventualmente
         troncato) deve essere prefisso del nome completo del CSV, o viceversa,
         confrontando la sequenza dei caratteri senza spazi. Cosi'
         "GIORDANETTI ALESSAN" abbina "GIORDANETTI ALESSANDRO".
    Se un nome del LUL e' prefisso ambiguo di piu' discenti, NON si abbina e si
    segnala, per evitare attribuzioni sbagliate.
    """
    def _tokens(nome):
        # insieme ordinato di token alfabetici maiuscoli, senza accenti spurî
        return [t for t in re.sub(r"[^A-Za-z ]", " ", str(nome).upper()).split() if t]

    def _key(nome):
        # chiave per match esatto: token ordinati alfabeticamente
        return " ".join(sorted(_tokens(nome)))

    def _match_token(a, b):
        # due token combaciano se uno e' prefisso dell'altro di almeno 3 lettere
        # (gestisce i troncamenti del PDF, es. FEDERIC ~ FEDERICO, GIOVA ~ GIOVANNI)
        if a == b:
            return True
        corto, lungo = (a, b) if len(a) <= len(b) else (b, a)
        return len(corto) >= 3 and lungo.startswith(corto)

    def _nomi_compatibili(tok_csv, tok_lul):
        # ogni token del nome piu' corto deve trovare un match (prefisso) in un
        # token ancora libero dell'altro nome. Robust a ordine invertito
        # (CSV "Nome Cognome" vs LUL "Cognome Nome") e a token troncati/mancanti.
        piccolo, grande = (tok_csv, tok_lul) if len(tok_csv) <= len(tok_lul) else (tok_lul, tok_csv)
        disponibili = list(grande)
        for t in piccolo:
            trovato = None
            for u in disponibili:
                if _match_token(t, u):
                    trovato = u
                    break
            if trovato is None:
                return False
            disponibili.remove(trovato)
        # almeno 2 token combaciati (cognome + nome) per evitare falsi positivi
        return len(piccolo) >= 2

    # Indici del LUL
    lul_per_nome = {}
    lul_tokens   = []   # (tokens, dipendente)
    for d in dipendenti_lul:
        lul_per_nome[_key(d["name"])] = d
        lul_tokens.append((_tokens(d["name"]), d))

    cf_a_lul     = {}
    non_abbinati = []
    abbinati_prefisso = []

    for _, riga in discenti_info.iterrows():
        cf       = riga["Codice Fiscale"]
        nome_csv = riga["nome"]
        kcsv     = _key(nome_csv)

        # 1) match esatto sui token ordinati
        if kcsv in lul_per_nome:
            cf_a_lul[cf] = lul_per_nome[kcsv]
            continue

        # 2) match token-per-token con prefisso (ordine invertito + troncamenti)
        tcsv = _tokens(nome_csv)
        candidati = []
        for tlul, d in lul_tokens:
            if _nomi_compatibili(tcsv, tlul):
                candidati.append(d)
        nomi_cand = {_key(c["name"]) for c in candidati}
        if len(nomi_cand) == 1:
            cf_a_lul[cf] = candidati[0]
            if _key(candidati[0]["name"]) != kcsv:
                abbinati_prefisso.append(f"{nome_csv} ~ {candidati[0]['name']}")
        else:
            non_abbinati.append(nome_csv)

    print(f"  Abbinati: {len(cf_a_lul)} / {len(discenti_info)}")
    if abbinati_prefisso:
        print(f"  Abbinati per nome troncato/parziale ({len(abbinati_prefisso)}):")
        for a in abbinati_prefisso:
            print(f"     - {a}")
    if non_abbinati:
        print(f"  Non abbinati (restano in Modalita' A): {', '.join(non_abbinati)}")
    return cf_a_lul


# =============================================================================
# CALCOLO ORE PER GIORNO
# =============================================================================

def calcola_giorno(cf, giorno, df_giornaliero, cf_a_lul, weekend_days, modalita_b):
    mask = (df_giornaliero["Codice Fiscale"] == cf) & (df_giornaliero["day"] == giorno)
    riga = df_giornaliero[mask]

    ore_prima = ore_dopo = ore_mattino = ore_tot = 0.0
    if not riga.empty:
        ore_prima   = riga["ore_prima"].iloc[0]
        ore_dopo    = riga["ore_dopo"].iloc[0]
        ore_mattino = riga["ore_mattino"].iloc[0]
        ore_tot     = riga["ore_tot"].iloc[0]

    # Applica tolleranza dopo le 18:00 e nel mattino (20 min/giorno per finestra non contano)
    tolleranza_ore  = TOLLERANZA_MINUTI / 60
    ore_dopo_eff    = 0.0 if ore_dopo    <= tolleranza_ore else ore_dopo
    ore_mattino_eff = 0.0 if ore_mattino <= tolleranza_ore else ore_mattino

    # Applica cap giornaliero: se le ore in orario normale superano 8h, l'eccesso non e' riconosciuto.
    exc_cap   = max(0.0, ore_prima - CAP_ORE_GIORNALIERO)
    ore_prima = min(ore_prima, CAP_ORE_GIORNALIERO)

    # ── MODALITA A: solo CSV ──────────────────────────────────────────────────
    if not modalita_b:
        if giorno in weekend_days:
            return {"ore_tot": ore_tot, "eff": 0.0,
                    "exc_we": ore_tot, "exc_dopo": 0.0, "exc_mattino": 0.0,
                    "exc_cap": 0.0, "exc_lav": 0.0,
                    "colore": COLORI["rosa"] if ore_tot > 0 else COLORI["grigio"]}
        else:
            if ore_dopo_eff > 0:
                colore = COLORI["viola"]
            elif ore_mattino_eff > 0:
                colore = COLORI["azzurro"]
            elif exc_cap > 0:
                colore = COLORI["arancio"]
            elif ore_tot > 0:
                colore = COLORI["verde"]
            else:
                colore = None
            return {"ore_tot": ore_tot, "eff": ore_prima,
                    "exc_we": 0.0, "exc_dopo": ore_dopo_eff,
                    "exc_mattino": ore_mattino_eff, "exc_cap": exc_cap, "exc_lav": 0.0,
                    "colore": colore}

    # ── MODALITA B: CSV + LUL ────────────────────────────────────────────────
    # Metodo (verificato a mano sui dati reali):
    #   1) Weekend          -> tutte le ore sono eccesso weekend, eff = 0
    #   2) Nessun abbinamento LUL (lul None) o giorno fuori dal LUL (lav_ore None)
    #                       -> come Modalita' A: nessun vincolo LAV
    #   3) Giorno NON lavorato (assente, oppure lav_ore == 0)
    #                       -> tutte le ore feriali sono eccesso LAV, eff = 0
    #   4) Giorno lavorato (lav_ore > 0)
    #                       -> ore valide = min(ore_prima, lav_ore);
    #                          l'eccedenza oltre le ore lavorate e' eccesso LAV.
    # In tutti i casi restano validi anche i tagli "serale" (ore_dopo_eff),
    # "mattutino" (ore_mattino_eff) e "cap 8h" (exc_cap) gia' calcolati sopra.
    lul       = cf_a_lul.get(cf)
    e_assente = (giorno in lul["assente"]) if lul else False
    lav_ore   = lul["lav"].get(giorno, None) if lul else None

    # 1) Weekend
    if giorno in weekend_days:
        return {"ore_tot": ore_tot, "eff": 0.0,
                "exc_we": ore_tot, "exc_dopo": 0.0, "exc_mattino": 0.0,
                "exc_cap": 0.0, "exc_lav": 0.0,
                "colore": COLORI["rosa"] if ore_tot > 0 else COLORI["grigio"]}

    # 2) Nessun vincolo LAV applicabile -> come Modalita' A
    if lul is None or (lav_ore is None and not e_assente):
        if ore_dopo_eff > 0:      colore = COLORI["viola"]
        elif ore_mattino_eff > 0: colore = COLORI["azzurro"]
        elif exc_cap > 0:         colore = COLORI["arancio"]
        elif ore_tot > 0:         colore = COLORI["verde"]
        else:                     colore = None
        return {"ore_tot": ore_tot, "eff": ore_prima,
                "exc_we": 0.0, "exc_dopo": ore_dopo_eff,
                "exc_mattino": ore_mattino_eff, "exc_cap": exc_cap, "exc_lav": 0.0,
                "colore": colore}

    # 3) Giorno NON lavorato (assenza dichiarata oppure ore lavorate = 0)
    if e_assente or lav_ore == 0.0:
        return {"ore_tot": ore_tot, "eff": 0.0, "exc_we": 0.0,
                "exc_dopo": ore_dopo_eff, "exc_mattino": ore_mattino_eff,
                "exc_cap": exc_cap, "exc_lav": ore_prima,
                "colore": COLORI["giallo"] if ore_tot > 0 else None}

    # 4) Giorno lavorato: cap alle ore effettivamente lavorate (LAV)
    eff     = min(ore_prima, lav_ore)
    exc_lav = max(0.0, ore_prima - lav_ore)

    if ore_tot == 0:           colore = None
    elif exc_lav > 0:          colore = COLORI["arancio"]
    elif exc_cap > 0:          colore = COLORI["arancio"]
    elif ore_dopo_eff > 0:     colore = COLORI["viola"]
    elif ore_mattino_eff > 0:  colore = COLORI["azzurro"]
    else:                      colore = COLORI["verde"]

    return {"ore_tot": ore_tot, "eff": eff, "exc_we": 0.0,
            "exc_dopo": ore_dopo_eff, "exc_mattino": ore_mattino_eff,
            "exc_cap": exc_cap, "exc_lav": exc_lav, "colore": colore}


# =============================================================================
# SCRITTURA FOGLIO EXCEL (un foglio per mese, sullo stesso Workbook)
# =============================================================================

def scrivi_foglio(wb, discenti_info, df_giornaliero, azienda,
                  mese, anno, cf_a_lul=None, discenti_globali=None):
    """
    Aggiunge un foglio al Workbook wb con i dati del mese indicato.
    """
    modalita_b    = cf_a_lul is not None
    totale_giorni = giorni_nel_mese(anno, mese)
    weekend_days  = giorni_weekend(anno, mese)
    modo_label    = "MODALITA B (CSV+LUL)" if modalita_b else "MODALITA A (solo CSV)"

    titolo = f"MONITORAGGIO ORE FORMAZIONE - {NOMI_MESI[mese]} {anno} - {azienda}"

    # Layout colonne
    COL_A       = 1
    COL_B       = 2
    COL_C       = 3
    COL_DAY_INI = 4
    COL_DAY_FIN = COL_DAY_INI + totale_giorni - 1

    # Colonne riepilogative — ordine fisso.
    # In MODALITA B viene aggiunta una colonna "Eccesso Ore LAV" tra
    # "Totale Eccesso" e "Eccesso Weekend"; in MODALITA A questa colonna
    # non e' presente perche' non c'e' la sorgente LUL.
    COL_TOT_EFF      = COL_DAY_FIN + 1   # Totale Ore Effettive (=SUM celle giorno)
    COL_TOT_ORE      = COL_DAY_FIN + 2   # Ore Totali           (valore Python)
    COL_TOT_EXC      = COL_DAY_FIN + 3   # Totale Eccesso       (= ore totali - effettive)
    if modalita_b:
        COL_EXC_LAV      = COL_DAY_FIN + 4   # Eccesso Ore LAV (solo modalita B)
        COL_EXC_WE       = COL_DAY_FIN + 5
        COL_EXC_DOPO     = COL_DAY_FIN + 6
        COL_EXC_MATTINO  = COL_DAY_FIN + 7
    else:
        COL_EXC_WE       = COL_DAY_FIN + 4
        COL_EXC_DOPO     = COL_DAY_FIN + 5
        COL_EXC_MATTINO  = COL_DAY_FIN + 6
    ULTIMA_COL       = COL_EXC_MATTINO

    RIGA_TITOLO = 1
    RIGA_LOGO   = 2
    RIGA_HEADER = 3
    RIGA_DATI   = 4

    if discenti_globali:
        lista_discenti = sorted(
            discenti_globali.items(),
            key=lambda x: chiave_ordinamento(x[1]["nome"])
        )
    else:
        lista_discenti = [
            (row["Codice Fiscale"], {"nome": row["nome"], "percorso": row["percorso"]})
            for _, row in discenti_info.sort_values(
                "nome", key=lambda s: s.apply(chiave_ordinamento)).iterrows()
        ]

    N           = len(lista_discenti)
    RIGA_TOTALE = RIGA_DATI + N

    # Stili (palette pulita, testo grigio scuro, sfondo chiaro)
    font_base    = Font(name="Arial", size=9, color="333333")
    font_header  = Font(name="Arial", size=9, bold=True, color="495057")
    font_titolo  = Font(name="Arial", size=12, bold=True, color=COLORI["titolo_testo"])
    font_totale  = Font(name="Arial", size=9, bold=True, color="333333")
    allin_centro = Alignment(horizontal="center", vertical="center")
    allin_sin    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    allin_wrap   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bordo_top    = Border(top=Side(style="thin", color=COLORI["bordo"]))

    ws = wb.create_sheet(title=f"{NOMI_MESI[mese]} {anno}")

    # ── Titolo ────────────────────────────────────────────────────────────────
    ws.row_dimensions[RIGA_TITOLO].height = 26
    c = ws.cell(RIGA_TITOLO, COL_A, value=titolo)
    c.font = font_titolo; c.fill = fill_cell(COLORI["titolo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=RIGA_TITOLO, start_column=COL_A,
                   end_row=RIGA_TITOLO,   end_column=ULTIMA_COL)

    # ── Riga loghi ────────────────────────────────────────────────────────────
    from openpyxl.drawing.image import Image as XLImage
    ws.row_dimensions[RIGA_LOGO].height = 38
    for path, anchor_col in [(LOGO_AZIENDA, COL_A), (LOGO_FNC, ULTIMA_COL - 1)]:
        if not os.path.isfile(path):
            continue
        try:
            img = XLImage(path)
            h_target = 50
            scale    = h_target / img.height
            img.width  = int(img.width  * scale)
            img.height = int(img.height * scale)
            img.anchor = f"{get_column_letter(anchor_col)}{RIGA_LOGO}"
            ws.add_image(img)
        except Exception:
            pass

    # ── Intestazioni ──────────────────────────────────────────────────────────
    ws.row_dimensions[RIGA_HEADER].height = 32

    def scrivi_header(col, testo):
        c = ws.cell(RIGA_HEADER, col, value=testo)
        c.font = font_header
        c.fill = fill_cell(COLORI["header"])
        c.alignment = allin_wrap

    scrivi_header(COL_A, "Utente")
    scrivi_header(COL_B, "Codice Fiscale")
    scrivi_header(COL_C, "Percorso Formativo")

    for d in range(1, totale_giorni + 1):
        col  = COL_DAY_INI + d - 1
        abbr = abbreviazione_giorno(anno, mese, d)
        c = ws.cell(RIGA_HEADER, col, value=f"{d}\n{abbr}")
        c.font = Font(name="Arial", size=8, bold=True, color="495057")
        c.alignment = allin_wrap
        c.fill = fill_cell(COLORI["grigio"] if abbr in ("S", "D") else COLORI["header"])

    # Colonne riepilogative — 6 colonne fisse in ordine specifico
    scrivi_header(COL_TOT_EFF,       "Totale\nOre Effettive")
    ws.cell(RIGA_HEADER, COL_TOT_EFF).comment = Comment(
        "Somma delle ore effettive del mese\n"
        "(= ore totali fruite - ore in eccesso).\n"
        "Coincide con la somma delle celle giornaliere.", "Monitoraggio")

    scrivi_header(COL_TOT_ORE,       "Ore\nTotali")
    ws.cell(RIGA_HEADER, COL_TOT_ORE).comment = Comment(
        "Somma delle ore totali fruite nel mese (effettive + eccesso).",
        "Monitoraggio")

    scrivi_header(COL_TOT_EXC,       "Totale\nEccesso")
    if modalita_b:
        ws.cell(RIGA_HEADER, COL_TOT_EXC).comment = Comment(
            "Ore totali in eccesso del mese\n"
            "(= Ore Totali - Totale Ore Effettive).\n"
            "Comprende eccesso ore LAV, weekend, dopo soglia serale,\n"
            "mattutino e oltre cap 8h/giorno.", "Monitoraggio")
    else:
        ws.cell(RIGA_HEADER, COL_TOT_EXC).comment = Comment(
            "Ore totali in eccesso del mese\n"
            "(= Ore Totali - Totale Ore Effettive).\n"
            "Comprende eccesso weekend, dopo soglia serale, mattutino\n"
            "e oltre cap 8h/giorno.", "Monitoraggio")

    if modalita_b:
        scrivi_header(COL_EXC_LAV,   "Eccesso\nOre LAV")
        ws.cell(RIGA_HEADER, COL_EXC_LAV).comment = Comment(
            "Ore di studio eccedenti le ore LAV (Libro Unico del Lavoro)\n"
            "dichiarate per quel giorno, oppure ore di studio in giorni\n"
            "di assenza dichiarata.", "Monitoraggio")

    scrivi_header(COL_EXC_WE,        "Eccesso\nWeekend")
    ws.cell(RIGA_HEADER, COL_EXC_WE).comment = Comment(
        "Ore di studio nei giorni di sabato e domenica.", "Monitoraggio")

    scrivi_header(COL_EXC_DOPO,      "Eccesso\nDopo 18:00")
    ws.cell(RIGA_HEADER, COL_EXC_DOPO).comment = Comment(
        "Ore di studio dopo le 18:00.\n"
        "Tolleranza 20 min/giorno (sotto questa soglia non contano).",
        "Monitoraggio")

    scrivi_header(COL_EXC_MATTINO,   "Eccesso\nMattutino")
    ws.cell(RIGA_HEADER, COL_EXC_MATTINO).comment = Comment(
        "Ore di studio comprese fra le 06:00 e le 07:40.\n"
        "Tolleranza 20 min/giorno (sotto questa soglia non contano).",
        "Monitoraggio")

    # ── Larghezze colonne ─────────────────────────────────────────────────────
    ws.column_dimensions[get_column_letter(COL_A)].width = 25
    ws.column_dimensions[get_column_letter(COL_B)].width = 17
    ws.column_dimensions[get_column_letter(COL_C)].width = 35
    for d in range(1, totale_giorni + 1):
        ws.column_dimensions[get_column_letter(COL_DAY_INI + d - 1)].width = 6.5
    col_riepilogo = [COL_TOT_EFF, COL_TOT_ORE, COL_TOT_EXC]
    if modalita_b:
        col_riepilogo.append(COL_EXC_LAV)
    col_riepilogo.extend([COL_EXC_WE, COL_EXC_DOPO, COL_EXC_MATTINO])
    for col in col_riepilogo:
        ws.column_dimensions[get_column_letter(col)].width = 13

    cf_a_riga = {}
    cf_fruitori = set(discenti_info["Codice Fiscale"])
    col_day_ini_letter = get_column_letter(COL_DAY_INI)
    col_day_fin_letter = get_column_letter(COL_DAY_FIN)
    col_tot_eff_letter = get_column_letter(COL_TOT_EFF)
    col_tot_ore_letter = get_column_letter(COL_TOT_ORE)

    # Mappa CF -> ore effettive del mese (per il foglio Riepilogo Generale,
    # che ora le scrive come VALORI invece che come formule cross-sheet).
    ore_eff_per_cf = {}

    for idx, (cf, info_disc) in enumerate(lista_discenti):
        r = RIGA_DATI + idx
        cf_a_riga[cf] = r

        # Il nome nei file e' gia' nel formato "COGNOME NOME": lo si usa cosi'
        # com'e' (solo in maiuscolo), senza invertire i token.
        cognome_nome = info_disc["nome"].strip().upper()

        ws.row_dimensions[r].height = 17
        c = ws.cell(r, COL_A, value=cognome_nome); c.font = font_base; c.alignment = allin_sin
        c = ws.cell(r, COL_B, value=cf);           c.font = font_base; c.alignment = allin_centro
        c = ws.cell(r, COL_C, value=info_disc["percorso"]); c.font = font_base; c.alignment = allin_sin

        # Accumulatori mensili:
        # - acc_tot:     somma ore totali del mese (per "Ore Totali")
        # - acc_eff:     somma ore effettive del mese (per Riepilogo)
        # - acc_we:      somma eccesso weekend
        # - acc_dopo:    somma eccesso dopo soglia serale
        # - acc_mattino: somma eccesso mattutino
        # - acc_lav:     somma eccesso ore LAV (solo modalita B, sempre 0 in A)
        # exc_cap viene comunque conteggiato nel Totale Eccesso via la formula
        # (Ore Totali - Totale Ore Effettive), quindi non serve accumularlo.
        acc_tot = acc_eff = acc_we = acc_dopo = acc_mattino = acc_lav = 0.0

        for d in range(1, totale_giorni + 1):
            col = COL_DAY_INI + d - 1
            res = calcola_giorno(cf, d, df_giornaliero, cf_a_lul or {}, weekend_days, modalita_b)

            ore_tot_daily   = res["ore_tot"]
            excess_daily    = (res["exc_we"]   + res["exc_dopo"]    + res["exc_mattino"]
                              + res.get("exc_cap", 0) + res.get("exc_lav", 0))
            ore_eff_daily   = ore_tot_daily - excess_daily

            # Valore cella: ore EFFETTIVE (= totali - eccesso).
            # Se ore_tot==0 -> cella vuota (giornata senza fruizione).
            # Se ore_tot>0 ma ore_eff==0 (es. weekend interamente in eccesso)
            # mostra esplicitamente 0:00:00 col colore di sfondo dell'eccesso.
            if ore_tot_daily > 0:
                c = ws.cell(r, col, value=ore_eff_daily / 24)
                c.number_format = DURATION_FORMAT
            else:
                c = ws.cell(r, col)
            c.font = Font(name="Arial", size=8, color="333333"); c.alignment = allin_centro
            if res["colore"]:
                c.fill = fill_cell(res["colore"])

            # Nota: SOLO se c'e' eccesso effettivamente contato (ore totali != ore effettive)
            # Mostra "Ore totali", "Eccesso" e il dettaglio per tipo di eccesso.
            if excess_daily > 1e-6:
                note_lines = [
                    f"Ore totali: {ore_decimali_a_hhmmss(ore_tot_daily)}",
                    f"Eccesso: {ore_decimali_a_hhmmss(excess_daily)}",
                ]
                # Dettaglio per tipo di eccesso (solo voci con valore > 0)
                if res.get("exc_we", 0) > 1e-6:
                    note_lines.append(
                        f"- Weekend: {ore_decimali_a_hhmmss(res['exc_we'])}")
                if res.get("exc_dopo", 0) > 1e-6:
                    note_lines.append(
                        f"- Dopo le {int(SOGLIA_ORA)}:00: "
                        f"{ore_decimali_a_hhmmss(res['exc_dopo'])}")
                if res.get("exc_mattino", 0) > 1e-6:
                    note_lines.append(
                        f"- Mattutino (06:00-07:40): "
                        f"{ore_decimali_a_hhmmss(res['exc_mattino'])}")
                if res.get("exc_cap", 0) > 1e-6:
                    note_lines.append(
                        f"- Oltre cap {int(CAP_ORE_GIORNALIERO)}h/giorno: "
                        f"{ore_decimali_a_hhmmss(res['exc_cap'])}")
                if res.get("exc_lav", 0) > 1e-6:
                    note_lines.append(
                        f"- Eccesso ore LAV: "
                        f"{ore_decimali_a_hhmmss(res['exc_lav'])}")
                c.comment = Comment("\n".join(note_lines), "Monitoraggio")

            acc_tot     += ore_tot_daily
            acc_eff     += ore_eff_daily
            acc_we      += res["exc_we"]
            acc_dopo    += res["exc_dopo"]
            acc_mattino += res.get("exc_mattino", 0)
            acc_lav     += res.get("exc_lav", 0)

        # Memorizza ore effettive per il Riepilogo (anche per i non-fruitori: 0)
        ore_eff_per_cf[cf] = acc_eff

        def scrivi_durata(col, val):
            if val and val > 0:
                c = ws.cell(r, col, value=val / 24)
                c.number_format = DURATION_FORMAT
            else:
                c = ws.cell(r, col)
            c.font = font_base; c.alignment = allin_centro

        def scrivi_formula(col, formula):
            c = ws.cell(r, col, value=formula)
            c.number_format = DURATION_FORMAT
            c.font = font_base; c.alignment = allin_centro

        if cf not in cf_fruitori:
            continue

        # ── Colonne riepilogative (ordine fisso) ─────────────────────────────
        # 1) Totale Ore Effettive = somma delle celle giornaliere (=SUM)
        scrivi_formula(COL_TOT_EFF,
                       f"=SUM({col_day_ini_letter}{r}:{col_day_fin_letter}{r})")
        # 2) Ore Totali = valore Python (somma di ore_tot giornaliere)
        scrivi_durata(COL_TOT_ORE, acc_tot)
        # 3) Totale Eccesso = Ore Totali - Totale Ore Effettive (formula trasparente)
        scrivi_formula(COL_TOT_EXC,
                       f"={col_tot_ore_letter}{r}-{col_tot_eff_letter}{r}")
        # 3-bis) Eccesso Ore LAV (solo modalita B)
        if modalita_b:
            scrivi_durata(COL_EXC_LAV, acc_lav)
        # 4) Eccesso Weekend
        scrivi_durata(COL_EXC_WE, acc_we)
        # 5) Eccesso Dopo soglia serale
        scrivi_durata(COL_EXC_DOPO, acc_dopo)
        # 6) Eccesso Mattutino
        scrivi_durata(COL_EXC_MATTINO, acc_mattino)

    # ── Riga TOTALE ───────────────────────────────────────────────────────────
    # Tutti i totali sono ottenuti tramite formule =SUM() sulle righe utenti,
    # in modo che l'azienda possa verificare i conteggi cliccando sulle celle.
    ws.row_dimensions[RIGA_TOTALE].height = 18
    c = ws.cell(RIGA_TOTALE, COL_A, value="TOTALE")
    c.font = font_totale
    c.fill = fill_cell(COLORI["totale"]); c.alignment = allin_sin
    c.border = bordo_top
    for col in [COL_B, COL_C]:
        cc = ws.cell(RIGA_TOTALE, col)
        cc.fill = fill_cell(COLORI["totale"])
        cc.border = bordo_top

    riga_primo_utente = RIGA_DATI
    riga_ultimo_utente = RIGA_TOTALE - 1

    # Le colonne riepilogative (in MODALITA B con anche EXC_LAV)
    col_riepilogo_tot = [COL_TOT_EFF, COL_TOT_ORE, COL_TOT_EXC]
    if modalita_b:
        col_riepilogo_tot.append(COL_EXC_LAV)
    col_riepilogo_tot.extend([COL_EXC_WE, COL_EXC_DOPO, COL_EXC_MATTINO])
    for col in col_riepilogo_tot:
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}{riga_primo_utente}:{col_letter}{riga_ultimo_utente})"
        c = ws.cell(RIGA_TOTALE, col, value=formula)
        c.font = font_totale
        c.alignment = allin_centro; c.fill = fill_cell(COLORI["totale"])
        c.number_format = DURATION_FORMAT
        c.border = bordo_top

    ws.freeze_panes = ws.cell(RIGA_DATI, COL_DAY_INI)

    # Statistiche di log (solo per output a console, non scritte sul foglio)
    tot_fruite_log = df_giornaliero["ore_tot"].sum()
    tolleranza_ore = TOLLERANZA_MINUTI / 60
    tot_dopo_log = 0.0
    for cf_log in discenti_info["Codice Fiscale"]:
        for d in range(1, totale_giorni + 1):
            mask = (df_giornaliero["Codice Fiscale"] == cf_log) & (df_giornaliero["day"] == d)
            riga = df_giornaliero[mask]
            if not riga.empty and d not in weekend_days:
                od = riga["ore_dopo"].iloc[0]
                if od > tolleranza_ore:
                    tot_dopo_log += od
    tot_we_log = sum(
        df_giornaliero[df_giornaliero["day"] == d]["ore_tot"].sum()
        for d in giorni_weekend(anno, mese)
    )

    print(f"  -> Foglio '{NOMI_MESI[mese]} {anno}' scritto | "
          f"Discenti: {N} | Ore fruite: {ore_decimali_a_hhmmss(tot_fruite_log)} | "
          f"Eccesso 18:00: {ore_decimali_a_hhmmss(tot_dopo_log)} | "
          f"Eccesso WE: {ore_decimali_a_hhmmss(tot_we_log)} | {modo_label}")

    return {
        "sheet_name":       ws.title,
        "anno":             anno,
        "mese":             mese,
        "cf_a_riga":        cf_a_riga,
        "col_tot_ore":      COL_TOT_ORE,
        "col_exc_dopo":     COL_EXC_DOPO,
        "col_exc_we":       COL_EXC_WE,
        "col_tot_exc":      COL_TOT_EXC,
        "col_ore_maturate": COL_TOT_EFF,
        "n_discenti":       N,
        # Dizionario CF -> ore effettive del mese (in ore decimali).
        # Usato dal foglio Riepilogo per scrivere VALORI (non formule).
        "ore_eff_per_cf":   ore_eff_per_cf,
        "discenti":         discenti_info[["Codice Fiscale", "nome", "percorso"]].to_dict("records"),
    }


# =============================================================================
# FOGLIO RIEPILOGO GENERALE (totali per discente, su tutti i mesi elaborati)
# =============================================================================

def ore_decimali_a_testo(ore_dec):
    """Formatta ore decimali come stringa 'X H YY M ZZ S' (es. '150 H 00 M 00 S').

    Le ore non hanno padding (cosi' '150 H ...' resta naturale), mentre
    minuti e secondi sono sempre a 2 cifre. Valori <=0 -> '0 H 00 M 00 S'.
    """
    if ore_dec is None or ore_dec <= 0:
        return "0 H 00 M 00 S"
    total_sec = int(round(ore_dec * 3600))
    h = total_sec // 3600
    m = (total_sec % 3600) // 60
    s = total_sec % 60
    return f"{h} H {m:02d} M {s:02d} S"


def scrivi_foglio_riepilogo(wb, fogli_info, azienda, discenti_globali=None):
    """
    Crea un foglio "Riepilogo Generale" con un discente per riga, una
    colonna per ogni mese elaborato e tre colonne finali:
      - Totale Ore Maturate    (valore, somma dei mesi)
      - Min(150 ore)           (valore, capped a 150 ore)
      - Ore (formato testuale) (testo "X H YY M ZZ S" del valore capped)

    Tutte le celle del Riepilogo usano formule Excel per massima trasparenza:
      - Colonne mensili     : ='NomeFoglio'!{COL_TOT_EFF}{riga}
      - Totale Ore Maturate : =SUM(colonne mensili)
      - Min(150 ore)        : =MIN(Totale, 150/24)
      - Ore (testo)         : formula TEXT()+MOD() -> "X H YY M ZZ S"

    fogli_info: lista di dict restituiti da scrivi_foglio() (uno per mese).
    discenti_globali: dict {cf: {"nome":..., "percorso":...}} con TUTTI i
        discenti del CSV. Se fornito, ogni discente compare in tabella anche
        se non ha maturato ore in nessun mese (riga con tutti 0:00:00).
    """
    if not fogli_info:
        return

    # Ordina i fogli per anno/mese
    fogli_info = sorted(fogli_info, key=lambda x: (x["anno"], x["mese"]))

    # Elenco discenti: usa quello globale se passato, altrimenti unione mensile
    if discenti_globali is None:
        discenti_globali = {}
        for f in fogli_info:
            for d in f["discenti"]:
                cf = d["Codice Fiscale"]
                if cf not in discenti_globali:
                    discenti_globali[cf] = {"nome": d["nome"], "percorso": d["percorso"]}

    discenti_lista = sorted(
        discenti_globali.items(),
        key=lambda x: chiave_ordinamento(x[1]["nome"])
    )

    n_mesi = len(fogli_info)
    N      = len(discenti_lista)

    # Layout colonne (aggiunta colonna "Ore Tolte" dopo Min(150))
    COL_NOME       = 1
    COL_CF         = 2
    COL_PERCORSO   = 3
    COL_MESE_INI   = 4
    COL_MESE_FIN   = COL_MESE_INI + n_mesi - 1
    COL_TOTALE     = COL_MESE_FIN + 1   # Totale Ore Maturate (valore)
    COL_MIN_150    = COL_MESE_FIN + 2   # Cap a 150 ore (valore)
    COL_ORE_TOLTE  = COL_MESE_FIN + 3   # Ore Totali - Min(150): ore non riconoscibili
    COL_TESTO      = COL_MESE_FIN + 4   # Stringa "X H YY M ZZ S" (valore testo)
    ULTIMA_COL     = COL_TESTO

    RIGA_TITOLO = 1
    RIGA_LOGO   = 2
    RIGA_HEADER = 3
    RIGA_DATI   = 4
    RIGA_TOTALE = RIGA_DATI + N

    # Stili
    font_base    = Font(name="Arial", size=10, color="333333")
    font_header  = Font(name="Arial", size=9, bold=True, color="495057")
    font_titolo  = Font(name="Arial", size=13, bold=True, color=COLORI["titolo_testo"])
    font_dato_b  = Font(name="Arial", size=10, bold=True, color="333333")
    allin_centro = Alignment(horizontal="center", vertical="center")
    allin_sin    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    allin_wrap   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bordo_top    = Border(top=Side(style="thin", color=COLORI["bordo"]))

    # Inserisci il foglio in PRIMA posizione
    ws = wb.create_sheet(title="Riepilogo Generale", index=0)

    # ── Titolo ────────────────────────────────────────────────────────────────
    ws.row_dimensions[RIGA_TITOLO].height = 28
    c = ws.cell(RIGA_TITOLO, COL_NOME, value=f"Riepilogo Generale - {azienda}")
    c.font = font_titolo; c.fill = fill_cell(COLORI["titolo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=RIGA_TITOLO, start_column=COL_NOME,
                   end_row=RIGA_TITOLO,   end_column=ULTIMA_COL)

    # ── Riga loghi ────────────────────────────────────────────────────────────
    from openpyxl.drawing.image import Image as XLImage
    ws.row_dimensions[RIGA_LOGO].height = 38
    for path, anchor_col in [(LOGO_AZIENDA, COL_NOME), (LOGO_FNC, ULTIMA_COL - 1)]:
        if not os.path.isfile(path):
            continue
        try:
            img = XLImage(path)
            h_target = 50
            scale    = h_target / img.height
            img.width  = int(img.width  * scale)
            img.height = int(img.height * scale)
            img.anchor = f"{get_column_letter(anchor_col)}{RIGA_LOGO}"
            ws.add_image(img)
        except Exception:
            pass

    # ── Intestazioni ──────────────────────────────────────────────────────────
    ws.row_dimensions[RIGA_HEADER].height = 38

    def scrivi_header(col, testo):
        c = ws.cell(RIGA_HEADER, col, value=testo)
        c.font = font_header; c.fill = fill_cell(COLORI["header"])
        c.alignment = allin_wrap

    scrivi_header(COL_NOME,     "Discente")
    scrivi_header(COL_CF,       "Codice Fiscale")
    scrivi_header(COL_PERCORSO, "Percorso Formativo")

    for j, f in enumerate(fogli_info):
        col = COL_MESE_INI + j
        scrivi_header(col, f"{NOMI_MESI[f['mese']]}\n{f['anno']}")

    scrivi_header(COL_TOTALE,  "Totale\nOre Maturate")
    scrivi_header(COL_MIN_150, f"Min({ORE_LIMITE_FINANZIATE} ore)")
    scrivi_header(COL_ORE_TOLTE, "Ore\nTolte")
    scrivi_header(COL_TESTO,   "Ore\n(testo)")

    # Tooltip esplicativi
    ws.cell(RIGA_HEADER, COL_TOTALE).comment = Comment(
        "Somma delle ore maturate (al netto degli eccessi) "
        "su tutti i mesi elaborati. Formula: =SUM(colonne mensili).",
        "Monitoraggio")
    ws.cell(RIGA_HEADER, COL_MIN_150).comment = Comment(
        f"Totale Ore Maturate cappato a {ORE_LIMITE_FINANZIATE} ore "
        f"(limite riconosciuto). Formula: =MIN(Totale, {ORE_LIMITE_FINANZIATE}/24).",
        "Monitoraggio")
    ws.cell(RIGA_HEADER, COL_ORE_TOLTE).comment = Comment(
        "Totale ore in eccesso scartate (non riconoscibili):\n"
        "= Ore Totali grezze fruite - Ore Effettive maturate.\n"
        "Include: eccesso weekend, dopo le 18:00, mattutino (06:00-07:40),\n"
        "oltre cap 8h/giorno e (in Modalità B) eccesso ore LAV.\n"
        "Somma dei 'Totale Eccesso' di ogni foglio mensile.",
        "Monitoraggio")
    ws.cell(RIGA_HEADER, COL_TESTO).comment = Comment(
        f"Stesso valore di 'Min({ORE_LIMITE_FINANZIATE} ore)' "
        f"in formato testuale 'X H YY M ZZ S' (es. '150 H 00 M 00 S'). "
        f"Formula Excel con TEXT() e MOD().",
        "Monitoraggio")

    # ── Larghezze colonne ─────────────────────────────────────────────────────
    ws.column_dimensions[get_column_letter(COL_NOME)].width     = 28
    ws.column_dimensions[get_column_letter(COL_CF)].width       = 17
    ws.column_dimensions[get_column_letter(COL_PERCORSO)].width = 35
    for col in range(COL_MESE_INI, COL_MESE_FIN + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions[get_column_letter(COL_TOTALE)].width  = 14
    ws.column_dimensions[get_column_letter(COL_MIN_150)].width = 14
    ws.column_dimensions[get_column_letter(COL_ORE_TOLTE)].width = 12
    ws.column_dimensions[get_column_letter(COL_TESTO)].width   = 22

    # ── Lettere colonne (usate nelle formule) ────────────────────────────────
    col_mese_ini_letter  = get_column_letter(COL_MESE_INI)
    col_mese_fin_letter  = get_column_letter(COL_MESE_FIN)
    col_totale_letter    = get_column_letter(COL_TOTALE)
    col_min150_letter    = get_column_letter(COL_MIN_150)
    col_ore_tolte_letter = get_column_letter(COL_ORE_TOLTE)

    # ── Righe dati ────────────────────────────────────────────────────────────
    for i, (cf, info) in enumerate(discenti_lista):
        r = RIGA_DATI + i
        # Il nome e' gia' "COGNOME NOME": usato cosi' com'e' (maiuscolo).
        cognome_nome = info["nome"].strip().upper()

        ws.row_dimensions[r].height = 19
        c = ws.cell(r, COL_NOME,     value=cognome_nome);     c.font = font_base; c.alignment = allin_sin
        c = ws.cell(r, COL_CF,       value=cf);               c.font = font_base; c.alignment = allin_centro
        c = ws.cell(r, COL_PERCORSO, value=info["percorso"]); c.font = font_base; c.alignment = allin_sin

        # Colonne mensili: MATCH+INDEX cross-sheet cercando il CF nella col B del foglio
        # Usa MATCH(CF, col_B_foglio, 0) per trovare la riga corretta indipendentemente
        # dall'ordinamento — evita il bug di inversione quando i nomi hanno accenti o
        # caratteri speciali che alterano il sort key.
        for j, f in enumerate(fogli_info):
            col    = COL_MESE_INI + j
            sname  = f["sheet_name"].replace("'", "''")
            col_cf_ms  = get_column_letter(2)                    # col B = Codice Fiscale
            col_eff_ms = get_column_letter(f["col_ore_maturate"])
            n_righe    = f.get("n_discenti", 200)                # range di ricerca abbondante
            # IFERROR(...,0) per i discenti assenti in quel mese
            formula_eff = (
                f"=IFERROR(INDEX('{sname}'!{col_eff_ms}{RIGA_DATI}:{col_eff_ms}{RIGA_DATI+n_righe},"
                f"MATCH({chr(34)}{cf}{chr(34)},'{sname}'!{col_cf_ms}{RIGA_DATI}:{col_cf_ms}{RIGA_DATI+n_righe},0)),0)"
            )
            c = ws.cell(r, col, value=formula_eff)
            c.number_format = DURATION_FORMAT
            c.font = font_base; c.alignment = allin_centro

        # COL_TOTALE: formula =SUM(colonne mensili)
        c = ws.cell(r, COL_TOTALE,
                    value=f"=SUM({col_mese_ini_letter}{r}:{col_mese_fin_letter}{r})")
        c.number_format = DURATION_FORMAT
        c.font = font_dato_b; c.alignment = allin_centro
        c.fill = fill_cell(COLORI["totale"])

        # COL_MIN_150: formula =MIN(Totale, limite/24)
        c = ws.cell(r, COL_MIN_150,
                    value=f"=MIN({col_totale_letter}{r},{ORE_LIMITE_FINANZIATE}/24)")
        c.number_format = DURATION_FORMAT
        c.font = font_dato_b; c.alignment = allin_centro
        c.fill = fill_cell(COLORI["totale"])

        # COL_ORE_TOLTE: somma MATCH+INDEX su COL_TOT_EXC di ogni foglio mensile
        parti_exc = []
        for f in fogli_info:
            sname      = f["sheet_name"].replace("'", "''")
            col_cf_ms  = get_column_letter(2)
            col_exc_ms = get_column_letter(f["col_tot_exc"])
            n_righe    = f.get("n_discenti", 200)
            parti_exc.append(
                f"IFERROR(INDEX('{sname}'!{col_exc_ms}{RIGA_DATI}:{col_exc_ms}{RIGA_DATI+n_righe},"
                f"MATCH({chr(34)}{cf}{chr(34)},'{sname}'!{col_cf_ms}{RIGA_DATI}:{col_cf_ms}{RIGA_DATI+n_righe},0)),0)"
            )
        formula_exc = "=" + "+".join(parti_exc) if parti_exc else 0
        c = ws.cell(r, COL_ORE_TOLTE, value=formula_exc)
        c.number_format = DURATION_FORMAT
        c.font = font_dato_b; c.alignment = allin_centro
        c.fill = fill_cell("FEF2F2")   # rosso tenuissimo

        # COL_TESTO: formula Excel "X H YY M ZZ S" dalla cella Min(150)
        testo_formula = (
            f'=TEXT(INT({col_min150_letter}{r}*24),"0")'
            f'&" H "&TEXT(INT(MOD({col_min150_letter}{r}*24,1)*60),"00")'
            f'&" M "&TEXT(INT(MOD({col_min150_letter}{r}*24*60,1)*60),"00")'
            f'&" S"'
        )
        c = ws.cell(r, COL_TESTO, value=testo_formula)
        c.font = font_dato_b; c.alignment = allin_centro
        c.fill = fill_cell(COLORI["totale"])

    ws.freeze_panes = ws.cell(RIGA_DATI, COL_MESE_INI)

    # ── Riga TOTALE ───────────────────────────────────────────────────────────
    bordo_top   = Border(top=Side(style="thin", color=COLORI["bordo"]))
    font_totale = Font(name="Arial", size=10, bold=True, color="333333")
    ws.row_dimensions[RIGA_TOTALE].height = 20
    c = ws.cell(RIGA_TOTALE, COL_NOME, value="TOTALE")
    c.font = font_totale; c.fill = fill_cell(COLORI["totale"])
    c.alignment = Alignment(horizontal="left", vertical="center"); c.border = bordo_top
    for col in [COL_CF, COL_PERCORSO]:
        cc = ws.cell(RIGA_TOTALE, col)
        cc.fill = fill_cell(COLORI["totale"]); cc.border = bordo_top

    riga_primo = RIGA_DATI
    riga_ultimo = RIGA_TOTALE - 1
    for col in list(range(COL_MESE_INI, COL_MESE_FIN + 1)) + [COL_TOTALE, COL_MIN_150, COL_ORE_TOLTE]:
        cl = get_column_letter(col)
        c = ws.cell(RIGA_TOTALE, col, value=f"=SUM({cl}{riga_primo}:{cl}{riga_ultimo})")
        c.number_format = DURATION_FORMAT
        c.font = font_totale; c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = fill_cell(COLORI["totale"]); c.border = bordo_top
    # Colonna testo: lasciata vuota nel totale
    cc = ws.cell(RIGA_TOTALE, COL_TESTO)
    cc.fill = fill_cell(COLORI["totale"]); cc.border = bordo_top

    print(f"\n>> Foglio 'Riepilogo Generale' creato | "
          f"Discenti: {N} | Mesi: {n_mesi}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 65)
    print("  MONITORAGGIO ORE FORMAZIONE E-LEARNING")
    print("=" * 65)

    csv_presente = os.path.isfile(CSV_FILE)
    # Rileva il LUL: preferenza al PDF (piu' affidabile sui nomi), poi XLSX
    global LUL_FILE
    if os.path.isfile(LUL_FILE_PDF):
        LUL_FILE = LUL_FILE_PDF
    elif os.path.isfile(LUL_FILE_XLSX):
        LUL_FILE = LUL_FILE_XLSX
    else:
        LUL_FILE = LUL_FILE_XLSX  # nome di default per i messaggi
    lul_presente = os.path.isfile(LUL_FILE)

    if not csv_presente:
        print(f"\nERRORE: file CSV non trovato -> {CSV_FILE}")
        return

    print(f"\n[OK] CSV trovato:  {CSV_FILE}")
    if lul_presente:
        tipo = "PDF" if LUL_FILE.lower().endswith(".pdf") else "XLSX"
        print(f"[OK] LUL trovato:  {LUL_FILE}  ({tipo})")
    else:
        print(f"[--] LUL non trovato ({LUL_FILE_PDF} o {LUL_FILE_XLSX}) -> elaborazione solo CSV")

    # Carica CSV completo
    df_completo = carica_csv_completo(CSV_FILE)

    # Determina mesi da elaborare
    if MESE is None:
        periodi = mesi_disponibili_nel_csv(df_completo)
        print(f"\nModalita: TUTTI I MESI nel CSV")
        print(f"Mesi trovati: {', '.join(f'{NOMI_MESI[m]} {a}' for a, m in periodi)}")
    else:
        # Parsing del formato MMAAAA (es. 112025 = Novembre 2025, 32026 = Marzo 2026)
        codice = str(MESE)
        anno_sel = int(codice[-4:])
        mese_sel = int(codice[:-4])
        if not (1 <= mese_sel <= 12):
            print(f"\nERRORE: mese non valido nel codice {MESE}. Usa formato MMAAAA, es. 32026 per Marzo 2026.")
            return
        periodi = [(anno_sel, mese_sel)]
        print(f"\nModalita: mese singolo -> {NOMI_MESI[mese_sel]} {anno_sel}")

    # Carica LUL (se presente)
    dipendenti_lul = []
    mese_lul       = None
    if lul_presente:
        try:
            dipendenti_lul, mese_lul = carica_lul_auto(LUL_FILE)
        except Exception as e:
            print(f"  Errore nella lettura del LUL: {e} -> procedo senza LUL")

    print(f"\n{'-' * 65}")
    print(f"  ELABORAZIONE")
    print(f"{'-' * 65}")

    # Crea un unico Workbook con un foglio per ogni mese
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # rimuove il foglio vuoto creato di default
    fogli_creati = 0
    fogli_info = []          # per il foglio Riepilogo Generale
    azienda_global = ""

    # =========================================================================
    # ELENCO DISCENTI GLOBALE
    # =========================================================================
    # Costruiamo discenti_globali da TUTTO il CSV (df_completo), non solo
    # dai mesi che verranno elaborati. In questo modo OGNI foglio mensile
    # conterra' la riga di TUTTI i discenti presenti nel CSV, anche di quelli
    # che in quel mese non hanno maturato ore (la loro riga sara' vuota:
    # solo nome, codice fiscale e percorso, senza dati giornalieri ne'
    # colonne riepilogative).
    # =========================================================================
    discenti_globali = {}
    discenti_unici_df = (
        df_completo.groupby("Codice Fiscale")
        .agg(nome    =("Nome Cognome", "first"),
             percorso=("Percorso",     lambda x: " | ".join(sorted(set(x.dropna().astype(str))))))
        .reset_index()
    )
    for _, row in discenti_unici_df.iterrows():
        discenti_globali[row["Codice Fiscale"]] = {
            "nome":     row["nome"],
            "percorso": row["percorso"],
        }
    # Riordina i nomi in "COGNOME NOME" usando il LUL come riferimento, cosi'
    # l'ordinamento e la visualizzazione nei fogli sono per cognome (A->Z).
    normalizza_ordine_nomi(discenti_globali, dipendenti_lul)
    print(f"\nDiscenti totali nel CSV: {len(discenti_globali)} "
          f"(saranno tutti presenti in ogni foglio mensile)")

    periodi_elaborati_dati = []

    for anno, mese in periodi:
        print(f"\n>> {NOMI_MESI[mese]} {anno}")

        df_giornaliero, discenti_info, azienda = elabora_mese_dal_csv(
            df_completo, mese, anno, SOGLIA_ORA)

        if discenti_info.empty:
            print(f"   Nessun dato per {NOMI_MESI[mese]} {anno}, salto.")
            continue

        azienda_global = azienda

        cf_a_lul = None
        if dipendenti_lul:
            # Il LUL si riferisce a UN SOLO mese. Se abbiamo rilevato il mese del
            # LUL (mese_lul), applichiamo la Modalita' B solo a quel mese; gli
            # altri mesi restano in Modalita' A (solo CSV). Se il mese non e'
            # stato rilevato (es. LUL in XLSX), manteniamo il comportamento
            # legacy applicandolo a tutti i mesi.
            if mese_lul is None or (mese, anno) == mese_lul:
                cf_a_lul = abbina_discenti_lul(discenti_info, dipendenti_lul)
            else:
                print(f"   LUL non riferito a questo mese "
                      f"(LUL = {mese_lul[0]:02d}/{mese_lul[1]}) -> Modalita' A per "
                      f"{NOMI_MESI[mese]} {anno}")

        periodi_elaborati_dati.append((anno, mese, df_giornaliero, discenti_info, azienda, cf_a_lul))

    for anno, mese, df_giornaliero, discenti_info, azienda, cf_a_lul in periodi_elaborati_dati:
        info = scrivi_foglio(wb, discenti_info, df_giornaliero, azienda,
                             mese, anno, cf_a_lul=cf_a_lul,
                             discenti_globali=discenti_globali)
        fogli_info.append(info)
        fogli_creati += 1

    # Crea il foglio Riepilogo Generale (in prima posizione)
    if fogli_info:
        scrivi_foglio_riepilogo(wb, fogli_info, azienda_global,
                                discenti_globali=discenti_globali)

    if fogli_creati > 0:
        wb.save(OUTPUT_FILE)
        print(f"\n{'=' * 65}")
        print(f"  COMPLETATO")
        print(f"  File creato: {OUTPUT_FILE}")
        print(f"  Fogli nel file: {fogli_creati} ({', '.join(f'{NOMI_MESI[m]} {a}' for a, m in periodi)})")
        print(f"{'=' * 65}\n")
    else:
        print("\nNessun dato trovato. Controlla il CSV e le impostazioni.")


if __name__ == "__main__":
    main()
