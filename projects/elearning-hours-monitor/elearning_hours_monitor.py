"""
=============================================================================
E-LEARNING TRAINING HOURS MONITORING
=============================================================================
The script automatically detects the input files present in the folder
and produces a SINGLE Excel file with one sheet for each processed month.

  Daily cells        -> ACTUAL hours (= total hours - excess hours)
  Notes on the cells -> "Ore totali" + "Eccesso" (only if there is excess)

  Summary columns (always in the same order, MODALITA A and MODALITA B):
      1) Totale Ore Effettive  (HH:MM:SS)
      2) Ore Totali            (HH:MM:SS)
      3) Totale Eccesso        (HH:MM:SS)  = Ore Totali - Totale Ore Effettive
      4) Eccesso Weekend       (HH:MM:SS)
      5) Eccesso Dopo 18:00    (HH:MM:SS)
      6) Eccesso Mattutino     (HH:MM:SS)  (06:00-07:40)

  MODALITA A - CSV only
  MODALITA B - CSV + LUL (Presenze XLSX): the LAV hours excess rule
      and the 8h/day cap are still applied internally and flow into
      the Totale Eccesso, without a dedicated column.

  MODALITA B - LUL parsing (V17):
      - leggi_mappa_colonne handles: "1 L", "1", openpyxl dates
      - parse_blocco_dipendente looks for LAV also in col 1, handles absences
        with a code (F, M, ROL...) even without hours >= 8
      - carica_lul looks for the 'ORE' row in all sheets of the workbook
      - calcola_giorno distinguishes: lul=None (no match -> same as MODALITA A),
        lav=None (day outside the LUL -> same as MODALITA A), lav=0 (day not
        worked -> absent), lav>0 (worked day -> capped at lav_ore)

PERIOD SELECTION:
  - Set MESE = month number (e.g. 3 for March) for a specific month
  - Set MESE = None to process ALL months present in the CSV
  In both cases a single Excel file is produced with one sheet per month.

DEPENDENCIES:
    pip install pandas openpyxl

USAGE:
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
# CONFIGURATION
# =============================================================================

CSV_FILE    = "Report_Accessi.csv"
# The LUL can be a PDF (times platform printout) or an XLSX.
# The script automatically detects the first available file among these names.
# IMPORTANT: the PDF is the most reliable source. Converting the PDF to
# Excel often LOSES the employee names (except the first one), so whenever
# possible use the PDF directly.
LUL_FILE_PDF  = "Presenze.pdf"
LUL_FILE_XLSX = "Presenze.xlsx"
LUL_FILE      = LUL_FILE_XLSX   # compatibility: set dynamically in main()
AZIENDA_BREVE = "aicomply"   # short company name for the output file
# The output file name automatically includes today's date and the company
# E.g.: Monitoraggio_15-04-2026.xlsx
OUTPUT_FILE = f"Monitoraggio_{datetime.date.today().strftime('%d-%m-%Y')}.xlsx"

# Month to process:
#   MESE = None    -> process ALL months and years present in the CSV (recommended)
#   MESE = 32026   -> process only March 2026     (3  + 2026)
#   MESE = 112025  -> process only November 2025  (11 + 2025)
#   MESE = 12026   -> process only January 2026   (1  + 2026)
# Formula: write the month number followed by the 4-digit year
# Note: with MESE = None the ANNO variable below is ignored
MESE = None

CSV_SEPARATOR      = ";"
SOGLIA_ORA         = 18.0
SOGLIA_MATTINO_INI = 6.0
SOGLIA_MATTINO_FIN = 7 + 40/60
TOLLERANZA_MINUTI  = 20
CAP_ORE_GIORNALIERO = 8.0    # maximum recognizable hours per learner in a single day
ORE_LIMITE_FINANZIATE = 150  # maximum recognizable hours per learner (for the MIN column in the Riepilogo)


# =============================================================================
# CONSTANTS
# =============================================================================

NOMI_MESI = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
              "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

COLORI = {
    # Cell states (very light shades to reduce visual "noise")
    "verde":   "F0F6EC",   # normal usage
    "arancio": "FBEAD8",   # LAV hours excess
    "viola":   "F2E8EE",   # excess after threshold
    "azzurro": "E8EFF5",   # morning excess (06:00-07:40)
    "rosa":    "FAE6E6",   # weekend usage
    "giallo":  "FBF4D9",   # absent day with usage
    "grigio":  "F2F2F2",   # weekend without usage
    # UI
    "header":  "F4F5F7",   # headers - very light gray
    "titolo":  "ECF0F4",   # title background - very light blue
    "titolo_testo": "2C3E50",  # title text color
    "totale":  "F1F3F5",   # total row - very light gray
    "bordo":   "BDC3C7",   # thin borders
}


# Excel format for durations (also above 24 hours)
DURATION_FORMAT = "[h]:mm:ss"


# =============================================================================
# GENERAL UTILITIES
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
    """Normalizes a name for LUL<->CSV MATCHING: sorts the tokens alphabetically
    so that "MARIO ROSSI" and "ROSSI MARIO" coincide. Do NOT use this
    function to sort the learners in the sheets (see chiave_ordinamento)."""
    token = re.sub(r"[^A-Za-z ]", "", nome.upper()).split()
    return " ".join(sorted(token))


def chiave_ordinamento(nome):
    """Key for the ALPHABETICAL SORTING of learners in the sheets, by SURNAME.

    In the files the name is written "COGNOME NOME" (e.g. "ROSSI FABIO"), so
    the cleaned uppercase string already sorts by surname. Unlike
    normalizza_nome, here the tokens are NOT re-sorted: the word order
    (surname first) is preserved."""
    return re.sub(r"[^A-Za-z ]", "", str(nome).upper()).strip()


def normalizza_ordine_nomi(discenti_globali, dipendenti_lul):
    """Reorders the learner names into "COGNOME NOME" format using the LUL
    as reference.

    In the CSV names may be written "Nome Cognome" (e.g. "Mario Rossi"),
    while in the LUL they are "Cognome Nome" (e.g. "ROSSI MARIO"). To sort and
    display learners by COGNOME the tokens must be in the right order.

    Strategy:
      - for each learner the corresponding tokens are searched in the LUL
        (prefix match, robust to PDF truncations and inverted order);
      - if the LUL employee is found, the LUL token order is adopted
        (surname first), completing truncated tokens with the full version
        taken from the CSV (e.g. LUL "GIOVA" -> "GIOVANNI" from the CSV is kept);
      - if the learner is NOT in the LUL, the fallback applies: the LAST word
        is assumed to be the surname and is moved to the front.
    The "nome" field of each learner is updated to "COGNOME NOME".
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

        # Look for the compatible LUL employee
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
            # Adopt the LUL order, but for each LUL token prefer the LONGEST
            # version between LUL and CSV (so PDF truncations are completed
            # with the full name from the CSV).
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
            # any unmatched CSV tokens (rare) are appended
            nuovi.extend(csv_disp)
            info["nome"] = " ".join(nuovi).upper()
        else:
            # Fallback (learner NOT in the LUL): "Nome ... Cognome" is assumed, so
            # the surname is at the end. Surnames with a particle are handled
            # (DE, DEL, DELLA, DI, DA, LO, LA, LE, VAN, VON, MC, ...): the
            # particle preceding the last word is part of the surname.
            PARTICELLE = {"DE","DEL","DELLA","DELLE","DELLO","DEI","DEGLI","DI","DA",
                          "DAL","DALLA","LO","LA","LE","LI","VAN","VON","MC","MAC",
                          "SAN","SANTA","SANT","D"}
            if len(tok_csv) >= 2:
                # how many trailing words make up the surname?
                n_cog = 1
                # include particles immediately preceding the surname
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
    """Inserts the two logos in the given row (default: row 2).
    Company logo on the left (cell A), FNC on the right (last visible column).
    Gracefully handles missing image files."""
    from openpyxl.drawing.image import Image as XLImage
    ws.row_dimensions[riga_logo].height = altezza_riga
    for path, anchor in [(LOGO_AZIENDA, "A2"), (LOGO_FNC, None)]:
        if not os.path.isfile(path):
            continue
        try:
            img = XLImage(path)
            # Scale keeping the aspect ratio at the target height
            h_target = altezza_riga * 1.33   # points -> approximate pixels
            scale    = h_target / img.height
            img.width  = int(img.width  * scale)
            img.height = int(img.height * scale)
            if anchor is None:
                # Compute the last column of the title merge to position FNC
                # (uses a column far enough to the right — the calling function
                # can pass the correct letter via parameter if needed)
                anchor = "B2"   # placeholder; will be overwritten by the caller
            img.anchor = anchor
            ws.add_image(img)
        except Exception:
            pass   # If Pillow/openpyxl cannot handle the file, skip silently


def ore_decimali_a_hhmmss(ore_dec):
    if not ore_dec or ore_dec <= 0:
        return None
    totale_secondi = int(round(ore_dec * 3600))
    h = totale_secondi // 3600
    m = (totale_secondi % 3600) // 60
    s = totale_secondi % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


# =============================================================================
# CSV PARSING
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
    # Use the real session interval as denominator (not the consumed duration),
    # to correctly compute the study fraction proportional to the time after the threshold.
    elapsed = t_fine - t_inizio
    frazione_dopo = (t_fine - soglia) / elapsed if elapsed > 0 else 0
    # Clamp for numerical safety
    frazione_dopo = max(0.0, min(1.0, frazione_dopo))
    ore_dopo = durata * frazione_dopo
    return durata - ore_dopo, ore_dopo


def ore_in_finestra_mattino(riga):
    """
    Computes the study hours (proportional to elapsed time) that fall
    in the morning window [SOGLIA_MATTINO_INI, SOGLIA_MATTINO_FIN].
    Uses the same proportional logic as splitta_ore_per_soglia: the study
    fraction in the window is proportional to the elapsed time fraction in
    the window, not to the consumed duration.
    """
    durata   = parse_durata(riga["Totale Ore"])
    t_inizio = ora_in_decimale(riga["Primo Accesso"])
    t_fine   = ora_in_decimale(riga["Ultimo Accesso"])
    # Session entirely outside the morning window
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
    """Formats a decimal duration in compact form (for cell notes).
    E.g. 1.5h -> '1 h 30 min', 0.42h -> '25 min', 0.005h -> '18 s'."""
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
    print(f"Loading CSV: {percorso_file}")
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
    # ore_prima = "valid" hours (normal schedule = between the end of the morning window and the threshold)
    # = ore_prima_raw (before the threshold) minus the part falling in the morning window
    df_mese["ore_prima"] = (df_mese["ore_prima_raw"] - df_mese["ore_mattino"]).clip(lower=0)
    df_mese["day"] = df_mese["Giorno"].dt.day

    # Grouping by (Codice Fiscale, day): SUMS the hours of ALL the sessions of
    # that day, INCLUDING different training paths of the same person. If a
    # learner follows two or more paths and on the same day has sessions on
    # more than one of them (even with overlapping schedules), the hours are
    # summed (simple sum: any time overlaps are counted).
    df_giornaliero = (
        df_mese.groupby(["Codice Fiscale", "day"])
        .agg(ore_prima  =("ore_prima",   "sum"),
             ore_dopo   =("ore_dopo",    "sum"),
             ore_mattino=("ore_mattino", "sum"),
             ore_tot    =("ore_tot",     "sum"))
        .reset_index()
    )

    # Learner list: one learner = one Codice Fiscale. If they have multiple paths,
    # all of them are listed, concatenated with " | ".
    discenti_info = (
        df_mese.groupby("Codice Fiscale")
        .agg(nome    =("Nome Cognome", "first"),
             percorso=("Percorso",     lambda x: " | ".join(sorted(set(x.dropna().astype(str))))))
        .reset_index()
    )
    discenti_info["sort_key"] = discenti_info["nome"].apply(chiave_ordinamento)
    discenti_info = discenti_info.sort_values("sort_key").reset_index(drop=True)

    azienda = df_mese["Azienda"].iloc[0]
    print(f"  {NOMI_MESI[mese]} {anno}: {len(discenti_info)} learners | Company: {azienda}")
    return df_giornaliero, discenti_info, azienda


# =============================================================================
# LUL PARSING
# =============================================================================

def leggi_mappa_colonne(riga_ore):
    """Builds the map {column_index: day_number} from the LUL header row.

    Handles the various formats companies use to number the days in the 'ORE' row:
      - "1 L", "2 M", ... (classic format with space + day letter)
      - "1", "2", ...     (integer only, without letter)
      - datetime.date     (some LULs return openpyxl dates)
      - Skips "ORE", "TOT" and empty cells.
    """
    col_giorno = {}
    for ci, valore in enumerate(riga_ore):
        if valore is None:
            continue
        # openpyxl date object -> use the day of the month
        if isinstance(valore, (datetime.date, datetime.datetime)):
            col_giorno[ci] = valore.day
            continue
        s = str(valore).strip()
        if not s or s in ("ORE", "TOT"):
            continue
        # "1 L", "15 M", etc.
        m = re.match(r"^(\d{1,2})\s+[A-Za-z]", s)
        if m:
            col_giorno[ci] = int(m.group(1))
            continue
        # Integer only (e.g. "1", "15")
        if re.match(r"^\d{1,2}$", s):
            col_giorno[ci] = int(s)
    return col_giorno


def estrai_nome_dipendente(testo):
    m = re.search(
        r"Dipendente:\s*\d+\s*-\s*([A-Z][A-Z ]+?)(?:\s{2,}|\n|$|Data)", testo)
    return m.group(1).strip() if m else "SCONOSCIUTO"


def parse_blocco_dipendente(righe, indice_inizio, col_giorno):
    """Extracts the LAV hours and the absence days from an employee block in the LUL.

    Improvements over the previous version:
    - Looks for the LAV row even when the label is in a non-first cell (e.g. col 1).
    - Handles absences with hours < 8 but with a significant absence code (e.g. "F", "M").
    - Handles the alternative ASSENZE format with multiple code rows.
    - Distinguishes between "day not present in the LUL" (employee not in that month)
      and "day with LAV = 0" (employee present but with no declared worked hours).
    - Does not count a day with lav_ore > 0 as "absent" even if it has an absence entry.
    """
    nome    = estrai_nome_dipendente(str(righe[indice_inizio][0]))
    lav     = {}
    assenti = set()

    # Absence codes indicating a non-worked day (in addition to the hours >= 8 criterion)
    CODICI_ASSENZA = {"F", "M", "MR", "P", "ROL", "EX", "AL", "ASP", "INF",
                      "MAL", "CIG", "CIGS", "0", "ART", "SOS", "PERM"}

    for j in range(indice_inizio + 1, min(indice_inizio + 120, len(righe))):
        riga = righe[j]
        if not riga:
            continue
        # Look for the label in column 0 or column 1 (some LULs have a
        # descriptive column in col 0 and the real label in col 1)
        v0 = riga[0]
        v1 = riga[1] if len(riga) > 1 else None
        label = None
        if isinstance(v0, str):
            label = v0.strip()
        elif isinstance(v1, str):
            label = v1.strip()

        # ── LAV row ──────────────────────────────────────────────────────────
        if label == "LAV":
            for ci, val in enumerate(riga):
                if ci in col_giorno:
                    if isinstance(val, (int, float)):
                        lav[col_giorno[ci]] = float(val)
                    elif val is None or val == "":
                        # Day present in the map but without a value ->
                        # explicitly mark as 0 (employee in that month,
                        # day with no declared worked hours)
                        lav.setdefault(col_giorno[ci], 0.0)

        # ── ASSENZE section ──────────────────────────────────────────────────
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
                # Certain absence: hours >= 8 full day
                if isinstance(ore_ass, (int, float)) and ore_ass >= 8:
                    assenti.add(giorno_num)
                # Absence by recognized code (half day too)
                elif codice_s in CODICI_ASSENZA:
                    # Only if it has no LAV hours > 0 already recorded
                    if lav.get(giorno_num, 0.0) == 0.0:
                        assenti.add(giorno_num)

        # ── End of block ─────────────────────────────────────────────────────
        elif j > indice_inizio + 2 and isinstance(label, str):
            if "PIATTAFORMA" in label or label.startswith("Azienda") or "Dipendente" in label:
                break

    return {"name": nome, "lav": lav, "assente": assenti}


def carica_lul(percorso_file):
    """Loads the LUL (Presenze.xlsx) and returns the list of employees with
    LAV hours and absence days for each day of the month.

    Improvements over the previous version:
    - Looks for the 'ORE' row in all rows (not only the first ones), useful for LULs
      with company headers of variable length.
    - Tries all sheets of the workbook if the map is not found in the active sheet.
    - Debug log of employees found and not found.
    """
    print(f"Loading LUL: {percorso_file}")
    wb = openpyxl.load_workbook(percorso_file, data_only=True)

    # Try the active sheet first, then all the others
    fogli_da_provare = [wb.active] + [wb[s] for s in wb.sheetnames if wb[s] != wb.active]

    righe      = None
    col_giorno = {}

    for ws in fogli_da_provare:
        righe_candidate = list(ws.iter_rows(values_only=True))
        # Look for the 'ORE' row (day header)
        for riga in righe_candidate:
            if riga and riga[0] == "ORE":
                col_giorno = leggi_mappa_colonne(riga)
                if col_giorno:
                    righe = righe_candidate
                    print(f"  Column map found in sheet '{ws.title}' "
                          f"({len(col_giorno)} days)")
                    break
        if col_giorno:
            break

    if not col_giorno or righe is None:
        print("  WARNING: unable to find the 'ORE' row with the day map in the LUL.")
        print("  Check that the LUL has a row with 'ORE' in column A and the days "
              "in the format '1 L', '2 M', ... (or number only) in the following columns.")
        return []

    # Locate the starts of the employee blocks.
    # A row starts a block if it contains "Dipendente:" in column A.
    # The position of "PIATTAFORMA" is NOT constrained (in real LULs it can
    # appear at the end of the row, beyond the first characters): it is enough
    # that the word is present in the row, or that the row starts with "Azienda".
    inizi_blocchi = [
        i for i, riga in enumerate(righe)
        if riga and isinstance(riga[0], str)
        and re.search(r"Dipendente\s*:", riga[0])
    ]

    if not inizi_blocchi:
        print("  WARNING: no employee block found in the LUL. "
              "Check the format (expected: row with 'Dipendente:' in column A).")
        return []

    dipendenti = [parse_blocco_dipendente(righe, i, col_giorno) for i in inizi_blocchi]
    # Filter out any empty blocks / blocks without a name
    dipendenti = [d for d in dipendenti if d["name"] != "SCONOSCIUTO" or d["lav"]]
    print(f"  Found {len(dipendenti)} employees in the LUL")
    return dipendenti


def carica_lul_da_pdf(percorso_file):
    """Loads the LUL directly from the times platform PDF (detailed attendance
    control printout) and returns the list of employees with LAV hours and
    absence days.

    Why read the PDF instead of the converted Excel:
    in the PDF each employee has their own page header with name and
    "Ore lavorate". The PDF->Excel conversion typically keeps only the first
    header and makes the other blocks anonymous: reading the PDF
    solves the problem at the root.

    Parsing technique:
    - one page = one employee;
    - from the header the name and "Ore lavorate" are extracted (for validation);
    - from the 'ORE' row each day-number is mapped to its X coordinate;
    - the 'LAV' row contains the worked hours, aligned by X to the nearest
      day column (columns in the PDF are not evenly spaced);
    - a day with LAV > 0 is worked; weekdays without a LAV value
      (holidays, leave, illness, ROL, permits) are marked 0.0 and
      treated downstream as not worked.
    """
    try:
        import pdfplumber
    except ImportError:
        print("  ERROR: the 'pdfplumber' library is required to read the LUL in PDF.")
        print("          Install it from the terminal with:  pip install pdfplumber")
        print("          (or:  python -m pip install pdfplumber)")
        print("          Without LUL the sheets will be generated in MODALITA A (CSV only).")
        return [], None
    from collections import defaultdict

    def _num(s):
        return float(str(s).replace(",", "."))

    print(f"Loading LUL (PDF): {percorso_file}")
    pdf = pdfplumber.open(percorso_file)
    dipendenti = []
    senza_ore  = 0
    mese_lul   = None   # (month_num, year) detected from the PDF header

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

        # Group the tokens by row (2px bins on the vertical axis)
        righe_y = defaultdict(list)
        for w in words:
            righe_y[round(w['top'] / 2) * 2].append(w)

        # Map day -> X center from the 'ORE' row
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

        # 'LAV' row -> worked hours per day (by X position)
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
                        if v > 24:      # TOT column at the end of the row
                            continue
                        g = _col_giorno((w['x0'] + w['x1']) / 2)
                        if g:
                            lav[g] = v
                break   # only the first 'LAV' row

        if not lav:
            senza_ore += 1
        # The weekdays of the month WITHOUT a LAV value (leave, illness, ROL,
        # permits, holidays) are explicitly marked 0.0: this way
        # calcola_giorno treats them as "not worked" (LAV excess) and not as
        # "day outside the LUL". Weekends are left untouched: the weekend
        # logic of calcola_giorno handles them. Month/year are derived from the header.
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
                    wd = datetime.date(anno_pdf, num_mese, d).weekday()  # 0=Mon..6=Sun
                    if wd < 5 and d not in lav:   # weekday and not already worked
                        lav[d] = 0.0
        dipendenti.append({"name": nome, "lav": lav, "assente": set(),
                           "ore_dichiarate": ore_dichiarate})

    # Validation: LAV sum == declared "Ore lavorate"
    incongruenti = []
    for d in dipendenti:
        tot = sum(d["lav"].values())
        if d["ore_dichiarate"] is not None and abs(tot - d["ore_dichiarate"]) > 0.5:
            incongruenti.append(f"{d['name']} (LAV={tot:.0f}h vs dich={d['ore_dichiarate']:.0f}h)")

    print(f"  Found {len(dipendenti)} employees in the LUL (PDF)")
    if mese_lul:
        print(f"  LUL month: {mese_lul[0]:02d}/{mese_lul[1]} "
              f"(MODALITA B will be applied ONLY to this month)")
    else:
        print("  WARNING: unable to detect the month from the LUL "
              "('Mese: ...' header not found).")
    if senza_ore:
        print(f"  Of which {senza_ore} without worked hours in the month (e.g. terminated/absent).")
    if incongruenti:
        print(f"  WARNING: {len(incongruenti)} employees with LAV sum different from the declared hours:")
        for s in incongruenti[:10]:
            print(f"     - {s}")
    return dipendenti, mese_lul


def carica_lul_auto(percorso_file):
    """Automatically chooses the parser based on the LUL file extension.

    Always returns a tuple (dipendenti, mese_lul), where mese_lul is
    (month_num, year) if detected, otherwise None. For the XLSX format the
    month is not detectable from the header, so mese_lul = None and MODALITA B
    is applied to all months (legacy behavior)."""
    ext = os.path.splitext(percorso_file)[1].lower()
    if ext == ".pdf":
        return carica_lul_da_pdf(percorso_file)
    return carica_lul(percorso_file), None


def abbina_discenti_lul(discenti_info, dipendenti_lul):
    """Matches the CSV learners to the LUL employees by name.

    Problem with times platform PDFs: the employee name is TRUNCATED to
    about 19 characters (e.g. "GIORDANETTI ALESSAN" instead of
    "GIORDANETTI ALESSANDRO", "LOMBARDINI BEATRIC" instead of
    "LOMBARDINI BEATRICE"). An exact comparison would fail for all long
    names, wrongly leaving them in MODALITA A (no LAV hours check).

    Matching strategy (in order):
      1) exact match on the normalized name (sorted tokens);
      2) SURNAME+NAME PREFIX match: the LUL name (possibly truncated)
         must be a prefix of the full CSV name, or vice versa,
         comparing the character sequence without spaces. This way
         "GIORDANETTI ALESSAN" matches "GIORDANETTI ALESSANDRO".
    If a LUL name is an ambiguous prefix of multiple learners, it is NOT
    matched and is reported, to avoid wrong attributions.
    """
    def _tokens(nome):
        # sorted set of uppercase alphabetic tokens, without stray accents
        return [t for t in re.sub(r"[^A-Za-z ]", " ", str(nome).upper()).split() if t]

    def _key(nome):
        # key for exact match: alphabetically sorted tokens
        return " ".join(sorted(_tokens(nome)))

    def _match_token(a, b):
        # two tokens match if one is a prefix of the other of at least 3 letters
        # (handles PDF truncations, e.g. FEDERIC ~ FEDERICO, GIOVA ~ GIOVANNI)
        if a == b:
            return True
        corto, lungo = (a, b) if len(a) <= len(b) else (b, a)
        return len(corto) >= 3 and lungo.startswith(corto)

    def _nomi_compatibili(tok_csv, tok_lul):
        # each token of the shorter name must find a (prefix) match in a
        # still-free token of the other name. Robust to inverted order
        # (CSV "Nome Cognome" vs LUL "Cognome Nome") and to truncated/missing tokens.
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
        # at least 2 matching tokens (surname + name) to avoid false positives
        return len(piccolo) >= 2

    # LUL indices
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

        # 1) exact match on the sorted tokens
        if kcsv in lul_per_nome:
            cf_a_lul[cf] = lul_per_nome[kcsv]
            continue

        # 2) token-by-token prefix match (inverted order + truncations)
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

    print(f"  Matched: {len(cf_a_lul)} / {len(discenti_info)}")
    if abbinati_prefisso:
        print(f"  Matched by truncated/partial name ({len(abbinati_prefisso)}):")
        for a in abbinati_prefisso:
            print(f"     - {a}")
    if non_abbinati:
        print(f"  Not matched (remain in MODALITA A): {', '.join(non_abbinati)}")
    return cf_a_lul


# =============================================================================
# HOURS CALCULATION PER DAY
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

    # Apply tolerance after 18:00 and in the morning (20 min/day per window do not count)
    tolleranza_ore  = TOLLERANZA_MINUTI / 60
    ore_dopo_eff    = 0.0 if ore_dopo    <= tolleranza_ore else ore_dopo
    ore_mattino_eff = 0.0 if ore_mattino <= tolleranza_ore else ore_mattino

    # Apply daily cap: if the normal-schedule hours exceed 8h, the excess is not recognized.
    exc_cap   = max(0.0, ore_prima - CAP_ORE_GIORNALIERO)
    ore_prima = min(ore_prima, CAP_ORE_GIORNALIERO)

    # ── MODALITA A: CSV only ────────────────────────────────────────────────
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
    # Method (manually verified on real data):
    #   1) Weekend          -> all hours are weekend excess, eff = 0
    #   2) No LUL match (lul None) or day outside the LUL (lav_ore None)
    #                       -> same as MODALITA A: no LAV constraint
    #   3) NOT worked day (absent, or lav_ore == 0)
    #                       -> all weekday hours are LAV excess, eff = 0
    #   4) Worked day (lav_ore > 0)
    #                       -> valid hours = min(ore_prima, lav_ore);
    #                          the amount beyond the worked hours is LAV excess.
    # In all cases the "evening" (ore_dopo_eff), "morning" (ore_mattino_eff)
    # and "8h cap" (exc_cap) cuts already computed above remain valid.
    lul       = cf_a_lul.get(cf)
    e_assente = (giorno in lul["assente"]) if lul else False
    lav_ore   = lul["lav"].get(giorno, None) if lul else None

    # 1) Weekend
    if giorno in weekend_days:
        return {"ore_tot": ore_tot, "eff": 0.0,
                "exc_we": ore_tot, "exc_dopo": 0.0, "exc_mattino": 0.0,
                "exc_cap": 0.0, "exc_lav": 0.0,
                "colore": COLORI["rosa"] if ore_tot > 0 else COLORI["grigio"]}

    # 2) No LAV constraint applicable -> same as MODALITA A
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

    # 3) NOT worked day (declared absence or worked hours = 0)
    if e_assente or lav_ore == 0.0:
        return {"ore_tot": ore_tot, "eff": 0.0, "exc_we": 0.0,
                "exc_dopo": ore_dopo_eff, "exc_mattino": ore_mattino_eff,
                "exc_cap": exc_cap, "exc_lav": ore_prima,
                "colore": COLORI["giallo"] if ore_tot > 0 else None}

    # 4) Worked day: capped at the actually worked hours (LAV)
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
# EXCEL SHEET WRITING (one sheet per month, on the same Workbook)
# =============================================================================

def scrivi_foglio(wb, discenti_info, df_giornaliero, azienda,
                  mese, anno, cf_a_lul=None, discenti_globali=None):
    """
    Adds a sheet to the Workbook wb with the data of the given month.
    """
    modalita_b    = cf_a_lul is not None
    totale_giorni = giorni_nel_mese(anno, mese)
    weekend_days  = giorni_weekend(anno, mese)
    modo_label    = "MODALITA B (CSV+LUL)" if modalita_b else "MODALITA A (CSV only)"

    titolo = f"MONITORAGGIO ORE FORMAZIONE - {NOMI_MESI[mese]} {anno} - {azienda}"

    # Column layout
    COL_A       = 1
    COL_B       = 2
    COL_C       = 3
    COL_DAY_INI = 4
    COL_DAY_FIN = COL_DAY_INI + totale_giorni - 1

    # Summary columns — fixed order.
    # In MODALITA B an "Eccesso Ore LAV" column is added between
    # "Totale Eccesso" and "Eccesso Weekend"; in MODALITA A this column
    # is not present because there is no LUL source.
    COL_TOT_EFF      = COL_DAY_FIN + 1   # Totale Ore Effettive (=SUM of day cells)
    COL_TOT_ORE      = COL_DAY_FIN + 2   # Ore Totali           (Python value)
    COL_TOT_EXC      = COL_DAY_FIN + 3   # Totale Eccesso       (= total hours - actual)
    if modalita_b:
        COL_EXC_LAV      = COL_DAY_FIN + 4   # Eccesso Ore LAV (MODALITA B only)
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

    # Styles (clean palette, dark gray text, light background)
    font_base    = Font(name="Arial", size=9, color="333333")
    font_header  = Font(name="Arial", size=9, bold=True, color="495057")
    font_titolo  = Font(name="Arial", size=12, bold=True, color=COLORI["titolo_testo"])
    font_totale  = Font(name="Arial", size=9, bold=True, color="333333")
    allin_centro = Alignment(horizontal="center", vertical="center")
    allin_sin    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    allin_wrap   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bordo_top    = Border(top=Side(style="thin", color=COLORI["bordo"]))

    ws = wb.create_sheet(title=f"{NOMI_MESI[mese]} {anno}")

    # ── Title ────────────────────────────────────────────────────────────────
    ws.row_dimensions[RIGA_TITOLO].height = 26
    c = ws.cell(RIGA_TITOLO, COL_A, value=titolo)
    c.font = font_titolo; c.fill = fill_cell(COLORI["titolo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=RIGA_TITOLO, start_column=COL_A,
                   end_row=RIGA_TITOLO,   end_column=ULTIMA_COL)

    # ── Logos row ────────────────────────────────────────────────────────────
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

    # ── Headers ──────────────────────────────────────────────────────────────
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

    # Summary columns — 6 fixed columns in a specific order
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

    # ── Column widths ─────────────────────────────────────────────────────────
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

    # Map CF -> actual hours of the month (for the Riepilogo Generale sheet,
    # which now writes them as VALUES instead of cross-sheet formulas).
    ore_eff_per_cf = {}

    for idx, (cf, info_disc) in enumerate(lista_discenti):
        r = RIGA_DATI + idx
        cf_a_riga[cf] = r

        # The name in the files is already in "COGNOME NOME" format: it is used
        # as-is (uppercase only), without swapping the tokens.
        cognome_nome = info_disc["nome"].strip().upper()

        ws.row_dimensions[r].height = 17
        c = ws.cell(r, COL_A, value=cognome_nome); c.font = font_base; c.alignment = allin_sin
        c = ws.cell(r, COL_B, value=cf);           c.font = font_base; c.alignment = allin_centro
        c = ws.cell(r, COL_C, value=info_disc["percorso"]); c.font = font_base; c.alignment = allin_sin

        # Monthly accumulators:
        # - acc_tot:     sum of total hours of the month (for "Ore Totali")
        # - acc_eff:     sum of actual hours of the month (for the Riepilogo)
        # - acc_we:      sum of weekend excess
        # - acc_dopo:    sum of excess after the evening threshold
        # - acc_mattino: sum of morning excess
        # - acc_lav:     sum of LAV hours excess (MODALITA B only, always 0 in A)
        # exc_cap is still counted in the Totale Eccesso via the formula
        # (Ore Totali - Totale Ore Effettive), so it does not need to be accumulated.
        acc_tot = acc_eff = acc_we = acc_dopo = acc_mattino = acc_lav = 0.0

        for d in range(1, totale_giorni + 1):
            col = COL_DAY_INI + d - 1
            res = calcola_giorno(cf, d, df_giornaliero, cf_a_lul or {}, weekend_days, modalita_b)

            ore_tot_daily   = res["ore_tot"]
            excess_daily    = (res["exc_we"]   + res["exc_dopo"]    + res["exc_mattino"]
                              + res.get("exc_cap", 0) + res.get("exc_lav", 0))
            ore_eff_daily   = ore_tot_daily - excess_daily

            # Cell value: ACTUAL hours (= total - excess).
            # If ore_tot==0 -> empty cell (day without usage).
            # If ore_tot>0 but ore_eff==0 (e.g. weekend entirely in excess)
            # it explicitly shows 0:00:00 with the excess background color.
            if ore_tot_daily > 0:
                c = ws.cell(r, col, value=ore_eff_daily / 24)
                c.number_format = DURATION_FORMAT
            else:
                c = ws.cell(r, col)
            c.font = Font(name="Arial", size=8, color="333333"); c.alignment = allin_centro
            if res["colore"]:
                c.fill = fill_cell(res["colore"])

            # Note: ONLY if there is actually counted excess (total hours != actual hours)
            # Shows "Ore totali", "Eccesso" and the detail by excess type.
            if excess_daily > 1e-6:
                note_lines = [
                    f"Ore totali: {ore_decimali_a_hhmmss(ore_tot_daily)}",
                    f"Eccesso: {ore_decimali_a_hhmmss(excess_daily)}",
                ]
                # Detail by excess type (only entries with value > 0)
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

        # Store the actual hours for the Riepilogo (also for non-users: 0)
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

        # ── Summary columns (fixed order) ─────────────────────────────────────
        # 1) Totale Ore Effettive = sum of the daily cells (=SUM)
        scrivi_formula(COL_TOT_EFF,
                       f"=SUM({col_day_ini_letter}{r}:{col_day_fin_letter}{r})")
        # 2) Ore Totali = Python value (sum of the daily ore_tot)
        scrivi_durata(COL_TOT_ORE, acc_tot)
        # 3) Totale Eccesso = Ore Totali - Totale Ore Effettive (transparent formula)
        scrivi_formula(COL_TOT_EXC,
                       f"={col_tot_ore_letter}{r}-{col_tot_eff_letter}{r}")
        # 3-bis) Eccesso Ore LAV (MODALITA B only)
        if modalita_b:
            scrivi_durata(COL_EXC_LAV, acc_lav)
        # 4) Weekend excess
        scrivi_durata(COL_EXC_WE, acc_we)
        # 5) Evening threshold excess
        scrivi_durata(COL_EXC_DOPO, acc_dopo)
        # 6) Morning excess
        scrivi_durata(COL_EXC_MATTINO, acc_mattino)

    # ── TOTALE row ───────────────────────────────────────────────────────────
    # All totals are obtained via =SUM() formulas on the user rows,
    # so that the company can verify the counts by clicking the cells.
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

    # The summary columns (in MODALITA B also with EXC_LAV)
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

    # Log statistics (console output only, not written to the sheet)
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

    print(f"  -> Sheet '{NOMI_MESI[mese]} {anno}' written | "
          f"Learners: {N} | Hours used: {ore_decimali_a_hhmmss(tot_fruite_log)} | "
          f"Excess after 18:00: {ore_decimali_a_hhmmss(tot_dopo_log)} | "
          f"Weekend excess: {ore_decimali_a_hhmmss(tot_we_log)} | {modo_label}")

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
        # Dictionary CF -> actual hours of the month (in decimal hours).
        # Used by the Riepilogo sheet to write VALUES (not formulas).
        "ore_eff_per_cf":   ore_eff_per_cf,
        "discenti":         discenti_info[["Codice Fiscale", "nome", "percorso"]].to_dict("records"),
    }


# =============================================================================
# RIEPILOGO GENERALE SHEET (totals per learner, across all processed months)
# =============================================================================

def ore_decimali_a_testo(ore_dec):
    """Formats decimal hours as an 'X H YY M ZZ S' string (e.g. '150 H 00 M 00 S').

    Hours are not padded (so '150 H ...' stays natural), while
    minutes and seconds are always 2 digits. Values <=0 -> '0 H 00 M 00 S'.
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
    Creates a "Riepilogo Generale" sheet with one learner per row, one
    column for each processed month and three final columns:
      - Totale Ore Maturate    (value, sum of the months)
      - Min(150 ore)           (value, capped at 150 hours)
      - Ore (testo)            (text "X H YY M ZZ S" of the capped value)

    All the Riepilogo cells use Excel formulas for maximum transparency:
      - Monthly columns     : ='SheetName'!{COL_TOT_EFF}{row}
      - Totale Ore Maturate : =SUM(monthly columns)
      - Min(150 ore)        : =MIN(Totale, 150/24)
      - Ore (testo)         : TEXT()+MOD() formula -> "X H YY M ZZ S"

    fogli_info: list of dicts returned by scrivi_foglio() (one per month).
    discenti_globali: dict {cf: {"nome":..., "percorso":...}} with ALL the
        learners of the CSV. If provided, every learner appears in the table
        even if they earned no hours in any month (row with all 0:00:00).
    """
    if not fogli_info:
        return

    # Sort the sheets by year/month
    fogli_info = sorted(fogli_info, key=lambda x: (x["anno"], x["mese"]))

    # Learner list: use the global one if provided, otherwise the monthly union
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

    # Column layout (adds the "Ore Tolte" column after Min(150))
    COL_NOME       = 1
    COL_CF         = 2
    COL_PERCORSO   = 3
    COL_MESE_INI   = 4
    COL_MESE_FIN   = COL_MESE_INI + n_mesi - 1
    COL_TOTALE     = COL_MESE_FIN + 1   # Totale Ore Maturate (value)
    COL_MIN_150    = COL_MESE_FIN + 2   # Cap at 150 hours (value)
    COL_ORE_TOLTE  = COL_MESE_FIN + 3   # Ore Totali - Min(150): non-recognizable hours
    COL_TESTO      = COL_MESE_FIN + 4   # "X H YY M ZZ S" string (text value)
    ULTIMA_COL     = COL_TESTO

    RIGA_TITOLO = 1
    RIGA_LOGO   = 2
    RIGA_HEADER = 3
    RIGA_DATI   = 4
    RIGA_TOTALE = RIGA_DATI + N

    # Styles
    font_base    = Font(name="Arial", size=10, color="333333")
    font_header  = Font(name="Arial", size=9, bold=True, color="495057")
    font_titolo  = Font(name="Arial", size=13, bold=True, color=COLORI["titolo_testo"])
    font_dato_b  = Font(name="Arial", size=10, bold=True, color="333333")
    allin_centro = Alignment(horizontal="center", vertical="center")
    allin_sin    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    allin_wrap   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    bordo_top    = Border(top=Side(style="thin", color=COLORI["bordo"]))

    # Insert the sheet in FIRST position
    ws = wb.create_sheet(title="Riepilogo Generale", index=0)

    # ── Title ────────────────────────────────────────────────────────────────
    ws.row_dimensions[RIGA_TITOLO].height = 28
    c = ws.cell(RIGA_TITOLO, COL_NOME, value=f"Riepilogo Generale - {azienda}")
    c.font = font_titolo; c.fill = fill_cell(COLORI["titolo"])
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=RIGA_TITOLO, start_column=COL_NOME,
                   end_row=RIGA_TITOLO,   end_column=ULTIMA_COL)

    # ── Logos row ────────────────────────────────────────────────────────────
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

    # ── Headers ──────────────────────────────────────────────────────────────
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

    # Explanatory tooltips
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

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions[get_column_letter(COL_NOME)].width     = 28
    ws.column_dimensions[get_column_letter(COL_CF)].width       = 17
    ws.column_dimensions[get_column_letter(COL_PERCORSO)].width = 35
    for col in range(COL_MESE_INI, COL_MESE_FIN + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions[get_column_letter(COL_TOTALE)].width  = 14
    ws.column_dimensions[get_column_letter(COL_MIN_150)].width = 14
    ws.column_dimensions[get_column_letter(COL_ORE_TOLTE)].width = 12
    ws.column_dimensions[get_column_letter(COL_TESTO)].width   = 22

    # ── Column letters (used in formulas) ────────────────────────────────────
    col_mese_ini_letter  = get_column_letter(COL_MESE_INI)
    col_mese_fin_letter  = get_column_letter(COL_MESE_FIN)
    col_totale_letter    = get_column_letter(COL_TOTALE)
    col_min150_letter    = get_column_letter(COL_MIN_150)
    col_ore_tolte_letter = get_column_letter(COL_ORE_TOLTE)

    # ── Data rows ────────────────────────────────────────────────────────────
    for i, (cf, info) in enumerate(discenti_lista):
        r = RIGA_DATI + i
        # The name is already "COGNOME NOME": used as-is (uppercase).
        cognome_nome = info["nome"].strip().upper()

        ws.row_dimensions[r].height = 19
        c = ws.cell(r, COL_NOME,     value=cognome_nome);     c.font = font_base; c.alignment = allin_sin
        c = ws.cell(r, COL_CF,       value=cf);               c.font = font_base; c.alignment = allin_centro
        c = ws.cell(r, COL_PERCORSO, value=info["percorso"]); c.font = font_base; c.alignment = allin_sin

        # Monthly columns: cross-sheet MATCH+INDEX looking up the CF in col B of the sheet
        # Uses MATCH(CF, sheet_col_B, 0) to find the correct row regardless of
        # the sorting — avoids the inversion bug when names have accents or
        # special characters that alter the sort key.
        for j, f in enumerate(fogli_info):
            col    = COL_MESE_INI + j
            sname  = f["sheet_name"].replace("'", "''")
            col_cf_ms  = get_column_letter(2)                    # col B = Codice Fiscale
            col_eff_ms = get_column_letter(f["col_ore_maturate"])
            n_righe    = f.get("n_discenti", 200)                # generous search range
            # IFERROR(...,0) for learners absent in that month
            formula_eff = (
                f"=IFERROR(INDEX('{sname}'!{col_eff_ms}{RIGA_DATI}:{col_eff_ms}{RIGA_DATI+n_righe},"
                f"MATCH({chr(34)}{cf}{chr(34)},'{sname}'!{col_cf_ms}{RIGA_DATI}:{col_cf_ms}{RIGA_DATI+n_righe},0)),0)"
            )
            c = ws.cell(r, col, value=formula_eff)
            c.number_format = DURATION_FORMAT
            c.font = font_base; c.alignment = allin_centro

        # COL_TOTALE: =SUM(monthly columns) formula
        c = ws.cell(r, COL_TOTALE,
                    value=f"=SUM({col_mese_ini_letter}{r}:{col_mese_fin_letter}{r})")
        c.number_format = DURATION_FORMAT
        c.font = font_dato_b; c.alignment = allin_centro
        c.fill = fill_cell(COLORI["totale"])

        # COL_MIN_150: =MIN(Totale, limit/24) formula
        c = ws.cell(r, COL_MIN_150,
                    value=f"=MIN({col_totale_letter}{r},{ORE_LIMITE_FINANZIATE}/24)")
        c.number_format = DURATION_FORMAT
        c.font = font_dato_b; c.alignment = allin_centro
        c.fill = fill_cell(COLORI["totale"])

        # COL_ORE_TOLTE: MATCH+INDEX sum over COL_TOT_EXC of each monthly sheet
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

        # COL_TESTO: Excel formula "X H YY M ZZ S" from the Min(150) cell
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

    # ── TOTALE row ───────────────────────────────────────────────────────────
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
    # Text column: left empty in the total
    cc = ws.cell(RIGA_TOTALE, COL_TESTO)
    cc.fill = fill_cell(COLORI["totale"]); cc.border = bordo_top

    print(f"\n>> Sheet 'Riepilogo Generale' created | "
          f"Learners: {N} | Months: {n_mesi}")


# =============================================================================
# MAIN
# =============================================================================

def main():
    print("=" * 65)
    print("  E-LEARNING TRAINING HOURS MONITORING")
    print("=" * 65)

    csv_presente = os.path.isfile(CSV_FILE)
    # Detect the LUL: PDF preferred (more reliable for names), then XLSX
    global LUL_FILE
    if os.path.isfile(LUL_FILE_PDF):
        LUL_FILE = LUL_FILE_PDF
    elif os.path.isfile(LUL_FILE_XLSX):
        LUL_FILE = LUL_FILE_XLSX
    else:
        LUL_FILE = LUL_FILE_XLSX  # default name for messages
    lul_presente = os.path.isfile(LUL_FILE)

    if not csv_presente:
        print(f"\nERROR: CSV file not found -> {CSV_FILE}")
        return

    print(f"\n[OK] CSV found:  {CSV_FILE}")
    if lul_presente:
        tipo = "PDF" if LUL_FILE.lower().endswith(".pdf") else "XLSX"
        print(f"[OK] LUL found:  {LUL_FILE}  ({tipo})")
    else:
        print(f"[--] LUL not found ({LUL_FILE_PDF} or {LUL_FILE_XLSX}) -> CSV-only processing")

    # Load the full CSV
    df_completo = carica_csv_completo(CSV_FILE)

    # Determine the months to process
    if MESE is None:
        periodi = mesi_disponibili_nel_csv(df_completo)
        print(f"\nMode: ALL MONTHS in the CSV")
        print(f"Months found: {', '.join(f'{NOMI_MESI[m]} {a}' for a, m in periodi)}")
    else:
        # Parsing of the MMYYYY format (e.g. 112025 = November 2025, 32026 = March 2026)
        codice = str(MESE)
        anno_sel = int(codice[-4:])
        mese_sel = int(codice[:-4])
        if not (1 <= mese_sel <= 12):
            print(f"\nERROR: invalid month in code {MESE}. Use MMYYYY format, e.g. 32026 for March 2026.")
            return
        periodi = [(anno_sel, mese_sel)]
        print(f"\nMode: single month -> {NOMI_MESI[mese_sel]} {anno_sel}")

    # Load the LUL (if present)
    dipendenti_lul = []
    mese_lul       = None
    if lul_presente:
        try:
            dipendenti_lul, mese_lul = carica_lul_auto(LUL_FILE)
        except Exception as e:
            print(f"  Error reading the LUL: {e} -> proceeding without LUL")

    print(f"\n{'-' * 65}")
    print(f"  PROCESSING")
    print(f"{'-' * 65}")

    # Create a single Workbook with one sheet for each month
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # remove the default empty sheet created automatically
    fogli_creati = 0
    fogli_info = []          # for the Riepilogo Generale sheet
    azienda_global = ""

    # =========================================================================
    # GLOBAL LEARNER LIST
    # =========================================================================
    # We build discenti_globali from the WHOLE CSV (df_completo), not only
    # from the months that will be processed. This way EVERY monthly sheet
    # will contain the row of ALL the learners present in the CSV, including
    # those who earned no hours in that month (their row will be empty:
    # only name, codice fiscale and path, without daily data or
    # summary columns).
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
    # Reorder the names into "COGNOME NOME" using the LUL as reference, so
    # the sorting and display in the sheets are by surname (A->Z).
    normalizza_ordine_nomi(discenti_globali, dipendenti_lul)
    print(f"\nTotal learners in the CSV: {len(discenti_globali)} "
          f"(all of them will appear in every monthly sheet)")

    periodi_elaborati_dati = []

    for anno, mese in periodi:
        print(f"\n>> {NOMI_MESI[mese]} {anno}")

        df_giornaliero, discenti_info, azienda = elabora_mese_dal_csv(
            df_completo, mese, anno, SOGLIA_ORA)

        if discenti_info.empty:
            print(f"   No data for {NOMI_MESI[mese]} {anno}, skipping.")
            continue

        azienda_global = azienda

        cf_a_lul = None
        if dipendenti_lul:
            # The LUL refers to a SINGLE month. If we detected the LUL month
            # (mese_lul), we apply MODALITA B only to that month; the
            # other months remain in MODALITA A (CSV only). If the month was
            # not detected (e.g. LUL in XLSX), we keep the legacy behavior
            # and apply it to all months.
            if mese_lul is None or (mese, anno) == mese_lul:
                cf_a_lul = abbina_discenti_lul(discenti_info, dipendenti_lul)
            else:
                print(f"   LUL not referring to this month "
                      f"(LUL = {mese_lul[0]:02d}/{mese_lul[1]}) -> MODALITA A for "
                      f"{NOMI_MESI[mese]} {anno}")

        periodi_elaborati_dati.append((anno, mese, df_giornaliero, discenti_info, azienda, cf_a_lul))

    for anno, mese, df_giornaliero, discenti_info, azienda, cf_a_lul in periodi_elaborati_dati:
        info = scrivi_foglio(wb, discenti_info, df_giornaliero, azienda,
                             mese, anno, cf_a_lul=cf_a_lul,
                             discenti_globali=discenti_globali)
        fogli_info.append(info)
        fogli_creati += 1

    # Create the Riepilogo Generale sheet (in first position)
    if fogli_info:
        scrivi_foglio_riepilogo(wb, fogli_info, azienda_global,
                                discenti_globali=discenti_globali)

    if fogli_creati > 0:
        wb.save(OUTPUT_FILE)
        print(f"\n{'=' * 65}")
        print(f"  COMPLETED")
        print(f"  File created: {OUTPUT_FILE}")
        print(f"  Sheets in file: {fogli_creati} ({', '.join(f'{NOMI_MESI[m]} {a}' for a, m in periodi)})")
        print(f"{'=' * 65}\n")
    else:
        print("\nNo data found. Check the CSV and the settings.")


if __name__ == "__main__":
    main()
