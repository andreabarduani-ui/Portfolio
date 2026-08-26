"""
====================================================================
 FUNDING CALL SCRAPER - Interprofessional Funds (v2)
 Goal: find NON-FNC3 calls with a budget above a threshold (default 800,000 EUR)
====================================================================

WHAT CHANGES COMPARED TO V1
-----------------------------------------------------------------
V1 scanned the funds' HOMEPAGES with word matching that was too
generic ("risorse", "lavoratori", "dipendenti", etc.), which produced
almost only noise (menu items, presentation sentences) and never
extracted an essential piece of data: the call's AMOUNT.

This v2:

  1) Targets the "elenco bandi/avvisi" (call/notice list) pages
     directly, with two corrections found by manually checking the
     sites (see SITI_TARGET below for details):
       - Fondo FBA: "Archivio Avvisi" is the CLOSED archive (latest
         call 2020). The page with the actually open calls is
         "Avvisi Aperti" -> replaced.
       - Fondo Conoscenza: "Avvisi Aperti" (Conto Sistema) was added
         alongside the "Programmazione 2026" page; it lists the
         individual active notices with a direct link.

  2) Does TWO-LEVEL scraping:
       level 1: list page   -> identifies the individual calls (links)
       level 2: call page   -> opens the detail and extracts its
                               text, amount and FNC3 presence
     This is because, on the verified sites, the list shows only
     title/date, while the amount ("Dotazione finanziaria") and any
     link to FNC3 appear ONLY in the detail.
     Real example (Fondo FBA):
       - in the list: "Avviso Competenze per l'innovazione" (no
         mention of FNC3)
       - in the detail: "Opportunita' di finanziamento con il Fondo
         Nuove Competenze - Terza Edizione" -> this is FNC3, exclude it.

  3) Extracts the AMOUNT with several patterns, in confidence order:
       high   -> "Dotazione finanziaria/complessiva/stanziamento... -> € X"
       medium -> "X milioni" / "X mln"
       low    -> generic "€ X.XXX.XXX" or "X.XXX.XXX euro" in the text
     Handles the Italian number format (dot = thousands separator,
     comma = decimals).

  4) Explicitly checks for FNC3 in the full text (list + detail),
     with variants: "FNC3", "FNC 3", "FNC-3",
     "FNC III", "Fondo Nuove Competenze" (terza edizione/3/III).

  5) Filters and highlights only the calls with:
       importo_rilevato > SOGLIA_IMPORTO  AND  NOT FNC3
     while still keeping, for transparency, a sheet with ALL the
     calls found (including the discarded ones), because automatic
     amount extraction from free text is never perfect: doubtful
     cases must always be checked by eye.

  6) Respects robots.txt (some sites, e.g. Fondimpresa, explicitly
     require it) and caps the total number of requests so as not to
     overload the funds' sites.

KNOWN LIMITATIONS (visually verify the results)
-----------------------------------------------------------------
  - Amount extraction from free text can generate false
    positives/negatives: for this reason each row also reports the
    "confidenza" and the text fragment the amount was extracted
    from, so you can check in 5 seconds whether it is reliable.
  - Calls published only as PDF are read with pdfplumber IF the
    library is installed (pip install pdfplumber); otherwise they
    are flagged as "PDF non analizzato" and must be checked
    manually.
  - Some sites return content via JavaScript: in that case the
    page will turn out empty or without calls even though they are
    visible in the browser. This is not a bug in the script: a
    headless browser would be needed (not included here to stay
    lightweight, as per the v1 setting).

USAGE
-----------------------------------------------------------------
    pip install requests beautifulsoup4 openpyxl pdfplumber --break-system-packages
    python scraper_bandi_fondi.py
    python scraper_bandi_fondi.py --soglia 500000        (different threshold)
    python scraper_bandi_fondi.py --includi-fnc3          (do not exclude FNC3)

    On Windows, if "python" is not recognized, use "py".
"""

import io
import re
import sys
import time
import random
import argparse
from datetime import datetime
from urllib.parse import urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import pdfplumber
    PDF_DISPONIBILE = True
except ImportError:
    PDF_DISPONIBILE = False


# --------------------------------------------------------------------
# 1) CONFIGURATION
# --------------------------------------------------------------------

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "it-IT,it;q=0.9",
    "Referer": "https://www.google.com",
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

TIMEOUT = 20

# Amount threshold (in euros) above which a call is a "target".
SOGLIA_IMPORTO_DEFAULT = 800_000

# Maximum number of detail pages followed for each site
# (to avoid overloading the sites and to keep the execution time
# under control). Increase it if a fund has many active notices.
MAX_DETTAGLI_PER_SITO = 15

# Maximum cap on total HTTP requests in one run (listings +
# details + robots.txt), as a further safety net.
MAX_RICHIESTE_TOTALI = 180

PAUSA_MIN_SITI, PAUSA_MAX_SITI = 2.0, 5.0
PAUSA_MIN_DETTAGLI, PAUSA_MAX_DETTAGLI = 1.0, 2.5

# --- List of funds and the "elenco bandi" (call list) pages to start from. ---
# The pages listed below are the ones provided, with two manually verified
# corrections (see the comments in the initial docstring):
SITI_TARGET = [
    ("FonARCom", "https://www.fonarcom.it/strumenti/"),
    ("Fonditalia", "https://www.fonditalia.org/archivio-e-graduatorie/"),
    ("Fondimpresa", "https://www.fondimpresa.it/i-canali-di-finanziamento"),
    ("Fondir", "https://www.fondir.it/imprese/piani-formativi"),
    ("Fondo For.Te.", "https://www.fondoforte.it/elenco-avvisi/"),
    ("Fondo Conoscenza - Programmazione", "https://www.fondoconoscenza.it/programmazione-2026/"),
    ("Fondo Conoscenza - Avvisi Aperti", "https://www.fondoconoscenza.it/avvisi-aperti/"),
    ("Fondo FBA - Avvisi Aperti", "https://fondofba.it/avvisi-aperti/"),
    ("Fondo Professioni", "https://www.fondoprofessioni.it/avvisi-per-la-formazione/"),
    ("Fondo FAPI", "https://www.fondopmi.com/avvisi/avvisi-e-graduatorie/"),
]

# --------------------------------------------------------------------
# 2) MATCHING PATTERNS
# --------------------------------------------------------------------

# A link in a list page is a "call candidate" if its text
# contains one of these words (much more specific than v1, which
# used generic words like "risorse" or "lavoratori").
PATTERN_TITOLO_BANDO = re.compile(
    r"\bavvis[oi]\b|\bband[oi]\b|\binvit[oi]\b|\blinea\s*\d|\bedizione\b|"
    r"\bpian[oi]\s+formativ|\bvoucher\b|\bconto\s+(?:sistema|formazione)\b",
    re.IGNORECASE,
)

# Titles/menu items too generic to be considered a specific call
# (they are category links, not links to the individual notice).
NAV_GENERICI = {
    "avviso", "avvisi", "bando", "bandi", "invito", "inviti",
    "avvisi aperti", "avvisi chiusi", "avvisi attivi", "avvisi pubblicati",
    "archivio avvisi", "graduatorie", "graduatoria", "bandi di gara",
    "piani formativi", "piano formativo", "elenco avvisi",
}

# Filter B: words indicating an old/closed/ranked call
# (year threshold lowered compared to v1: a notice published in
# 2024 can stay open for years, so it must not be discarded just
# because it does not contain a year >= current year).
PAROLE_SCARTO = [
    "scaduto", "chiuso definitivamente", "concluso", "terminato",
    "graduatoria approvata", "esaurito", "revocato",
]
ANNO_MINIMO = 2023  # only discard blocks whose years are ALL earlier

# --- Amount extraction -------------------------------------------------
# High confidence: the text explicitly states the budget.
RE_DOTAZIONE = re.compile(
    r"(?:dotazione\s+finanziaria|dotazione\s+economica|dotazione\s+complessiva|"
    r"dotazione\s+dell['’]avviso|stanziamento\s+complessivo|"
    r"risorse\s+disponibili|valore\s+complessivo|importo\s+complessivo)"
    r"[^€\d]{0,20}€?\s*([\d.,]{4,})",
    re.IGNORECASE,
)
# Medium confidence: amounts expressed "in millions".
RE_MILIONI = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*(?:milioni(?:\s+di\s+euro)?|mln)\b",
    re.IGNORECASE,
)
# Low confidence: any "€ X.XXX.XXX" or "X.XXX.XXX euro" in the text.
RE_EURO_PREFISSO = re.compile(r"€\s*([\d]{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d{4,}(?:,\d{1,2})?)")
RE_EURO_SUFFISSO = re.compile(
    r"\b([\d]{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d{4,}(?:,\d{1,2})?)\s*(?:euro|eur)\b",
    re.IGNORECASE,
)

# --- FNC3 detection ----------------------------------------------------
PATTERN_FNC3 = re.compile(
    r"fnc\s*3\b|fnc[\s\-]?iii\b|"
    r"fondo\s+nuove\s+competenze\s*(?:[-–—]\s*)?(?:3\b|iii\b|terza\s+edizione)",
    re.IGNORECASE,
)

CONTENITORI = ("li", "article", "section", "div")


# --------------------------------------------------------------------
# 3) NETWORK FUNCTIONS
# --------------------------------------------------------------------

class ContatoreRichieste:
    """Keeps count of the total HTTP requests so the cap is not exceeded."""
    def __init__(self, massimo):
        self.massimo = massimo
        self.usate = 0

    def puo_procedere(self):
        return self.usate < self.massimo

    def incrementa(self):
        self.usate += 1


CONTATORE = ContatoreRichieste(MAX_RICHIESTE_TOTALI)
_ROBOTS_CACHE = {}


def crea_sessione():
    sessione = requests.Session()
    sessione.headers.update(HEADERS)
    return sessione


def permesso_da_robots(sessione, url):
    """
    Checks the site's robots.txt before downloading a page.
    If the file does not exist or is unreadable, permission is
    presumed (conservative behaviour only when the ban is explicit).
    """
    parsed = urlparse(url)
    base = f"{parsed.scheme}://{parsed.netloc}"
    if base not in _ROBOTS_CACHE:
        rp = None
        try:
            if CONTATORE.puo_procedere():
                CONTATORE.incrementa()
                risposta = sessione.get(base + "/robots.txt", timeout=10)
                if risposta.status_code == 200:
                    rp = RobotFileParser()
                    rp.parse(risposta.text.splitlines())
        except requests.exceptions.RequestException:
            rp = None
        _ROBOTS_CACHE[base] = rp
    rp = _ROBOTS_CACHE[base]
    if rp is None:
        return True
    return rp.can_fetch(HEADERS["User-Agent"], url)


def scarica_pagina(sessione, url):
    """Downloads an HTML page handling errors without interrupting the program."""
    if not CONTATORE.puo_procedere():
        return None, "Tetto richieste raggiunto"
    if not permesso_da_robots(sessione, url):
        return None, "Bloccato da robots.txt"

    try:
        CONTATORE.incrementa()
        risposta = sessione.get(url, timeout=TIMEOUT)
        risposta.raise_for_status()
        risposta.encoding = risposta.apparent_encoding
        return BeautifulSoup(risposta.text, "html.parser"), "OK"
    except requests.exceptions.HTTPError as e:
        codice = e.response.status_code if e.response is not None else "?"
        if codice == 403:
            return None, "Bloccato (403)"
        return None, f"Errore HTTP {codice}"
    except requests.exceptions.Timeout:
        return None, "Timeout"
    except requests.exceptions.ConnectionError:
        return None, "Connessione fallita"
    except requests.exceptions.RequestException as e:
        return None, f"Errore generico: {e}"


def estrai_testo_pdf(sessione, url, max_pagine=6):
    """Extracts the text of the first pages of a PDF (requires pdfplumber)."""
    if not PDF_DISPONIBILE or not CONTATORE.puo_procedere():
        return ""
    try:
        CONTATORE.incrementa()
        risposta = sessione.get(url, timeout=TIMEOUT)
        risposta.raise_for_status()
        testo = []
        with pdfplumber.open(io.BytesIO(risposta.content)) as pdf:
            for pagina in pdf.pages[:max_pagine]:
                testo.append(pagina.extract_text() or "")
        return "\n".join(testo)
    except Exception:
        return ""


def normalizza_url(href, url_base):
    if not href:
        return url_base
    href = href.strip()
    if href.startswith("http"):
        return href
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        parti = url_base.split("/")
        dominio = "/".join(parti[:3])
        return dominio + href
    return url_base.rstrip("/") + "/" + href


# --------------------------------------------------------------------
# 4) EXTRACTION: MAIN CONTENT, BLOCKS, AMOUNTS, FNC3
# --------------------------------------------------------------------

def estrai_contenuto_principale(soup):
    """
    Narrows the analysis to the main content of the page (when it can
    be identified), so the amount/FNC3 search is not diluted by the
    menu/footer noise repeated on every page of the site.
    """
    candidati = [
        soup.find("main"),
        soup.find("article"),
        soup.find("div", class_=re.compile(r"entry-content|post-content|content-area|td-post-content", re.I)),
        soup.find("div", id=re.compile(r"^content$|^main$", re.I)),
    ]
    for elemento in candidati:
        if elemento is not None and len(elemento.get_text(strip=True)) > 200:
            return elemento.get_text(" ", strip=True)
    return soup.get_text(" ", strip=True)


def trova_blocco(elemento, max_caratteri=3000):
    """Climbs up to the container (li/article/section/div) that encloses the link."""
    blocco = elemento
    nodo = elemento.parent
    while nodo is not None and nodo.name not in ("body", "html", "[document]"):
        if nodo.name in CONTENITORI:
            blocco = nodo
            break
        nodo = nodo.parent

    testo = blocco.get_text(" ", strip=True)
    if len(testo) > max_caratteri:
        genitore = elemento.parent
        testo = genitore.get_text(" ", strip=True) if genitore is not None else elemento.get_text(" ", strip=True)
    return testo


def blocco_da_scartare(testo_blocco):
    """Filter to discard evidently old/closed calls up front."""
    tl = testo_blocco.lower()
    for parola in PAROLE_SCARTO:
        if parola in tl:
            return True, f"word '{parola}'"

    anni = [int(a) for a in re.findall(r"\b((?:19|20)\d{2})\b", testo_blocco)]
    if anni and not any(a >= ANNO_MINIMO for a in anni):
        return True, f"year {max(anni)} earlier than {ANNO_MINIMO}"
    return False, ""


def _numero_italiano_a_float(testo_numero):
    """Converts '1.200.000,50' or '800000' to float, handling the IT format."""
    s = testo_numero.strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        # Comma only: treat it as a decimal separator if the last part has <=2 digits.
        parti = s.split(",")
        if len(parti[-1]) <= 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        s = s.replace(".", "") if s.count(".") > 1 else s
    try:
        return float(s)
    except ValueError:
        return None


def estrai_importi(testo):
    """
    Searches for euro amounts in the text. Returns a list of tuples
    (valore, confidenza, frammento) sortable by reliability.
    confidenza in {"alta", "media", "bassa"}.
    """
    trovati = []

    for m in RE_DOTAZIONE.finditer(testo):
        val = _numero_italiano_a_float(m.group(1))
        if val and val >= 1000:
            trovati.append((val, "alta", m.group(0)[:90]))

    for m in RE_MILIONI.finditer(testo):
        numero = m.group(1).replace(".", "").replace(",", ".")
        try:
            val = float(numero) * 1_000_000
            trovati.append((val, "media", m.group(0)[:90]))
        except ValueError:
            pass

    for pattern in (RE_EURO_PREFISSO, RE_EURO_SUFFISSO):
        for m in pattern.finditer(testo):
            val = _numero_italiano_a_float(m.group(1))
            if val and val >= 1000:
                trovati.append((val, "bassa", m.group(0)[:90]))

    return trovati


def importo_migliore(testo):
    """Picks the most reliable amount found in the text (or None)."""
    importi = estrai_importi(testo)
    if not importi:
        return None, None, ""
    priorita = {"alta": 3, "media": 2, "bassa": 1}
    importi.sort(key=lambda t: (priorita[t[1]], t[0]), reverse=True)
    return importi[0]


def e_fnc3(testo):
    return bool(PATTERN_FNC3.search(testo))


def e_candidato_valido(testo_link, href, url_base):
    """Discards navigation/category links that are too generic."""
    t = testo_link.strip().lower()
    if t in NAV_GENERICI or len(t) < 6:
        return False
    if href.startswith(("javascript:", "#", "mailto:", "tel:")):
        return False
    link = normalizza_url(href, url_base)
    if link.rstrip("/") == url_base.rstrip("/"):
        return False
    return True


def trova_candidati_bando(soup, url_base):
    """Identifies, in the list page, the links that look like individual calls."""
    candidati = []
    visti = set()
    for a in soup.find_all("a", href=True):
        testo_link = a.get_text(" ", strip=True)
        if not testo_link or not PATTERN_TITOLO_BANDO.search(testo_link):
            continue
        href = a["href"]
        if not e_candidato_valido(testo_link, href, url_base):
            continue
        link = normalizza_url(href, url_base)
        if link in visti:
            continue
        visti.add(link)
        candidati.append({
            "titolo": testo_link,
            "link": link,
            "blocco": trova_blocco(a),
        })
    return candidati


# --------------------------------------------------------------------
# 5) ANALYSIS OF A SINGLE CALL (level 2: detail page)
# --------------------------------------------------------------------

def analizza_bando(sessione, candidato, soglia, escludi_fnc3):
    """
    Opens (when possible) the call's detail page, combines the list
    text with the detail text, and extracts the amount and FNC3
    presence. Returns the complete result dictionary.
    """
    link = candidato["link"]
    testo_dettaglio = ""
    stato_dettaglio = "Non seguito"

    if link.lower().endswith(".pdf"):
        if PDF_DISPONIBILE:
            testo_dettaglio = estrai_testo_pdf(sessione, link)
            stato_dettaglio = "OK (PDF)" if testo_dettaglio else "PDF non leggibile"
        else:
            stato_dettaglio = "PDF non analizzato (installa pdfplumber)"
    else:
        soup_dettaglio, stato_dettaglio = scarica_pagina(sessione, link)
        if soup_dettaglio is not None:
            testo_dettaglio = estrai_contenuto_principale(soup_dettaglio)

    testo_completo = candidato["blocco"] + " \n " + testo_dettaglio
    valore, confidenza, frammento = importo_migliore(testo_completo)
    fnc3 = e_fnc3(testo_completo)

    is_target = (
        valore is not None
        and valore > soglia
        and (not fnc3 or not escludi_fnc3)
    )

    return {
        "titolo": candidato["titolo"],
        "link": link,
        "stato_dettaglio": stato_dettaglio,
        "importo": valore,
        "confidenza": confidenza or "",
        "frammento_importo": frammento,
        "fnc3": fnc3,
        "target": is_target,
    }


# --------------------------------------------------------------------
# 6) PER-SITE ORCHESTRATION
# --------------------------------------------------------------------

def scansiona_sito(sessione, nome_sito, url, soglia, escludi_fnc3):
    print(f"\n--> {nome_sito}\n    {url}")

    soup, stato = scarica_pagina(sessione, url)
    if soup is None:
        print(f"    [!] {stato}")
        return {"nome": nome_sito, "url": url, "stato": stato, "bandi": []}

    candidati = trova_candidati_bando(soup, url)
    print(f"    Found {len(candidati)} call candidates in the list.")

    bandi = []
    seguiti = 0
    for candidato in candidati:
        scarta, motivo = blocco_da_scartare(candidato["blocco"])
        if scarta:
            print(f"        [discarded up front: {motivo}] {candidato['titolo'][:60]}")
            continue

        if seguiti >= MAX_DETTAGLI_PER_SITO or not CONTATORE.puo_procedere():
            print(f"        [not expanded: cap reached] {candidato['titolo'][:60]}")
            continue

        risultato = analizza_bando(sessione, candidato, soglia, escludi_fnc3)
        bandi.append(risultato)
        seguiti += 1

        bandiera = "TARGET" if risultato["target"] else ("FNC3" if risultato["fnc3"] else "-")
        importo_fmt = f"€ {risultato['importo']:,.0f}".replace(",", ".") if risultato["importo"] else "n.d."
        print(f"        [{bandiera:7}] {candidato['titolo'][:55]:55} amount={importo_fmt}")

        time.sleep(random.uniform(PAUSA_MIN_DETTAGLI, PAUSA_MAX_DETTAGLI))

    return {"nome": nome_sito, "url": url, "stato": stato, "bandi": bandi}


# --------------------------------------------------------------------
# 7) EXCEL EXPORT (3 sheets: Riepilogo, Tutti i Bandi, Target)
# --------------------------------------------------------------------

def build_excel(scansioni, soglia, percorso="report_bandi_fondi.xlsx"):
    HDR_FILL = PatternFill("solid", fgColor="1F4E79")
    HDR_FONT = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
    BORDER = Border(left=Side(style="thin"), right=Side(style="thin"),
                     top=Side(style="thin"), bottom=Side(style="thin"))
    ALT_FILL = PatternFill("solid", fgColor="D6E4F0")
    TARGET_FILL = PatternFill("solid", fgColor="C6EFCE")
    CELL_ALIGN = Alignment(vertical="top", wrap_text=True)
    LINK_FONT = Font(name="Calibri", size=10, color="1F4E79", underline="single")
    oggi = datetime.now().strftime("%d/%m/%Y %H:%M")

    def intesta(ws, intestazioni):
        for col, titolo in enumerate(intestazioni, start=1):
            c = ws.cell(1, col, value=titolo)
            c.font, c.fill, c.alignment, c.border = HDR_FONT, HDR_FILL, HDR_ALIGN, BORDER
        ws.row_dimensions[1].height = 28
        ws.freeze_panes = "A2"

    def imposta_larghezze(ws, larghezze):
        for lettera, larghezza in larghezze.items():
            ws.column_dimensions[lettera].width = larghezza

    wb = Workbook()

    # ---------------- Sheet 1: Riepilogo ----------------
    ws1 = wb.active
    ws1.title = "Riepilogo"
    intesta(ws1, ["Fondo", "URL elenco", "Stato", "Bandi analizzati", "Target trovati"])
    for i, s in enumerate(scansioni, start=2):
        n_target = sum(1 for b in s["bandi"] if b["target"])
        valori = [s["nome"], s["url"], s["stato"], len(s["bandi"]), n_target]
        for col, val in enumerate(valori, start=1):
            c = ws1.cell(i, col, value=val)
            c.border, c.alignment = BORDER, CELL_ALIGN
            c.font = Font(name="Calibri", size=10, bold=(col == 1))
            if col == 2 and s["url"]:
                c.hyperlink, c.font = s["url"], LINK_FONT
            if col == 5 and n_target > 0:
                c.fill = TARGET_FILL
            elif i % 2 == 0:
                c.fill = ALT_FILL
    imposta_larghezze(ws1, {"A": 30, "B": 48, "C": 22, "D": 16, "E": 16})

    nota_riga = len(scansioni) + 3
    ws1.cell(nota_riga, 1, value=f"Scansione del {oggi}  |  Soglia importo: € {soglia:,.0f}".replace(",", "."))
    ws1.cell(nota_riga, 1).font = Font(name="Calibri", italic=True, size=9, color="666666")

    # ---------------- Sheet 2: Tutti i Bandi ----------------
    ws2 = wb.create_sheet("Tutti i Bandi")
    intesta(ws2, ["Fondo", "Titolo", "Importo (€)", "Confidenza", "Frammento importo",
                  "FNC3", "Target", "Stato dettaglio", "Link"])
    riga = 2
    for s in scansioni:
        for b in s["bandi"]:
            valori = [
                s["nome"], b["titolo"],
                round(b["importo"]) if b["importo"] else None,
                b["confidenza"], b["frammento_importo"],
                "Sì" if b["fnc3"] else "No",
                "Sì" if b["target"] else "No",
                b["stato_dettaglio"], b["link"],
            ]
            for col, val in enumerate(valori, start=1):
                c = ws2.cell(riga, col, value=val)
                c.border, c.alignment = BORDER, CELL_ALIGN
                c.font = Font(name="Calibri", size=10, bold=(col == 1))
                if col == 9 and b["link"]:
                    c.hyperlink, c.font = b["link"], LINK_FONT
                if col == 3 and val:
                    c.number_format = "#,##0"
                if b["target"]:
                    c.fill = TARGET_FILL
                elif riga % 2 == 0:
                    c.fill = ALT_FILL
            riga += 1
    if riga == 2:
        ws2.cell(2, 1, value="Nessun bando analizzato.")
    imposta_larghezze(ws2, {"A": 24, "B": 42, "C": 16, "D": 12, "E": 42,
                            "F": 8, "G": 9, "H": 18, "I": 46})

    # ---------------- Sheet 3: Target (>threshold, non-FNC3) ----------------
    ws3 = wb.create_sheet(f"Target sopra {int(soglia)//1000}k")
    intesta(ws3, ["Fondo", "Titolo", "Importo (€)", "Confidenza", "Link"])
    target = [
        (s["nome"], b) for s in scansioni for b in s["bandi"] if b["target"]
    ]
    target.sort(key=lambda t: t[1]["importo"], reverse=True)
    riga = 2
    for nome_fondo, b in target:
        valori = [nome_fondo, b["titolo"], round(b["importo"]), b["confidenza"], b["link"]]
        for col, val in enumerate(valori, start=1):
            c = ws3.cell(riga, col, value=val)
            c.border, c.alignment, c.fill = BORDER, CELL_ALIGN, TARGET_FILL
            c.font = Font(name="Calibri", size=11, bold=(col in (1, 3)))
            if col == 3:
                c.number_format = "#,##0"
            if col == 5 and val:
                c.hyperlink, c.font = val, LINK_FONT
        riga += 1
    if riga == 2:
        ws3.cell(2, 1, value="Nessun bando sopra soglia, non-FNC3, trovato in questa scansione.")
        ws3.cell(2, 1).font = Font(name="Calibri", italic=True, size=10, color="666666")
    imposta_larghezze(ws3, {"A": 26, "B": 46, "C": 16, "D": 12, "E": 50})

    wb.save(percorso)
    print(f"\nExcel saved to: {percorso}")
    return len(target)


def salva_report_txt(scansioni, soglia, percorso="report_bandi.txt"):
    oggi = datetime.now().strftime("%d/%m/%Y %H:%M")
    target = sorted(
        ((s["nome"], b) for s in scansioni for b in s["bandi"] if b["target"]),
        key=lambda t: t[1]["importo"], reverse=True,
    )
    with open(percorso, "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n")
        f.write("REPORT BANDI - Fondi Interprofessionali\n")
        f.write(f"Scansione del {oggi}  |  Soglia: € {soglia:,.0f}\n".replace(",", "."))
        f.write("=" * 70 + "\n\n")

        f.write(f"BANDI TARGET (> soglia, non-FNC3): {len(target)}\n")
        f.write("-" * 70 + "\n")
        if not target:
            f.write("Nessuno trovato in questa scansione.\n\n")
        for nome_fondo, b in target:
            f.write(f"[{nome_fondo}] {b['titolo']}\n")
            f.write(f"  Importo: € {b['importo']:,.0f}  (confidenza: {b['confidenza']})\n".replace(",", "."))
            f.write(f"  {b['link']}\n\n")

        f.write("\n" + "=" * 70 + "\n")
        f.write("DETTAGLIO PER FONDO (tutti i bandi analizzati)\n")
        f.write("=" * 70 + "\n\n")
        for s in scansioni:
            f.write("-" * 70 + "\n")
            f.write(f"{s['nome']}  [{s['stato']}]\n{s['url']}\n")
            f.write("-" * 70 + "\n")
            if not s["bandi"]:
                f.write("  (nessun bando analizzato)\n\n")
                continue
            for b in s["bandi"]:
                importo_fmt = f"€ {b['importo']:,.0f}".replace(",", ".") if b["importo"] else "n.d."
                fnc3_fmt = "FNC3" if b["fnc3"] else "non-FNC3"
                f.write(f"  - {b['titolo']}  |  {importo_fmt}  |  {fnc3_fmt}\n")
                f.write(f"    {b['link']}\n")
            f.write("\n")
    print(f"Text report saved to: {percorso}")


# --------------------------------------------------------------------
# 8) MAIN PROGRAM
# --------------------------------------------------------------------

def analizza_argomenti():
    parser = argparse.ArgumentParser(
        description="Searches Interprofessional Funds calls that are non-FNC3 above an amount threshold."
    )
    parser.add_argument("--soglia", type=float, default=SOGLIA_IMPORTO_DEFAULT,
                         help=f"Minimum threshold in euros (default: {SOGLIA_IMPORTO_DEFAULT})")
    parser.add_argument("--includi-fnc3", action="store_true",
                         help="If set, does NOT exclude FNC3-related calls from the target sheet")
    return parser.parse_args()


def main():
    args = analizza_argomenti()
    soglia = args.soglia
    escludi_fnc3 = not args.includi_fnc3

    print("=" * 70)
    print(" FUNDING CALL SCRAPER - Interprofessional Funds")
    print(f" Amount threshold: € {soglia:,.0f}  |  Excludes FNC3: {escludi_fnc3}".replace(",", "."))
    print("=" * 70)

    if not PDF_DISPONIBILE:
        print("[i] pdfplumber not installed: calls published only as PDF "
              "will not be read (pip install pdfplumber --break-system-packages).")

    sessione = crea_sessione()
    scansioni = []

    for indice, (nome_sito, url) in enumerate(SITI_TARGET, start=1):
        if not CONTATORE.puo_procedere():
            print(f"\n[!] Cap of {MAX_RICHIESTE_TOTALI} requests reached: scan interrupted.")
            break
        print(f"\n[{indice}/{len(SITI_TARGET)}]", end="")
        scansioni.append(scansiona_sito(sessione, nome_sito, url, soglia, escludi_fnc3))

        if indice < len(SITI_TARGET):
            time.sleep(random.uniform(PAUSA_MIN_SITI, PAUSA_MAX_SITI))

    print("\n" + "=" * 70)
    n_target = build_excel(scansioni, soglia)
    salva_report_txt(scansioni, soglia)

    print("=" * 70)
    if n_target:
        print(f"FOUND {n_target} call(s) above € {soglia:,.0f} and not linked to FNC3.".replace(",", "."))
        print("See the 'Target...' sheet in the Excel for details, and ALWAYS VERIFY")
        print("by eye the text fragment the amount was extracted from.")
    else:
        print("No target calls found with the current parameters.")
        print("Suggestions: increase MAX_DETTAGLI_PER_SITO, check the")
        print("'Tutti i Bandi' sheet for cases close to the threshold, or add more")
        print("'avvisi aperti' pages for funds not yet covered in SITI_TARGET.")
    print("=" * 70)


if __name__ == "__main__":
    main()
