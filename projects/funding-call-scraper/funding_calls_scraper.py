"""
====================================================================
 SCRAPER BANDI - Fondi Interprofessionali (v2)
 Obiettivo: trovare bandi NON-FNC3 con dotazione > soglia (default 800.000 EUR)
====================================================================

COSA CAMBIA RISPETTO ALLA V1
-----------------------------------------------------------------
La v1 scansionava le HOMEPAGE dei fondi con un matching per parole
troppo generico ("risorse", "lavoratori", "dipendenti", ecc.), che
produceva quasi solo rumore (voci di menu, frasi di presentazione)
e non estraeva mai un dato fondamentale: l'IMPORTO del bando.

Questa v2:

  1) Punta direttamente alle pagine "elenco bandi/avvisi" indicate,
     con due correzioni emerse da una verifica manuale dei siti
     (vedi SITI_TARGET piu' sotto per il dettaglio):
       - Fondo FBA: "Archivio Avvisi" e' lo storico CHIUSO (ultimo
         bando 2020). La pagina con i bandi davvero aperti e'
         "Avvisi Aperti" -> sostituita.
       - Fondo Conoscenza: alla pagina "Programmazione 2026" e'
         stata affiancata "Avvisi Aperti" (Conto Sistema), che
         elenca i singoli avvisi attivi con link diretto.

  2) Fa uno scraping A DUE LIVELLI:
       livello 1: pagina-elenco  -> individua i singoli bandi (link)
       livello 2: pagina-bando   -> apre il dettaglio e ne estrae
                                    testo, importo e presenza FNC3
     Questo perche', nei siti verificati, l'elenco mostra solo
     titolo/data, mentre l'importo ("Dotazione finanziaria") e
     l'eventuale legame con FNC3 compaiono SOLO nel dettaglio.
     Esempio reale (Fondo FBA):
       - in elenco: "Avviso Competenze per l'innovazione" (nessun
         accenno a FNC3)
       - nel dettaglio: "Opportunita' di finanziamento con il Fondo
         Nuove Competenze - Terza Edizione" -> e' FNC3, va escluso.

  3) Estrae l'IMPORTO con piu' pattern, in ordine di confidenza:
       alta  -> "Dotazione finanziaria/complessiva/stanziamento... -> € X"
       media -> "X milioni" / "X mln"
       bassa -> "€ X.XXX.XXX" o "X.XXX.XXX euro" generico nel testo
     Gestisce il formato italiano dei numeri (punto = migliaia,
     virgola = decimali).

  4) Verifica esplicitamente la presenza di FNC3 nel testo completo
     (elenco + dettaglio), con varianti: "FNC3", "FNC 3", "FNC-3",
     "FNC III", "Fondo Nuove Competenze" (terza edizione/3/III).

  5) Filtra e mette in evidenza solo i bandi con:
       importo_rilevato > SOGLIA_IMPORTO  E  NON FNC3
     mantenendo comunque, per trasparenza, un foglio con TUTTI i
     bandi trovati (anche quelli scartati), perche' l'estrazione
     automatica di un importo da testo libero non e' mai perfetta:
     vanno sempre controllati a vista i casi dubbi.

  6) Rispetta robots.txt (alcuni siti, es. Fondimpresa, lo richiedono
     esplicitamente) e mette un tetto al numero di richieste totali
     per non sovraccaricare i siti dei fondi.

LIMITI NOTI (da verificare a vista sui risultati)
-----------------------------------------------------------------
  - L'estrazione di importi da testo libero puo' generare falsi
    positivi/negativi: per questo ogni riga riporta anche la
    "confidenza" e il frammento di testo da cui e' stato estratto
    l'importo, cosi' puoi verificare in 5 secondi se e' attendibile.
  - I bandi pubblicati solo in PDF vengono letti con pdfplumber SE
    la libreria e' installata (pip install pdfplumber); altrimenti
    vengono segnalati come "PDF non analizzato" e vanno controllati
    a mano.
  - Alcuni siti restituiscono contenuto via JavaScript: in quel caso
    la pagina risultera' vuota o priva di bandi anche se nel browser
    si vedono. Non e' un bug dello script: serve un browser headless
    (non incluso qui per restare leggero, come da impostazione v1).

USO
-----------------------------------------------------------------
    pip install requests beautifulsoup4 openpyxl pdfplumber --break-system-packages
    python scraper_bandi_fondi.py
    python scraper_bandi_fondi.py --soglia 500000        (soglia diversa)
    python scraper_bandi_fondi.py --includi-fnc3          (non escludere FNC3)

    Su Windows, se "python" non e' riconosciuto, usa "py".
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
# 1) CONFIGURAZIONE
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

# Soglia di importo (in euro) sopra la quale un bando e' "target".
SOGLIA_IMPORTO_DEFAULT = 800_000

# Numero massimo di pagine di dettaglio seguite per ciascun sito
# (per non sovraccaricare i siti e tenere il tempo di esecuzione
# sotto controllo). Aumentalo se un fondo ha tanti avvisi attivi.
MAX_DETTAGLI_PER_SITO = 15

# Tetto massimo di richieste HTTP totali in un'esecuzione (listing +
# dettagli + robots.txt), come ulteriore rete di sicurezza.
MAX_RICHIESTE_TOTALI = 180

PAUSA_MIN_SITI, PAUSA_MAX_SITI = 2.0, 5.0
PAUSA_MIN_DETTAGLI, PAUSA_MAX_DETTAGLI = 1.0, 2.5

# --- Elenco dei fondi e delle pagine "elenco bandi" da cui partire. ---
# Le pagine indicate sotto sono quelle fornite, con due correzioni
# verificate manualmente (vedi commenti nel docstring iniziale):
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
# 2) PATTERN DI RICONOSCIMENTO
# --------------------------------------------------------------------

# Un link in pagina-elenco e' un "candidato bando" se il suo testo
# contiene una di queste parole (molto piu' specifico della v1, che
# usava parole generiche come "risorse" o "lavoratori").
PATTERN_TITOLO_BANDO = re.compile(
    r"\bavvis[oi]\b|\bband[oi]\b|\binvit[oi]\b|\blinea\s*\d|\bedizione\b|"
    r"\bpian[oi]\s+formativ|\bvoucher\b|\bconto\s+(?:sistema|formazione)\b",
    re.IGNORECASE,
)

# Titoli/voci di menu troppo generici da NON considerare un bando
# specifico (sono link di categoria, non al singolo avviso).
NAV_GENERICI = {
    "avviso", "avvisi", "bando", "bandi", "invito", "inviti",
    "avvisi aperti", "avvisi chiusi", "avvisi attivi", "avvisi pubblicati",
    "archivio avvisi", "graduatorie", "graduatoria", "bandi di gara",
    "piani formativi", "piano formativo", "elenco avvisi",
}

# Filtro B: parole che indicano un bando vecchio/chiuso/in graduatoria
# (soglia anno abbassata rispetto alla v1: un Avviso pubblicato nel
# 2024 puo' restare aperto per anni, quindi non va scartato solo
# perche' non contiene un anno >= anno corrente).
PAROLE_SCARTO = [
    "scaduto", "chiuso definitivamente", "concluso", "terminato",
    "graduatoria approvata", "esaurito", "revocato",
]
ANNO_MINIMO = 2023  # scarta solo blocchi i cui anni sono TUTTI precedenti

# --- Estrazione importi -------------------------------------------------
# Alta confidenza: il testo dichiara esplicitamente la dotazione.
RE_DOTAZIONE = re.compile(
    r"(?:dotazione\s+finanziaria|dotazione\s+economica|dotazione\s+complessiva|"
    r"dotazione\s+dell['’]avviso|stanziamento\s+complessivo|"
    r"risorse\s+disponibili|valore\s+complessivo|importo\s+complessivo)"
    r"[^€\d]{0,20}€?\s*([\d.,]{4,})",
    re.IGNORECASE,
)
# Media confidenza: importi espressi "a milioni".
RE_MILIONI = re.compile(
    r"(\d+(?:[.,]\d+)?)\s*(?:milioni(?:\s+di\s+euro)?|mln)\b",
    re.IGNORECASE,
)
# Bassa confidenza: qualunque "€ X.XXX.XXX" o "X.XXX.XXX euro" nel testo.
RE_EURO_PREFISSO = re.compile(r"€\s*([\d]{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d{4,}(?:,\d{1,2})?)")
RE_EURO_SUFFISSO = re.compile(
    r"\b([\d]{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d{4,}(?:,\d{1,2})?)\s*(?:euro|eur)\b",
    re.IGNORECASE,
)

# --- Rilevamento FNC3 ----------------------------------------------------
PATTERN_FNC3 = re.compile(
    r"fnc\s*3\b|fnc[\s\-]?iii\b|"
    r"fondo\s+nuove\s+competenze\s*(?:[-–—]\s*)?(?:3\b|iii\b|terza\s+edizione)",
    re.IGNORECASE,
)

CONTENITORI = ("li", "article", "section", "div")


# --------------------------------------------------------------------
# 3) FUNZIONI DI RETE
# --------------------------------------------------------------------

class ContatoreRichieste:
    """Tiene il conto delle richieste HTTP totali per non eccedere il tetto."""
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
    Verifica il robots.txt del sito prima di scaricare una pagina.
    Se il file non esiste o non e' leggibile, si presume consentito
    (comportamento conservativo solo quando il divieto e' esplicito).
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
    """Scarica una pagina HTML gestendo errori senza interrompere il programma."""
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
    """Estrae il testo delle prime pagine di un PDF (richiede pdfplumber)."""
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
# 4) ESTRAZIONE: CONTENUTO PRINCIPALE, BLOCCHI, IMPORTI, FNC3
# --------------------------------------------------------------------

def estrai_contenuto_principale(soup):
    """
    Restringe l'analisi al contenuto principale della pagina (quando
    individuabile), per non annacquare la ricerca di importo/FNC3 nel
    rumore di menu/footer ripetuti su ogni pagina del sito.
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
    """Risale al contenitore (li/article/section/div) che racchiude il link."""
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
    """Filtro per scartare a priori bandi evidentemente vecchi/chiusi."""
    tl = testo_blocco.lower()
    for parola in PAROLE_SCARTO:
        if parola in tl:
            return True, f"parola '{parola}'"

    anni = [int(a) for a in re.findall(r"\b((?:19|20)\d{2})\b", testo_blocco)]
    if anni and not any(a >= ANNO_MINIMO for a in anni):
        return True, f"anno {max(anni)} antecedente al {ANNO_MINIMO}"
    return False, ""


def _numero_italiano_a_float(testo_numero):
    """Converte '1.200.000,50' o '800000' in float, gestendo il formato IT."""
    s = testo_numero.strip()
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        # Solo virgola: trattala come decimale se l'ultima parte ha <=2 cifre.
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
    Cerca importi in euro nel testo. Ritorna una lista di tuple
    (valore, confidenza, frammento) ordinabile per attendibilita'.
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
    """Sceglie l'importo piu' attendibile trovato nel testo (o None)."""
    importi = estrai_importi(testo)
    if not importi:
        return None, None, ""
    priorita = {"alta": 3, "media": 2, "bassa": 1}
    importi.sort(key=lambda t: (priorita[t[1]], t[0]), reverse=True)
    return importi[0]


def e_fnc3(testo):
    return bool(PATTERN_FNC3.search(testo))


def e_candidato_valido(testo_link, href, url_base):
    """Scarta link di navigazione/categoria troppo generici."""
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
    """Individua, nella pagina-elenco, i link che sembrano singoli bandi."""
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
# 5) ANALISI DI UN SINGOLO BANDO (livello 2: pagina di dettaglio)
# --------------------------------------------------------------------

def analizza_bando(sessione, candidato, soglia, escludi_fnc3):
    """
    Apre (se possibile) la pagina di dettaglio del bando, combina il
    testo dell'elenco con quello del dettaglio, ed estrae importo e
    presenza FNC3. Ritorna il dizionario completo del risultato.
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
# 6) ORCHESTRAZIONE PER SITO
# --------------------------------------------------------------------

def scansiona_sito(sessione, nome_sito, url, soglia, escludi_fnc3):
    print(f"\n--> {nome_sito}\n    {url}")

    soup, stato = scarica_pagina(sessione, url)
    if soup is None:
        print(f"    [!] {stato}")
        return {"nome": nome_sito, "url": url, "stato": stato, "bandi": []}

    candidati = trova_candidati_bando(soup, url)
    print(f"    Trovati {len(candidati)} candidati-bando nell'elenco.")

    bandi = []
    seguiti = 0
    for candidato in candidati:
        scarta, motivo = blocco_da_scartare(candidato["blocco"])
        if scarta:
            print(f"        [scartato a priori: {motivo}] {candidato['titolo'][:60]}")
            continue

        if seguiti >= MAX_DETTAGLI_PER_SITO or not CONTATORE.puo_procedere():
            print(f"        [non approfondito: tetto raggiunto] {candidato['titolo'][:60]}")
            continue

        risultato = analizza_bando(sessione, candidato, soglia, escludi_fnc3)
        bandi.append(risultato)
        seguiti += 1

        bandiera = "TARGET" if risultato["target"] else ("FNC3" if risultato["fnc3"] else "-")
        importo_fmt = f"€ {risultato['importo']:,.0f}".replace(",", ".") if risultato["importo"] else "n.d."
        print(f"        [{bandiera:7}] {candidato['titolo'][:55]:55} importo={importo_fmt}")

        time.sleep(random.uniform(PAUSA_MIN_DETTAGLI, PAUSA_MAX_DETTAGLI))

    return {"nome": nome_sito, "url": url, "stato": stato, "bandi": bandi}


# --------------------------------------------------------------------
# 7) EXPORT EXCEL (3 fogli: Riepilogo, Tutti i Bandi, Target)
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

    # ---------------- Foglio 1: Riepilogo ----------------
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

    # ---------------- Foglio 2: Tutti i Bandi ----------------
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

    # ---------------- Foglio 3: Target (>soglia, non-FNC3) ----------------
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
    print(f"\nExcel salvato in: {percorso}")
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
    print(f"Report testuale salvato in: {percorso}")


# --------------------------------------------------------------------
# 8) PROGRAMMA PRINCIPALE
# --------------------------------------------------------------------

def analizza_argomenti():
    parser = argparse.ArgumentParser(
        description="Cerca bandi dei Fondi Interprofessionali non-FNC3 sopra una soglia di importo."
    )
    parser.add_argument("--soglia", type=float, default=SOGLIA_IMPORTO_DEFAULT,
                         help=f"Soglia minima in euro (default: {SOGLIA_IMPORTO_DEFAULT})")
    parser.add_argument("--includi-fnc3", action="store_true",
                         help="Se presente, NON esclude i bandi legati a FNC3 dal foglio target")
    return parser.parse_args()


def main():
    args = analizza_argomenti()
    soglia = args.soglia
    escludi_fnc3 = not args.includi_fnc3

    print("=" * 70)
    print(" SCRAPER BANDI - Fondi Interprofessionali")
    print(f" Soglia importo: € {soglia:,.0f}  |  Esclude FNC3: {escludi_fnc3}".replace(",", "."))
    print("=" * 70)

    if not PDF_DISPONIBILE:
        print("[i] pdfplumber non installato: i bandi pubblicati solo in PDF "
              "non verranno letti (pip install pdfplumber --break-system-packages).")

    sessione = crea_sessione()
    scansioni = []

    for indice, (nome_sito, url) in enumerate(SITI_TARGET, start=1):
        if not CONTATORE.puo_procedere():
            print(f"\n[!] Tetto di {MAX_RICHIESTE_TOTALI} richieste raggiunto: scansione interrotta.")
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
        print(f"TROVATI {n_target} bando/i sopra € {soglia:,.0f} e non legati a FNC3.".replace(",", "."))
        print("Vedi il foglio 'Target...' nell'Excel per i dettagli, e VERIFICA SEMPRE")
        print("a vista il frammento di testo da cui e' stato estratto l'importo.")
    else:
        print("Nessun bando target trovato con i parametri attuali.")
        print("Suggerimenti: aumenta MAX_DETTAGLI_PER_SITO, controlla il foglio")
        print("'Tutti i Bandi' per casi vicini alla soglia, o aggiungi altre pagine")
        print("'avvisi aperti' ai fondi non ancora coperti in SITI_TARGET.")
    print("=" * 70)


if __name__ == "__main__":
    main()
