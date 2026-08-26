"""
adhesion_letter_compiler.py
===========================
Compila automaticamente la "Lettera Adesione.docx" partendere dai dati di un
file Excel ("Dati da inserire.xlsx").

LOGICA DI ACCOPPIAMENTO (posizionale):
- La prima persona (riga 3 dell'Excel) si accoppia con la prima azienda
  (riga 20), la seconda persona (riga 4) con la seconda azienda (riga 21),
  e cosi' via ("primo col primo, secondo col secondo").
- Se ci sono piu' aziende che persone, le lettere senza persona avranno i
  box della persona fisica vuoti (e viceversa).

Per ogni coppia genera una copia del modello Word compilata.

CENTRATURA:
- I valori vengono centrati orizzontalmente dentro i box del modulo,
  usando tab stop di tipo 'center' posizionati al centro esatto di ciascun
  box (posizioni ricavate dalla geometria del PDF originale).

USO:
    py adhesion_letter_compiler.py

I nomi dei file sono configurabili nelle costanti in basso. Il modello
originale non viene MAI modificato.
"""

import sys
import re
import copy
from datetime import datetime, date
from pathlib import Path

import openpyxl
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================ CONFIGURAZIONE ============================
CARTELLA = Path(__file__).resolve().parent
MODELLO_DOCX = CARTELLA / "Lettera Adesione.docx"
DATI_XLSX = CARTELLA / "Dati da inserire.xlsx"
PREFISSO_OUTPUT = "Lettera Adesione_"
FOGLIO_DATI = "Foglio1"

# Indici (0-based) dei paragrafi nel modello Word che contengono i campi.
P_IL_SOTTOSCRITTO = 18
P_NATO_A          = 20
P_RESIDENTE_IN    = 22
P_CAP_CF          = 24
P_DENOMINAZIONE   = 28
P_PIVA_CF         = 30
P_RAGIONE_SOC     = 32
P_TIPOLOGIA_ENTE  = 34
P_CON_SEDE_PROV   = 37
P_VIA_CAP         = 39

# --- Posizioni dei tab stop (twips) ---
# Ricavate dalla geometria del PDF originale.
#
# Campi SINGOLI (1 box largo): un solo tab 'center' al centro del box.
#   es. Denominazione: center @ centro del box [138-542].
#
# Campi DOPPI (2 box affiancati con due etichette):
#   3 tab stop nell'ordine:
#     [center@centro_box_sx , left@pos_etichetta_dx , center@centro_box_dx]
#   cosi' l'etichetta di destra ("il", "via", "C.F.", "Prov.", "CAP") mantiene
#   la posizione originale (tab left, ~x=333pt) e NON viene decentrata, mentre
#   i valori vengono centrati nei rispettivi box.
#
# Sintassi tupla: (tipo, posizione_twips). Lista di tuple per paragrafo.
def _c(pos):  return ('center', pos)   # tab center
def _l(pos):  return ('left', pos)     # tab left (etichetta dx originale)

TABS = {
    P_IL_SOTTOSCRITTO: [_c(5631)],
    # 'Nato a/il': 2 tab prima di 'il'. Sequenza tab consumati:
    #   TAB1 -> center@box_sx (valore 'Parma (PR)' centrato)
    #   TAB2 -> left@etichetta_dx (fermo a sinistra, NON sposta oltre)
    #   TAB3 (dopo 'il') -> center@box_dx (valore data centrato)
    # Per far cio' i tab stop devono essere: [center@box_sx, left@box_sx, left@etichetta_dx, center@box_dx]
    # Il 2o (left@box_sx) ferma il 2o tab subito dopo il valore senza avanzare.
    P_NATO_A:        [_c(3391), _l(5528), _c(7892)],
    P_RESIDENTE_IN:  [_c(3400), _l(5528), _c(7898)],
    P_CAP_CF:        [_c(3414), _l(5528), _c(7880)],
    P_DENOMINAZIONE: [_c(5669)],
    P_PIVA_CF:       [_c(3443), _l(5528), _c(7899)],
    P_RAGIONE_SOC:   [_c(5631)],
    P_TIPOLOGIA_ENTE:[_c(5642)],
    P_CON_SEDE_PROV: [_c(3414), _l(5528), _c(7871)],
    P_VIA_CAP:       [_c(3427), _l(5528), _c(7880)],
}

# Righe Excel
RIGA_PRIMA_PERSONA = 3      # persona fisica a partire dalla riga 3
RIGA_PRIMA_AZIENDA = 20     # aziende a partire dalla riga 20
COL_PERSONA = {             # colonna -> chiave dict
    'nome': 2, 'nato_a': 3, 'data_nascita': 4, 'residente': 5,
    'via': 6, 'cap': 7, 'cf': 8,
}
COL_AZIENDA = {
    'denominazione': 2, 'piva': 3, 'cf': 4, 'ragione_sociale': 5,
    'tipologia': 7, 'via': 8, 'prov': 10, 'cap': 11,
}
# ========================================================================


# ----------------------- Utility di lettura Excel -----------------------
def normalizza(valore):
    if valore is None:
        return ""
    if isinstance(valore, (datetime, date)):
        return valore.strftime("%d/%m/%Y")
    if isinstance(valore, float) and valore.is_integer():
        return str(int(valore))
    return str(valore).strip()


def come_piva_cf(valore):
    """P.IVA / C.F. lette come numero: restituisce stringa senza decimali."""
    if valore is None:
        return ""
    if isinstance(valore, (int, float)):
        return str(int(valore))
    return str(valore).strip()


def come_provincia(valore):
    """Estrae la sigla di provincia da valori tipo 'Roma | RM' o 'ROMA (RM)'."""
    s = normalizza(valore)
    if not s:
        return ""
    m = re.search(r"\(([A-Z]{2})\)", s)
    if m:
        return m.group(1)
    m = re.search(r"[|/]\s*([A-Z]{2})\b", s)
    if m:
        return m.group(1)
    if len(s) == 2 and s.isalpha():
        return s.upper()
    return s


def leggi_persone(ws):
    persone = []
    r = RIGA_PRIMA_PERSONA
    while True:
        nome = ws.cell(row=r, column=COL_PERSONA['nome']).value
        # fermiamoci se la riga e' completamente vuota E siamo oltre la prima
        if nome is None and r > RIGA_PRIMA_PERSONA:
            break
        if nome is not None:
            persone.append({
                'nome':         normalizza(ws.cell(row=r, column=COL_PERSONA['nome']).value),
                'nato_a':       normalizza(ws.cell(row=r, column=COL_PERSONA['nato_a']).value),
                'data_nascita': normalizza(ws.cell(row=r, column=COL_PERSONA['data_nascita']).value),
                'residente':    normalizza(ws.cell(row=r, column=COL_PERSONA['residente']).value),
                'via':          normalizza(ws.cell(row=r, column=COL_PERSONA['via']).value),
                'cap':          normalizza(ws.cell(row=r, column=COL_PERSONA['cap']).value),
                'cf':           normalizza(ws.cell(row=r, column=COL_PERSONA['cf']).value),
            })
        else:
            # riga vuota tra le persone: segnaposto per accoppiamento posizionale
            persone.append(None)
        # ci fermiamo quando incontriamo la riga 5 (testo fisso "In qualita...")
        if r + 1 == 5:
            break
        r += 1
        if r > ws.max_row:
            break
    return persone


def leggi_aziende(ws):
    aziende = []
    r = RIGA_PRIMA_AZIENDA
    while r <= ws.max_row:
        denom = ws.cell(row=r, column=COL_AZIENDA['denominazione']).value
        if denom is not None:
            aziende.append({
                'denominazione':   normalizza(denom),
                'piva':            come_piva_cf(ws.cell(row=r, column=COL_AZIENDA['piva']).value),
                'cf':              come_piva_cf(ws.cell(row=r, column=COL_AZIENDA['cf']).value),
                'ragione_sociale': normalizza(ws.cell(row=r, column=COL_AZIENDA['ragione_sociale']).value),
                'tipologia':       normalizza(ws.cell(row=r, column=COL_AZIENDA['tipologia']).value),
                'via':             normalizza(ws.cell(row=r, column=COL_AZIENDA['via']).value),
                'prov':            come_provincia(ws.cell(row=r, column=COL_AZIENDA['prov']).value),
                'cap':             normalizza(ws.cell(row=r, column=COL_AZIENDA['cap']).value),
            })
        r += 1
    return aziende


# ----------------------- Utility di manipolazione Word ------------------
def _set_tabs(paragraph, tab_defs):
    """Imposta i tab stop del paragrafo e azzera eventuali indentazioni
    (w:ind) che sposterebbero l'origine dei tab stop."""
    pPr = paragraph._p.get_or_add_pPr()
    # rimuovi indentazione esistente (causa sfasamento tab stop)
    old_ind = pPr.find(qn('w:ind'))
    if old_ind is not None:
        pPr.remove(old_ind)
    old = pPr.find(qn('w:tabs'))
    if old is not None:
        pPr.remove(old)
    tabs = OxmlElement('w:tabs')
    for tipo, pos in tab_defs:
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), tipo)
        tab.set(qn('w:pos'), str(pos))
        tabs.append(tab)
    pPr.append(tabs)


def _nuova_run(template_run, testo):
    """Crea una run copiando le proprieta' (rPr) del template e impostando testo."""
    new_r = copy.deepcopy(template_run._r)
    for el in list(new_r):
        if el.tag in (qn('w:t'), qn('w:br'), qn('w:tab')):
            new_r.remove(el)
    if testo == "":
        return None
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = testo
    new_r.append(t)
    return new_r


def _nuova_tab_run(template_run):
    new_r = copy.deepcopy(template_run._r)
    for el in list(new_r):
        if el.tag in (qn('w:t'), qn('w:br'), qn('w:tab')):
            new_r.remove(el)
    new_r.append(OxmlElement('w:tab'))
    return new_r


def _template_run(paragraph):
    for r in paragraph.runs:
        if r._r.find(qn('w:rPr')) is not None:
            return r
    return paragraph.runs[0] if paragraph.runs else None


def _svuota_runs(paragraph):
    """Rimuove tutte le run del paragrafo mantenendo pPr (formattazione paragrafo)."""
    for r in list(paragraph.runs):
        r._r.getparent().remove(r._r)


def _appendi_run(paragraph, tmpl_run, testo):
    """Aggiunge una run con 'testo' in fondo al paragrafo, ereditando rPr."""
    new_r = copy.deepcopy(tmpl_run._r)
    for el in list(new_r):
        if el.tag in (qn('w:t'), qn('w:br'), qn('w:tab')):
            new_r.remove(el)
    if testo:
        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = testo
        new_r.append(t)
    paragraph._p.append(new_r)


def _appendi_tab(paragraph, tmpl_run):
    new_r = copy.deepcopy(tmpl_run._r)
    for el in list(new_r):
        if el.tag in (qn('w:t'), qn('w:br'), qn('w:tab')):
            new_r.remove(el)
    new_r.append(OxmlElement('w:tab'))
    paragraph._p.append(new_r)


def ricostruisci_campo(paragraph, segmenti):
    """Ricostruisce da zero il contenuto (run) del paragrafo.
    'segmenti' e' una lista di tuple (tipo, testo):
        ('t', 'testo')   -> run di testo
        ('tab', '')      -> una tabulazione
    I tab stop del paragrafo (gia' impostati) verranno consumati nell'ordine.
    """
    tmpl = _template_run(paragraph)
    _svuota_runs(paragraph)
    for tipo, testo in segmenti:
        if tipo == 't':
            _appendi_run(paragraph, tmpl, testo)
        elif tipo == 'tab':
            _appendi_tab(paragraph, tmpl)


# ----------------------- Compilazione -----------------------
def _campo_singolo(paragraph, etichetta, valore):
    """Campo con 1 box: [etichetta] TAB [valore centrato]."""
    seg = [('t', etichetta), ('tab', '')]
    if valore:
        seg.append(('t', valore))
    ricostruisci_campo(paragraph, seg)


def _campo_doppio(paragraph, etichetta_sx, valore_sx, etichetta_dx, valore_dx,
                  tab_prima_etichetta_dx=1):
    """Campo con 2 box. Struttura:
       [etichetta_sx] TAB [valore_sx centrato] (TAB xN) [etichetta_dx] TAB [valore_dx centrato]

    tab_prima_etichetta_dx: numero di tab tra il valore sx e l'etichetta dx.
        Nel modello alcune righe hanno 1 tab (Residente/via, CAP/C.F., P.IVA/C.F.,
        Con sede/Prov., via/CAP), la riga 'Nato a/il' ne ha 2.

    Tab stop del paragrafo devono essere:
        [center@box_sx, (left filler...) , left@etichetta_dx, center@box_dx]
    I tab in eccesso prima dell'etichetta dx devono avere un tab stop 'left' di
    riempimento alla stessa posizione per non spostare il cursore.
    """
    seg = [('t', etichetta_sx), ('tab', '')]
    if valore_sx:
        seg.append(('t', valore_sx))
    for _ in range(tab_prima_etichetta_dx):
        seg.append(('tab', ''))        # tab per raggiungere l'etichetta dx
    seg.append(('t', etichetta_dx))
    if valore_dx:
        seg.append(('tab', ''))        # salta a box_dx (tab center)
        seg.append(('t', valore_dx))
    ricostruisci_campo(paragraph, seg)


def compila(modello_path, persona, azienda, output_path):
    doc = Document(str(modello_path))
    p = doc.paragraphs

    # --- imposta tab stop (misti center/left) su tutti i campi ---
    for idx, tab_defs in TABS.items():
        _set_tabs(p[idx], tab_defs)

    # --- PERSONA FISICA (se presente) ---
    if persona:
        # 'Il sottoscritto' e' un solo box: etichetta composta 'Il sottoscritto'
        _campo_singolo(p[P_IL_SOTTOSCRITTO], 'Il sottoscritto', persona['nome'])
        _campo_doppio(p[P_NATO_A],       'Nato a',  persona['nato_a'],       'il',   persona['data_nascita'])
        _campo_doppio(p[P_RESIDENTE_IN], 'Residente in', persona['residente'], 'via', persona['via'])
        _campo_doppio(p[P_CAP_CF],       'CAP',     persona['cap'],          'C.F.', persona['cf'])

    # --- AZIENDA (se presente) ---
    if azienda:
        _campo_singolo(p[P_DENOMINAZIONE], 'Denominazione',  azienda['denominazione'])
        _campo_doppio(p[P_PIVA_CF],        'P.IVA', azienda['piva'], 'C.F.', azienda['cf'])
        _campo_singolo(p[P_RAGIONE_SOC],   'Ragione sociale', azienda['ragione_sociale'])
        _campo_singolo(p[P_TIPOLOGIA_ENTE],'Tipologia ente',  azienda['tipologia'])
        # 'Con sede / Prov.' e 'via / CAP' sono due righe (Con sede ... legale in / via CAP)
        # Il campo "via CAP" contiene via (box sx) e CAP (box dx)
        _campo_doppio(p[P_VIA_CAP],        'via',    azienda['via'], 'CAP',  azienda['cap'])
        # 'Con sede Prov.': il box sx contiene la via della sede, ma nel modello
        # la via e' sulla riga successiva ("legale in via..."); qui mettiamo solo Prov.
        _campo_doppio(p[P_CON_SEDE_PROV],  'Con sede', '', 'Prov.', azienda['prov'])

    doc.save(str(output_path))


def nome_file_output(denominazione, indice):
    safe = re.sub(r'[\\/:*?"<>|]', '_', denominazione or "").strip()
    safe = re.sub(r'\s+', ' ', safe)
    if not safe:
        safe = f"soggetto_{indice+1}"
    return CARTELLA / f"{PREFISSO_OUTPUT}{safe}.docx"


def main():
    if not MODELLO_DOCX.exists():
        print(f"ERRORE: modello non trovato: {MODELLO_DOCX}")
        sys.exit(1)
    if not DATI_XLSX.exists():
        print(f"ERRORE: file dati non trovato: {DATI_XLSX}")
        sys.exit(1)

    wb = openpyxl.load_workbook(str(DATI_XLSX), data_only=True)
    ws = wb[FOGLIO_DATI]
    persone = leggi_persone(ws)
    aziende = leggi_aziende(ws)

    print(f"Persone lette: {sum(1 for x in persone if x)} (su {len(persone)} righe)")
    print(f"Aziende lette: {len(aziende)}")

    if not aziende:
        print("Nessuna azienda trovata nell'Excel. Nulla da generare.")
        sys.exit(0)

    n_coppie = len(aziende)
    for i in range(n_coppie):
        az = aziende[i]
        persona = persone[i] if i < len(persone) else None
        out = nome_file_output(az['denominazione'], i)
        compila(MODELLO_DOCX, persona, az, out)
        nome_p = persona['nome'] if persona else "(nessuna persona)"
        print(f"  [{i+1}] {az['denominazione']:40s}  persona: {nome_p}")
        print(f"       -> {out.name}")

    print(f"\nFatto. {n_coppie} documento/i generato/i.")


if __name__ == "__main__":
    main()
