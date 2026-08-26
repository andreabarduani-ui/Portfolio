"""
adhesion_letter_compiler.py
===========================
Automatically fills in the "Lettera Adesione.docx" starting from the data in
an Excel file ("Dati da inserire.xlsx").

PAIRING LOGIC (positional):
- The first person (row 3 of the Excel) is paired with the first company
  (row 20), the second person (row 4) with the second company (row 21),
  and so on ("first with first, second with second").
- If there are more companies than persons, the letters without a person will
  have empty natural-person boxes (and vice versa).

For each pair it generates a filled-in copy of the Word template.

CENTERING:
- Values are centered horizontally inside the form boxes,
  using 'center' tab stops positioned at the exact center of each
  box (positions derived from the geometry of the original PDF).

USAGE:
    py adhesion_letter_compiler.py

The file names are configurable in the constants below. The original
template is NEVER modified.
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


# ============================ CONFIGURATION ============================
CARTELLA = Path(__file__).resolve().parent
MODELLO_DOCX = CARTELLA / "Lettera Adesione.docx"
DATI_XLSX = CARTELLA / "Dati da inserire.xlsx"
PREFISSO_OUTPUT = "Lettera Adesione_"
FOGLIO_DATI = "Foglio1"

# (0-based) indices of the paragraphs in the Word template that contain the fields.
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

# --- Tab stop positions (twips) ---
# Derived from the geometry of the original PDF.
#
# SINGLE fields (1 wide box): a single 'center' tab at the center of the box.
#   e.g. Denominazione: center @ center of box [138-542].
#
# DOUBLE fields (2 side-by-side boxes with two labels):
#   3 tab stops in this order:
#     [center@center_left_box , left@pos_right_label , center@center_right_box]
#   so the right-hand label ("il", "via", "C.F.", "Prov.", "CAP") keeps
#   the original position (left tab, ~x=333pt) and is NOT re-centered, while
#   the values are centered in their respective boxes.
#
# Tuple syntax: (type, position_twips). List of tuples per paragraph.
def _c(pos):  return ('center', pos)   # center tab
def _l(pos):  return ('left', pos)     # left tab (original right-hand label)

TABS = {
    P_IL_SOTTOSCRITTO: [_c(5631)],
    # 'Nato a/il': 2 tabs before 'il'. Sequence of consumed tabs:
    #   TAB1 -> center@box_sx (value 'Parma (PR)' centered)
    #   TAB2 -> left@etichetta_dx (stops at left, does NOT move further)
    #   TAB3 (after 'il') -> center@box_dx (date value centered)
    # To do this the tab stops must be: [center@box_sx, left@box_sx, left@etichetta_dx, center@box_dx]
    # The 2nd (left@box_sx) stops the 2nd tab right after the value without advancing.
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

# Excel rows
RIGA_PRIMA_PERSONA = 3      # natural persons starting from row 3
RIGA_PRIMA_AZIENDA = 20     # companies starting from row 20
COL_PERSONA = {             # column -> dict key
    'nome': 2, 'nato_a': 3, 'data_nascita': 4, 'residente': 5,
    'via': 6, 'cap': 7, 'cf': 8,
}
COL_AZIENDA = {
    'denominazione': 2, 'piva': 3, 'cf': 4, 'ragione_sociale': 5,
    'tipologia': 7, 'via': 8, 'prov': 10, 'cap': 11,
}
# ========================================================================


# ----------------------- Excel reading utilities ------------------------
def normalizza(valore):
    if valore is None:
        return ""
    if isinstance(valore, (datetime, date)):
        return valore.strftime("%d/%m/%Y")
    if isinstance(valore, float) and valore.is_integer():
        return str(int(valore))
    return str(valore).strip()


def come_piva_cf(valore):
    """P.IVA / C.F. read as a number: returns the string without decimals."""
    if valore is None:
        return ""
    if isinstance(valore, (int, float)):
        return str(int(valore))
    return str(valore).strip()


def come_provincia(valore):
    """Extracts the province code from values like 'Roma | RM' or 'ROMA (RM)'."""
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
        # stop if the row is completely empty AND we are past the first one
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
            # empty row among the persons: placeholder for positional pairing
            persone.append(None)
        # we stop when we reach row 5 (fixed text "In qualita...")
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


# ----------------------- Word manipulation utilities --------------------
def _set_tabs(paragraph, tab_defs):
    """Sets the paragraph's tab stops and clears any indentation
    (w:ind) that would shift the origin of the tab stops."""
    pPr = paragraph._p.get_or_add_pPr()
    # remove existing indentation (it causes tab stop misalignment)
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
    """Creates a run copying the template's properties (rPr) and setting the text."""
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
    """Removes all runs of the paragraph while keeping pPr (paragraph formatting)."""
    for r in list(paragraph.runs):
        r._r.getparent().remove(r._r)


def _appendi_run(paragraph, tmpl_run, testo):
    """Appends a run with 'testo' (text) at the end of the paragraph, inheriting rPr."""
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
    """Rebuilds the paragraph content (runs) from scratch.
    'segmenti' is a list of (type, text) tuples:
        ('t', 'testo')   -> text run
        ('tab', '')      -> a tabulation
    The paragraph's tab stops (already set) will be consumed in order.
    """
    tmpl = _template_run(paragraph)
    _svuota_runs(paragraph)
    for tipo, testo in segmenti:
        if tipo == 't':
            _appendi_run(paragraph, tmpl, testo)
        elif tipo == 'tab':
            _appendi_tab(paragraph, tmpl)


# ----------------------- Compilation -----------------------
def _campo_singolo(paragraph, etichetta, valore):
    """Field with 1 box: [etichetta] TAB [centered valore]."""
    seg = [('t', etichetta), ('tab', '')]
    if valore:
        seg.append(('t', valore))
    ricostruisci_campo(paragraph, seg)


def _campo_doppio(paragraph, etichetta_sx, valore_sx, etichetta_dx, valore_dx,
                  tab_prima_etichetta_dx=1):
    """Field with 2 boxes. Structure:
       [etichetta_sx] TAB [centered valore_sx] (TAB xN) [etichetta_dx] TAB [centered valore_dx]

    tab_prima_etichetta_dx: number of tabs between the left value and the right label.
        In the template some rows have 1 tab (Residente/via, CAP/C.F., P.IVA/C.F.,
        Con sede/Prov., via/CAP), the 'Nato a/il' row has 2.

    The paragraph's tab stops must be:
        [center@box_sx, (left filler...) , left@etichetta_dx, center@box_dx]
    Excess tabs before the right label must have a filler 'left' tab stop
    at the same position so the cursor is not moved.
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

    # --- set tab stops (mixed center/left) on all fields ---
    for idx, tab_defs in TABS.items():
        _set_tabs(p[idx], tab_defs)

    # --- NATURAL PERSON (if present) ---
    if persona:
        # 'Il sottoscritto' is a single box: composite label 'Il sottoscritto'
        _campo_singolo(p[P_IL_SOTTOSCRITTO], 'Il sottoscritto', persona['nome'])
        _campo_doppio(p[P_NATO_A],       'Nato a',  persona['nato_a'],       'il',   persona['data_nascita'])
        _campo_doppio(p[P_RESIDENTE_IN], 'Residente in', persona['residente'], 'via', persona['via'])
        _campo_doppio(p[P_CAP_CF],       'CAP',     persona['cap'],          'C.F.', persona['cf'])

    # --- COMPANY (if present) ---
    if azienda:
        _campo_singolo(p[P_DENOMINAZIONE], 'Denominazione',  azienda['denominazione'])
        _campo_doppio(p[P_PIVA_CF],        'P.IVA', azienda['piva'], 'C.F.', azienda['cf'])
        _campo_singolo(p[P_RAGIONE_SOC],   'Ragione sociale', azienda['ragione_sociale'])
        _campo_singolo(p[P_TIPOLOGIA_ENTE],'Tipologia ente',  azienda['tipologia'])
        # 'Con sede / Prov.' and 'via / CAP' are two rows (Con sede ... legale in / via CAP)
        # The "via CAP" field contains via (left box) and CAP (right box)
        _campo_doppio(p[P_VIA_CAP],        'via',    azienda['via'], 'CAP',  azienda['cap'])
        # 'Con sede Prov.': the left box contains the registered office street, but in
        # the template the street is on the next row ("legale in via..."); here we only put Prov.
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
        print(f"ERROR: template not found: {MODELLO_DOCX}")
        sys.exit(1)
    if not DATI_XLSX.exists():
        print(f"ERROR: data file not found: {DATI_XLSX}")
        sys.exit(1)

    wb = openpyxl.load_workbook(str(DATI_XLSX), data_only=True)
    ws = wb[FOGLIO_DATI]
    persone = leggi_persone(ws)
    aziende = leggi_aziende(ws)

    print(f"Persons read: {sum(1 for x in persone if x)} (out of {len(persone)} rows)")
    print(f"Companies read: {len(aziende)}")

    if not aziende:
        print("No companies found in the Excel file. Nothing to generate.")
        sys.exit(0)

    n_coppie = len(aziende)
    for i in range(n_coppie):
        az = aziende[i]
        persona = persone[i] if i < len(persone) else None
        out = nome_file_output(az['denominazione'], i)
        compila(MODELLO_DOCX, persona, az, out)
        nome_p = persona['nome'] if persona else "(no person)"
        print(f"  [{i+1}] {az['denominazione']:40s}  person: {nome_p}")
        print(f"       -> {out.name}")

    print(f"\nDone. {n_coppie} document(s) generated.")


if __name__ == "__main__":
    main()
