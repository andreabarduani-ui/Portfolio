"""Offline demo tests for the top-3 proof tools (fictional data only).

No network, no secrets, no PII. Each test imports the real tool module
from projects/*/ by file path and exercises pure functions.
"""

import importlib.util
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]


def load(name, relpath):
    path = REPO / relpath
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_elearning_parsing_4_4():
    mod = load("elearning_hours_monitor", "projects/elearning-hours-monitor/elearning_hours_monitor.py")
    durations = ["2 h 30 min 00 s", "1 h 45 min 30 s", "0 h 25 min 10 s", "3 h 00 min 00 s"]
    parsed = [mod.parse_durata(d) for d in durations]
    assert len(parsed) == 4
    assert all(v > 0 for v in parsed)
    row = {"Totale Ore": "4 h 00 min 00 s", "Primo Accesso": "10:00:00", "Ultimo Accesso": "20:00:00"}
    prima, dopo = mod.splitta_ore_per_soglia(row, 18.0)
    assert prima > 0 and dopo > 0
    assert mod.normalizza_nome("Mario Rossi") == mod.normalizza_nome("ROSSI MARIO")


def test_funding_threshold():
    mod = load("funding_calls_scraper", "projects/funding-call-scraper/funding_calls_scraper.py")
    texts = [
        "Dotazione finanziaria complessiva \u20ac 2.500.000 per il bando.",
        "Il bando stanzia 1,1 milioni di euro per la formazione.",
        "Budget \u20ac 950000 - scadenza 2026.",
        "Contributo 120000 euro (sotto soglia).",
    ]
    assert mod.importo_migliore(texts[0])[0] == 2500000.0
    assert mod.importo_migliore(texts[0])[1] == "alta"
    assert mod.importo_migliore(texts[1])[0] == 1100000.0
    assert mod.importo_migliore(texts[3])[0] == 120000.0
    soglia = 800000.0
    above = [t for t in texts if (mod.importo_migliore(t)[0] or 0) > soglia]
    assert len(above) == 3
    assert mod.e_fnc3("Opportunita con Fondo Nuove Competenze - Terza Edizione") is True
    assert mod.e_fnc3("Bando ordinario per la formazione continua") is False


def test_regulation_page_hits():
    mod = load("regulation_search", "projects/regulation-search/regulation_search.py")
    pages = [
        (1, "General introduction to the fictional training programme. No relevant rules here."),
        (2, "I costi ammissibili includono il costo orario definito nell'Allegato B."),
        (3, "Variation requests must be submitted at least 30 days before the end of the programme."),
        (4, "State aid under de minimis regulation 1407/2013 applies."),
    ]
    res = mod.trova_corrispondenze(pages, mod.CATEGORIE["costi_ammissibili"])
    assert len(res) >= 1
    assert sorted(set(x["pagina"] for x in res)) == [2]
    res2 = mod.trova_corrispondenze(pages, ["variation", "deadline"])
    assert any(x["pagina"] == 3 for x in res2)
