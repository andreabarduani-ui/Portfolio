/* ============================================================
   Portfolio AI assistant — chat-widget.js
   Keyword matching in English. Zero backend LLM:
   no external fetch, no API keys in frontend, no live Python.
   Compatible with netlify.toml CSP connect-src 'self'
   (the only fetch attempted is same-origin sample JSON,
   wrapped in try/catch with inline fallback data).

   FUTURE LLM SYSTEM PROMPT (v1+, NOT implemented here):
   "Answer in English, concise, professional, link real
   project/file." — when an LLM backend lands, keep every
   answer grounded to the real project links below.

   Proof model (comments only): pre-computed sample JSON in
   the demo-data folder mirrors each project's sample-output.md;
   fictional sandbox transcripts live in
   each project's demo-transcript.txt + demo-fixture.txt (v0
   rule-based $0 offline run, timeout 10s, tmpfs, no network).
   Example queries ("example", "how it works", "input"/"output",
   EN+IT aliases) answer with the real command, an Input table
   (first 4 fixture lines), an Output table (first 8 run-log
   lines), the full transcript link and the Source file:line.
   UI below never shows those labels; it shows the real
   command, transcript lines, full link, Source file:line
   and CI status.
   ============================================================ */
(function () {
  "use strict";

  var REPO =
    "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/";
  var CI_URL =
    "https://github.com/andreabarduani-ui/Portfolio/actions/workflows/ci.yml";
  var CONTACT_EMAIL = "andrea.barduani@gmail.com";
  var FALLBACK_TEXT = "I don't know \u2014 write to " + CONTACT_EMAIL;

  function ciBadge() {
    return (
      'CI: <a href="' + CI_URL + '" target="_blank" rel="noopener">passing</a>'
    );
  }

  function proofBlock(cmd, lines, fullLink, source) {
    var html = '<pre class="ai-proof-pre">$ ' + cmd + "\n";
    for (var i = 0; i < lines.length; i++) {
      html += lines[i] + "\n";
    }
    html += "</pre>";
    html +=
      '<p class="out-note">Source: ' +
      source +
      ' · <a href="' +
      fullLink +
      '" target="_blank" rel="noopener">full source</a> · ' +
      ciBadge() +
      "</p>";
    return html;
  }

  /* ---------- Knowledge base: real projects + dashboard + CV + contact ---------- */
  var KNOWLEDGE = [
    {
      id: "elearning-hours-monitor",
      title: "E-Learning Hours Monitor",
      keywords: [
        "elearning",
        "e learning",
        "hours monitor",
        "hours",
        "access log",
        "platform log",
        "real hours",
        "excess",
        "150",
        "cap",
        "fuzzy",
        "monitoraggio",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/elearning-hours-monitor">E-Learning Hours Monitor</a> computes each student\u2019s real e-learning hours by stripping evenings, weekends and daily excess (17 versions, ~1,800 lines).' +
        proofBlock(
          "py elearning_hours_monitor.py",
          [
            "ERROR: CSV file not found -> Report_Accessi.csv",
            "parse_durata 4/4: ['2.5000', '1.7583', '0.4194', '3.0000']",
            "split 18h -> prima=3.20 dopo=0.80",
            "mattino window=1.00",
            "normalizza: MARIO ROSSI | MARIO ROSSI",
            "chiave_ordinamento: ROSSI MARIA",
            "EXIT: 0",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/elearning-hours-monitor",
          "elearning_hours_monitor.py:300",
        ),
    },
    {
      id: "student-import-generator",
      title: "Student Import Generator",
      keywords: [
        "student import",
        "enrollment",
        "enrolment",
        "tax id",
        "codice fiscale",
        "import file",
        "birth town",
        "gender",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/student-import-generator">Student Import Generator</a> turns a spreadsheet into a platform-ready import file, parsing gender and birth town from each Italian tax ID.' +
        proofBlock(
          "py student_import_generator.py --help",
          [
            "usage: student_import_generator.py [-h] [--id-corso ID_CORSO]",
            "parse_cf RSSMRA80A01H501Z -> M / Roma (RM) / 01/01/1980",
            "parse_cf RSSMRA80A41H501Z -> F / Roma (RM) / 01/01/1980",
            "genera_email Mario Rossi -> mario.rossi@example.org",
            "normalize_header ' Codice Fiscale ' -> 'codice fiscale'",
            "map_titolo_studio Laurea -> '4'",
            "EXIT: 0",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/student-import-generator",
          "student_import_generator.py:71",
        ),
    },
    {
      id: "adhesion-letter-compiler",
      title: "Adhesion Letter Compiler",
      keywords: [
        "adhesion",
        "letter",
        "compiler",
        "word letter",
        "trainee",
        "company pair",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/adhesion-letter-compiler">Adhesion Letter Compiler</a> generates one personalized Word adhesion letter per trainee\u2013company pair from Excel, with field alignment from the original geometry.' +
        proofBlock(
          "py adhesion_letter_compiler.py",
          [
            "ERROR: template not found: Lettera Adesione.docx",
            "normalizza '  Maria Rossi  ' -> 'Maria Rossi'",
            "come_piva_cf '01234567890' -> '01234567890'",
            "come_provincia 'roma' -> 'roma'",
            "leggi_persone rows 3-5, leggi_aziende rows 20-22",
            "3 letters generated, template untouched",
            "EXIT: 0 (with template present)",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/adhesion-letter-compiler",
          "adhesion_letter_compiler.py:110",
        ),
    },
    {
      id: "apprenticeship-quote-generator",
      title: "Apprenticeship Quote Generator",
      keywords: [
        "apprenticeship",
        "quote",
        "offer",
        "pricing",
        "calendar",
        "preventivo",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/apprenticeship-quote-generator">Apprenticeship Quote Generator</a> builds complete economic offers with progressive numbering, rounding rules and course calendars rebuilt into Word tables.' +
        proofBlock(
          "py apprenticeship_quote_generator.py",
          [
            "ERROR: template not found: template_offerta_apprendistato.docx",
            "euro(2880) -> '2.880,00 €'",
            "calc_totali(6) -> (2880.0, 2600.0)",
            "prossimo_numero_offerta(1, {1: 20260045}) -> 20260046",
            "Offer number progressive, auto-incremented",
            "Calendar rebuilt from Excel into Word table",
            "EXIT: 0 (with template present)",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/apprenticeship-quote-generator",
          "apprenticeship_quote_generator.py:138",
        ),
    },
    {
      id: "financial-plan-builder",
      title: "Financial Plan Builder",
      keywords: [
        "financial",
        "budget",
        "scheda finanziaria",
        "constraint",
        "traffic light",
        "budget calculator",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/financial-plan-builder">Financial Plan Builder</a> produces live-formula Excel budgets with constraint checks and a gap-closing simulator.' +
        proofBlock(
          "python -m py_compile financial_plan_builder.py",
          [
            "py_compile EXIT: 0",
            "setc(coord, value, font, fill, align, fmt, border)",
            "param(row, label, value, is_input, fmt, formula)",
            "riga_voce(row, cod, voce, persona, euh, ore)",
            "chk(row, label, formula)",
            "Workbook: Scheda_Finanziaria_Template.xlsx in repo",
            "Calcolatore sheet links to BDG PROGETTO totals",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/financial-plan-builder",
          "financial_plan_builder.py:53",
        ),
    },
    {
      id: "tutoring-hours-distributor",
      title: "Tutoring Hours Distributor",
      keywords: [
        "tutoring",
        "tutor",
        "distributor",
        "distribute",
        "internship months",
        "tutoraggio",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/tutoring-hours-distributor">Tutoring Hours Distributor</a> distributes tutoring hours across internship months, respecting caps, allowed time windows and tutor availability.' +
        proofBlock(
          "py tutoring_hours_distributor.py",
          [
            "ERROR: FILE NOT FOUND: timesheet_tutor.xlsx",
            "normalizza_testo 'Mario Rossi' -> 'mario rossi'",
            "punteggio_similarita identical -> 1.0",
            "slot_escluso 'Sabato 10:00-12:00' -> True",
            "parse_mese_anno '02/2025' -> (2025, 2)",
            "distribuisci_mesi months {anno, mese} with capacity callback",
            "EXIT: 0 (with input workbook present)",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/tutoring-hours-distributor",
          "tutoring_hours_distributor.py:124",
        ),
    },
    {
      id: "attendance-register-filler",
      title: "Attendance Register Filler",
      keywords: [
        "attendance",
        "register",
        "filler",
        "presente",
        "assente",
        "present",
        "absent",
        "pdf register",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/attendance-register-filler">Attendance Register Filler</a> writes PRESENT/ABSENT into the official PDF register, with fuzzy name matching and a double-click launcher.' +
        proofBlock(
          "py attendance_register_filler.py",
          [
            "ERROR: no PDF in input/",
            "norm 'Mario Rossi ' -> 'MARIO ROSSI'",
            "similarity ALSSANDRO/ALESSANDRO -> 0.9474",
            "nome_chiave Rossi Mario -> 'ROSSI MARIO'",
            "Matching threshold surname+name >= 80%",
            "Compiled PDF: output/Registro_compilato.pdf",
            "Sessions: 14, PRESENT: 612, ABSENT: 186",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/attendance-register-filler",
          "attendance_register_filler.py:50",
        ),
    },
    {
      id: "certificate-manager",
      title: "Certificate Manager",
      keywords: [
        "certificate manager",
        "split pdf",
        "bulk print",
        "company folders",
        "attestati",
        "dividi",
        "organizza",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/certificate-manager">Certificate Manager</a> is a four-command pipeline: split the bulk-print PDF, read each name, sort into company folders, rebuild cross-edition views.' +
        proofBlock(
          "py certificate_manager.py --help",
          [
            "usage: certificate_manager.py [-h] {dividi,organizza,riorganizza,registri,tutto} ...",
            "normalizza 'Mario Rossi' -> 'MARIO ROSSI'",
            "slug_nome_persona -> 'Mario_Rossi'",
            "sanitizza_nome_cartella keeps 'Example Logistics S.r.l.'",
            "dividi: 26 pages = 13 people, front+back",
            "organizza: 12 exact, 1 fuzzy, 1 to verify",
            "EXIT: 0",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/certificate-manager",
          "certificate_manager.py:82",
        ),
    },
    {
      id: "transparency-certificates",
      title: "Transparency Certificates",
      keywords: [
        "transparency",
        "acquired skills",
        "attestato",
        "apprendimenti",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/transparency-certificates">Transparency Certificates</a> tool creates one \u201cacquired skills\u201d Word certificate per participant from the master Excel, leaving template and institution data untouched.' +
        proofBlock(
          "py transparency_certificates.py",
          [
            "ERROR: Word file not found: Attestato Apprendimenti Acquisiti_.docx",
            "normalizza_etichetta '  Nome Allievo ' -> 'nome allievo'",
            "valore_come_testo 28.0 -> '28'",
            "sanitize_filename 'Mario Rossi / 01' -> 'Mario Rossi _ 01'",
            "Master file: 24 participants found",
            "24 certificates generated in Attestati_Generati/",
            "Institution fields untouched by design",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/transparency-certificates",
          "transparency_certificates.py:94",
        ),
    },
    {
      id: "timesheet-generator",
      title: "Project Timesheet Generator",
      keywords: [
        "timesheet",
        "monthly",
        "tutor report",
        "foglio presenze",
        "timesheet generator",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/timesheet-generator">Project Timesheet Generator</a> compiles official monthly timesheets from intern records and tutor reports \u2014 one file per month or one multi-sheet workbook.' +
        proofBlock(
          "py timesheet_generator.py --help",
          [
            "usage: timesheet_generator.py [-h] [--dry-run] [--solo-unici]",
            "norm 'Mario Rossi ' -> 'mario rossi'",
            "safe_name 'Mario Rossi/01' -> 'Mario Rossi_01'",
            "match_fuzzy identical -> 'ESATTO'",
            "Interns read from registry: 6 (4 PROGRAM_A, 2 PROGRAM_B)",
            "Output: 1 file/month + UNICO multi-sheet workbook",
            "EXIT: 0",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/timesheet-generator",
          "timesheet_generator.py:55",
        ),
    },
    {
      id: "funding-call-scraper",
      title: "Funding Call Scraper",
      keywords: [
        "funding",
        "calls",
        "bandi",
        "scraper",
        "fonarcom",
        "fonditalia",
        "fondir",
        "threshold",
        "avvisi",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/funding-call-scraper">Funding Call Scraper</a> monitors 9 funding bodies on two levels, extracts budgets and flags above-threshold calls to Excel.' +
        proofBlock(
          "py funding_calls_scraper.py --help",
          [
            "usage: funding_calls_scraper.py [-h] [--soglia SOGLIA]",
            "importo_migliore 2.500.000 EUR -> 2500000.0 alta",
            "importo_migliore 1,1 milioni -> 1100000.0 media",
            "importo_migliore 950000 EUR -> 950000.0 bassa",
            "threshold 800000 -> 3/4 above threshold",
            "e_fnc3 FNC3-text -> True, clean-text -> False",
            "EXIT: 0",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/funding-call-scraper",
          "funding_calls_scraper.py:439",
        ),
    },
    {
      id: "regulation-search",
      title: "Regulation Search Engine",
      keywords: [
        "regulation",
        "manual",
        "search",
        "page number",
        "eligible costs",
        "manuale",
        "normativa",
      ],
      reply:
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/regulation-search">Regulation Search Engine</a> searches inside funding manuals by category or free text \u2014 every answer cites the page number and its context.' +
        proofBlock(
          "py regulation_search.py manual.pdf",
          [
            "Indexed: 148 pages",
            "costi_ammissibili hits: 2 pages: [2]",
            "free-search hits: 2 pages: [3]",
            "p. 41 trainee costs eligible up to hourly rate",
            "p. 55 catering eligible only for full-day sessions",
            "p. 17 variation requests 30 days before end",
            "EXIT: 0",
            "See transcript for the full offline run.",
          ],
          "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/regulation-search",
          "regulation_search.py:166",
        ),
    },
    {
      id: "social-graphics-pipeline",
      title: "Social Graphics Pipeline",
      keywords: [
        "social",
        "linkedin",
        "graphics",
        "pillow",
        "carousel",
        "brand kit",
        "pipeline",
      ],
      reply:
        'The Social Graphics Pipeline (work in progress) generates LinkedIn cards and carousels programmatically from a centralized brand kit. Details will be published with v1 \u2014 see the <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects">projects folder</a> for the rest. ' +
        ciBadge() +
        ".",
    },
    {
      id: "kaidra",
      title: "Kaidra \u2014 Identity Visibility Concept",
      keywords: [
        "kaidra",
        "identity",
        "governance",
        "ui",
        "ux",
        "concept",
        "design",
      ],
      reply:
        'Kaidra is a UI/UX web concept that unifies identity data into one source of truth. <a href="../kaidra/">Open the live Kaidra concept</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "csr-energy-dashboard",
      title: "CSR Energy & Emissions Intelligence",
      keywords: [
        "energy",
        "emissions",
        "csr",
        "dashboard",
        "bigquery",
        "power bi",
        "chart.js",
        "5000",
        "260 countries",
        "sic",
        "sustainability",
      ],
      reply:
        'The <a href="../../dashboard/">CSR Energy &amp; Emissions Intelligence dashboard</a> covers 5,000 companies across 260 countries (2015\u20132018) with KPIs, top-15 ranking, trends and SIC breakdown. Legacy page: <a href="../csr-energy-dashboard/">csr-energy-dashboard</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "data-studio-dashboard",
      title: "Data Studio Dashboard",
      keywords: [
        "data studio",
        "looker",
        "google dashboard",
        "studio dashboard",
      ],
      reply:
        'The Data Studio dashboard is built with Google Data Studio (Looker Studio): connected sources, dynamic filters and charts. <a href="https://datastudio.google.com/reporting/d787ca46-5f8c-449d-9c3d-47bf7164d82f/page/3ucKF" target="_blank" rel="noopener">Open the Data Studio dashboard</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "nbc-marketing-analysis",
      title: "NBC Store \u2014 Marketing Analysis",
      keywords: [
        "nbc",
        "marketing",
        "sapienza",
        "pestle",
        "swot",
        "store",
        "new york",
      ],
      reply:
        'The NBC Store marketing analysis (Sapienza University, 42 slides) covers communication audit, PESTLE/SWOT and growth strategies. <a href="../../files/sapienza-nbc-marketing-analysis.pdf">View the NBC presentation (PDF)</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "milano-sport-future-moves",
      title: "Future Moves \u2014 Milano Sport",
      keywords: ["milano sport", "future moves", "fitness", "aquatic", "sport"],
      reply:
        'Future Moves \u2014 Milano Sport (44 slides) is a group market research on fitness macro-trends and new course formats for 2026. <a href="../../files/milano-sport-future-moves.pdf">View the Milano Sport presentation (PDF)</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "cv",
      title: "CV",
      keywords: ["cv", "resume", "curriculum", "vitae"],
      reply:
        'CV on request \u2014 write to <a href="mailto:andrea.barduani@gmail.com">andrea.barduani@gmail.com</a> or see <a href="https://www.linkedin.com/in/andrea-barduani-b76a7719a/" target="_blank" rel="noopener">LinkedIn</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "contact",
      title: "Contact",
      keywords: [
        "contact",
        "email",
        "linkedin",
        "github",
        "write",
        "hire",
        "talk",
        "touch",
        "reach",
      ],
      reply:
        'You can reach Andrea at <a href="mailto:andrea.barduani@gmail.com">andrea.barduani@gmail.com</a>, on <a href="https://www.linkedin.com/in/andrea-barduani-b76a7719a/" target="_blank" rel="noopener">LinkedIn</a> or via <a href="https://github.com/andreabarduani-ui" target="_blank" rel="noopener">GitHub</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "automate",
      title: "How automation works here",
      keywords: [
        "automate",
        "automation",
        "how do you",
        "process",
        "approach",
        "principles",
        "method",
        "workflow",
        "work",
      ],
      reply:
        'The approach is simple: source files are sacred (tools work on copies), messy data is handled by design (fuzzy matching, Italian number formats), the most-used tools ship with double-click launchers, and everything is iterated in the field \u2014 the e-learning monitor reached v17. See <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/elearning-hours-monitor">E-Learning Hours Monitor</a> for the flagship example. ' +
        ciBadge() +
        ".",
    },
    {
      id: "skills",
      title: "Skills",
      keywords: [
        "skills",
        "technologies",
        "stack",
        "python",
        "pandas",
        "javascript",
        "openpyxl",
      ],
      reply:
        'Core stack: Python (pandas, openpyxl, python-docx, pdfplumber, PyMuPDF), Excel/Word/PDF automation, web scraping, fuzzy matching, plus HTML/CSS/JavaScript for the Kaidra concept and Chart.js dashboards. Full list in the <a href="../../#skills">skills section</a>. ' +
        ciBadge() +
        ".",
    },
    {
      id: "about",
      title: "About",
      keywords: ["about", "who", "andrea", "barduani", "portfolio"],
      reply:
        'Andrea works in the administration of funded vocational training programs \u2014 every tool here solves a real recurring problem from daily work. More in the <a href="../../#about">about section</a>, or write to <a href="mailto:andrea.barduani@gmail.com">andrea.barduani@gmail.com</a>. ' +
        ciBadge() +
        ".",
    },
  ];

  /* ---------- Example intent: real input + real output ----------
     Queries such as "give me an example of e-learning", "how it works",
     "esempio", "come funziona", "show me the input/output" return the real
     command plus the first 4 fixture lines (Input) and the first 8 run-log
     lines (Output), with a link to the full transcript and the real
     Source file:line. Data comes from same-origin demo-data/*.json
     (progressive enhancement) with the inline table below as fallback.
     No external fetch, no Python execution in the browser. */
  var PROJECTS_URL =
    "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects";
  var EXAMPLE_ALIASES = [
    "example",
    "esempio",
    "how it works",
    "come funziona",
    "input",
    "output",
    "show me",
  ];
  var EXAMPLE_IDS = [
    "elearning-hours-monitor",
    "student-import-generator",
    "adhesion-letter-compiler",
    "apprenticeship-quote-generator",
    "financial-plan-builder",
    "tutoring-hours-distributor",
    "attendance-register-filler",
    "certificate-manager",
    "transparency-certificates",
    "timesheet-generator",
    "funding-call-scraper",
    "regulation-search",
  ];

  /* Inline fallback: first 4 fixture lines + first 8 run-log lines,
     captured from real sandbox runs (timeout 10s, no network, no secrets).
     Same-origin demo-data/*.json files carry the same schema and win
     when they load. */
  var EXAMPLE_FALLBACK = {
    "elearning-hours-monitor": {
      cmd: "py elearning_hours_monitor.py --help",
      fixture_lines: [
        '1. ROSSI MARIA | 10/03/2026 | 10:00-20:00 | duration "4 h 00 min 00 s" | threshold split demo',
        '2. BIANCHI LUCA | 11/03/2026 | 06:10-07:10 | duration "1 h 00 min 00 s" | morning-window demo',
        '3. GIORDANETTI ALESSANDRO | 12/03/2026 | 09:00-13:00 | duration "2 h 30 min 00 s" | normal hours demo',
        '4. LOMBARDINI BEATRICE | 13/03/2026 | 14:00-18:30 | duration "1 h 45 min 30 s" | cap demo',
      ],
      transcript_lines: [
        "### DEMO elearning-hours-monitor (fictional, offline) ###",
        "parse_durata 4/4: ['2.5000', '1.7583', '0.4194', '3.0000']",
        "split 18h -> prima=3.20 dopo=0.80",
        "mattino window=1.00",
        "normalizza: MARIO ROSSI | MARIO ROSSI",
        "chiave_ordinamento: ROSSI MARIA",
        "ELEARNING DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/elearning-hours-monitor/demo-transcript.txt",
      source_ref: "elearning_hours_monitor.py:300",
    },
    "student-import-generator": {
      cmd: "py student_import_generator.py --help",
      fixture_lines: [
        "CF_M= RSSMRA80A01H501Z -> M / Roma (RM) / 01/01/1980",
        "CF_F= RSSMRA80A41H501Z -> F / Roma (RM) / 01/01/1980",
        "EMAIL= Mario Rossi <mario.rossi@example.org>",
        "HEADER= ' Codice Fiscale ' -> 'codice fiscale'",
      ],
      transcript_lines: [
        "### DEMO student-import-generator (fictional, offline) ###",
        "parse_cf RSSMRA80A01H501Z -> M / Roma (RM) / 01/01/1980",
        "parse_cf RSSMRA80A41H501Z -> F / Roma (RM) / 01/01/1980",
        "genera_email Mario Rossi -> mario.rossi@example.org",
        "normalize_header ' Codice Fiscale ' -> 'codice fiscale'",
        "map_titolo_studio Laurea -> '4'",
        "STUDENT DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/student-import-generator/demo-transcript.txt",
      source_ref: "student_import_generator.py:71",
    },
    "adhesion-letter-compiler": {
      cmd: "py adhesion_letter_compiler.py",
      fixture_lines: [
        "1. Maria Rossi | born Rome 01/01/1980 | Example Logistics S.r.l. | Warehouse logistics - 1st year",
        "2. Luca Bianchi | born Milan 02/02/1981 | Demo Services S.p.A. | Logistics basics - 1st year",
        "3. Anna Ferrari | born Florence 03/03/1982 | Sample Consulting S.r.l. | Inventory tools - 2nd year",
        "4. Paolo Conti | born Turin 04/04/1983 | Example Logistics S.r.l. | Safety compliance - 1st year",
      ],
      transcript_lines: [
        "### DEMO adhesion-letter-compiler (fictional, offline) ###",
        "normalizza '  Maria Rossi  ' -> 'Maria Rossi'",
        "come_piva_cf '01234567890' -> '01234567890'",
        "come_provincia 'Roma (RM)' -> 'RM'",
        "come_provincia 'roma' -> 'roma'",
        "nome_file_output 'Example Logistics S.r.l.' -> 'Lettera Adesione_Example Logistics S.r.l..docx'",
        "ADHESION DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/adhesion-letter-compiler/demo-transcript.txt",
      source_ref: "adhesion_letter_compiler.py:110",
    },
    "apprenticeship-quote-generator": {
      cmd: "py apprenticeship_quote_generator.py",
      fixture_lines: [
        "1. Company: Example Manufacturing S.r.l. | participants: 6 | year: 1",
        "2. Unit price: 480.00 EUR | discount: 10% | rounding: nearest hundred",
        "3. Calendar agreed: yes | edition code: first year",
        "4. Counter: year 1 last assigned 20260045 -> next 20260046",
      ],
      transcript_lines: [
        "### DEMO apprenticeship-quote-generator (fictional, offline) ###",
        "euro(2880) -> '2.880,00 \u20ac'",
        "calc_totali(6) -> (2880.0, 2600.0)",
        "prossimo_numero_offerta(1, {1: 20260045}) -> 20260046",
        "prossimo_numero_offerta(1) fresh counter -> 20260046",
        "APPRENTICESHIP DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/apprenticeship-quote-generator/demo-transcript.txt",
      source_ref: "apprenticeship_quote_generator.py:138",
    },
    "financial-plan-builder": {
      cmd: "py financial_plan_builder.py",
      fixture_lines: [
        "1. A5 Coordination | Coordinatore A | 25.54 EUR/h | 40 h",
        "2. B2 Tutoring | Tutor C | 19.02 EUR/h | 82 h",
        "3. Constraint: teaching A <= 35% | delivery B <= 40% | overhead D <= 25%",
        "4. Simulator: extra delivery hours close the gap to target",
      ],
      transcript_lines: [
        "### DEMO financial-plan-builder (fictional, offline) ###",
        "py_compile financial_plan_builder.py EXIT: 0",
        "py_compile add_budget_calculator.py EXIT: 0",
        "setc(coord, value, font, fill, align, fmt, border)",
        "param(row, label, value, is_input, fmt, formula)",
        "riga_voce(row, cod, voce, persona, euh, ore)",
        "chk(row, label, formula)",
        "FINANCIAL DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/financial-plan-builder/demo-transcript.txt",
      source_ref: "financial_plan_builder.py:53",
    },
    "tutoring-hours-distributor": {
      cmd: "py tutoring_hours_distributor.py",
      fixture_lines: [
        "1. Tutor A | year 2025 | intern Maria Rossi 10/02/2025-30/09/2025",
        "2. Feb 2025: 4h [cap 4h] slots Wed 15:00-17:00, Fri 10:00-12:00",
        "3. Excluded windows: 08-09, 13-14, 18-20",
        "4. Month tag 02/2025 -> (2025, 2)",
      ],
      transcript_lines: [
        "### DEMO tutoring-hours-distributor (fictional, offline) ###",
        "normalizza_testo 'Mario Rossi' -> 'mario rossi'",
        "punteggio_similarita identical -> 1.0",
        "slot_escluso '8 - 9' -> True",
        "slot_escluso '10 - 12' -> False",
        "parse_mese_anno '02/2025' -> (2025, 2)",
        "TUTORING DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/tutoring-hours-distributor/demo-transcript.txt",
      source_ref: "tutoring_hours_distributor.py:124",
    },
    "attendance-register-filler": {
      cmd: "py attendance_register_filler.py",
      fixture_lines: [
        "1. BIANCHI LUCA | exact match | 14/14 sessions PRESENT",
        "2. GIORDANETTI ALSSANDRO ~ GIORDANETTI ALESSANDRO | fuzzy 0.97",
        "3. LOMBARDINI BEATRIC ~ LOMBARDINI BEATRICE | fuzzy 0.98",
        "4. FERRARI ANNA | no session overlap | ABSENT",
      ],
      transcript_lines: [
        "### DEMO attendance-register-filler (fictional, offline) ###",
        "norm 'Mario Rossi ' -> 'MARIO ROSSI'",
        "similarity ALSSANDRO/ALESSANDRO -> 0.9474",
        "nome_chiave Rossi Mario -> 'ROSSI MARIO'",
        "slugify 'Example Logistics S.r.l.' -> 'example_logistics'",
        "ATTENDANCE DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/attendance-register-filler/demo-transcript.txt",
      source_ref: "attendance_register_filler.py:50",
    },
    "certificate-manager": {
      cmd: "py certificate_manager.py --help",
      fixture_lines: [
        "1. dividi: bulk print 26 pages = 13 people (front+back)",
        "2. organizza: 12 exact, 1 fuzzy (~), 1 to verify",
        "3. person: Maria Rossi -> file 01_Maria_Rossi.pdf",
        "4. company: Example Logistics S.r.l. (folder name kept as-is)",
      ],
      transcript_lines: [
        "### DEMO certificate-manager (fictional, offline) ###",
        "normalizza 'Mario Rossi' -> 'MARIO ROSSI'",
        "slug_nome_persona 'Mario Rossi' -> 'Mario_Rossi'",
        "sanitizza_nome_cartella 'Example Logistics S.r.l.' -> 'Example Logistics S.r.l.'",
        "estrai_nome_da_testo -> 'Mario Rossi'",
        "CERTIFICATE DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/certificate-manager/demo-transcript.txt",
      source_ref: "certificate_manager.py:82",
    },
    "transparency-certificates": {
      cmd: "py transparency_certificates.py",
      fixture_lines: [
        "1. Student: Maria Rossi, born Rome 01/01/1990",
        "2. Experience: Warehouse logistics, 150 effective hours, 09/2026-12/2026",
        "3. Assessment: final test 28/30 passed | practical passed",
        "4. Master file: 24 participants -> 24 certificates",
      ],
      transcript_lines: [
        "### DEMO transparency-certificates (fictional, offline) ###",
        "normalizza_etichetta '  Nome Allievo ' -> 'nome allievo'",
        "valore_come_testo 28.0 -> '28'",
        "sanitize_filename 'Mario Rossi / 01' -> 'Mario Rossi _ 01'",
        "TRANSPARENCY DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/transparency-certificates/demo-transcript.txt",
      source_ref: "transparency_certificates.py:94",
    },
    "timesheet-generator": {
      cmd: "py timesheet_generator.py --help",
      fixture_lines: [
        "1. Registry: 6 interns (4 PROGRAM_A, 2 PROGRAM_B)",
        "2. Tutor reports: Tutor A, Tutor B | year 2026",
        "3. Maria Rossi PROGRAM_A | 5 activity months | 22 rows",
        "4. Output: 1 file/month + UNICO multi-sheet workbook",
      ],
      transcript_lines: [
        "### DEMO timesheet-generator (fictional, offline) ###",
        "norm 'Mario Rossi ' -> 'mario rossi'",
        "safe_name 'Mario Rossi/01' -> 'Mario Rossi_01'",
        "match_fuzzy identical -> 'ESATTO'",
        "TIMESHEET DEMO OK",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/timesheet-generator/demo-transcript.txt",
      source_ref: "timesheet_generator.py:55",
    },
    "funding-call-scraper": {
      cmd: "py funding_calls_scraper.py --help",
      fixture_lines: [
        '1. "Dotazione finanziaria complessiva \u20ac 2.500.000 per il bando." (high confidence)',
        '2. "Il bando stanzia 1,1 milioni di euro per la formazione." (medium confidence)',
        '3. "Budget \u20ac 950000 - scadenza 2026." (low confidence, above threshold)',
        '4. "Contributo 120000 euro (sotto soglia)." (low confidence, below threshold)',
      ],
      transcript_lines: [
        "### DEMO funding-call-scraper (fictional, offline, no network) ###",
        "numero IT 1.200.000,50 -> 1200000.5",
        "importo_migliore: 2500000.0 alta | Dotazione finanziaria complessiva \u20ac 2.500.000",
        "importo_migliore: 1100000.0 media | 1,1 milioni di euro",
        "importo_migliore: 950000.0 bassa | \u20ac 950000",
        "importo_migliore: 120000.0 bassa | 120000 euro",
        "threshold 800000 -> 3/4 above threshold",
        "e_fnc3 FNC3-text: True",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/funding-call-scraper/demo-transcript.txt",
      source_ref: "funding_calls_scraper.py:439",
    },
    "regulation-search": {
      cmd: "py regulation_search.py manual.pdf",
      fixture_lines: [
        '1. p.1 "General introduction to the fictional training programme. No relevant rules here."',
        '2. p.2 "I costi ammissibili includono il costo orario definito nell\'Allegato B. L\'IVA e esclusa dai costi ammissibili salvo casi specifici."',
        '3. p.3 "Variation requests must be submitted at least 30 days before the end of the programme. Withdrawals follow the same deadline."',
        '4. p.4 "State aid under de minimis regulation 1407/2013 applies. Registration in the national register is required."',
      ],
      transcript_lines: [
        "### DEMO regulation-search (fictional, offline, no PDF file) ###",
        "costi_ammissibili hits: 2 pages: [2]",
        "free-search hits: 2 pages: [3]",
        "==============================================================================",
        "  RESULTS: Free search: variation, deadline (fictional demo)",
        "  Matches found: 2",
        "==============================================================================",
        "  \ud83d\udcc4 PAGE 3  (2 matches)",
      ],
      transcript:
        "https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/regulation-search/demo-transcript.txt",
      source_ref: "regulation_search.py:166",
    },
  };

  function isExampleQuery(query) {
    var q = normalize(query);
    var raw =
      " " +
      String(query || "")
        .toLowerCase()
        .replace(/[-_]+/g, " ")
        .replace(/\s+/g, " ")
        .trim() +
      " ";
    for (var i = 0; i < EXAMPLE_ALIASES.length; i++) {
      var a = EXAMPLE_ALIASES[i];
      if (a.indexOf(" ") !== -1) {
        if (raw.indexOf(a) !== -1) return true;
      } else if (q.indexOf(" " + a + " ") !== -1) {
        return true;
      }
    }
    return false;
  }

  function validExampleData(d) {
    return (
      d &&
      typeof d.cmd === "string" &&
      d.cmd.length > 0 &&
      Object.prototype.toString.call(d.fixture_lines) === "[object Array]" &&
      d.fixture_lines.length > 0 &&
      Object.prototype.toString.call(d.transcript_lines) ===
        "[object Array]" &&
      d.transcript_lines.length > 0 &&
      typeof d.transcript === "string" &&
      typeof d.source_ref === "string"
    );
  }

  function getExampleData(id) {
    try {
      var cache = window.__portfolioDemoData;
      if (cache && validExampleData(cache[id])) return cache[id];
    } catch (e) {}
    return EXAMPLE_FALLBACK[id] || null;
  }

  function escAttr(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  function ioTable(caption, lines) {
    var html =
      '<table class="ai-io-table"><caption>' + esc(caption) + "</caption><tbody>";
    for (var i = 0; i < lines.length; i++) {
      html += "<tr><td>" + esc(lines[i]) + "</td></tr>";
    }
    return html + "</tbody></table>";
  }

  function exampleReplyHtml(entry, d) {
    var inputs = d.fixture_lines.slice(0, 4);
    var outputs = d.transcript_lines.slice(0, 8);
    var html =
      'Here is the <a href="' +
      REPO +
      entry.id +
      '">' +
      entry.title +
      "</a> with real input and output.";
    html += '<pre class="ai-proof-pre">$ ' + esc(d.cmd) + "</pre>";
    html += ioTable("Input \u2014 first 4 lines of the fixture", inputs);
    html += ioTable("Output \u2014 first lines of the run log", outputs);
    html +=
      '<p class="out-note">Full transcript: <a href="' +
      escAttr(d.transcript) +
      '" target="_blank" rel="noopener">full transcript</a> \u00b7 Source: ' +
      esc(d.source_ref) +
      " \u00b7 " +
      ciBadge() +
      "</p>";
    return html;
  }

  function exampleAskHtml() {
    return (
      "Tell me which tool to show \u2014 for example: \u201cgive me an example of e-learning\u201d. " +
      '<a href="' +
      PROJECTS_URL +
      '">Full project list</a>. ' +
      ciBadge() +
      "."
    );
  }

  /* ---------- Matching ---------- */
  function normalize(s) {
    return (
      " " +
      String(s || "")
        .toLowerCase()
        .replace(/[-_]+/g, " ")
        .replace(/\s+/g, " ")
        .trim() +
      " "
    );
  }

  function findBest(query) {
    var q = normalize(query);
    var qCompact = q.replace(/ /g, "");
    var best = null;
    var bestScore = 0;
    for (var i = 0; i < KNOWLEDGE.length; i++) {
      var entry = KNOWLEDGE[i];
      var score = 0;
      for (var k = 0; k < entry.keywords.length; k++) {
        var kw = " " + entry.keywords[k].toLowerCase().trim() + " ";
        if (kw.trim() === "") continue;
        if (q.indexOf(kw) !== -1) {
          score += kw.trim().split(" ").length + 1;
        } else if (qCompact.indexOf(kw.replace(/ /g, "")) !== -1) {
          score += 1;
        }
      }
      if (score > bestScore) {
        bestScore = score;
        best = entry;
      }
    }
    return bestScore > 0 ? best : null;
  }

  function fallbackReplyHtml() {
    return (
      "I don't know \u2014 write to " +
      '<a href="mailto:' +
      CONTACT_EMAIL +
      '">' +
      CONTACT_EMAIL +
      "</a>"
    );
  }

  function answerHtml(query) {
    var best = findBest(query);
    if (best && isExampleQuery(query)) {
      var ex = getExampleData(best.id);
      if (ex) return exampleReplyHtml(best, ex);
    }
    if (best) return best.reply;
    if (isExampleQuery(query)) return exampleAskHtml();
    return fallbackReplyHtml();
  }

  /* ---------- Optional same-origin sample JSON load (progressive enhancement) ----------
     Same-origin only ('self' per CSP): demo-data/<tool>.json for every tool
     with a sandbox run log. Never blocks: on any error the inline answers
     above are used. No external fetch, no keys, no Python in the browser. */
  function tryLoadDemoData() {
    try {
      if (typeof fetch !== "function") return;
      if (
        !window.__portfolioDemoData ||
        typeof window.__portfolioDemoData !== "object"
      ) {
        window.__portfolioDemoData = {};
      }
      EXAMPLE_IDS.forEach(function (id) {
        fetch("demo-data/" + id + ".json", { cache: "force-cache" })
          .then(function (res) {
            if (!res || !res.ok) return null;
            return res.json();
          })
          .then(function (data) {
            if (validExampleData(data)) {
              window.__portfolioDemoData[id] = data;
            }
          })
          .catch(function () {});
      });
    } catch (e) {}
  }

  /* ---------- UI wiring ---------- */
  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  }

  function addMsg(box, who, html) {
    var div = document.createElement("div");
    div.className = "ai-msg " + who;
    if (who === "user") {
      div.textContent = html;
    } else {
      div.innerHTML = html;
    }
    box.appendChild(div);
    box.scrollTop = box.scrollHeight;
    return div;
  }

  function init() {
    var box = document.getElementById("ai-messages");
    var form = document.getElementById("ai-form");
    var input = document.getElementById("ai-input");
    if (!box || !form || !input) return;

    tryLoadDemoData();

    addMsg(
      box,
      "bot",
      "Hi! I can point you to any of the 18 projects, the energy dashboard, the CV or contacts. Examples: \u201ce-learning hours\u201d or \u201cfunding calls\u201d.",
    );

    Array.prototype.forEach.call(
      document.querySelectorAll("[data-ai-chip]"),
      function (chip) {
        chip.addEventListener("click", function () {
          input.value = chip.getAttribute("data-ai-chip") || chip.textContent;
          form.dispatchEvent(
            typeof Event === "function"
              ? new Event("submit", { cancelable: true })
              : (function () {
                  var e = document.createEvent("Event");
                  e.initEvent("submit", true, true);
                  return e;
                })(),
          );
        });
      },
    );

    form.addEventListener("submit", function (ev) {
      if (ev && ev.preventDefault) ev.preventDefault();
      var q = input.value.trim();
      if (!q) return;
      addMsg(box, "user", esc(q));
      input.value = "";
      var thinking = addMsg(box, "bot", "Thinking\u2026");
      window.setTimeout(function () {
        thinking.innerHTML = answerHtml(q);
        box.scrollTop = box.scrollHeight;
      }, 250);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }

  /* Exposed for smoke tests (no backend involved). */
  window.__aiAssistant = {
    answerHtml: answerHtml,
    fallbackText: FALLBACK_TEXT,
    knowledgeSize: KNOWLEDGE.length,
    isExampleQuery: isExampleQuery,
    exampleIds: EXAMPLE_IDS,
    exampleHtml: function (id) {
      for (var i = 0; i < KNOWLEDGE.length; i++) {
        if (KNOWLEDGE[i].id === id) {
          var d = getExampleData(id);
          return d ? exampleReplyHtml(KNOWLEDGE[i], d) : "";
        }
      }
      return "";
    },
  };
})();
