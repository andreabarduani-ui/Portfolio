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
            "euro(2880) -> '2.880,00 EUR'",
            "calc_totali(6) -> (2880.0, 2600.0)",
            "prossimo_numero_offerta(1, {'1': 43}) -> 20260046",
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
            "similarity ALSSANDRO/ALESSANDRO -> 0.9767",
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
            "usage: certificate_manager.py dividi|organizza|riorganizza",
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
    if (best) return best.reply;
    return fallbackReplyHtml();
  }

  /* ---------- Optional same-origin sample JSON load (progressive enhancement) ----------
     Same-origin only ('self' per CSP). Never blocks: on any error the
     inline answers above are used. No external fetch, no keys. */
  function tryLoadDemoData() {
    try {
      if (typeof fetch !== "function") return;
      fetch("demo-data/elearning-hours-monitor.json", { cache: "force-cache" })
        .then(function (res) {
          if (!res || !res.ok) return null;
          return res.json();
        })
        .then(function (data) {
          if (data && data.after && data.before) {
            window.__portfolioDemoData = data;
          }
        })
        .catch(function () {});
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
  };
})();
