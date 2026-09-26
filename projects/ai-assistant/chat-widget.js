/* ============================================================
   Portfolio AI assistant (v0) — chat-widget.js
   Rule-based EN keyword matching. $0, zero backend LLM:
   no external fetch, no API keys in frontend, no live Python.
   Compatible with netlify.toml CSP connect-src 'self'
   (the only fetch attempted is same-origin demo-data JSON,
   wrapped in try/catch with inline fallback data).

   FUTURE LLM SYSTEM PROMPT (v1+, NOT implemented in v0):
   "Answer in English, concise, professional, link real
   project/file." — when an LLM backend lands, keep every
   answer grounded to the real project links below.
   ============================================================ */
(function () {
  "use strict";

  var REPO =
    "https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/";
  var CONTACT_EMAIL = "andrea.barduani@gmail.com";
  var FALLBACK_TEXT = "I don't know \u2014 write to " + CONTACT_EMAIL;

  /* ---------- Inline demo (mirrors demo-data/*.json, no fetch needed) ---------- */
  var DEMO_ELEARNING = {
    title: "E-Learning Hours Monitor \u2014 before / after (pre-computed)",
    rows: [
      ["ROSSI MARIA", "71:40:00 total", "9:25:00 excess", "62:15:00 effective"],
      [
        "BIANCHI LUCA",
        "49:00:00 total",
        "0:30:00 excess",
        "48:30:00 effective",
      ],
      [
        "GIORDANETTI ALESSANDRO",
        "63:45:00 total",
        "8:40:00 excess",
        "55:05:00 effective",
      ],
    ],
    note: "Effective = total minus evenings after 19:00, weekends, 06:00\u201307:40 and the 8h/day cap. Full pre-computed data: demo-data/elearning-hours-monitor.json.",
  };

  /* ---------- Knowledge base: 18 real projects + dashboard + CV + contact ---------- */
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/elearning-hours-monitor">E-Learning Hours Monitor</a> computes each student\u2019s real e-learning hours by stripping evenings, weekends and daily excess (17 versions, ~1,800 lines). Try the chip \u201cShow backend demo\u201d for before/after numbers.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/student-import-generator">Student Import Generator</a> turns a spreadsheet into a platform-ready import file, parsing gender and birth town from each Italian tax ID.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/adhesion-letter-compiler">Adhesion Letter Compiler</a> generates one personalized Word adhesion letter per trainee\u2013company pair from Excel, with pixel-accurate field alignment.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/apprenticeship-quote-generator">Apprenticeship Quote Generator</a> builds complete economic offers with progressive numbering, rounding rules and course calendars rebuilt into Word tables.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/financial-plan-builder">Financial Plan Builder</a> produces live-formula Excel budgets with traffic-light constraint checks and a gap-closing simulator.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/tutoring-hours-distributor">Tutoring Hours Distributor</a> distributes tutoring hours across internship months, respecting caps, allowed time windows and tutor availability.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/attendance-register-filler">Attendance Register Filler</a> writes PRESENT/ABSENT into the official PDF register, with fuzzy name matching and a double-click launcher.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/certificate-manager">Certificate Manager</a> is a four-command pipeline: split the bulk-print PDF, read each name, sort into company folders, rebuild cross-edition views.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/transparency-certificates">Transparency Certificates</a> tool creates one \u201cacquired skills\u201d Word certificate per participant from the master Excel, leaving template and institution data untouched.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/timesheet-generator">Project Timesheet Generator</a> compiles official monthly timesheets from intern records and tutor reports \u2014 one file per month or one multi-sheet workbook.',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/funding-call-scraper">Funding Call Scraper</a> monitors 9 funding bodies on two levels, extracts budgets and flags above-threshold calls to Excel (pre-computed sample: demo-data/funding-call-scraper.json).',
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
        'The <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/regulation-search">Regulation Search Engine</a> searches inside funding manuals by category or free text \u2014 every answer cites the page number and its context.',
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
        'The Social Graphics Pipeline (work in progress) generates LinkedIn cards and carousels programmatically from a centralized brand kit. Demo coming when v1 ships \u2014 see the <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects">projects folder</a> for the rest.',
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
        'Kaidra is a UI/UX web concept that unifies identity data into one source of truth. <a href="../kaidra/">Open the live Kaidra concept</a>.',
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
        'The <a href="../../dashboard/">CSR Energy &amp; Emissions Intelligence dashboard</a> covers 5,000 companies across 260 countries (2015\u20132018) with KPIs, top-15 ranking, trends and SIC breakdown. Legacy page: <a href="../csr-energy-dashboard/">csr-energy-dashboard</a>.',
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
        'The Data Studio dashboard is built with Google Data Studio (Looker Studio): connected sources, dynamic filters and charts. <a href="https://datastudio.google.com/reporting/d787ca46-5f8c-449d-9c3d-47bf7164d82f/page/3ucKF" target="_blank" rel="noopener">Open the Data Studio dashboard</a>.',
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
        'The NBC Store marketing analysis (Sapienza University, 42 slides) covers communication audit, PESTLE/SWOT and growth strategies. <a href="../../files/sapienza-nbc-marketing-analysis.pdf">View the NBC presentation (PDF)</a>.',
    },
    {
      id: "milano-sport-future-moves",
      title: "Future Moves \u2014 Milano Sport",
      keywords: ["milano sport", "future moves", "fitness", "aquatic", "sport"],
      reply:
        'Future Moves \u2014 Milano Sport (44 slides) is a group market research on fitness macro-trends and new course formats for 2026. <a href="../../files/milano-sport-future-moves.pdf">View the Milano Sport presentation (PDF)</a>.',
    },
    {
      id: "cv",
      title: "CV",
      keywords: ["cv", "resume", "curriculum", "vitae"],
      reply:
        'You can <a href="../../files/Andrea_Barduani_CV.pdf">download the CV (PDF)</a>. For anything missing there, write to <a href="mailto:andrea.barduani@gmail.com">andrea.barduani@gmail.com</a>.',
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
        'You can reach Andrea at <a href="mailto:andrea.barduani@gmail.com">andrea.barduani@gmail.com</a>, on <a href="https://www.linkedin.com/in/andrea-barduani-b76a7719a/" target="_blank" rel="noopener">LinkedIn</a> or via <a href="https://github.com/andreabarduani-ui" target="_blank" rel="noopener">GitHub</a>.',
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
        'The approach is simple: source files are sacred (tools work on copies), messy data is handled by design (fuzzy matching, Italian number formats), the most-used tools ship with double-click launchers, and everything is iterated in the field \u2014 the e-learning monitor reached v17. See <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/elearning-hours-monitor">E-Learning Hours Monitor</a> for the flagship example.',
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
        'Core stack: Python (pandas, openpyxl, python-docx, pdfplumber, PyMuPDF), Excel/Word/PDF automation, web scraping, fuzzy matching, plus HTML/CSS/JavaScript for the Kaidra concept and Chart.js dashboards. Full list in the <a href="../../#skills">skills section</a>.',
    },
    {
      id: "about",
      title: "About",
      keywords: ["about", "who", "andrea", "barduani", "portfolio"],
      reply:
        'Andrea works in the administration of funded vocational training programs \u2014 every tool here solves a real recurring problem from daily work. More in the <a href="../../#about">about section</a>, or write to <a href="mailto:andrea.barduani@gmail.com">andrea.barduani@gmail.com</a>.',
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

  function isDemoRequest(query) {
    var q = normalize(query);
    return (
      q.indexOf(" backend demo ") !== -1 ||
      q.indexOf(" show demo ") !== -1 ||
      q.indexOf(" before after ") !== -1 ||
      q.indexOf(" before ") !== -1 ||
      q.indexOf(" example numbers ") !== -1 ||
      q.indexOf(" show backend ") !== -1
    );
  }

  function demoReplyHtml() {
    var html =
      "<b>E-Learning Hours Monitor \u2014 before / after (pre-computed, fictional data).</b>" +
      '<table class="mini-table"><tr><th>Student</th><th>Before (total)</th><th>Excess stripped</th><th>After (effective)</th></tr>';
    for (var i = 0; i < DEMO_ELEARNING.rows.length; i++) {
      var r = DEMO_ELEARNING.rows[i];
      html +=
        "<tr><td>" +
        r[0] +
        '</td><td class="num">' +
        r[1] +
        '</td><td class="num">' +
        r[2] +
        '</td><td class="num">' +
        r[3] +
        "</td></tr>";
    }
    html +=
      '</table><p class="out-note">' +
      DEMO_ELEARNING.note +
      ' Source: <a href="https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/elearning-hours-monitor">E-Learning Hours Monitor</a>.</p>';
    return html;
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
    if (isDemoRequest(query)) return demoReplyHtml();
    var best = findBest(query);
    if (best) return best.reply;
    return fallbackReplyHtml();
  }

  /* ---------- Optional same-origin demo-data load (progressive enhancement) ----------
     Same-origin only ('self' per CSP). Never blocks: on any error the
     inline DEMO_ELEARNING above is used. No external fetch, no keys. */
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
      "Hi! I can point you to any of the 18 projects, the energy dashboard, the CV or contacts. Try \u201ce-learning hours\u201d or tap a chip below.",
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

  /* Exposed for smoke/Playwright tests (no backend involved). */
  window.__aiAssistant = {
    answerHtml: answerHtml,
    fallbackText: FALLBACK_TEXT,
    knowledgeSize: KNOWLEDGE.length,
  };
})();
