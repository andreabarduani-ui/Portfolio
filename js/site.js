/* ============================================================
   Portfolio — site.js
   Extracted from index.html (same functions, zero dependencies).
   DEMOS + playTerm + panels + heroLoop + chips + countUp +
   scroll-spy + reveal + toTop preserved verbatim.
   Addition: light/dark [data-theme] toggle (persisted).
   Works from file:// and Netlify. No npm needed.
   ============================================================ */
(function () {
  'use strict';
  var REDUCED = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var REPO = 'https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/';

  /* ---------- Theme toggle (new, non-breaking) ---------- */
  function initTheme() {
    var KEY = 'portfolio-theme';
    var root = document.documentElement;
    var btn = document.getElementById('themeToggle');
    var saved = null;
    try { saved = localStorage.getItem(KEY); } catch (e) {}
    if (saved === 'dark' || saved === 'light') root.setAttribute('data-theme', saved);
    else if (!root.getAttribute('data-theme')) root.setAttribute('data-theme', 'light');
    if (btn) {
      var sync = function () {
        var dark = root.getAttribute('data-theme') === 'dark';
        btn.textContent = dark ? '☀ Light' : '◐ Dark';
        btn.setAttribute('aria-pressed', String(dark));
      };
      sync();
      btn.addEventListener('click', function () {
        var next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        root.setAttribute('data-theme', next);
        try { localStorage.setItem(KEY, next); } catch (e) {}
        sync();
      });
    }
  }

  function mk(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  /* ============ Demo data (from each project's published sample output) ============ */
  var DEMOS = {

    'student-import-generator': {
      file: 'student_import_generator.py',
      lines: [
        ['cmd', 'py student_import_generator.py'],
        ['ok', 'Read 4 students from Richiesta_dati_allievi.xlsx'],
        ['ok', 'Parsed tax IDs: 4/4 — gender + birth place resolved'],
        ['ok', 'Written: modello-importazione-corsisti.xlsx (4 rows, ready for import)']
      ],
      out: '<table class="mini-table"><tr><th>Cognome</th><th>Nome</th><th>Codice Fiscale</th><th>Sesso</th><th>Comune</th><th>Prov.</th></tr>' +
           '<tr><td>ROSSI</td><td>MARIA</td><td class="num">RSSMRA80A01H501Z</td><td>F</td><td>ROMA</td><td>RM</td></tr>' +
           '<tr><td>BIANCHI</td><td>LUCA</td><td class="num">BNCLCU85C22D612X</td><td>M</td><td>MILANO</td><td>MI</td></tr>' +
           '<tr><td>FERRARI</td><td>ANNA</td><td class="num">FRRNNA92B41A944W</td><td>F</td><td>FIRENZE</td><td>FI</td></tr></table>' +
           '<p class="out-note">Sesso, Comune and Province are derived automatically by parsing the tax ID (gender digits + Belfiore municipality code).</p>'
    },

    'adhesion-letter-compiler': {
      file: 'adhesion_letter_compiler.py',
      lines: [
        ['cmd', 'py adhesion_letter_compiler.py'],
        ['ok', 'Read 3 people (rows 3-5) and 3 companies (rows 20-22)'],
        ['ok', '3 letters generated — original template untouched']
      ],
      out: '<div class="doc-preview"><div class="doc-title">Admission to the training program</div>' +
           '<div class="row"><b>Trainee</b><span>Maria Rossi — born in Rome, 01/01/1980</span></div>' +
           '<div class="row"><b>Company</b><span>Example Logistics S.r.l.</span></div>' +
           '<div class="row"><b>Course</b><span>Warehouse logistics — 1st year</span></div>' +
           '<div class="row"><b>Date</b><span>15/09/2026</span></div></div>' +
           '<div class="file-chips"><span class="file-chip">Lettera Adesione_Example Logistics.docx</span><span class="file-chip">Lettera Adesione_Demo Services.docx</span><span class="file-chip">Lettera Adesione_Sample Consulting.docx</span></div>' +
           '<p class="out-note">Fields are horizontally centered via tab stops computed from the original PDF geometry.</p>'
    },

    'apprenticeship-quote-generator': {
      file: 'apprenticeship_quote_generator.py',
      lines: [
        ['cmd', 'py apprenticeship_quote_generator.py'],
        ['prompt', 'Calendar agreed? (s/n): s'],
        ['prompt', 'Participants: 6'],
        ['prompt', 'Year (1/2): 1'],
        ['prompt', 'Company: Example Manufacturing S.r.l.'],
        ['result', 'Offer number:  20260044  (progressive, auto-incremented)'],
        ['result', 'Total:          2.880,00 EUR'],
        ['result', 'Best offer:     2.600,00 EUR  (x 0.90, rounded to 100)'],
        ['ok', 'Written: output/Offerta_Example_Manufacturing_20260827.docx']
      ],
      out: '<table class="mini-table"><tr><th>Date</th><th>Time</th><th>Hours</th><th>Module</th><th>Mode</th></tr>' +
           '<tr><td class="num">05/10/2026</td><td class="num">09:00–13:00</td><td class="num">4</td><td>Workplace safety</td><td>Classroom</td></tr>' +
           '<tr><td class="num">07/10/2026</td><td class="num">14:00–18:00</td><td class="num">4</td><td>Logistics basics</td><td>Classroom</td></tr>' +
           '<tr><td class="num">12/10/2026</td><td class="num">09:00–13:00</td><td class="num">4</td><td>…</td><td>…</td></tr></table>' +
           '<p class="out-note">Course calendar rebuilt from Excel into a Word table — embedded in the offer while the calendar is still being agreed.</p>'
    },

    'financial-plan-builder': {
      file: 'financial_plan_builder.py',
      lines: [
        ['cmd', 'py financial_plan_builder.py'],
        ['ok', 'Scheda finanziaria built — official cost items A1–A10, B1–B10'],
        ['ok', 'Live formulas wired: change any input, every total recalculates'],
        ['cmd', 'py add_budget_calculator.py'],
        ['ok', 'Constraint checker added (Calcolatore sheet, traffic lights)']
      ],
      out: '<table class="mini-table"><tr><th>Rule</th><th>Status</th><th>Value</th></tr>' +
           '<tr><td>Teaching costs A ≤ 35%</td><td><span class="pill ok">OK</span></td><td class="num">31.2%</td></tr>' +
           '<tr><td>Delivery costs B ≥ 40%</td><td><span class="pill ok">OK</span></td><td class="num">44.7%</td></tr>' +
           '<tr><td>Overhead D ≤ 25%</td><td><span class="pill bad">OVER</span></td><td class="num">26.1% → add 18 delivery hours to fix</td></tr></table>' +
           '<p class="out-note">The repo ships a real generated workbook: <a href="https://github.com/andreabarduani-ui/Portfolio/blob/main/projects/financial-plan-builder/Scheda_Finanziaria_Template.xlsx" target="_blank" rel="noopener">Scheda_Finanziaria_Template.xlsx</a>.</p>'
    },

    'elearning-hours-monitor': {
      file: 'elearning_hours_monitor.py',
      lines: [
        ['cmd', 'py elearning_hours_monitor.py'],
        ['ok', 'Input detected: Report_Accessi.csv + Presenze.pdf (LUL, mode B)'],
        ['ok', 'Months processed: 6 — students: 28'],
        ['dim', 'Excess stripped: evenings >19:00 · weekends · 06:00–07:40 · >8h/day'],
        ['ok', 'Written: Monitoraggio_26-08-2026.xlsx']
      ],
      out: '<table class="mini-table"><tr><th>Student</th><th>Accrued (effective)</th><th>Cap</th><th>Status</th></tr>' +
           '<tr><td>ROSSI MARIA</td><td class="num">148:15</td><td class="num">150h</td><td><span class="pill ok">in cap</span></td></tr>' +
           '<tr><td>BIANCHI LUCA</td><td class="num">150:00</td><td class="num">150h</td><td><span class="pill warn">cap reached</span></td></tr>' +
           '<tr><td>LOMBARDINI BEATRICE</td><td class="num">153:30</td><td class="num">150h</td><td><span class="pill bad">+3:30 over cap</span></td></tr></table>' +
           '<p class="out-note">"Riepilogo Generale" sheet — accrued hours vs the 150-hour per-student cap. One sheet per month, cell notes record the excess breakdown.</p>'
    },

    'tutoring-hours-distributor': {
      file: 'tutoring_hours_distributor.py',
      lines: [
        ['cmd', 'py tutoring_hours_distributor.py'],
        ['ok', 'Tutor A — year 2025'],
        ['out', 'Maria Rossi   Feb 4h [cap 4h] · Mar 4h · Apr 4h · May 4h   → 16h'],
        ['out', 'Luca Bianchi  Mar–Jul 19h assigned, 1h residue flagged'],
        ['result', 'SUMMARY: 2 interns, 35h assigned, 1h non-assignable (weekend window)'],
        ['ok', 'Written: Proposta ore tutoraggio.xlsx + Report ore tutoraggio.txt']
      ],
      out: '<div class="tree">TUTOR: Tutor A — year 2025\n\n' +
           '<span class="dim">Intern:</span> Maria Rossi <span class="dim">(10/02 → 30/09, matched fuzzily 0.94)</span>\n' +
           '  Feb 2025: 4h  <span class="dim">[cap 4h]  slots: Wed 15:00-17:00, Fri 10:00-12:00</span>\n' +
           '  Mar 2025: 4h  <span class="dim">[cap 4h]</span>\n' +
           '  TOTAL ASSIGNED: 16h — non-assignable residue: 0h</div>' +
           '<p class="out-note">The proposal workbook mirrors the official timesheet grid — the two-step workflow guarantees no original file is ever modified.</p>'
    },

    'attendance-register-filler': {
      file: 'attendance_register_filler.py',
      lines: [
        ['cmd', 'py attendance_register_filler.py'],
        ['out', '[1/5] Register found: Registro.pdf (3 pages/session, 57 students)'],
        ['out', '[3/5] Reading 14 session Excel reports'],
        ['out', '[4/5] Matching students (surname+name, similarity ≥ 80%)'],
        ['dim', '      GIORDANETTI ALSSANDRO ~ GIORDANETTI ALESSANDRO  fuzzy 0.97'],
        ['out', '[5/5] Compiled PDF: output/Registro_compilato.pdf'],
        ['result', 'Sessions: 14 — PRESENT labels: 612 — ABSENT labels: 186']
      ],
      out: '<table class="mini-table"><tr><th>#</th><th>Student</th><th>Entry</th><th>Exit</th></tr>' +
           '<tr><td class="num">1</td><td>BIANCHI LUCA</td><td><span class="pill ok">PRESENTE</span></td><td><span class="pill ok">PRESENTE</span></td></tr>' +
           '<tr><td class="num">2</td><td>FERRARI ANNA</td><td><span class="pill bad">ASSENTE</span></td><td><span class="pill bad">ASSENTE</span></td></tr>' +
           '<tr><td class="num">3</td><td>GIORDANETTI ALESSANDRO</td><td><span class="pill ok">PRESENTE</span></td><td><span class="pill ok">PRESENTE</span></td></tr></table>' +
           '<p class="out-note">Inputs are archived in elaborati/&lt;timestamp&gt;/, never deleted; a JSON report records every match decision for auditing.</p>'
    },

    'certificate-manager': {
      file: 'certificate_manager.py',
      lines: [
        ['cmd', 'py certificate_manager.py dividi'],
        ['ok', 'Input: stampa_massiva.pdf (26 pages = 13 people, front+back)'],
        ['out', 'attestati_output/  01_Maria_Rossi.pdf … 13 files (names read from text)'],
        ['cmd', 'py certificate_manager.py organizza'],
        ['ok', 'Sorted into company folders — 12 exact, 1 fuzzy (~), 1 to verify']
      ],
      out: '<div class="tree"><span class="dir">CORSO10/</span>\n' +
           '  <span class="dir">Example Logistics S.r.l./</span> 01_Maria_Rossi.pdf\n' +
           '  <span class="dir">Demo Services S.p.A./</span>     02_Luca_Bianchi.pdf\n' +
           '  <span class="dir">Sample Consulting S.r.l./</span> 03_Anna_Ferrari.pdf\n' +
           '  <span class="dir">_DA_VERIFICARE/</span>          07_persona_7.pdf  <span class="dim">(name not in directory)</span></div>' +
           '<p class="out-note">Fuzzy matches are flagged with ~ in the log, so a human double-checks only the uncertain ones. riorganizza rebuilds the cross-edition company view.</p>'
    },

    'transparency-certificates': {
      file: 'transparency_certificates.py',
      lines: [
        ['cmd', 'py transparency_certificates.py'],
        ['ok', 'Template: Attestato Apprendimenti Acquisiti_.docx (4 tables)'],
        ['ok', 'Master file: File madre.xlsx — 24 participants found'],
        ['ok', '24 certificates generated in Attestati_Generati/'],
        ['ok', 'Institution/responsible data fields: untouched (by design)']
      ],
      out: '<div class="doc-preview"><div class="doc-title">Attestato Apprendimenti Acquisiti</div>' +
           '<div class="row"><b>Student</b><span>Maria Rossi — born in Rome, 01/01/1990</span></div>' +
           '<div class="row"><b>Experience</b><span>Warehouse logistics, 150 effective hours, 09/2026–12/2026</span></div>' +
           '<div class="row"><b>Skills</b><span>Goods handling · Safety compliance · Inventory IT tools</span></div>' +
           '<div class="row"><b>Assessment</b><span>Written test passed (28/30) · Practical evaluation passed</span></div></div>' +
           '<p class="out-note">Only the empty value cells of the template are filled: layout, headers and signature blocks stay exactly as in the original model.</p>'
    },

    'timesheet-generator': {
      file: 'timesheet_generator.py',
      lines: [
        ['cmd', 'py timesheet_generator.py --dry-run'],
        ['ok', 'Interns read from registry: 6 (4 PROGRAM_A, 2 PROGRAM_B)'],
        ['out', 'Maria Rossi    PROGRAM_A  5 months  est. 22 rows'],
        ['out', 'Anna Ferrari   PROGRAM_B  7 months  est. 30 rows'],
        ['dim', '(no files written — dry run)'],
        ['ok', 'py timesheet_generator.py  →  1 file/month + UNICO multi-sheet workbook']
      ],
      out: '<table class="mini-table"><tr><th>Day</th><th>Morning</th><th>Afternoon</th><th>Total</th><th>Activity</th></tr>' +
           '<tr><td class="num">03/09/2026</td><td class="num">09:00–13:00</td><td class="num">—</td><td class="num">4h</td><td>Warehouse procedures</td></tr>' +
           '<tr><td class="num">04/09/2026</td><td class="num">09:00–13:00</td><td class="num">14:00–16:00</td><td class="num">6h</td><td>Inventory tools</td></tr>' +
           '<tr><td class="num">10/09/2026</td><td class="num">—</td><td class="num">14:00–18:00</td><td class="num">4h</td><td>Safety refresher</td></tr></table>' +
           '<p class="out-note">The official template is extended with the project-name row on a copy ("v2") — the original template file is never modified.</p>'
    },

    'funding-call-scraper': {
      file: 'funding_calls_scraper.py',
      lines: [
        ['cmd', 'py funding_calls_scraper.py'],
        ['ok', 'Scraping 9 funding bodies (2 levels: list page → call detail)'],
        ['ok', 'robots.txt respected on all sites'],
        ['out', 'FonARCom ......... 4 calls read, 1 above threshold'],
        ['out', 'Fonditalia ....... 7 calls read, 2 above threshold'],
        ['result', '38 calls total — report_bandi_fondi.xlsx written']
      ],
      out: '<table class="mini-table"><tr><th>Fund</th><th>Budget</th><th>Confidence</th><th>Restricted</th></tr>' +
           '<tr><td>Fonditalia</td><td class="num">€2,500,000</td><td><span class="pill ok">High</span></td><td>No</td></tr>' +
           '<tr><td>FonARCom</td><td class="num">€1,100,000</td><td><span class="pill warn">Medium</span></td><td>Yes</td></tr>' +
           '<tr><td>Fondir</td><td class="num">€950,000</td><td><span class="pill bad">Low</span></td><td>No</td></tr></table>' +
           '<p class="out-note">"Target" sheet (calls above €800,000). Confidence reflects how explicitly each budget is stated; Italian number formats parsed (dot = thousands).</p>'
    },

    'regulation-search': {
      file: 'regulation_search.py',
      lines: [
        ['cmd', 'py regulation_search.py manual.pdf'],
        ['out', 'Indexed: 148 pages'],
        ['prompt', 'Choice: 3  (eligible costs)'],
        ['out', 'p. 41  «trainee costs are eligible up to the hourly rate'],
        ['out', '        defined in Annex B…»          [match: eligible costs]'],
        ['out', 'p. 55  «catering is eligible only for full-day sessions…»'],
        ['result', '3 hits on 3 pages — every result cites the page number']
      ],
      out: '<div class="tree"><span class="dim">$</span> py regulation_search.py manual.pdf\n\n' +
           '  CATEGORIES\n' +
           '   1. Plan variations    2. Reporting deadlines\n' +
           '   3. Eligible costs     4. Training modes\n' +
           '   5. State aid regimes  0. Free text search\n\n' +
           '<span class="dim">Choice: 0 →</span> "variations deadline"\n' +
           '  p. 17  «variation requests must be submitted at least 30\n' +
           '          days before the end of the program…»</div>' +
           '<p class="out-note">The page number is always returned, so the answer can be verified on the source PDF immediately.</p>'
    }
  };

  /* ============ Terminal typing engine ============ */
  function buildLine(data) {
    var type = data[0], text = data[1];
    var line = mk('div', 'tline ' + type);
    var pref = { cmd: '$', ok: '>>', prompt: '?' }[type];
    if (pref) line.appendChild(mk('span', 'ps', pref + ' '));
    line.appendChild(mk('span', 'tx', text));
    return line;
  }

  function playTerm(body, demo, token) {
    body.textContent = '';
    var cursor = mk('div', 'tline cursor');
    body.appendChild(cursor);
    var alive = function () { return token.alive; };

    if (REDUCED) {
      demo.lines.forEach(function (l) { body.insertBefore(buildLine(l), cursor); });
      return Promise.resolve();
    }

    return demo.lines.reduce(function (p, data) {
      return p.then(function () {
        if (!alive()) return;
        if (data[0] === 'cmd') {
          var line = buildLine(data);
          var tx = line.querySelector('.tx');
          tx.textContent = '';
          body.insertBefore(line, cursor);
          var chars = data[1].split('');
          return chars.reduce(function (q, ch) {
            return q.then(function () {
              if (!alive()) return;
              tx.textContent += ch;
              return sleep(13);
            });
          }, Promise.resolve()).then(function () { return sleep(300); });
        }
        body.insertBefore(buildLine(data), cursor);
        return sleep(data[0] === 'result' ? 150 : data[0] === 'ok' ? 115 : 70);
      });
    }, Promise.resolve());
  }

  /* ============ Build expandable panels ============ */
  function initPanels() {
    document.querySelectorAll('[data-demo]').forEach(function (card) {
      var slug = card.dataset.demo;
      var demo = DEMOS[slug];
      if (!demo) return;
      var slot = card.querySelector('.panel-slot');
      var runBtn = card.querySelector('.run');
      if (!slot) return;

      var panel = mk('div', 'panel');
      panel.id = 'panel-' + slug;
      panel.hidden = true;

      var tabs = mk('div', 'tabs');
      var tTerm = mk('button', 'tab active', 'Terminal');
      var tOut = mk('button', 'tab', 'Output');
      tTerm.type = tOut.type = 'button';
      tTerm.dataset.tab = 'term'; tOut.dataset.tab = 'out';
      tabs.appendChild(tTerm); tabs.appendChild(tOut);

      var term = mk('div', 'term');
      var bar = mk('div', 'term-bar');
      bar.appendChild(mk('span', 'dot r'));
      bar.appendChild(mk('span', 'dot y'));
      bar.appendChild(mk('span', 'dot g'));
      var title = mk('span', 'term-title', demo.file);
      bar.appendChild(title);
      var tbody = mk('div', 'term-body');
      term.appendChild(bar); term.appendChild(tbody);

      var out = mk('div', 'out');
      out.hidden = true;
      out.innerHTML = demo.out;
      var foot = mk('p', 'out-note');
      var link = mk('a', null, 'Full sample output on GitHub →');
      link.href = REPO + slug; link.target = '_blank'; link.rel = 'noopener';
      foot.appendChild(document.createTextNode('Illustrative excerpt, fictional data · '));
      foot.appendChild(link);
      out.appendChild(foot);

      panel.appendChild(tabs); panel.appendChild(term); panel.appendChild(out);
      slot.appendChild(panel);
      if (runBtn) runBtn.setAttribute('aria-controls', panel.id);

      var token = { alive: false };

      function selectTab(which) {
        tTerm.classList.toggle('active', which === 'term');
        tOut.classList.toggle('active', which === 'out');
        term.hidden = which !== 'term';
        out.hidden = which !== 'out';
      }
      tTerm.addEventListener('click', function () { selectTab('term'); });
      tOut.addEventListener('click', function () { selectTab('out'); });

      if (runBtn) runBtn.addEventListener('click', function () {
        var open = panel.hidden;
        panel.hidden = !open;
        runBtn.setAttribute('aria-expanded', String(open));
        runBtn.classList.toggle('open', open);
        runBtn.querySelector('.lbl').textContent = open ? 'Hide' : 'See it run';
        if (open) {
          token.alive = false;
          token = { alive: true };
          playTerm(tbody, demo, token);
        } else {
          token.alive = false;
        }
      });
    });
  }

  /* ============ Hero terminal loop ============ */
  function initHero() {
    var HERO = [
      { file: 'elearning_hours_monitor.py', lines: DEMOS['elearning-hours-monitor'].lines },
      { file: 'attendance_register_filler.py', lines: DEMOS['attendance-register-filler'].lines },
      { file: 'funding_calls_scraper.py', lines: DEMOS['funding-call-scraper'].lines },
      { file: 'student_import_generator.py', lines: DEMOS['student-import-generator'].lines }
    ];
    var heroTerm = document.querySelector('.hero-term');
    var body = document.querySelector('.hero-term .term-body');
    var title = document.querySelector('.hero-term .term-title');
    if (!heroTerm || !body) return;

    if (REDUCED) {
      title.textContent = HERO[0].file;
      playTerm(body, HERO[0], { alive: true });
      return;
    }
    var visible = true;
    if ('IntersectionObserver' in window) {
      new IntersectionObserver(function (en) { visible = en[0].isIntersecting; }).observe(heroTerm);
    }

    var i = 0;
    (function next() {
      title.textContent = HERO[i].file;
      playTerm(body, HERO[i], { alive: true }).then(function () {
        setTimeout(function () {
          var wait = function () {
            if (visible) { i = (i + 1) % HERO.length; next(); }
            else setTimeout(wait, 400);
          };
          wait();
        }, 3000);
      });
    })();
  }

  /* ============ Filter chips (python|web|data|strategy, default python, single grid) ============ */
  function normFilter(f) {
    return (f === 'python' || f === 'web' || f === 'data' || f === 'strategy') ? f : 'python';
  }
  function applyFilter(chips, f, pushHash) {
    f = normFilter(f);
    chips.forEach(function (c) {
      var on = c.dataset.filter === f;
      c.classList.toggle('active', on);
      c.setAttribute('aria-pressed', String(on));
    });
    var cards = Array.prototype.slice.call(document.querySelectorAll('#projects .card'));
    cards.forEach(function (card) {
      var show = card.dataset.phase === f;
      if (!REDUCED && !show) { card.classList.add('filtering'); }
      card.hidden = !show;
      if (show) { card.classList.remove('filtering'); void card.offsetWidth; }
    });
    // no .group-head in DOM anymore: single filtered grid, context lives in chips
    var i = 0;
    cards.forEach(function (card) {
      if (!card.hidden) { card.style.setProperty('--i', String(i % 6)); i++; }
    });
    if (pushHash && history.replaceState) {
      history.replaceState(null, '', '#projects-' + f);
    }
  }
  function initChips() {
    var chips = document.querySelectorAll('#projects .chip');
    chips.forEach(function (chip) {
      chip.addEventListener('click', function () { applyFilter(chips, chip.dataset.filter, true); });
    });
    var m = (location.hash || '').match(/^#projects(?:-(python|web|data|strategy))?$/);
    applyFilter(chips, m ? (m[1] || 'python') : 'python', false);
    window.addEventListener('hashchange', function () {
      var mm = (location.hash || '').match(/^#projects(?:-(python|web|data|strategy))?$/);
      if (mm) { applyFilter(chips, mm[1] || 'python', false); }
    });
  }

  /* ============ Stats count-up ============ */
  function countUp(el) {
    var target = +el.dataset.count;
    if (REDUCED) { el.textContent = target; return; }
    var t0 = null, dur = 1100;
    function step(now) {
      if (!t0) t0 = now;
      var p = Math.min(1, (now - t0) / dur);
      el.textContent = Math.round(target * (1 - Math.pow(1 - p, 3)));
      if (p < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
  }
  function initStats() {
    var statsDone = false;
    if (!('IntersectionObserver' in window)) {
      document.querySelectorAll('.stat b').forEach(countUp);
      return;
    }
    var statsIo = new IntersectionObserver(function (en) {
      if (en[0].isIntersecting && !statsDone) {
        statsDone = true;
        document.querySelectorAll('.stat b').forEach(function (el, i) {
          setTimeout(function () { countUp(el); }, i * 130);
        });
        statsIo.disconnect();
      }
    });
    var statsEl = document.querySelector('.stats');
    if (statsEl) statsIo.observe(statsEl);
  }

  /* ============ Scroll-spy ============ */
  function initSpy() {
    if (!('IntersectionObserver' in window)) return;
    var spy = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (!en.isIntersecting) return;
        document.querySelectorAll('.nav ul a').forEach(function (a) {
          a.classList.toggle('active', a.hash === '#' + en.target.id);
        });
      });
    }, { rootMargin: '-40% 0px -55% 0px' });
    ['projects', 'skills', 'about', 'contact'].forEach(function (id) {
      var s = document.getElementById(id);
      if (s) spy.observe(s);
    });
  }

  /* ============ Reveal + stagger + AOS (CDN, offline fallback) ============ */
  function initReveal() {
    document.querySelectorAll('#projects .card').forEach(function (card, idx) {
      card.style.setProperty('--i', String(idx % 6));
    });
    if (!('IntersectionObserver' in window) || REDUCED) {
      document.querySelectorAll('.reveal').forEach(function (el) { el.classList.add('in'); });
      return;
    }
    var ro = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (e.isIntersecting) { e.target.classList.add('in'); ro.unobserve(e.target); }
      });
    }, { threshold: 0.06 });
    document.querySelectorAll('.reveal').forEach(function (el) { ro.observe(el); });
    // AOS pattern from references/animations/aos — once, offset 80
    var boot = function () {
      if (REDUCED) return;
      var s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/aos/2.3.4/aos.js';
      s.defer = true;
      s.onload = function () {
        try { if (window.AOS) { window.AOS.init({ once: true, offset: 80, duration: 600 }); } } catch (e) {}
      };
      document.head.appendChild(s);
    };
    if ('requestIdleCallback' in window) { requestIdleCallback(boot); }
    else { setTimeout(boot, 1200); }
  }

  /* ============ Tilt (pointer:fine) + scroll progress ============ */
  function initMotion() {
    if (!REDUCED && window.matchMedia('(pointer: fine)').matches) {
      document.querySelectorAll('#projects .card').forEach(function (card) {
        var raf = null;
        card.addEventListener('pointermove', function (ev) {
          if (raf) return;
          raf = requestAnimationFrame(function () {
            var r = card.getBoundingClientRect();
            var px = (ev.clientX - r.left) / r.width - 0.5;
            var py = (ev.clientY - r.top) / r.height - 0.5;
            card.style.setProperty('--ry', (px * 5).toFixed(2) + 'deg');
            card.style.setProperty('--rx', (-py * 5).toFixed(2) + 'deg');
            card.classList.add('tilt');
            raf = null;
          });
        });
        card.addEventListener('pointerleave', function () {
          card.classList.remove('tilt');
          card.style.setProperty('--rx', '0deg');
          card.style.setProperty('--ry', '0deg');
        });
      });
    }
  }

  /* ============ Progress bar + Back to top ============ */
  function initToTop() {
    var toTop = document.getElementById('toTop');
    if (!toTop) return;
    var prog = document.getElementById('scrollProgress');
    var onScroll = function () {
      toTop.classList.toggle('show', window.scrollY > 700);
      if (prog && !REDUCED) {
        var h = document.documentElement.scrollHeight - window.innerHeight;
        var p = h > 0 ? Math.min(1, window.scrollY / h) : 0;
        prog.style.transform = 'scaleX(' + p.toFixed(4) + ')';
      }
    };
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
    toTop.addEventListener('click', function () {
      window.scrollTo({ top: 0, behavior: REDUCED ? 'auto' : 'smooth' });
    });
  }

  initTheme();
  initPanels();
  initHero();
  initChips();
  initStats();
  initSpy();
  initReveal();
  initMotion();
  initToTop();
})();
