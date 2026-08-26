# Automation Projects — Source Code

Python source code for the 13 automation tools shown on the
[portfolio website](https://andreabarduani-ui.github.io/Portfolio/).
Each tool solves a real, recurring problem in the administration of funded
vocational training programs.

> **Privacy note** — these are public showcase copies: company identifiers,
> client names, personal names and internal paths have been generalized.
> The logic is fully intact and the scripts run as-is with the input files
> named in each project's README.

## Projects by phase

### Phase 1 — Program Launch
| Project | What it does |
|---|---|
| [student-import-generator](student-import-generator/) | Builds the platform-ready enrollment file, enriching records from the Italian tax ID (gender, birth town/province via embedded municipality database) |
| [adhesion-letter-compiler](adhesion-letter-compiler/) | Fills Word adhesion-letter templates from Excel, one letter per trainee–company pair |
| [apprenticeship-quote-generator](apprenticeship-quote-generator/) | Generates apprenticeship course offers: pricing rules, progressive numbering, calendar tables from Excel |
| [financial-plan-builder](financial-plan-builder/) | Live-formula Excel budget with official cost items and traffic-light constraint checks |

### Phase 2 — Delivery
| Project | What it does |
|---|---|
| [elearning-hours-monitor](elearning-hours-monitor/) | Computes real e-learning hours per student, splitting out excess time (evenings, weekends, daily cap) — v17 |
| [tutoring-hours-distributor](tutoring-hours-distributor/) | Distributes tutoring hours across internship months under real constraints |
| [attendance-register-filler](attendance-register-filler/) | Fills the official PDF attendance register from session reports, with fuzzy name matching |

### Phase 3 — Reporting
| Project | What it does |
|---|---|
| [certificate-manager](certificate-manager/) | Splits bulk-print certificate PDFs, sorts them per company, rebuilds cross-edition views |
| [transparency-certificates](transparency-certificates/) | Generates end-of-course "acquired skills" certificates, one per participant |
| [timesheet-generator](timesheet-generator/) | Compiles monthly project timesheets from intern records and tutor reports |

### Cross-cutting utilities
| Project | What it does |
|---|---|
| [funding-call-scraper](funding-call-scraper/) | Two-level scraper of 9 interprofessional funding bodies, Excel report of calls above threshold |
| [regulation-search](regulation-search/) | Interactive search inside funding manuals: predefined categories + free text, with page numbers |

## Tech stack

Python · pandas · openpyxl · python-docx · pdfplumber · PyMuPDF · pypdf ·
requests · BeautifulSoup · Pillow · PyYAML
