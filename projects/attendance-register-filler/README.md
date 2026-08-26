# Attendance Register Filler

**Phase 2 — Delivery** · Python, pdfplumber, PyMuPDF, openpyxl

Fills the official **PDF attendance register** of training courses from the
session attendance Excel reports. The most-used tool of the whole portfolio:
colleagues run it with a double click, no technical skills required.

## What it does

- Discovers the register PDF and the session reports in `input/`
- Extracts the student list **from the register itself** (the single source
  of truth), auto-detecting the page layout (two layouts supported)
- Matches students to each session's report with normalized
  **surname+name comparison and ≥80% similarity** — typos in names are
  handled automatically
- Writes **PRESENT/ABSENT** in the exact entry/exit cells of each session page
- Archives processed inputs in `elaborati/<timestamp>/` (never deletes),
  writes the compiled PDF to `output/` and opens it for review
- Ships with a `COMPILA.bat` launcher: checks files, installs dependencies if
  missing, runs, archives, opens the result

## Run

```bash
py attendance_register_filler.py          # simple
py attendance_register_filler.py --archive
# or just double-click COMPILA.bat in the original deployment
```

## Value

Compiling 57 students × dozens of session pages by hand each month was pure
typing work with real error risk. It is now automatic, consistent and
auditable.

> Public showcase copy: company and personal identifiers have been generalized.

## Sample output

An illustrative output example (fictional data) is available in [sample-output.md](sample-output.md).
