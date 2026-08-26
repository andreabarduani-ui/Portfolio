# Timesheet Generator

**Phase 3 — Reporting** · Python, openpyxl, PyYAML

Compiles the official monthly timesheets of publicly-funded internship
programs (attachment A.7 format) from intern records and tutor hour reports.

## What it does

- Reads the intern registry (name, program, internship period) and the tutor
  timesheets
- For each intern with activity in the reporting year, generates:
  1. one Excel file per month (`Timesheet_<Name>_<YYYY-MM>.xlsx`)
  2. a single workbook with one sheet per month (`Timesheet_<Name>_UNICO_<YYYY>.xlsx`)
- Extends the official template with the required project-name row **on a
  copy** ("v2") — the original template is never modified
- Program assignment per intern is driven by a YAML configuration

## Run

```bash
py timesheet_generator.py              # full run
py timesheet_generator.py --dry-run    # extraction + report, no files
py timesheet_generator.py --solo-unici # only the single-workbook output
```

## Value

One file per intern per month, manually copied from the template and filled
from scattered sources — now generated in batch with a dry-run mode for
verification before writing anything.

> Public showcase copy: company and personal identifiers have been generalized.

## Sample output

An illustrative output example (fictional data) is available in [sample-output.md](sample-output.md).
