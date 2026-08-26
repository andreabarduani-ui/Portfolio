# Apprenticeship Quote Generator

**Phase 1 — Program Launch** · Python, python-docx, openpyxl

Generates complete economic offers for private apprenticeship courses, from a
faithful copy of the official template (logo, header, footer, fonts preserved).

## What it does

- Interactive input: calendar agreed or not, number of participants, year
  (first/second), recipient company
- Price calculation: unit price × participants, "best offer" rounding to the
  nearest hundred
- **Progressive offer numbering** per year, persisted in a JSON counter
- If the calendar is not agreed yet: reads the course calendar Excel and
  rebuilds it as a formatted 6-column Word table (Date | Hours | Credits |
  Module | Teacher | Mode), including the edition code detected from the
  sheet header
- Output: `output/Offerta_<Company>_<YYYYMMDD>.docx`

## Run

```bash
py apprenticeship_quote_generator.py
```

## Value

Each quote went from manual editing of a Word file (with numbering mistakes
and stale dates) to a guided, consistent, automatically-numbered generation.

> Public showcase copy: company and personal identifiers have been generalized.

## Sample output

An illustrative output example (fictional data) is available in [sample-output.md](sample-output.md).
