# E-Learning Hours Monitor (v17)

**Phase 2 — Delivery** · Python, pandas, openpyxl, fuzzy matching

Computes each student's **actual** e-learning hours from platform access logs,
automatically stripping out time that is not countable. ~1,800 lines, refined
over 17 versions of real-world use.

## What it does

- Auto-detects input files: platform access CSV + attendance PDF/Excel
- Computes per-student, per-month **effective hours** = total − excess
- Detects excess time: **after 19:00, weekends, early mornings (06:00–07:40),
  beyond the 8h/day cap** (10-minute tolerance)
- Handles the platform PDF's quirk of **truncating names at ~19 characters**,
  via token-based fuzzy matching ("ROSSI MARIO" vs "MARIO ROSSI" are unified)
- Output workbook: one sheet per month + a general summary with the 150-hour
  per-student cap, logos and styling applied automatically

## Run

```bash
py elearning_hours_monitor.py
```

## Value

Reconciling two inconsistent sources by hand for dozens of students each month
was the single most error-prone task in the whole pipeline. This tool ended that.

> Public showcase copy: company and personal identifiers have been generalized.

## Sample output

An illustrative output example (fictional data) is available in [sample-output.md](sample-output.md).
