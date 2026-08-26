# Tutoring Hours Distributor

**Phase 2 — Delivery** · Python, openpyxl, fuzzy matching

Distributes tutoring hours across the months of an internship, producing a
proposal that respects every operational constraint.

## What it does

- Reads tutor timesheets (days × time slots) and the intern registry
  (start/end dates), matching interns with **fuzzy name matching**
- Always distributes **starting from the internship start date**
- Respects: **4 hours/month cap**, allowed time windows, weekends excluded,
  tutor's actual availability
- Availability-aware allocation: hours that don't fit a saturated month flow
  into the next one (v3's key fix over v2, which stacked hours in the wrong
  month)
- Output: a proposal workbook with the same layout as the real timesheet +
  a text report — originals are **never modified**; the operator reviews the
  proposal and corrects the real timesheet by hand (two-step workflow)

## Run

```bash
py tutoring_hours_distributor.py
```

## Value

Month-by-month hour allocation used to take an afternoon per tutor and was
still wrong on edge cases (late-starting internships). Now it's a proposal
to review, not a spreadsheet to build.

> Public showcase copy: company and personal identifiers have been generalized.
