# Sample Output — Tutoring Hours Distributor

> Illustrative excerpt with **fictional data**. The tool produces
> `Proposta ore tutoraggio.xlsx` (same layout as the real timesheet)
> and `Report ore tutoraggio.txt`.

Report excerpt:

```
TUTOR: Tutor A — year 2025

Intern: Maria Rossi (internship 10/02/2025 → 30/09/2025, matched fuzzily 0.94)
  Feb 2025: 4h  [cap 4h]  slots: Wed 15:00-17:00, Fri 10:00-12:00
  Mar 2025: 4h  [cap 4h]
  Apr 2025: 4h  [cap 4h]
  May 2025: 4h  [cap 4h]
  TOTAL ASSIGNED: 16h — non-assignable residue: 0h

Intern: Luca Bianchi (internship 01/03/2025 → 31/07/2025)
  Mar 2025: 4h | Apr 2025: 4h | May 2025: 4h | Jun 2025: 4h | Jul 2025: 3h
  TOTAL ASSIGNED: 19h — residue 1h (July saturated, flagged in report)

SUMMARY: 2 interns, 35h assigned, 1h non-assignable (weekend window)
```

The proposal workbook mirrors the timesheet grid (days × time slots) so the
operator can review and manually transfer it to the official timesheet —
the two-step workflow guarantees no original file is ever modified.
