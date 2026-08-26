# Sample Output — Timesheet Generator

> Illustrative excerpt with **fictional data**. The tool produces, per intern
> with activity in the reporting year: one Excel per month
> (`Timesheet_<Name>_<YYYY-MM>.xlsx`) and a single multi-sheet workbook
> (`Timesheet_<Name>_UNICO_<YYYY>.xlsx`).

Console (`--dry-run` mode):

```
>> Interns read from registry: 6 (4 PROGRAM_A, 2 PROGRAM_B)
>> Tutor timesheets: 2 (Tutor A, Tutor B)
>> Reporting year: 2026

  Maria Rossi     PROGRAM_A  activity months: 5  est. rows: 22
  Luca Bianchi    PROGRAM_A  activity months: 3  est. rows: 14
  Anna Ferrari    PROGRAM_B  activity months: 7  est. rows: 30
  (no files written --dry-run)
```

One monthly sheet (fictional):

| Day | Morning | Afternoon | Total | Activity note |
|---|---|---|---|---|
| 03/09/2026 | 09:00–13:00 | — | 4h | Warehouse procedures |
| 04/09/2026 | 09:00–13:00 | 14:00–16:00 | 6h | Inventory tools |
| 10/09/2026 | — | 14:00–18:00 | 4h | Safety refresher |

The official template is extended with the project-name row **on a copy**
("v2") — the original template file is never modified.
