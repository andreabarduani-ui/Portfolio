# Sample Output — E-Learning Hours Monitor

> Illustrative excerpt with **fictional data**. The tool produces
> `Monitoraggio_<date>.xlsx`: one sheet per month + a general summary.

Monthly sheet (e.g. "Marzo 2026") — per student:

| Student | Effective hours | Total hours | Total excess | Evening excess | Weekend excess | Early-morning excess |
|---|---|---|---|---|---|---|
| ROSSI MARIA | 62:15:00 | 71:40:00 | 9:25:00 | 5:10:00 | 4:15:00 | 0:00:00 |
| BIANCHI LUCA | 48:30:00 | 49:00:00 | 0:30:00 | 0:30:00 | 0:00:00 | 0:00:00 |
| GIORDANETTI ALESSANDRO | 55:05:00 | 63:45:00 | 8:40:00 | 3:20:00 | 2:00:00 | 3:20:00 |

Daily cells contain **effective hours**; a cell note records the original
total and the excess breakdown when present.

"Riepilogo Generale" sheet — accrued hours vs the 150-hour per-student cap:

| Student | Accrued (effective) | Cap | Status |
|---|---|---|---|
| ROSSI MARIA | 148:15:00 | 150 | 🟢 in cap |
| BIANCHI LUCA | 150:00:00 | 150 | ⚪ cap reached |
| LOMBARDINI BEATRICE | 153:30:00 | 150 | 🔴 +3:30 over cap |

Console:

```
>> Input detected: Report_Accessi.csv + Presenze.pdf (LUL, mode B)
>> Months processed: 6 — students: 28
>> Excess time stripped: evenings >19:00, weekends, 06:00-07:40, >8h/day
>> Written: Monitoraggio_26-08-2026.xlsx
```
