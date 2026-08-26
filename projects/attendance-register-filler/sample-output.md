# Sample Output — Attendance Register Filler

> Illustrative excerpt with **fictional data**. The tool produces
> `output/Registro_compilato.pdf` (+ JSON report + page previews).

Console run:

```
[1/5] Register found: Registro.pdf (3 pages/session, 57 students)
[2/5] Layout detected: B (single clean row in student list)
[3/5] Reading 14 session Excel reports
[4/5] Matching students to sessions (surname+name, similarity >= 80%)
      - "BIANCHI LUCA"       matched exactly          (14/14 sessions)
      - "GIORDANETTI ALSSANDRO" ~ "GIORDANETTI ALESSANDRO"  fuzzy 0.97
      - "LOMBARDINI BEATRIC"    ~ "LOMBARDINI BEATRICE"     fuzzy 0.98
[5/5] Compiled PDF: output/Registro_compilato.pdf

Sessions: 14 — PRESENT labels: 612 — ABSENT labels: 186
```

Result on each register page (fictional):

| # | Student | Entry | Exit |
|---|---|---|---|
| 1 | BIANCHI LUCA | PRESENTE | PRESENTE |
| 2 | FERRARI ANNA | ASSENTE | ASSENTE |
| 3 | GIORDANETTI ALESSANDRO | PRESENTE | PRESENTE |

Input files are archived in `elaborati/<timestamp>/`, never deleted, and the
JSON report records every match decision for auditing.
