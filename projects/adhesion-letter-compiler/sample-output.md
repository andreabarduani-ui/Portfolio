# Sample Output — Adhesion Letter Compiler

> Illustrative excerpt with **fictional data**. The tool produces one Word
> letter per company: `Lettera Adesione_<Company>.docx`.

Input pairing (positional: first person ↔ first company) and result:

| # | Person | Company | Output file |
|---|---|---|---|
| 1 | Maria Rossi | Example Logistics S.r.l. | Lettera Adesione_Example Logistics.docx |
| 2 | Luca Bianchi | Demo Services S.p.A. | Lettera Adesione_Demo Services.docx |
| 3 | Anna Ferrari | Sample Consulting S.r.l. | Lettera Adesione_Sample Consulting.docx |

Each generated letter has the template boxes filled and **horizontally
centered** (tab stops computed from the original PDF geometry):

```
        ADMISION TO THE TRAINING PROGRAM

   Trainee:        Maria Rossi
   Born in:        Rome, 01/01/1980
   Company:        Example Logistics S.r.l.
   Course:         Warehouse logistics — 1st year
   Date:           15/09/2026
```

Console:

```
>> Read 3 people (rows 3-5) and 3 companies (rows 20-22)
>> 3 letters generated — original template untouched
```
