# Sample Output — Financial Plan Builder

> This repository includes a **real generated file**:
> [Scheda_Finanziaria_Template.xlsx](Scheda_Finanziaria_Template.xlsx)
> (produced by `py financial_plan_builder.py`, with fictional example data).

Structure of the generated workbook:

- **Input cells (yellow)**: target budget, percentage split, people, €/h, hours
- **Live formulas**: change any input and every total recalculates
- Official cost items A1–A10 and B1–B10

Example rows (illustrative):

| Item | Description | Person | €/h | Hours | Total |
|---|---|---|---|---|---|
| A5 | Coordination | Coordinatore A | 25.54 | 40 | 1,021.60 |
| B2 | Tutoring (2025) | Tutor C | 19.02 | 82 | 1,559.64 |
| B10 | Classroom monitoring | Monitor D | 20.30 | 52 | 1,055.60 |

Constraint checks with traffic lights (from the `Calcolatore` sheet added by
`add_budget_calculator.py`):

| Rule | Status | Value |
|---|---|---|
| Teaching costs A ≤ 35% | 🟢 OK | 31.2% |
| Delivery costs B ≥ 40% | 🟢 OK | 44.7% |
| Overhead D ≤ 25% | 🔴 OVER | 26.1% → add 18 delivery hours to fix |
