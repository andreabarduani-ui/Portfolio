# Financial Plan Builder

**Phase 1 — Program Launch** · Python, openpyxl

Two complementary tools for building the financial budget of funded training
plans, with live formulas and constraint checking.

## financial_plan_builder.py

Generates `Scheda_Finanziaria_Template.xlsx` with:

- The funding body's official cost items (A1–A10, B1–B10)
- **Live formulas**: change an hour or a rate and everything recalculates
- Yellow input cells, automatic totals
- **Traffic-light constraint checks** (teaching ≤ 35%, delivery ≥ 40%,
  overhead ≤ 25%) and a simulator that tells how many hours are missing to
  reach the target

## add_budget_calculator.py

Adds a "Calcolatore" sheet to an existing budget workbook, linking to its
totals **without touching the original sheets**: current status, constraint
lights, target-gap simulator.

## Run

```bash
py financial_plan_builder.py
py add_budget_calculator.py
```

## Value

Budgets used to be checked by hand against the fund's percentage rules. The
calculator sheet turns compliance into a glance-at-a-traffic-light check.

> Public showcase copy: company and personal identifiers have been generalized.
