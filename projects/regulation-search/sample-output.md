# Sample Output — Regulation Search

> Illustrative session with a **fictional** manual (no PII). Live sandbox
> proof: [demo-transcript.txt](demo-transcript.txt) (WSL2, tmpfs,
> timeout 10s, no secrets) with fixture [demo-fixture.txt](demo-fixture.txt).

## INPUT

- Any regulation PDF, indexed page by page (fictional manual, 148 pages)
- Predefined categories (variations, reporting, eligible costs, training
  modes, state aid, ...) + free-text search

## COMMAND

```bash
py regulation_search.py manual.pdf
# or without arguments: the script asks for the path interactively
```

## OUTPUT

```
$ py regulation_search.py manual.pdf

  Indexed: 148 pages

  CATEGORIES
   1. Plan variations
   2. Reporting (deadlines, documents)
   3. Eligible costs
   4. Training modes
   5. State aid regimes
   0. Free text search

  Choice: 3   (eligible costs)

  --- ELIGIBLE COSTS --------------------------------

  p. 41  «...trainee costs are eligible up to the hourly rate
          defined in Annex B...»            [match: eligible costs]

  p. 55  «...catering is an eligible cost only for full-day
          classroom sessions...»            [match: eligible]

  p. 92  «...documentation for cost eligibility must be retained
          for 10 years...»                  [match: cost eligibility]

  ----------------------------------------------------
  3 hits on 3 pages

  Choice: 0 -> "variations deadline"

  p. 17  «...variation requests must be submitted at least 30
          days before the end of the program...»
          [match: variations + deadline]
```

Every result includes the **page number**, so the answer can be verified on
the source PDF immediately.
