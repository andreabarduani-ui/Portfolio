# Sample Output — Apprenticeship Quote Generator

> Illustrative excerpt with **fictional data**. The tool produces
> `output/Offerta_<Company>_<YYYYMMDD>.docx`.

Console session (interactive):

```
  >> Calendar agreed? (s/n): s
  >> Participants: 6
  >> Year (1/2): 1
  >> Company: Example Manufacturing S.r.l.

  Offer number:   20260044        (progressive, auto-incremented)
  Unit price:     480,00 EUR x 6 participants
  Total:          2.880,00 EUR
  Best offer:     2.600,00 EUR    (x 0.90, rounded to nearest hundred)
```

When the calendar is **not** agreed yet, the offer also embeds the course
calendar rebuilt from Excel into a 6-column Word table:

| Date | Time | Hours | Module | Teacher | Mode |
|---|---|---|---|---|---|
| 05/10/2026 | 09:00–13:00 | 4 | Workplace safety | Tutor A | Classroom |
| 07/10/2026 | 14:00–18:00 | 4 | Logistics basics | Tutor B | Classroom |
| 12/10/2026 | 09:00–13:00 | 4 | … | … | … |

The offer number counter is persisted in `contatore_offerte.json`, separate
per year, so numbering never collides.
