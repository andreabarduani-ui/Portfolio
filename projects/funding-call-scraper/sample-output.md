# Sample Output — Funding Call Scraper

> Illustrative excerpt. Fund websites are public; call titles below are
> **fictional examples** of the report format.

Console:

```
>> Scraping 9 funding bodies (2 levels: list page -> call detail)
>> robots.txt respected on all sites

  FonARCom ......... 4 calls read, 1 above threshold
  Fonditalia ....... 7 calls read, 2 above threshold
  Fondimpresa ...... 5 calls read, 0 above threshold
  ... (9 bodies, 38 calls total)
```

`report_bandi_fondi.xlsx` — "Target" sheet (calls above €800,000):

| Fund | Call | Budget | Confidence | Restricted program |
|---|---|---|---|---|
| Fonditalia | *Call title (fictional)* | €2,500,000 | High | No |
| FonARCom | *Call title (fictional)* | €1,100,000 | Medium | Yes |
| Fondir | *Call title (fictional)* | €950,000 | Low | No |

Budget extraction confidence: **High** = explicit "financial endowment → €X";
**Medium** = "X millions" wording; **Low** = generic amount found in text
(Italian number format parsed: dot = thousands, comma = decimals).
