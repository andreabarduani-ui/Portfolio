# Funding Call Scraper

**Cross-cutting utility** · Python, requests, BeautifulSoup, pandas

Monitors the websites of the main Italian interprofessional training funds and
finds **open funding calls above a budget threshold**, exported as a
ready-to-read Excel report.

## What it does

- **Two-level scraping**: level 1 discovers call links on list pages; level 2
  opens each call page and extracts the full text, because budgets and
  program references only appear in the detail pages
- Extracts the funding amount with **three confidence patterns** (explicit
  "financial endowment → €X", "X millions", generic amounts), correctly
  parsing the Italian number format (dot as thousands separator)
- Flags calls above a configurable threshold (default €800,000) and whether
  they belong to a restricted program family
- Respects `robots.txt`
- Output: text report + Excel workbook with Riepilogo / All calls / Target sheets

## Run

```bash
py funding_calls_scraper.py
```

## Value

Checking nine fund websites by hand every week became a single command that
produces a filtered, sorted report of the opportunities that matter.

> Public showcase copy: fund websites monitored are public institutions.

## Sample output

An illustrative output example (fictional data) is available in [sample-output.md](sample-output.md).
