# Portfolio AI Assistant (v0)

Rule-based English-only assistant for Andrea Barduani's portfolio: 18 projects, energy dashboard, CV and contacts. $0, no backend, no keys — runs entirely in your browser.

v0 is English-only, rule-based, isolated: no keys, no backend, $0 static demo with pre-computed demo-data JSON.
Planned: v1 LLM backend with grounded answers, v2 Telegram /demo on allowlisted WSL2 (real execution isolated, never via public site).

## What it is

- `index.html` — isolated assistant page (EN only), links back to portfolio.
- `chat-widget.js` — keyword matching, inline fallback demo, optional same-origin `demo-data/*.json` fetch (CSP `connect-src 'self'`).
- `chat-widget.css` — minimal styles reusing site variables.
- `demo-data/` — static pre-computed outputs (no live Python, no scraping):
  - `elearning-hours-monitor.json` (from `projects/elearning-hours-monitor/sample-output.md`)
  - `funding-call-scraper.json` (from `projects/funding-call-scraper/sample-output.md`)

## Run locally

Just open `index.html` in a browser — no build, no keys, no backend.

## Sources

- E-learning monitor: https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/elearning-hours-monitor
- Funding scraper: https://github.com/andreabarduani-ui/Portfolio/tree/main/projects/funding-call-scraper
