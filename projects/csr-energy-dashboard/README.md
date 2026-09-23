# CSR Energy & Emissions Intelligence

Interactive dashboard on corporate energy use and emissions intensity across listed US energy-intensive sectors.

- **What:** single-file `index.html` (Chart.js via CDN) + `data/*.json` (KPIs, top-15 intensity ranking, 5-year trend, SIC breakdown). Works offline for KPIs/table with embedded fallback; charts need internet once for the Chart.js CDN.
- **Sources:** World Bank WDI (country energy baselines) and SEC EDGAR company filings aggregated for the thesis pipeline (BigQuery + Power BI upstream, Chart.js web export here).
- **Note:** personal design project — visuals and layout by Andrea Barduani; figures are illustrative of the thesis aggregation workflow.
- **Regenerate from thesis repo:** copy `portfolio/csr-dashboard-v2.html` → `index.html` and `portfolio/data/*.json` → `data/` from `D:\Workspace\tesi-csr\`.
- **Live:** https://andreabarduani-ui.github.io/Portfolio/projects/csr-energy-dashboard/
