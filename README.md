# Excel Intelligence Dashboard

A browser-only React/Vite dashboard that reads `.xlsx`, `.xls`, and `.csv` files locally.

## Highlights
- Automatic header and column-type detection
- Multi-sheet workbooks
- Date, category and search filters generated from the data
- KPI cards, trends, monthly totals, rankings and mix charts
- Local rule-based business insights
- Exact values in cards and chart tooltips
- Cleaned CSV/Excel export
- Print-to-PDF support
- Excel parsing runs in a Web Worker so the main UI remains responsive during import
- No backend and no API key required

## GitHub Pages
Set **Settings → Pages → Build and deployment → Source** to **GitHub Actions**. Push to `main` and the workflow builds and deploys the `dist` folder.

## Important
The dashboard processes the workbook in the browser. It does not provide a cloud AI/LLM service. Automatic insights are deterministic, local analytics rules.
