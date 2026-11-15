# MP WCD Commissioner Dashboard

An interactive, mobile-friendly web dashboard that reads its data straight from an
Excel workbook (`data/dashboard.xlsx`). It recomputes every KPI, ranking and chart
from the raw sheets. When you change the workbook, the dashboard picks up the
change automatically: no rebuild and no page reload.

## Run it

```bash
python3 serve.py          # then open http://localhost:8000
```

Any static web server works (GitHub Pages, Netlify, nginx, IIS). Opening
`index.html` directly from disk (`file://`) can't fetch the workbook, but you can
still load a workbook using the **Data source** button.

## Keeping data up to date

There are three ways to feed the dashboard:

| Source | How | Auto-refresh |
|---|---|---|
| **Hosted workbook** (default) | Replace or edit `data/dashboard.xlsx` on the server (or commit a new one to the repo) | Checked every 15 s and whenever the tab regains focus |
| **Live local file** | Data source → *Choose Excel file* (Chrome / Edge) | Checked every 2 s. Just press **Save** in Excel |
| **One-off upload** | Drag & drop an `.xlsx` anywhere on the page | No (static copy) |

When new data arrives, a toast reports what changed (for example "3 district rows
changed"). The changed KPI tiles flash and show a ▲/▼ delta, and the charts
animate to their new values.

To point at a different file, use `?src=` in the URL, e.g. `index.html?src=data/nov-2026.xlsx`.

## Workbook format

The dashboard finds columns by **header name**. Column order doesn't matter,
and it ignores extra columns. It uses these sheets:

| Sheet | Key columns |
|---|---|
| District Performance | Division, District, Total AWCs, Active Children (0-6), Pregnant Women, Lactating Mothers, Children Measured, Measurement Efficiency, SAM/MAM % of Measured, THR / HCM Coverage, Home Visit Completion, AWC Operational Rate, AWW/AWH Sanctioned/Working/Vacant, Basic Infrastructure Coverage, ANC Coverage, Anaemia % of Pregnant Women |
| Nutrition | District, SAM Cases, MAM Cases, Severely/Moderately Stunted %, New SAM Cases This Month, SAM Follow-up Completed, SAM Referred |
| Maternal Health | District, ANC 1–4 Coverage, Pregnant Women with Anaemia (Count), Pending Due ANC (Count), Optimum Gestational Weight Gain Coverage |
| AWC Details | AWC Code, Division, District, Project, Sector, AWC Name, children/measurement/SAM/MAM, Building/Toilet/Drinking Water status, Days Open, Home visits, AWW/AWH Status |
| Alerts | District, Alert Category, Trigger / Issue, Priority, Responsible Officer, Action Status |
| Read Me | "Reporting period" row is shown in the header |

You can enter percentages as fractions (`0.92`) or whole numbers (`92`). Adding
or removing districts, AWCs or alerts works as long as the header rows stay in
place. The *Dashboard Summary* and *Division Summary* sheets aren't read: the
dashboard computes those totals itself, so they can never go stale.

RAG bands (coverage indicators): **≥ 90 %** on track, **75–90 %** watch, **< 75 %** off track.

## Features

- **Overview**: animated KPI tiles, indicator health (RAG counts), division comparison vs state average, top/bottom districts, and an alert summary.
- **Districts**: a ranking chart for any indicator, a sortable and searchable table, CSV export, and a drill-down drawer for each district (indicators vs state, maternal, nutrition, alerts, AWC list).
- **Nutrition, Maternal, AWCs, Alerts**: views for each of those topics, with filters.
- A global **division filter**, a light/dark theme, skeleton loading, a progress bar, toasts and reduced-motion support.
- **Mobile**: bottom tab bar, stacked cards, swipeable tables with a sticky first column, and a full-screen drawer.

## Files

```
index.html        page shell
css/styles.css    styles, themes, animations, responsive layout
js/data.js        Excel parsing and aggregation
js/app.js         views, charts, live sync
vendor/           SheetJS 0.18.5, Chart.js 4.4.4 (bundled, works offline)
data/dashboard.xlsx
serve.py          no-cache static server
```
