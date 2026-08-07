# IWPB Singapore Driller — KPI Dashboard

A self-contained, single-file web app that ingests the IWPB Singapore driller
extract and shows a **RAG-rated KPI summary at MICA Level 2** with full
drill-down. No server, no build step, no network calls — open `index.html` in
any modern browser and everything (including XLSX parsing via an embedded copy
of SheetJS) runs locally.

## Using it

1. Open `index.html` in a browser (double-click works).
2. Drop the driller file (`.xlsx` or `.csv`) onto the upload zone. The ingest
   report confirms the detected columns, KPI count and period coverage.
3. The dashboard opens on the **KPI summary**.

## What it expects in the file

A single header row with:

- **Dimensions** — `MICA`, `MICA_Level_3/2/1`, `Product_code`,
  `Product_Level_3/2/1`, `Segment_code`, `CG_Level_2/1`, `Function_code`,
  `Function_Level_2/1`, `Entity code`. Header matching is tolerant of
  case/underscore/spacing differences; missing optional columns just disable
  the related filter.
- **Periods** — either of two layouts:
  1. Two blocks of monthly columns (`Jan Actual … Dec Actual` prior year,
     then `Jan Actual …` current year), quarterly columns (`1Q25 Actual`,
     `1Q26 Actual`, …) whose 2-digit years label the blocks, plus optional
     `… FC` forecast months and YTD columns.
  2. CIB-style: title rows above the table, an **Actuals / Forecast /
     Target band row** above the headers, a single year of months
     (`Jan-26 … Dec-26`) whose actual-vs-forecast split comes from the
     band, quarters without year digits, YTD, FY and FY-Target columns.
     The data sheet may sit behind Control/notes tabs — every sheet is
     tried until one matches. With no prior-year months, the primary
     comparison becomes **YTD actuals vs YTD forecast** — the FY
     forecast column phased over the elapsed months (FY forecast ×
     months ÷ 12) — with the FY outlook (Jan–Dec months), FY target and
     FY prior-year actuals shown alongside in the trajectory panel. The
     forecast comparison applies to every KPI, costs and ECL included,
     with direction-aware ratings.

  Values may be numbers, `1,234` text, `(1,234)` bracket negatives, `-`,
  or blank.

## Views

- **KPI summary** — driller-style financial-dashboard tiles: one per MICA
  Level 1 rollup (Revenue, Costs, ECL, …) plus a Total P&L tile (files
  with a single Level 1 value fall back to Level 2 tiles), each with the KPI name and headline YTD value,
  then compact comparison rows (`vs PY`, `vs Fcst`, `vs Target`) showing
  the variance amount and % with direction triangles — the triangle tracks
  the business movement (revenue up ▲, cost up ▲) and the colour the
  favourability (green favourable, red unfavourable). The **⤓ PPT** button
  beside Export to Excel produces the same tile grid as a single
  widescreen slide, built natively in PowerPoint.
- **Financial Summary** — the management-reporting layout: three column
  bands over the P&L cascade taken from the driller's own hierarchy.
  *Month* (the three months ending at the selected one, then Variance vs
  Fcst / vs Target / vs PM / vs PY), *QTD* (the quarter to date, then vs
  Target / vs PQ / vs PY) and *YTD* (year to date, then Variance in
  currency and in % vs Target and vs PY). Rows cascade child lines into a
  bold subtotal per top level and a grand total, in the tile-priority
  order. Everything moves with the YTD-through month — pick MAY and the
  months become Mar/Apr/May, the quarter Q226 to date, the year-to-date
  Jan–May. Comparisons a file cannot serve show as `–` with a note saying
  which and why. Excel and PowerPoint exports carry the same headers and
  figures. Configurable via `fsum_title`, `fsum_section_label`,
  `fsum_section_dim` / `fsum_section_order` (band the table by a column,
  e.g. P&L / Balance Sheet / Key Metrics), `fsum_levels`, `fsum_ow` (the
  "o/w" prefix on the deepest level) and `fsum_total_label`.
- **Variance period** — a left-pane selector (Period & RAG) that sets the
  comparison every view is rated on: YTD vs PY, MoM, Month vs Target, YTD
  vs Target, YTD vs Forecast, FY vs Target, FY vs PY. The option labels
  are built from the selected YTD-through month, so they re-date
  themselves — pick MAY and MoM reads *MAY-26 vs APR-26*. Tiles,
  trajectory, Mix analysis, breakdowns, dashboard cards, the management
  commentary, the MI Assistant and both exports all follow the selection;
  **Auto** keeps the file's default basis (`calc_cmp_priority`). The
  performance trajectory redraws with it: a month-level basis (MoM, Month
  vs Target) plots the monthly shape, a YTD or FY basis the cumulative
  path, each against the base the period names — prior-year, phased
  forecast or phased target.
- **MI Assistant** — a governed Q&A page (left-pane entry): ten approved
  questions (performance, lines behind forecast, FY target, top/dragging
  products, costs, ECL, required run-rate, prior year, contribution mix)
  answered deterministically in the browser from the filtered data on the
  standard comparison basis. Free-text questions are matched to the
  governed set; anything outside it is declined. Nothing leaves the page.
- **Simulation Assistant** — what-if scenarios on the forecast months. Build
  any number of ordered rules, each scoped to any dimension value (or all
  rows) with three adjustment types: % change, add amount (spread over the
  chosen months, pro-rata across matching rows), or set the monthly total.
  The prior forecast stays untouched; the page shows prior vs simulated FY
  outlook, the delta, both against the FY target, a monthly chart
  (actuals, prior forecast, simulated forecast) and a per-Level-1 impact
  table, with Excel export. Scenarios save by name (localStorage) for
  reload; the work-in-progress scenario survives refreshes.
- **Data** — the filtered source rows, paginated with search.

Clicking a KPI card (or a KPI in the left pane) focuses the **Mix
analysis** panel on that KPI — deeper slicing lives in the Query builder
and the chart builder, which cover the old drill-down and more.

The KPI summary also ends with a **Mix analysis** panel: a flexible chart
driven entirely by dropdowns — KPI scope, view-by dimension, measure and
chart type (donut / column / horizontal bar) — that follows the global
filters and remembers its last configuration.

- **My dashboard** — a personal, drag-and-drop dashboard. Add widgets (KPI
  card, monthly trend, MICA L3 breakdown, product breakdown) for any KPI or
  the total, drag widgets by their header to rearrange, drag a KPI from the
  left pane onto the canvas to add it as a card, and remove widgets with ✕.
  The layout persists in the browser (localStorage) across sessions. Pinned
  charts carry two dropdowns in their header — view-by dimension and chart
  type — so any chart can be reshaped in place without rebuilding it.

The design follows HSBC management-reporting conventions: a light left pane
with a red brand block, section headers and pink-highlighted active items; a
red uppercase eyebrow over large page titles; flat white stat tiles with big
numbers (negatives in red brackets); and red monochrome charts with a grey
prior-year series. The left pane hosts the dashboards navigation and a KPI
shortcut list (each MICA Level 2 line with its RAG dot and YTD value — click
to analyse it in Mix analysis, drag onto My dashboard to add it).
Below the KPI list, the left pane holds the controls: period settings
(YTD-through month, include-forecast, RAG tolerance), then a checkbox
filter for EVERY dimension column in the file — each column expands to its distinct values
with All/None shortcuts and a search box for long lists, and any
combination applies to every view, chart and export. Below that, an
Output fields section with checkboxes chooses which dimension columns and
period column groups appear in the Data view and in Excel exports, and a
**Drivers** section holds known business drivers — a line or dimension
value plus its explanation (e.g. `Loans → mortgage repricing +50bps`) —
which the AI commentary weaves into its narrative wherever that line
surfaces as a mover (Mix analysis, chart builder, pinned dashboard
charts). Drivers can be typed in the pane or shipped in the config file's
`Drivers` section (Key = line or value, Value = driver note). RAG compares YTD actuals with the same
prior-year months: green at or above prior year, amber within the selected
tolerance below, red beyond it. Ratings are direction-aware — an increase
in revenue rates Favourable while an increase in costs or ECL rates
Unfavourable, whichever sign convention the file stores costs in (signed
negatives or positive magnitudes).

## Configuration

The app's behavioural rules are externalised into an uploadable
configuration — **left pane → Configuration → Upload config**. Two formats
are accepted, both parsed locally:

- an **xlsx workbook with a sheet per section** — `Settings`,
  `TilePriority`, `Dimensions`, `Views`
- a **flat CSV** with columns `Sheet,Key,Value,Extra1,Extra2`
  (`IWPB_SG_dashboard_config.csv` in this folder is the template, also
  downloadable in-app)

What each section governs:

- **Settings** — app title/subtitle/eyebrow, landing page title, total
  tile label, tile grouping level (`auto`/`level1`/`level2`), collapse of
  non-priority tiles, default RAG tolerance, the regex patterns that
  recognise Forecast and Target columns, the cost/ECL direction patterns,
  the hidden-comparison patterns for target-only files, the My
  dashboard page size, and `variance_periods` — the ordered list of
  comparison bases offered in the left pane (any of `auto`, `ytd_py`,
  `ytd_fc`, `ytd_tgt`, `mom`, `mth_py`, `mth_tgt`, `mth_fc`, `fy_tgt`,
  `fy_py`; entries a file cannot serve are hidden automatically).
- **Calculations** — the calculation rules themselves:
  `calc_ytd_actuals` (`column` = as-of YTD column authoritative /
  `months` = always sum months), `calc_fy_forecast` (`column` = FY
  forecast column first / `outlook` = always Jan–Dec months),
  `calc_ytd_forecast` phasing (`even` = elapsed÷12 / `profile` = the
  file's own monthly shape / `weights` = custom), `calc_forecast_weights`
  (12 comma-separated weights for `weights` mode, also drives the trend
  reference lines), `calc_cmp_priority` (ordered comparison basis, e.g.
  `py,fc,tgt`, used when the variance period is `auto`),
  `calc_variance_period` (the basis selected on load, e.g. `ytd_py`),
  `calc_rag_green_at` (variance needed to rate Favourable,
  e.g. `0.02` requires +2%), `calc_rag_direction` (Y/N — direction-aware
  cost/ECL rating).
- **TilePriority** — the ordered name patterns pinning the top KPI tiles.
- **Dimensions** — per dimension: display label override, whether it
  appears in the left-pane filters (`Extra1` = Y/N) and whether it is
  offered as a simulation rule scope (`Extra2` = Y/N).
- **Views** — enable/disable each page (summary, fsum, custom, builder,
  query, table, assist, sim).

The applied config persists in the browser and re-parses the cached data
file immediately; **Reset to defaults** reverts everything. Blank values
fall back to defaults, and an unreadable config never breaks the app.

## Sample data

`sample/IWPB_SG_Driller_sample.xlsx` mirrors the real file's 15 dimension +
28 period column layout with synthetic values, for testing without exposing
real data. Regenerate with:

```
cd sample && python3 make_sample.py   # needs openpyxl
```

## Drag & Drop Charts

The **Drag & drop charts** page is a field-based chart builder. Drag any
dimension chip (MICA levels, Product levels/code, Segment, Function, Entity)
into the **Group by** well — or click a chip to assign it — and optionally a
second field into **Split by** for stacked segments. Pick the KPI scope, the
measure (YTD current year, YTD prior year, or variance) and the chart type;
the preview renders live and respects the right-pane filters.

Eleven chart types are available everywhere a chart-type dropdown appears
(Mix analysis, the builder, and every pinned chart widget): Column,
Horizontal bar, Lollipop, Line, Area, Pie, Donut, Waterfall (category
build-up to total), Pareto (ranked bars + cumulative share), Treemap, and
Scatter CY vs PY (points above the diagonal grew year-on-year). Split by
adds stacked variants of column and horizontal bar. **Pin to My dashboard** saves the chart as a widget you
can drag to rearrange, and each pinned chart keeps its own chart-type
dropdown in its header. Top 8 categories are shown (6 for donuts) with the
rest folded into Other.

## Query Builder

The **Query builder** page reviews fields and data directly, in the
management-reporting drag-and-drop pattern: a searchable Available Fields
panel (grouped MICA / Product / Segment / Function / Entity plus every
period column) with +F / +O / +S shortcuts, a **Filters** well with
per-field operators (contains / equals / not equals / blank for text; > <
= etc. for period columns) and value suggestions, an ordered **Output
Columns** well (drag entries to change the result column sequence) and a
**Sort By** well with direction toggles. Queries can be saved, loaded and
deleted by name. Run Query renders the result grid (first 500 rows) and
exports the full result set to styled Excel. The chart builder's wells
also accept **multiple Group by and multiple Split by fields** — values
concatenate into hierarchical categories.

## Excel export

Everything exports back to Excel, generated locally in the browser:

- **KPI summary → ⤓ Export to Excel** — one workbook with the KPI summary
  (values, variance, RAG), the Mix analysis chart data and all filtered
  source rows.
- **Mix analysis / chart builder → ⤓ Excel** — the displayed chart's data.
- **My dashboard** — every widget has a ⤓ button exporting that widget's
  data (KPI figures, monthly series, breakdown or chart categories).
- **Data → ⤓ Export all rows** — every filtered source row (not just the
  400 shown), with the original column headers.

Exports honour the active filters, and each file is stamped with the filter
context and date. Workbooks are formatted: header rows carry the HSBC red
fill with bold white text, titles are bold, and column widths auto-fit the
content.

## PowerPoint export

Every chart exports to a widescreen .pptx deck (generated locally):
**⤓ PPT** buttons on the Mix analysis panel, and the chart builder preview
export one slide each. My dashboard has two deck exports: **⤓ PPT (this
page)** takes just the widgets on the page you're viewing, and **⤓ PPT
all pages** takes every widget across all pages — charts as
high-resolution images, KPI cards and breakdowns as styled tables — with
titles, the red accent rule and the filter context in the footer.
