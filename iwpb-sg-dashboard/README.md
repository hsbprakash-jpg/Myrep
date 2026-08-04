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
- **Periods** — two blocks of monthly columns (`Jan Actual … Dec Actual` for
  the prior year, then `Jan Actual …` for the current year), quarterly columns
  (`1Q25 Actual`, `1Q26 Actual`, …) whose 2-digit years label the two blocks,
  plus optional `… FC` forecast months and YTD columns. Values may be numbers,
  `1,234` text, `(1,234)` bracket negatives, `-`, or blank.

## Views

- **KPI summary** — a hero tile per MICA Level 1 (Revenue) and a card per
  MICA Level 2 line (Net Interest Income, Net Fee Income, Trading Income, …):
  YTD current-year actuals, variance vs the same prior-year period, a
  current-year sparkline and a RAG pill (▲ On track / ● Watch / ▼ Off track).
- **Drill-down** — click any card: monthly trend (current vs prior year, the
  forecast month drawn dashed with a hollow marker, hover for exact values,
  "View as table" for the numbers), a MICA Level 3 breakdown (click a line to
  isolate it), and a product breakdown by Product Level 2 expandable to the
  Level 3 products.
- **Data** — the filtered source rows.

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
to jump straight into its drill-down, drag onto My dashboard to add it).
Filters (entity, segment, function, YTD-through month, include-forecast, RAG
tolerance) live in a right-side panel and apply to every view. RAG compares YTD actuals with the same
prior-year months: green at or above prior year, amber within the selected
tolerance below, red beyond it.

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
measure (YTD current year, YTD prior year, or variance) and the chart type
(column, horizontal bar, donut); the preview renders live and respects the
right-pane filters. **Pin to My dashboard** saves the chart as a widget you
can drag to rearrange, and each pinned chart keeps its own chart-type
dropdown in its header. Top 8 categories are shown (6 for donuts) with the
rest folded into Other.
