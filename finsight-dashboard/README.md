# IWPB Singapore Driller — FinSight KPI dashboard

`FinSightFinal.html` is the single-file KPI dashboard (SheetJS, PptxGenJS and
the whole app embedded). This folder adds an **Interactive charts** page to
it, drawn with [Apache ECharts](https://echarts.apache.org/) 5.5.1.

## Interactive charts page

A new entry in the left nav, after *Drag & drop charts*.

- **Same numbers as the builder.** The page calls the builder's own
  `chartData()` for its figures, so scope, YTD month, units, comparison bases
  and the top-8-plus-Other fold are exactly what Drag & drop charts shows.
- **Controls**: Group by (any dimension level, a period axis, Metric or
  Comparison), optional Split by, KPI scope, Measure, and the chart type
  (Column, Stacked column, Horizontal bar, Line).
- **Interaction**: a tooltip on every mark, a legend that hides and shows
  series, a save-as-image button, and **click-to-drill**: clicking a category
  pins it as a page filter (every other page follows, and the filter shows in
  the subtitle) and moves the chart to the next level of that hierarchy
  (MICA, Product, Segment, Function, or Business › Region › Country).
  A breadcrumb shows the path; *Undo drill* steps back, *Clear drill* unwinds.
- **Exports**: Excel and PPT go through the app's own `chartAOA` / `exportPPT`
  paths. The chart is rendered as SVG so the deck gets a picture.
- Colours come from the page's own CSS palette (`--cat1…7`, `--stk1…5`), so
  light and dark modes match the rest of the app. Comparison bases are never
  stacked; stacked columns need a dimension split.
- The view answers to the config pack like the others: `views.ichart = N`
  hides it, `Labels.nav_ichart` renames it.

## How it is wired

The addition is two `<script>` blocks appended before `</body>`, following the
file's convention of self-contained patch blocks:

1. the ECharts common build (~0.66 MB), and
2. `interactive-charts.js` (kept here as a readable copy), which mounts the
   nav button and the page, wraps `setView`, `render` and `applyCfgDom`, and
   otherwise only calls existing functions.

Nothing in the original code was edited.

## Files

| Path | Purpose |
|---|---|
| `FinSightFinal.html` | The dashboard with the Interactive charts page |
| `interactive-charts.js` | Readable copy of the appended page code |
| `sample/make_sample.py` | Generates a synthetic IWPB-shaped extract (needs `openpyxl`) |
| `sample/IWPB_SG_sample.xlsx` | Output of the generator, for demo/testing |
