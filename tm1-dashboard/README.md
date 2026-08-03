# GPS Driller — TM1 Hierarchy Dashboard

A single-file, fully client-side HTML dashboard for exploring flat TM1 hierarchy
extracts (e.g. `GPS_Driller.xlsx`, sheet `GPS Data`) with collapsible /
expandable drill-down, styled in HSBC brand colours.

**No server, no install, no data leaves your machine.** Open `index.html` in a
browser, drop your Excel file on it, and drill.

## Features

- **Excel loader** — drag & drop or browse for `.xlsx` / `.xls` / `.xlsm` /
  `.csv`; pick the worksheet if there are several (defaults to the one that
  looks like `GPS Data`). Parsing is done in-browser by
  [SheetJS](https://sheetjs.com/) (0.18.5), embedded **inside** `index.html` —
  the dashboard is one self-contained file you can copy or email anywhere, and
  it works fully offline. A status banner reports progress, row counts, and
  any load errors (bad file, empty sheet, unrecognised headers).
- **Robust value parsing** — numeric cells, text numbers with thousands
  separators (`1,234.50`), accounting negatives (`(2.26)`), and TM1 blank
  markers (`-`) are all handled.
- **Calculated-subtotal exclusion** — rows whose dimension cells contain
  “calculate”/“calculated” are treated as derived consolidations and left
  out of all totals (they would double-count the leaves). The load banner
  and footnote report how many were excluded; a sidebar checkbox re-includes
  them for reconciliation.
- **Semantic layer (auto-detected)** — every column is classified on load:
  - *Period measures*: headers matching TM1 period patterns —
    `Jan'26-Act`, `Q1'25-Act`, `H1'26-Act`, `FY'25-Act`, `YTD Apr'26-Act` —
    are parsed into **year / granularity (Month, Quarter, Half, Full Year,
    YTD) / scenario (Act → Actual, Bud, Fcst…)** and given friendly labels.
  - *Dimensions*: everything else (`Unique Ref`, `Ac Type`, `Seg`, `Markets`,
    `Cntry name`, `Prd`/`Prd1`, `Ac Line`/`Ac Line1`, `Entity`, `Account`,
    `RTN`, `CG code`, `Product 1`, …).
  The sidebar shows the resulting model: dimension count, measure count,
  years and granularities found.
- **Drill-down pivot** — configurable row hierarchy (default
  `Ac Type › Ac Line › Prd › Product 1`); reorder / add / remove levels.
  Every node expands and collapses; `Expand all` / `Collapse all`; per-node
  child counts; values aggregate (sum) up the tree; sticky grand-total row.
- **Period selection** — measures grouped by year with per-year toggle-all,
  plus presets: *Latest FY vs PY* and *Current-year months*.
- **Variance** — one click adds `Δ Var` and `Δ %` between the first two
  selected periods (e.g. FY'26 vs FY'25).
- **Display scale** — show figures in units, thousands (k) or millions (m).
  Display-only division: aggregation always runs on source values, column
  headers gain a `(k)`/`(m)` suffix, and the footnote states the divisor.
  `Δ %` is unaffected by scale.
- **Search** — filters the tree, keeping ancestors of matches and forcing the
  matched paths open.
- **KPI cards** — grand totals for the selected periods.
- **HSBC theme** — HSBC red `#db0011`, black headers, hexagon-inspired mark,
  negative numbers in red parentheses `(28.11)`, light & dark modes.

## Expected file shape

A flat TM1 rollup export with one header row: attribute/dimension columns on
the left, then period columns. Exactly the shape of the `GPS Data` sheet:

```
Unique Ref | Unique Ref1 | Ac Type | Seg | Markets | Cntry name | Prd | Prd1
| Ac Line | Ac Line1 | Entity | Account | RTN | CG code | Product 1
| Jan'26-Act … Dec'26-Act | YTD Apr'26-Act | Q1'26-Act … | H1/H2'26-Act
| FY'26-Act | (same grid for '25) | YTD Dec'19-Act | YTD Dec'20-Act
```

Extra or differently-named dimension columns are fine — anything that doesn't
parse as a period is offered as a drill level. The header row doesn't need to
be row 1; the loader scans the first 10 rows and picks the most plausible one.

## Files

| Path | Purpose |
|---|---|
| `index.html` | The entire dashboard, self-contained (UI + semantic layer + pivot engine + embedded SheetJS) |
| `sample/make_sample.py` | Generates a synthetic extract with the same structure |
| `sample/GPS_Driller_sample.xlsx` | Output of the generator, for demo/testing |

## Try it

```bash
# open directly
open tm1-dashboard/index.html        # macOS
start tm1-dashboard\index.html       # Windows

# then drop sample/GPS_Driller_sample.xlsx (or your real GPS_Driller.xlsx) on it
```

## Notes

- Aggregation is a straight **sum of leaf rows** per node. If your extract
  already contains rollup rows *and* leaf rows for the same numbers, filter to
  leaves first (or drill on the leaf-level columns) to avoid double counting.
- Blank cells (`-` in TM1 exports) are treated as null, not zero.
- Branding uses HSBC's public palette for an internal-tool look; it is not an
  official HSBC asset.
