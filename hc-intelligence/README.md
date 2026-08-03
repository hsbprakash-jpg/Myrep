# HC Intelligence

A single-file, fully client-side workforce headcount analytics workspace. Open
`index.html` in a browser — no server or build step required. All Excel files
are processed locally in the browser; nothing is uploaded anywhere.

## What it does

Three Excel inputs activate the workspace:

1. **Semantic question config** — a workbook with a `Questions_Config` sheet
   (plus optional `Application_Config`, `Dynamic_Query_Settings`,
   `Position Questions Config` and `Lead_Targets_Config` sheets) that drives
   the governed questions, labels, colours and targets.
2. **Prior-month workforce file** — the previous reporting month.
3. **Current-month workforce file** — the latest reporting month.

From those it produces:

- **HC Dashboard** — KPI cards, year-end forecast, month-on-month commentary,
  position movement and position total tables.
- **Visual Dashboard** — location/GCB/MR-Lead mixes, hiring status, movement
  waterfall and year-end forecast charts, scopable to one MR Lead.
- **Dynamic Dashboard & Drag-and-Drop Charts** — a chart builder over the
  loaded data with 20+ chart types.
- **Movement Detail** — row-level joiners, leavers, position changes and
  vacancy transitions.
- **Workforce / Positions Assistants** — governed Q&A driven by the config.
- **Position Traceability** — month-by-month position history from a
  position master workbook, cached in IndexedDB so it survives reloads.

Workforce files need at least an `Employee ID` (or `Employee Name`) and a
`Position Number` column; `STATUS_RD` (Physical/Vacant), `JOINERS`, `LEAVERS`,
`MR_LEAD_RD`, `Work Location` and `GCB` unlock the richer views. Header
matching is alias-based and case-insensitive.

## Dependencies

Three libraries are loaded from jsDelivr at runtime: SheetJS (`xlsx`) for
reading workbooks, ExcelJS for styled exports and html2canvas for chart
snapshots. If any of them fail to load (offline or a blocked CDN), the Home
screen shows a clear warning banner instead of failing with a cryptic error
when a file is selected.
