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

- **Dimensions** — `MICA`, `MICA_Level_5/4/3/2/1`, `Product_code`,
  `Product_Level_3/2/1`, `Segment_code`, `CG_Level_2/1`, `Function_code`,
  `Function_Level_2/1`, `Entity code`. Header matching is tolerant of
  case/underscore/spacing differences; missing optional columns just disable
  the related filter. Beyond the built-in list, **any other column in the
  driller can be declared as a dimension** through the config's Dimensions
  section (see below) and then behaves like the rest everywhere — filters,
  Financial Summary cascade, drag & drop charts, dashboards, commentary.
- **Periods** — either of two layouts:
  1. Monthly columns for the prior and current year. When the labels name
     their year (`Jan-25`, `Jan 2026`) the columns can sit **anywhere in
     the file in any order** — the year on the label assigns them. Yearless
     labels (`Jan Actual`) fall back to the sequence rule: two blocks in
     calendar order, prior year first. Quarterly columns (`1Q25 Actual`,
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

**Multi-country workbooks** — a regional driller can carry **one country
per tab**: every tab that parses as a driller becomes a country (named
after its tab), while Control/Mapping/notes tabs are skipped as before.
The countries merge into one model with a `Country` dimension, aligned by
column label so tabs may lay their columns out in any order; any tab
missing columns of the first data tab is reported in the ingest summary.
A **Countries strip** appears in the left pane: a Region row — the sum of
every country — then one row per country with its P&L YTD and RAG; click
a row to scope every page, chart, export and simulation to that country
(the Country filter group mirrors the choice). The strip stays quiet:
the whole section folds from its header like the other left-pane
sections, the tree stops at business lines and regions (business rows
carry a persistent expand/collapse caret), and countries live in a
**country dropdown** beneath the tree — a picker showing the current
selection with its RAG dot and value, opening a searchable panel of
countries grouped by business line and region. Picking a country scopes
the whole app (keeping the business-line scope), a clear row restores
the wider scope, and scoping from anywhere else keeps the picker in
sync. Region consolidation keeps
all the standing rules: YTD from YTD columns, balances as closing
positions in `US$bn`, P&L in `US$m`, and P&L and balance sheet never
totalled together. Country also works as an ordinary column
(`Country`/`Country_Name`/`Country_Code`) in a single-tab file. Settings:
`country_tabs` (`auto` merges all parsing tabs, `first` restores the old
first-tab-only behaviour) and `region_label` (the Region row's caption).

**Country hierarchy (multi-region)** — a `CountryHierarchy` config
section defines the roll-up; nothing is hard-coded. Paste the group
sheet in verbatim: a header row then `ID | Level 1 | Level 2 | Level 3`
— or the five-column form with `Level 0` (IWPB / CIB / GROUP), which is
retained per entry (e.g. `C14 | IWPB | Global | Asia | Singapore`; a
Level-2-only row such as Holdings becomes its own region; flat
`Key=country, Value=region` rows also work). Without the section, a
multi-tab file shows the flat Region → Country strip.

**Business lines (Level 0 / group roll-up)** — when tabs are named
`<business> <country>` (`IWPB Singapore`, `CIB Singapore`) and the
hierarchy carries `Level 0`, the same country may appear once per
business line and the roll-up gains a tier: **Group — all business
lines → business line → region → country**, with the group total the
sum across business lines and a group-level country (business line
"All" + the country in the top bar) summing that country across
businesses. `Business line` is a full dimension — filters, chart axes,
simulation scopes — and the top bar gains a business dropdown, with
region and country lists narrowing to the selection. A plain country
tab name resolves to its (single) hierarchy entry as before, and
`group_label` renames the top row. With the map uploaded,
the left-pane strip becomes three levels — the Level 1 name (Global) on
top, each region in bold with its consolidated P&L YTD, and its
countries indented beneath — and clicking any row scopes the whole app
to that node. `Region` becomes a full dimension everywhere: filters,
chart axes, query builder, simulation scopes and "View impact by".
Countries not in the map group under "Unmapped" so nothing silently
drops out of the consolidation. The top bar (next to Units) carries the
same scope as two compact dropdowns — Global/region and country, the
country list narrowing to the chosen region — always in sync with the
left-pane strip and the Filters pane.

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
- **One way of writing a movement, in every view.** The KPI tiles and their
  trajectory strip, the KPI scorecard, the auto-written commentary, the
  executive decisions, the attention chips, the management commentary, every
  breakdown and chart data table, chart annotations and labels, the
  simulation's impact table and headline, and every MI Assistant answer all
  render variances through the same pair of rules: the
  **triangle is the direction the line actually moved**, and the **colour is
  what that movement means** — a cost or ECL going up is ▲ red, revenue or a
  balance going up is ▲ green, a cost coming down is ▼ green. Percentages
  are written as magnitudes (`▲15.6%`), never as a signed number that flips
  meaning between a file storing costs as negatives and one storing them as
  positives, so the text and the legend can never disagree with each other or
  with the tile they sit beside. Two figures that are demands rather than
  movements — the monthly run-rate a target needs, and the momentum of the
  last three months — say **above** or **below** the current pace in words
  and carry no favourability colour, because a bigger requirement is not a
  better or a worse result. The management commentary keeps the pack's F/A
  wording with the triangle joining the letter — `▲346A` / `▼18F`.
  **The Financial Summary is the one view without arrows**: it replicates the
  reporting pack, where fifteen variance columns sit on every line and a
  triangle in each one is clutter rather than signal. There the pack's own
  reading stands — the colour carries favourability (a cost underspend green,
  an ECL overshoot red) and the bracket carries the sign — in the grid, in
  its narrative's F/A wording and on its PowerPoint slide alike. The
  favourability test itself lives in one place
  (`movFav` / `movUp`), which every one of those surfaces calls, so they
  cannot drift apart again; a chart category or a scenario line, which has
  rows but no KPI record, derives the same flags from its own rows.
- **The result line is opened up** — on a driller whose only P&L rollup is a
  single Level 1 line (a PBT beside Deposits and Loans), the tile grid would
  otherwise show that one result and the balance sheet, and nothing of what
  drove it. Its Level 2 lines lead the grid in statement order — revenue,
  the cost lines, the charge — with the result behind them and the
  balance-sheet lines last. Files with several P&L rollups are unaffected.
  Switch it off with `tile_expand_result`.
- **Naming the tiles** — `tile_lines` lists exactly which lines become KPI
  tiles, in the order given: `Revenue,PBT,Loans,Deposits,Costs`. Each name
  is looked up in the MICA hierarchy — Level 1 first, then Level 2
  (`tile_line_levels`) — and taken from the first level that carries it, so
  a summary can mix a Level 1 rollup with Level 2 lines inside it without
  forcing everything to one level. Nothing resolves against a product,
  segment or function that happens to share a line's name. The list is shown as
  given — no priority reordering, no collapsing behind "show all" — a name
  the file does not carry is called out rather than silently dropped, and
  the grand-total tile is left off because the lines were chosen by hand.
  The PowerPoint tile grid and the landing narrative follow the same list.
  Leave it blank for the default behaviour (`tile_level` + `tilePriority`).
- **Hide a tile you don't want on the page** — every KPI tile carries a small
  **✕** in its top-right corner. Hiding parks the tile rather than deleting
  it: the grid ends with a dashed *"N hidden tiles — show"* card, one click
  brings them back as dimmed tiles, and the **↺** on a dimmed tile restores
  it to the grid. The choice is remembered in the browser (alongside the
  drag-to-reorder order) and the PowerPoint tile grid follows the page, so a
  parked line is off the one-pager too. Nothing is recalculated by hiding —
  a parked tile is only out of view, and the reporting pack, the Financial
  Summary and the KPI scorecard still carry the line in full. Set
  `tile_hidden` (a `;`- or `,`-separated list of tile names) to have a build
  open with those tiles parked out of the box; the reader's own first hide or
  restore takes over from then on.
- **Move a tile where you want it** — drag any KPI tile onto another and it
  lands there, before or after depending on which half of the target you drop
  on. The order holds whatever produced the tiles: a config-named set
  (`Metrics` / `tile_lines`), a priority-ordered set, or a plain MICA
  grouping — a dragged order is an explicit act and leads, with the declared
  order filling in behind it, so a line the stored order has never seen keeps
  its declared place rather than jumping the queue. It survives re-rendering,
  re-ingesting the file and reopening the browser, covers collapsed and
  parked tiles as well as the ones on screen, and the PowerPoint one-pager
  follows it. A quiet **↺ Tile order** button appears beside Export to Excel
  once something has moved — it puts the file's own order back — and stays
  out of sight until then, so it never takes a tile's place in the grid.
- **Balance scopes read in billions everywhere** — any tile, chart or
  table whose scope is a balance line switches to the balance unit: Mix
  analysis and builder charts scoped to Deposits or Loans plot and label
  in `US$bn`, dashboard cards and breakdown tables restate to closings in
  bn (a balance's YTD figure is the file's own YTD column when it carries
  one at the selected month — the same figure the Financial Summary shows —
  and the month's closing balance otherwise), and the MI Assistant's tables show each line in its own unit. A
  chart over a mixed scope (the whole file) stays in the file's unit.
- **Collapsible commentary** — the Management commentary folds behind its
  header: click the caret (or the title) to collapse it to a single line
  and click again to reopen. The state is remembered in the browser, and
  `commentary_collapsed: Y` starts it folded for everyone. Folding is
  display-only — the text, edits and Word/Copy exports are untouched.
- **A calmer landing page** — the KPI summary leads with one narrative
  line, the tile grid, a single row of RAG-and-attention chips, then the
  two analysis panels. The greeting line is off by default
  (`show_greeting: Y` restores it), best/weakest and row-count chips are
  gone (the narrative and left pane already carry them), the trajectory
  strip caps at four figures, the per-panel filter note appears only when
  filters are actually applied, and editing chrome — View detail, RETRIEVE
  / REGENERATE — reveals on hover.
- **Units on the tiles** — every KPI tile names its unit beside the KPI
  (`US$m` on the P&L lines, `US$bn` on balances), the page lede states the
  split ("P&L in US$m · balances in US$bn"), and the one-slide PPT export
  mirrors the same grid, units, expansion and totals rule as the page.
  Because the tiles carry their own units, the *Financial Performance*
  header carries none.
- **Summary layout** — the tiles sit in equal-height cards; below them the
  performance trajectory takes the left column and an **Executive
  decisions** panel the right, the two stretched to the same height. The
  panel is computed, never narrated: the lines outside tolerance rank
  worst-first as decision cards, each with a severity flag, the variance
  as its headline impact figure (233A / +11.5%), and the inferences the
  data supports — position vs the basis, the full-year gap to target, and
  the run-rate the remaining months must average — plus a *Review
  trajectory* action that points the chart at that line. A count badge
  carries the number of lines outside tolerance, cards cap at
  `decisions_max` (default 3) with the overflow noted, and a file with
  everything inside tolerance says so instead. The
  Management commentary sits full-width beneath the pair. Clicking a tile
  (or a KPI in the left rail) points the trajectory and its actions at
  that line. The mix-analysis panel no longer shows on this page, though
  its view still feeds the one-page PPT export.
- **No total when a total means nothing** — `total_tile` and `fsum_total`
  default to `auto`: a computed total is printed only while no MICA Level 1
  line in the file is a balance-sheet line, recognised by
  `balance_sheet_patterns` (deposits, loans, advances, mortgages, balances,
  assets, liabilities, RWA, AUM, NNM …). Those patterns are their own
  setting, independent of `TilePriority`, so re-ordering the tile
  priorities cannot switch totalling back on. Balance-sheet lines are also
  never hoisted by tile priority and always sort below the P&L, and a file
  carrying them never collapses its tile grid. A driller carrying Deposits or Loans beside a result gets no
  total tile, no total row and no total in its commentary, because adding a
  balance to a result is not a figure; a P&L-only driller keeps its bottom
  line as before. `Y` or `N` force it either way (though `Y` can never force a
  total on a scope mixing balance-sheet and P&L lines — that guard is
  absolute), and naming the lines in
  `tile_lines` or declaring statements in `fsum_sections` also drops it.
  Lines the priority patterns do not recognise are ordered as they appear
  in the driller rather than by size, so a balance-sheet line is not
  hoisted above the P&L. Lines *within* a statement — the Level 2 lines
  under a rollup — always read in the driller's own sequence, so revenue
  sits with the cost lines in statement order rather than being reshuffled
  by whichever moved most; the "driven by" openers still name the biggest
  movers first. With no total, the whole KPI summary follows the
  same line: the tiles carry the Level 1 lines only, the narrative and the
  trajectory open on the first of them, leader and laggard are ranked
  inside that line rather than across statements, Mix analysis opens scoped
  to it, and the Management commentary writes one narrative per Level 1
  line instead of one led by a total. With no total, the landing narrative and the
  performance trajectory are about a **MICA line** instead: a picker on the
  trajectory panel lists Level 1 and Level 2 lines, and the page opens on
  the first declared statement's line (the P&L result) — the trio, the
  chart, the RAG pill and the commentary all follow the chosen line.
- **The total tile** — the first KPI tile adds up every row in scope. If a
  line in the file carries the same name as `total_label` the tile takes
  `total_all_label` instead ("Total (all lines)"), so the grand total and
  the line of that name are never two tiles with one name. Set
  `total_tile` to `N` to drop the tile altogether — worth doing on a file
  that mixes statements, where adding balance-sheet balances to a P&L
  result gives a number with no meaning.
- **Units** — a strip at the top of every page declaring the unit every
  figure is stated in, with a selector. The file's own unit comes from
  `unit_label` (e.g. `US$m`); *as reported* leaves the figures alone, and
  `k` / `mm` / `bn` restate every figure against the chosen unit — a line
  of 11,489 US$m reads as 11.5 US$bn, or 11,489,000 US$k. Figures never carry their own `m`/`bn`
  suffix on top of the file's unit, which would scale the same number
  twice. Excel and PowerPoint exports follow the selection (percentages
  and the raw driller-row export are never restated), and simulation
  amounts are typed in whatever unit is on screen. Configurable via
  `display_units` (the unit selected on load), `unit_options` (which are
  offered) and `unit_decimals` (blank = 0 as reported, 1 when restated).
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
  which and why. The QTD figure comes from the file's **quarterly column**
  and the year-to-date from the **YTD column** — months are only added when
  no such column exists (`calc_ytd_actuals: months` forces sums). YTD
  headers are recognised in any common shape: `Jun YTD-26`, `YTD Jun-26`,
  `YTD-Jun'26`, a quarter-named `Q2 YTD-26` (read as through the
  quarter-end month), or a bare `YTD` (read as the file's running YTD,
  through the last actual month) — under an Actuals, Forecast or Target
  band alike. A file may carry **one YTD column per month** (`Apr YTD-26`,
  `May YTD-26`, `Jun YTD-26` …, actuals and forecast alike): all are kept
  and the one matching the selected YTD-through month serves that month —
  moving the month selector moves between columns, never back to summing
  months. A file whose actuals exist only as YTD columns (all month
  columns marked Forecast) takes its month list and default month from
  the YTD columns themselves. The same
  rule covers the prior year: a PY YTD column (`Jun YTD-25`) matching the
  selected month is authoritative for every vs-PY figure, with PY months
  summed only when the file carries no such column. Target comparisons are
  YTD-only: a **YTD target column** (`Jun YTD-26` under a Target band)
  matching the selected month drives the `vs Tgt` row on KPI cards, the
  scorecard and the Target comparison basis — there is no `FY vs Tgt` row
  and no phasing, so a file carrying only an FY target shows no target
  comparison (balances still compare the closing balance to the target
  balance, which is already same-grain). Comparison charts
  compute their bases on actuals only, so the YTD column stays
  authoritative there even with forecast months toggled into the page. Lines
  matched by `balance_sheet_patterns` are treated as stocks throughout:
  month, QTD, YTD and FY are closing balances, and target/forecast
  baselines are the full-year columns as-is, never phased fractions or
  summed months, and they read in **billions** while the P&L stays in the
  file's unit — the table splits into a `P&L (US$m)` and a
  `Balance Sheet (US$bn)` band automatically, and the commentary, Excel,
  PPT and Word exports follow (`bs_unit` names the balance unit, blank
  switches it off; `fsum_bs_label` names the band). A **Full Year band** closes the table — the live FY
  forecast (Jan–Dec actuals + forecast months) with variance in currency
  and percent vs the prior forecast column, the FY target and prior-year
  actuals — so once the selected month crosses the last actual and the
  month columns read Forecast, the comparison that matters is on the page,
  and the commentary reports it with numbers and percentages. The **PPT
  export** is a single aligned slide: banded headers, a fixed column grid,
  rows sized to fit the page, favourability colours in the variance cells,
  and the commentary — coloured F/A figures included — beneath the table.
  A **cascade builder** sits above the
  table: the levels the summary rolls through are chips — drag a dimension
  in from the palette (Product under MICA Level 1, or between Level 1 and
  Level 2 — MICA Levels 4 and 5 and any config-declared dimension are
  offered when the file carries them), drag chips to reorder, ✕ to remove
  — up to six levels (`fsum_max_levels` lowers the cap if a tighter page
  is wanted), remembered in the browser, with the roll-up selector, commentary
  and all three exports following the cascade. A level that doesn't fan
  out (a lone child repeating its parent) is collapsed, but the walk keeps
  descending so a deeper level that does split is never lost. The **rows
  themselves drag too**: pull any line onto a sibling to reorder it — drop
  on the top half to land above it, bottom half below — and the line's
  whole block moves with it (its o/w children re-render beneath it, and a
  group total carries everything above it). A row only accepts siblings at
  the same level under the same parent, so statements stay apart and a
  child can't leave its parent. The order persists in the browser, every
  export follows it, and a *Reset row order* link under the table restores
  the default. Every parent row carries a **− / + toggle**: collapse a
  line and its children fold away while the parent's own figures stay
  put, at any depth — the fold state persists in the browser, exports
  follow the folded view, and an *Expand all* link under the table
  reopens everything. **Row labels rename in place**: double-click a
  line's name, type the label you want (blank restores the original) —
  the rename keys off the underlying line, so it follows the row through
  reordering and shows in the table, the commentary and all three
  exports; a *Reset labels* link under the table clears them. The whole
  arrangement can be kept as a **named view**:
  *Save view…* beside the roll-up selector snapshots the cascade levels,
  roll-up depth, dragged row order, fold state and renamed labels under
  a name you give it, the
  picker switches between saved views, ✕ deletes the selected one, and
  the last applied view is remembered in the browser and restored on the
  next visit. The **⟲ Default** button clears the working view outright —
  labels, order, folds and cascade — and returns the page to the
  Account Hierarchy L1 cascade (`fsum_default_levels`, default
  `accH1,accH2,accH3`); saved views survive it and can be reapplied. When the cascade runs on hierarchy levels, lines the mapping
  could not place any deeper than the block's parent no longer masquerade
  as a sibling named after that parent — they show as an **Other** line
  (`fsum_residual_label`) at the end of the block — and the block's total
  row takes the name of the one hierarchy parent every line rolls up to,
  so a P&L block cascaded at Account Level 2 reads Revenue / ECL / Total
  Operating Expense / Other with **PBT** as its total, in the app and in
  all three exports (`fsum_total_label` still overrides). Parent rows
  lead their block by default — PBT on top, then Revenue with its lines
  beneath it (`fsum_parent_row: bottom` restores the classic
  bottom-total layout where children build up to the parent). A **roll-up selector** beside the export buttons sets how
  deep the cascade reads — Level 1 rollups only, to Level 2, or the full
  o/w Level 3 detail — and the Excel, PPT and Word exports follow it
  (`fsum_detail` sets the level on load). Variance colours follow
  favourability, not raw sign: a lower cost reads green and lower revenue
  red, on either sign convention. Every variance in the page's commentary is written the way
  the pack writes it — `18F`, `346A`, `14%A` — favourable or adverse taken
  from the line's own direction, so a smaller cost reads F and a deeper loss
  reads A. Excel and PowerPoint exports carry the same headers and
  figures. The page is deliberately independent of the **Variance period**
  selector — it shows every basis side by side already, so only the
  YTD-through month moves it. Configurable via `fsum_title`, `fsum_section_label`,
  `fsum_levels`, `fsum_ow` (the "o/w" prefix on the deepest level) and
  `fsum_total_label`.
- **Keeping statements apart** — a driller carrying both a P&L and balance
  sheet must not add them together. `fsum_sections` assigns the lines of a
  column (`fsum_section_dim`, MICA Level 1 by default) to named blocks:

  ```
  Settings,fsum_sections,"P&L=PBT | Balance Sheet=Deposits,Loans",,
  Settings,fsum_no_total,Balance Sheet,,
  ```

  The Financial Summary then shows the P&L lines cascading into PBT, and
  Deposits and Loans as separate balance-sheet lines beneath — with no
  total across the two, and none within a block named in `fsum_no_total`
  where adding the lines together would mean nothing. Lines in no block
  fall into a trailing block named by `fsum_other_label`. The Management
  commentary follows the same split, writing one narrative per block.
  `fsum_section_dim` on its own (without `fsum_sections`) still bands the
  table by every value of a column, ordered by `fsum_section_order`.

  Without declaring statements, two simpler keys do the same job in one
  block: `fsum_line_order` fixes the order of the top-level lines
  (`PBT,Deposits,Loans` puts the balance-sheet lines below the P&L), and
  `fsum_total` set to `N` removes every total row from the table, the
  commentary and both exports.
- **Management commentary** — the reporting pack's own wording, generated
  from the file. A heading naming the comparison, an opener giving the
  result and what moved it, then each line of the statement with its
  variance in F/A notation and the movements behind it:

  ```
  (JUN) FY Forecast Vs FY26 target (ex notables):
  (JUN) FY Forecast PBT is $(4,067)m down vs FY26 target driven by lower
  Revenue ($109mA / 1.3%), lower Total Direct Cost ($110mF / 3.2%) …
  Revenue of $8,379m is $109mA / 1.3% (ex notables):
    • NII - Interest Income $36mA / 0.9% driven by …
  ```

  The comparison follows the selected **Variance period**, so the same
  generator writes the vs-forecast, vs-target and vs-prior-year versions.
  Balance-sheet lines are reported after the statement, never inside it.
  `commentary_format` picks `slide` (this) or `house` (the earlier
  cascading style), `commentary_levels` sets the cascade,
  `commentary_dim` the dimension the "driven by" clauses cut by,
  `commentary_drivers` how many are named, and `commentary_headline` /
  `commentary_note` the wording of the heading. Every F/A figure is
  coloured by favourability — F green, A red — in the page and in the Word
  export alike (`commentary_fa_colours` switches it off). **⤓ Word** writes the
  commentary as it stands on the page — your edits included — to a real
  .docx, keeping the bold figures, the underlined heading and the bullets.
  The Financial Summary's own **⤓ Word** puts that page's commentary and
  the management commentary in one document. The file is assembled in the
  browser (a .docx is a zip of XML parts, written uncompressed), so it
  needs no library and nothing leaves the page.
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
  forecast or phased target. Mix analysis, the chart builder and pinned
  dashboard charts re-size on the period too, until a measure is picked by
  hand — that choice then sticks. Month-by-month visuals (the Monthly
  trend widget, the mini-charts under KPI cards, the YTD actuals vs
  forecast widget) stay monthly by definition and do not re-base.
- **MI Assistant** — a governed Q&A page (left-pane entry) driven by
  **plain language**: type the question in your own words and a
  comprehension layer works out the intent (how are we doing, why did it
  move, what is behind, break it down by X, monthly trend, run-rate,
  costs, ECL, target) together with anything the question names — a line,
  product, business line, region or country. Names are matched against the
  values actually in the loaded file, longest first, with the Global
  Business dimensions (business line, region, country) winning over the
  same text in an ordinary driller column — so "iwpb" reads as the
  business line, not the CG Level 1 code. Case does not matter except for
  codes that are also ordinary words (`US`, `UK`, `IT`), which must be
  capitalised. Naming two values of one dimension ("Singapore vs
  Malaysia", "IWPB vs CIB") is read as a comparison and ranks just those. **It asks back rather than
  guessing**: a bare name offers the readings available for it, a term
  living in two dimensions asks which was meant, a basis the file cannot
  serve offers the one it can, and an unrecognised question offers what
  this file can answer — each as one-click options. A short follow-up
  ("and Malaysia?") inherits the previous question. Behind it sit the
  approved analyses (performance, lines behind forecast, FY target,
  top/dragging products, costs, ECL, required run-rate, prior year,
  contribution mix, movement drivers, dimension breakdowns, monthly
  trend, geography comparisons)
  answered deterministically in the browser from the filtered data on the
  standard comparison basis, plus geography comparisons — countries,
  regions and business lines side by side with share and RAG. Answers
  respect the active Global Business scope (stated in each answer's
  note), and naming a business line, region or country in a question
  ("How is Singapore doing?", "CIB costs") answers for that slice
  without changing the page scope. Free-text questions are matched to
  the governed set; anything outside it is declined. Nothing leaves the
  page.
- **Simulation Assistant** — what-if scenarios on the forecast months. Build
  any number of ordered rules, each scoped to any dimension value (or all
  rows) with three adjustment types: % change, add amount (spread over the
  chosen months, pro-rata across matching rows), or set the monthly total.
  Each rule also carries an **Apply to** choice — P&L + balances (default),
  P&L only, or Balances only — because a product scope such as
  Product Level 4 = Deposits legitimately tags both the deposit balances
  and the deposit NII lines; pinning the rule to one statement stops the
  other side from moving, and the assistant warns when a rule's scope
  straddles both statements. The prior forecast stays untouched; the page
  shows prior vs simulated FY outlook, the delta, both against the FY
  target, a monthly chart (actuals, prior forecast, simulated forecast)
  and an impact table listing only the lines the scenario actually moves,
  with Excel export. A scenario whose scope is entirely balance-sheet rows
  reports on the closing-balance basis in `US$bn` — the last forecast
  month's closing position, never a sum of monthly balances — and mixed
  scopes total the P&L side only, with balances stated separately.
  Scenarios save by name (localStorage) for reload; the work-in-progress
  scenario survives refreshes.
- **Data** — the filtered source rows, paginated with search. Every text
  column header carries a **filter button**: it opens a value picker
  (searchable, Select all / Clear, tick the values to keep) that combines
  with the page search and the left-pane filters; active funnels highlight
  and a Clear column filters button drops them all at once. The picker
  only offers values that can still appear under the other active filters.

Clicking a KPI card (or a KPI in the left pane) focuses the **Mix
analysis** panel on that KPI — deeper slicing lives in the Query builder
and the chart builder, which cover the old drill-down and more.

The KPI summary also ends with a **Mix analysis** panel: a flexible chart
driven entirely by dropdowns — KPI scope, view-by dimension, measure and
chart type (donut / column / horizontal bar) — that follows the global
filters and remembers its last configuration.

- **My dashboard** — a personal, drag-and-drop dashboard. A widget is
  composed on the add bar from four choices: **type** (KPI card, monthly
  trend, YTD actuals vs forecast, breakdown by field, custom chart, KPI
  scorecard), **KPI** (any metric, MICA line, or the total), **field** and
  **graph**. The field picker lists every dimension the loaded file carries —
  MICA levels, product, segment, function, entity, country, region, business
  line — and appears for the types where a field means something (breakdown
  and custom chart); the graph list is the one belonging to the type chosen,
  so a card offers horizontal bar / sparkline / column / table, a trend or
  trajectory offers line / area / column / bar, and a breakdown offers table /
  bar / column / donut / pie / Pareto / treemap. A breakdown drawn as a table
  keeps the RAG grid; drawn as anything else it renders as that chart.
- **Every widget can be re-composed in place.** Its header carries the same
  pickers it was built with — field and graph on a breakdown, graph on a
  card, trend or trajectory, layout on the scorecard, view-by and chart type
  on a pinned chart — so nothing has to be deleted and rebuilt to change what
  it shows. The choice is saved with the widget, and the PowerPoint deck
  draws each widget the way the page draws it. Drag widgets by their header
  to rearrange, drag a KPI from the left pane onto the canvas to add it as a
  card, and remove widgets with ✕. The layout persists in the browser
  (localStorage) across sessions.
- A breakdown whose field mixes P&L and balance-sheet lines shows its lines
  and **no total** — on the page and in the Excel export alike — because the
  two statements are never added together.

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
- a **flat CSV** with columns `Sheet,Key,Value,Extra1,Extra2,Extra3`
  (`IWPB_SG_dashboard_config.csv` in this folder is the template, also
  downloadable in-app)

What each section governs:

- **Settings** — app title/subtitle/eyebrow (`app_title` also drives the
  browser-tab title, the upload-page wording, the greeting, the export
  headers, the red logo badge and the Financial Summary page title, so the
  whole app rebrands from the config; `logo_mark` optionally sets the
  badge letters, otherwise they derive from the title's initials;
  `export_prefix` optionally sets the download filename prefix, otherwise
  it derives from the title; `fsum_title` still overrides the Financial
  Summary heading explicitly), landing page title, total
  tile label, tile grouping level (`auto`/`level1`/`level2`), collapse of
  non-priority tiles, default RAG tolerance, the regex patterns that
  recognise Forecast and Target columns, the cost/ECL direction patterns,
  the hidden-comparison patterns for target-only files, the My
  dashboard page size, and `variance_periods` — the ordered list of
  comparison bases offered in the left pane (any of `auto`, `ytd_py`,
  `ytd_fc`, `ytd_tgt`, `mom`, `mth_py`, `mth_tgt`, `mth_fc`, `fy_tgt`,
  `fy_py`; entries a file cannot serve are hidden automatically).
- **Hierarchy check** — the ingest report states whether the cascade adds
  up: it names any row that matched no line (they group under Unmapped)
  and any line that keeps rows of its own while also having children, so
  the children would not add to it. A clean file reads "every row lands
  on a line, and every parent equals its children".
- **AccountHierarchy** — the reporting pack's own line cascade:
  `ID | Level 1 … Level 5 | Match`, where Level 1 is the statement group
  (`PBT ex Notables`, `Balance Sheet`, `Key Metrics`) and `Match` lists the
  driller line names that roll into the deepest level. Uploaded rows
  replace the defaults wholesale.
- **Key metrics as memo lines** — `memo_patterns` marks lines that are
  neither P&L nor balance sheet (Premier Customers, CER %, FTE, NNM/NND/
  NNIA, Wealth balances). They are **never totalled into the P&L or the
  balance sheet**, are never restated onto the balance unit, and get their
  own Financial Summary band named by `fsum_memo_label` (default
  "Key Metrics") — so the pack's three blocks appear as
  `P&L (US$m)`, `Balance Sheet (US$bn)`, `Key Metrics`.
- **Calculations** — the calculation rules themselves:
  `calc_ytd_actuals` (`column` = a YTD column matching the selected
  month is authoritative everywhere, months are never added when one
  exists — including in files that also carry prior-year months;
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
  offered as a simulation rule scope (`Extra2` = Y/N — every dimension
  defaults to Y, and the derived hierarchy levels are always offered, so a
  scenario can scope on any non-numeric field in the file; a blank cell
  inherits the default, an explicit N excludes). A row whose key is
  **not** one of the built-ins declares a **new dimension**: `Extra3`
  holds the text to match against the driller's column headers (falling
  back to the label, then the key — matching ignores case, spaces and
  underscores). E.g. `Dimensions,channel,Channel,Y,N,channel` turns a
  `Channel` column into a full dimension across every view.
- **AccountHierarchy / ProductHierarchy** — the reporting hierarchies as
  mapping tables, one row per node: `Key` an ID, then up to five levels
  across `Value…Extra4` (both run five deep — the account side mirrors
  the reporting pack: PBT → Revenue → Banking NII / Fees and Other Income
  → Retail Banking / Wealth / Others → Investment Distribution / Private
  Bank / Asset Management / Insurance, with ECLs, Operating Expenses →
  Direct Cost ex VP CC GT / Global Teams / Variable Pay ex CC / Indirect
  Costs, and Balance Sheet → Deposits, Loans and Advances, Wealth
  Balances → Wealth Deposits / Wealth Invested Assets → NNIA, NNM → NND;
  product — IWPB → Retail/Wealth/Others down to Payroll or Saving
  Accounts). Revenue foots as Banking NII + Fees and Other Income,
  Operating Expenses as its four cost lines, and PBT as Revenue + ECLs +
  Operating Expenses — the same arithmetic as the pack slide. The app binds them at load: each driller row's line names
  (matched against the columns in `account_hierarchy_from` /
  `product_hierarchy_from`, deepest match wins) place the row in each
  hierarchy, and every level becomes a derived dimension — `Acct Hier
  L1–L4`, `Prod Hier L1–L5` — usable in the drag & drop chart builder,
  dynamic dashboard widgets, Financial Summary cascade, commentary
  Driven-by and Mix analysis exactly like a driller column. Beyond a
  node's own name, `Extra5` **stitches the node to the driller's MICA
  lines**: a `|`-separated list of line names that belong to it
  (`Banking NII` ← `NII - Interest Income`, `Loans and Advances` ←
  `Loans|Customer Loans`, `Direct Cost` ← `Total Direct Cost|Staff
  Costs`, …). A stitched match always beats an incidental name
  equality, so `Deposits` lands on Customer Deposits even though a
  deeper node happens to be called Deposits too. Rows naming nothing in
  a hierarchy fall into `hierarchy_unmapped_label` (`Unmapped`). Uploading rows for a section replaces the built-in
  mapping wholesale; a section row with a blank first level clears it.
  The Financial Performance page itself binds through `tile_dims` (e.g.
  `accH1,accH2`, or just `accH2`): the KPI tiles, left-rail KPI list,
  trajectory line picker, narrative ranking and the one-slide PPT all
  regroup by those dimensions instead of MICA Level 1/2.
- **Metrics** — the reporting pack's scorecard metrics as named scopes:
  each row is `Metrics,<name>,<scope>` where the scope selects the rows
  the metric speaks about, as `dimension=value` pairs ANDed with `;` —
  any dimension works, hierarchy levels included
  (`accH3=Wealth Balances;cgL2=Private Bank` is *PB Wealth Balance*).
  Twelve ship by default: Revenue, Banking NII, Wealth Fees, PBT, Direct
  Cost ex VP CC GT, ECL, Deposits, Loans and Advances, Retail & Premier /
  PB Wealth Balances, NNM and FY ROTE (the last needs a ROTE line in the
  driller). A metric a file cannot serve simply doesn't appear. Metrics
  are first-class scopes everywhere — and by default they ARE the KPI
  cards: whenever at least two metrics resolve against the loaded file,
  the Financial Performance grid and the left-rail KPI list show the
  scorecard metrics (in catalogue order, each with its own unit and
  variances) instead of plain MICA grouping — `tile_metrics: N` restores
  the old grouping, and an explicit `tile_lines` (separate with `;` so
  comma-bearing names stay whole) always wins. From each card the usual
  flows make graphs and tables: click for the trajectory, drag a
  left-rail metric onto **My dashboard** for a card/spark/table widget
  scoped to it — dashboard KPI cards carry the same multi-comparison
  stack as the tiles and the reporting pack (the selected basis, vs PY,
  vs Fcst and FY vs Tgt, whichever the file can serve, each as coloured
  amount + %) above the mini chart — and the "Metrics" group heads the
  scope pickers in Mix analysis and the drag & drop chart builder. A
  **KPI scorecard table** widget puts every metric on one page, in the
  reporting pack's own format by default: a grid of bordered boxes, each
  headed by the metric's name (with its unit) and value, with the
  comparison bases beneath as coloured triangle + amount + % rows — a
  header switch flips to a flat table (RAG dot, value, unit, comparison
  column pairs) instead. Boxes and rows click through to the KPI
  summary, the widget exports to Excel, and it never totals across
  statements (each metric keeps its own unit and no narrative is drawn).
  **Any metric can be hidden**, exactly as a KPI tile can: the **✕** on a
  box (or beside a row name in the flat table) parks it, the widget ends
  with a *"N hidden metrics — show"* button that brings them back dimmed,
  and the **↺** on a parked metric restores it. The choice is remembered in
  the browser, kept separate from the parked KPI tiles so the two surfaces
  stay independent, and carried into the widget's Excel export and its
  PowerPoint slide so the pack matches the page. Set `scorecard_hidden` (a
  `;`- or `,`-separated list of metric names) to open with metrics already
  parked; the reader's first hide or restore takes over from then on.
  The dashboard's PPT export renders the scorecard in the same pack-box
  format — native PowerPoint shapes (bordered box per KPI, name + headline
  value, coloured triangle variance rows, three boxes per row, paginating
  onto extra slides when needed), editable after export; a widget switched
  to the Flat table layout exports as a table slide instead.
  In the drag & drop chart builder a **Metric** field chip fans one
  category per catalogue metric — drop it in Group by (it always groups,
  never splits — a drop on Split by lands in Group by) and any chart
  type compares the metrics side by side, split-able by any other
  dimension, with balance metrics on their closing basis. A **Comparison**
  field chip puts actuals against forecast, target and prior year on one
  chart: grouped alone it draws one bar per basis (YTD actuals, YTD
  forecast, YTD PY, then FY forecast, FY target, FY PY — whichever the
  file serves, never summed, and never phased: the YTD forecast basis
  appears only when the file carries a genuine YTD forecast column for
  the selected month, and targets compare at the FY grain only); dropped
  as a split against Metric or any dimension it clusters those bases side
  by side per category, scorecard-style. The table and Excel views add
  Variance and Var % per base, each computed at its own grain exactly as
  the KPI cards do — YTD bases against YTD actuals, FY bases against the
  FY forecast/outlook, never across grains — and skip the Total row when
  grouping by overlapping metrics. On a mixed P&L + balance-sheet scope the comparison
  covers the P&L side only (pick a balance-sheet KPI scope to compare
  balances: closing balance vs FY target and PY closing). Everything resolves
  through one scope machinery, exports included. Uploaded Metrics rows replace the
  built-in list wholesale.
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
measure and the chart type; the preview renders live and respects the
right-pane filters. The measure dropdown offers every comparison the file
serves, whatever else it carries: YTD actuals, paired **actual vs PY / vs
YTD forecast / vs YTD target** (side-by-side bars with variance labels),
the single bases and their variances, and the FY measures — a prior-year
file no longer hides the forecast and target measures.

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
