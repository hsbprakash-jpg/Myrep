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

## Enrich commentary

An **Enrich commentary** section in the left pane, below Commentary
templates, plus an **⤒ Enrich** button in the Management commentary header.
Anything written about the numbers — a write-up, review notes, a transcript —
can be added, one file or many at once, by the picker or by dropping the files
on the panel or on the commentary block. The sentences enrich the commentary
the moment the files are read; review is optional.

- **Files read**: Word (`.docx`, unzipped natively in the browser), plain
  text (`.txt`, `.md`), and transcripts (`.vtt`, or a text export with
  `Name: …` lines). PDF is not read; save it as Word or text first. In a
  mixed selection the files that cannot be read are skipped and named.
- **How sentences are matched**: each file is cut into sentences, and every
  sentence that names a line the extract carries is matched by the same
  matcher the config's free commentary uses (name, alias, acronym, short
  form, telling word). A document title is skipped; attendee, agenda and
  action lines are skipped; a sentence that opens on a pronoun ("It's driven
  by…") follows the sentence before it. In a transcript only sentences with a
  driver cue (driven by, offset, timing, one-off, higher, lower …) start
  ticked. A sentence that names no line is kept, unticked, as whole-view
  commentary the reviewer can turn on.
- **Review, when wanted**: the *Review* link in the status line, or per file
  in the list, or *Review all*, opens one table of every sentence with the
  line each was matched to (changeable), how it matched, the other lines it
  names and any ambiguous term, plus a **source label** per file. Untick what
  does not belong and Save. An ambiguous term is settled by an Alias row in
  the config pack.
- **Where the sentences appear**: accepted sentences join the loaded
  commentary, so the House commentary, the Financial Summary, All
  commentaries, the dashboard widgets and the Excel / Word / PPT exports
  weave them in exactly as they weave in the config's paragraphs: appended to
  the sentence about that line, or leading the block when the block does not
  print the line by name. Scope-following applies too: a sentence naming
  Private Bank is held back while the page is filtered to Premier.
- **Attribution**: the sources are named once, in a line under each
  commentary block — "Commentary enriched from: Sep review notes, call
  transcript" — so the narrative itself stays clean. A checkbox adds the
  source label to every clause as well.
- **Precedence**: config commentary first by default; tick *prefer these
  sources* to put the uploads first.
- **Storage**: sentences live in this browser's localStorage, separate from
  the config, so a config re-upload does not wipe them. *Forget* removes one
  file, *Forget all sources* removes everything.

Sample sources to try are in `sample/commentary-sources/`.

## How it is wired

The addition is two `<script>` blocks appended before `</body>`, following the
file's convention of self-contained patch blocks:

1. the ECharts common build (~0.66 MB),
2. `interactive-charts.js` (kept here as a readable copy), which mounts the
   nav button and the page, wraps `setView`, `render` and `applyCfgDom`, and
   otherwise only calls existing functions, and
3. `enrich-commentary.js`, which mounts the side-pane section, the header
   button and the review dialog, wraps `comInLoaded` (the sentences join the
   loaded commentary), `comDigest` (the optional inline marker) and `render`
   (the sources line under each block), and otherwise calls `comMatchList`,
   `comVague` and `comIndex` as the config path does.

Nothing in the original code was edited.

## Files

| Path | Purpose |
|---|---|
| `FinSightFinal.html` | The dashboard with the Interactive charts page and Enrich commentary |
| `interactive-charts.js` | Readable copy of the Interactive charts page code |
| `enrich-commentary.js` | Readable copy of the Enrich commentary code |
| `sample/commentary-sources/` | Sample sources: a Word write-up, a `.vtt` transcript, a chat-style `.txt` |
| `sample/make_sample.py` | Generates a synthetic IWPB-shaped extract (needs `openpyxl`) |
| `sample/IWPB_SG_sample.xlsx` | Output of the generator, for demo/testing |
