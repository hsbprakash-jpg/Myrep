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
templates, plus an **⤒ Enrich** button on every AI commentary block, on the
same row as Retrieve and Regenerate.
Anything written about the numbers — a write-up, review notes, a transcript —
can be added, one file or many at once, by the picker or by dropping the files
on the panel or on the commentary block. The sentences enrich the commentary
the moment the files are read; review is optional.

**The uploads are the only written commentary the page uses.** The config
pack's Commentary sheet is no longer read for narrative: with nothing
uploaded, every block reads its own figures; with uploads, their sentences
are what the page weaves in. (The config pack's Aliases, Labels and
commentary *templates* still apply — those shape matching and style, not
the words.)

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
- **Scope and roll-up**: every upload is stamped with the scope the page was
  in at the time — business, region, country, with the parents filled in
  from the data, so a Singapore upload knows it belongs to ASEAN and IWPB.
  At that scope, or a narrower one, the sentences read as written. At a
  wider scope they roll up automatically as one clause per line that names
  each place: at Group level, Fee income reads "Singapore: came in below
  forecast because … · Hong Kong: fell short of forecast because …". Under a
  different country they are held back. An upload made with no scope set
  applies to the whole book. The file list and the review dialog show the
  scope each file was uploaded under.
- **Attribution**: the sources are named once, in a line under each
  commentary block — "Commentary enriched from: Sep review notes, call
  transcript" — so the narrative itself stays clean. A checkbox adds the
  source label to every clause as well.
- **Storage**: sentences live in this browser's localStorage, separate from
  the config, so a config re-upload does not wipe them. *Forget* removes one
  file, *Forget all sources* removes everything.

Sample sources to try are in `sample/commentary-sources/`.

### AI review

Nothing to configure. When the page opens it looks for a model server on
this desktop — Ollama at port 11434 or LM Studio at port 1234 — and, if one
answers, picks the first chat-capable model it lists. One line under the
Enrich commentary section says what will decide, for example
"AI review: llama3.1:8b-instruct on this desktop (Ollama) decides which
sentences to use", with *Switch off* and *Look again* beside it. If nothing
is found, the line says so and the rules decide alone.

When a model is in play, the rules run first and then the model decides for
every sentence whether it is a driver, which line it belongs to and how sure
it is. Answers at or above the confidence threshold (0.7) are stitched with
no ticking; doubtful ones are left unticked with the model's reason shown in
the AI column of Review. If the model stops answering, the rules' decisions
stand and the status line says so.

- A desktop model keeps everything on this machine. Ollama must be started
  with `OLLAMA_ORIGINS=*` so a page opened from a file may call it; LM Studio
  needs CORS switched on. An 8B-class instruct model is enough for this task.
- **Advanced** (collapsed by default) holds everything else: mark the
  source on each clause; a
  custom model endpoint, model name and key — an approved OpenAI-style
  gateway, or Claude via the Anthropic Messages API (structured output,
  default model `claude-opus-5`), recognised from the URL; the confidence
  threshold; a reword toggle that rewrites spoken sentences into pack style
  (the original is kept and shown on hover); and a *Test the model* button.
  A custom endpoint sends sentences off the machine, so use one only where
  that is allowed; a key typed into the page stays in this browser's
  localStorage.
- The config pack's Settings sheet can supply team defaults: `ai_endpoint`,
  `ai_model`, `ai_key`, `ai_threshold`, `ai_rewrite`.
- `sample/mock-model-server.js` is a stand-in model for testing the flow
  without a real one: `node sample/mock-model-server.js 11434` answers the
  model listing, OpenAI-style chat and the Anthropic shape.

## How it is wired

The addition is two `<script>` blocks appended before `</body>`, following the
file's convention of self-contained patch blocks:

1. the ECharts common build (~0.66 MB),
2. `interactive-charts.js` (kept here as a readable copy), which mounts the
   nav button and the page, wraps `setView`, `render` and `applyCfgDom`, and
   otherwise only calls existing functions, and
3. `enrich-commentary.js`, which mounts the side-pane section and the
   review dialog, wraps `aiComment` (the Enrich button on each commentary
   block's Retrieve / Regenerate row), wraps `comInLoaded` (the sentences join the
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
| `sample/commentary-sources/` | Sample sources: a Word write-up, a `.vtt` transcript, a chat-style `.txt`, Hong Kong notes for the roll-up |
| `sample/mock-model-server.js` | Stand-in model server for testing AI review without a real model |
| `sample/make_sample.py` | Generates a synthetic IWPB-shaped extract (needs `openpyxl`); add a second argument for two countries |
| `sample/IWPB_SG_sample.xlsx` | Output of the generator, one country, for demo/testing |
| `sample/IWPB_2countries_sample.xlsx` | Two countries (Singapore / ASEAN, Hong Kong / North Asia), for the roll-up |
