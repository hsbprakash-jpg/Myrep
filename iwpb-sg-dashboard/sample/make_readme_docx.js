const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  LevelFormat, TableOfContents, PageNumber, Footer, BorderStyle,
  Table, TableRow, TableCell, WidthType, ShadingType,
} = require('docx');

const md = fs.readFileSync('/home/user/Myrep/iwpb-sg-dashboard/README.md', 'utf8');
const lines = md.split('\n');

// ---- block grouping: headings, bullets (with wrapped continuations),
// numbered items, code fences, plain paragraphs -----------------------------
const blocks = [];
let i = 0;
while (i < lines.length) {
  const ln = lines[i];
  if (/^\s*```/.test(ln)) {
    const ind = ln.match(/^\s*/)[0].length;
    const code = [];
    i++;
    while (i < lines.length && !/^\s*```/.test(lines[i])) { code.push(lines[i].slice(ind)); i++; }
    i++;
    blocks.push({ t: 'code', lines: code });
  } else if (/^# /.test(ln)) { blocks.push({ t: 'h1', text: ln.slice(2).trim() }); i++; }
  else if (/^### /.test(ln)) { blocks.push({ t: 'h3', text: ln.slice(4).trim() }); i++; }
  else if (/^## /.test(ln)) { blocks.push({ t: 'h2', text: ln.slice(3).trim() }); i++; }
  else if (/^\s*\|/.test(ln)) {
    // a markdown table, header row and all — the separator row is the giveaway
    const rows = [];
    while (i < lines.length && /^\s*\|/.test(lines[i])) {
      const cells = lines[i].trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim());
      if (!cells.every(c => /^:?-{2,}:?$/.test(c))) rows.push(cells);
      i++;
    }
    if (rows.length) blocks.push({ t: 'table', rows });
  }
  else if (/^\s*> /.test(ln)) {
    const quote = [];
    while (i < lines.length && /^\s*> ?/.test(lines[i])) { quote.push(lines[i].replace(/^\s*> ?/, '').trim()); i++; }
    blocks.push({ t: 'quote', text: quote.join(' ') });
  }
  else if (/^- /.test(ln)) {
    let text = ln.slice(2).trim();
    i++;
    // a wrapped bullet keeps going however deep the indent — the README wraps
    // its prose at seventy characters and indents sub-points further still
    while (i < lines.length && /^\s{2,}\S/.test(lines[i]) && !/^\s*- /.test(lines[i])
           && !/^\s*\|/.test(lines[i]) && !/^\s*> /.test(lines[i]) && !/^\s*```/.test(lines[i])) {
      text += ' ' + lines[i].trim(); i++;
    }
    blocks.push({ t: 'li', text });
  } else if (/^\d+\. /.test(ln)) {
    let text = ln.replace(/^\d+\. /, '').trim();
    i++;
    while (i < lines.length && /^   +\S/.test(lines[i]) && !/^\d+\. /.test(lines[i])) {
      text += ' ' + lines[i].trim(); i++;
    }
    blocks.push({ t: 'num', text });
  } else if (ln.trim() === '') { i++; }
  else {
    let text = ln.trim();
    i++;
    const indented = /^  /.test(ln);
    while (i < lines.length && lines[i].trim() !== '' &&
           !/^(#|##|- |\d+\. )/.test(lines[i]) && !/^\s*```/.test(lines[i])) {
      text += ' ' + lines[i].trim(); i++;
    }
    blocks.push({ t: 'p', text, indented });
  }
}

/* The README wraps its prose at seventy-odd characters, so one paragraph can
   run for twenty source lines and arrive here as a wall of text. Where the
   author started a new sub-topic with a bold lead-in — "…as planned. **The
   Financial Summary is the one view without arrows**: …" — that is a paragraph
   break in everything but name, so the document takes it as one. Nothing is
   reworded and no sentence moves; the text is only given air. */
function airy(text) {
  const LONG = 700, MIN = 280;
  if (text.length < LONG) return [text];
  // a break before a bold lead-in is the author's own; a plain sentence end is
  // the fallback, and both are only ever taken between sentences
  // "e.g." and its friends end in a full stop without ending a sentence
  const ABBR = /(?:e\.g\.|i\.e\.|etc\.|vs\.|cf\.|no\.|\b[A-Za-z]\.)$/i;
  const marks = (re) => { const out = []; let m; const r = new RegExp(re, 'g');
    while ((m = r.exec(text)) !== null)
      if (!ABBR.test(text.slice(Math.max(0, m.index - 6), m.index))) out.push(m.index);
    return out; };
  const bold = marks('(?<=[.!?]) (?=\\*\\*)');
  const any = marks('(?<=[.!?]) (?=[A-Z\\u201c\\*`])');
  const out = [];
  let start = 0;
  while (text.length - start > LONG) {
    const want = start + Math.min(LONG, Math.max(MIN, (text.length - start) / 2));
    const pick = (list) => list.filter(i => i > start + MIN && i < start + LONG * 1.6)
      .sort((a2, b2) => Math.abs(a2 - want) - Math.abs(b2 - want))[0];
    const cut = pick(bold) ?? pick(any);
    if (cut == null) break;
    out.push(text.slice(start, cut).trim());
    start = cut + 1;
  }
  out.push(text.slice(start).trim());
  return out.filter(Boolean);
}

// ---- inline markdown -> TextRuns: **bold**, `code`, *italic* ---------------
function runs(text, extra = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|`[^`]+`|\*[^*\s][^*]*\*)/g;
  let last = 0, m;
  const push = (t, o) => { if (t) out.push(new TextRun({ text: t, ...extra, ...o })); };
  while ((m = re.exec(text)) !== null) {
    push(text.slice(last, m.index), {});
    const tok = m[0];
    if (tok.startsWith('**')) push(tok.slice(2, -2), { bold: true });
    else if (tok.startsWith('`')) push(tok.slice(1, -1), { font: 'Consolas', shading: { type: 'clear', fill: 'F2F1ED' } });
    else push(tok.slice(1, -1), { italics: true });
    last = m.index + tok.length;
  }
  push(text.slice(last), {});
  return out;
}

const RED = 'DB0011', INK = '1A1A19', GREY = '52514E';
const children = [];
let first = true;
for (const b of blocks) {
  if (b.t === 'h1') {
    children.push(new Paragraph({
      heading: HeadingLevel.TITLE,
      spacing: { after: 120 },
      children: [new TextRun({ text: b.text, bold: true, color: INK })],
    }));
    children.push(new Paragraph({
      spacing: { after: 240 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 24, color: RED } },
      children: [],
    }));
    if (first) {
      children.push(new Paragraph({
        children: [new TextRun({ text: 'Contents', bold: true, size: 26, color: INK })],
        spacing: { before: 120, after: 120 },
      }));
      children.push(new TableOfContents('Contents', { hyperlink: true, headingStyleRange: '1-3' }));
      first = false;
    }
  } else if (b.t === 'h2') {
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 320, after: 140 },
      children: [new TextRun({ text: b.text, bold: true, color: RED })],
    }));
  } else if (b.t === 'h3') {
    children.push(new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 260, after: 110 },
      children: [new TextRun({ text: b.text, bold: true, color: INK })],
    }));
  } else if (b.t === 'quote') {
    // an example of what the page writes: set apart, in the page's own red rule
    children.push(new Paragraph({
      spacing: { before: 60, after: 160 },
      indent: { left: 360 },
      border: { left: { style: BorderStyle.SINGLE, size: 12, color: RED, space: 12 } },
      children: runs(b.text, { italics: true }),
    }));
  } else if (b.t === 'table') {
    const head = b.rows[0], body = b.rows.slice(1);
    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          tableHeader: true,
          children: head.map(c => new TableCell({
            shading: { type: ShadingType.CLEAR, fill: 'F7E9EA' },
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [new Paragraph({ children: runs(c || ' ', { bold: true, color: RED }) })],
          })),
        }),
        ...body.map(r => new TableRow({
          children: r.map(c => new TableCell({
            margins: { top: 60, bottom: 60, left: 100, right: 100 },
            children: [new Paragraph({ spacing: { after: 0 }, children: runs(c || ' ') })],
          })),
        })),
      ],
    }));
    children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
  } else if (b.t === 'li') {
    const parts = airy(b.text);
    children.push(new Paragraph({
      numbering: { reference: 'bullets', level: 0 },
      spacing: { after: parts.length > 1 ? 100 : 80 },
      children: runs(parts[0]),
    }));
    // what followed the lead-in keeps the bullet's indent without a second dot
    for (const extra of parts.slice(1))
      children.push(new Paragraph({
        indent: { left: 360 }, spacing: { after: 100 }, children: runs(extra),
      }));
  } else if (b.t === 'num') {
    children.push(new Paragraph({
      numbering: { reference: 'nums', level: 0 },
      spacing: { after: 80 },
      children: runs(b.text),
    }));
  } else if (b.t === 'code') {
    for (const cl of b.lines) {
      children.push(new Paragraph({
        spacing: { after: 0 },
        shading: { type: 'clear', fill: 'F2F1ED' },
        children: [new TextRun({ text: cl || ' ', font: 'Consolas', size: 18 })],
      }));
    }
    children.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  } else {
    for (const part of airy(b.text))
      children.push(new Paragraph({ spacing: { after: 140 },
        indent: b.indented ? { left: 360 } : undefined, children: runs(part) }));
  }
}

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Calibri', size: 21, color: INK } },
      title: { run: { size: 44, font: 'Calibri Light' } },
      heading1: { run: { size: 28, font: 'Calibri' }, paragraph: {} },
      heading2: { run: { size: 24, font: 'Calibri' }, paragraph: {} },
    },
  },
  numbering: {
    config: [
      { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 200 } } } }] },
      { reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] },
    ],
  },
  features: { updateFields: true },
  sections: [{
    properties: { page: { margin: { top: 1080, bottom: 1080, left: 1180, right: 1180 } } },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({ text: 'IWPB Singapore Driller — README   ', color: GREY, size: 16 }),
            new TextRun({ children: [PageNumber.CURRENT], color: GREY, size: 16 }),
          ],
        })],
      }),
    },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('IWPB_SG_Dashboard_README.docx', buf);
  console.log('written', buf.length, 'bytes');
});
