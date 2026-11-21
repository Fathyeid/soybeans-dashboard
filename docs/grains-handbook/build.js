// Builds the Grains Trading Handbook (.docx) from the markup files in ./content
// Markup:
//   @PART title | subtitle      -> part divider page
//   # Chapter                   -> Heading 1 (new page)
//   ## / ###                    -> Heading 2 / 3
//   - item / "  - item"         -> bullets (2 levels)
//   1) item                     -> numbered list (restarts per list)
//   | a | b |                   -> table (first row = header; |---| rows skipped)
//   > text                      -> tip box   (نصيحة من السوق)
//   !> text                     -> warning box (تنبيه)
//   ?> text                     -> example box (مثال عملي)
//   $$ text                     -> centered formula line (LTR)
//   **bold** inline
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, PageBreak, Header, Footer, PageNumber, TableOfContents,
  LevelFormat, TabStopType,
} = require("docx");

const FONT = "Arial";
const BODY = 26; // 13pt
const CONTENT_W = 9306;
const C = { navy: "1F3864", green: "2E7D32", gold: "B8860B", red: "B71C1C", blue: "1565C0", grey: "F2F2F2" };

function runs(text, base = {}) {
  const out = [];
  const parts = text.split(/(\*\*[^*]+\*\*)/g).filter(Boolean);
  for (const p of parts) {
    const bold = p.startsWith("**") && p.endsWith("**");
    out.push(new TextRun({
      text: bold ? p.slice(2, -2) : p, bold: bold || base.bold, color: base.color,
      size: base.size || BODY, font: { ascii: FONT, hAnsi: FONT, cs: FONT },
      rightToLeft: base.ltr ? false : true, boldComplexScript: bold || base.bold,
      sizeComplexScript: base.size || BODY,
    }));
  }
  return out;
}

const isLatin = t => !/[\u0600-\u06FF]/.test(t);
function para(text, opts = {}) {
  if (opts.ltr === undefined && isLatin(text) && text.trim()) { opts = { ...opts, ltr: true, align: opts.align || AlignmentType.LEFT }; }
  return new Paragraph({
    bidirectional: !opts.ltr, alignment: opts.align || AlignmentType.JUSTIFIED,
    spacing: { after: opts.after ?? 120, line: 336, before: opts.before ?? 0 },
    children: runs(text, opts), ...(opts.extra || {}),
  });
}

const BOX = {
  ">": { label: "نصيحة من السوق", color: C.green, fill: "EAF5EA" },
  "!>": { label: "تنبيه مهم", color: C.red, fill: "FDECEA" },
  "?>": { label: "مثال عملي", color: C.blue, fill: "E8F1FB" },
};

function box(kind, lines) {
  const b = BOX[kind];
  const cellChildren = [para(b.label, { bold: true, color: b.color, after: 80 })];
  for (const l of lines) {
    if (/^\s*-\s+/.test(l)) cellChildren.push(new Paragraph({
      bidirectional: true, numbering: { reference: "bullets", level: 0 }, spacing: { after: 60, line: 320 },
      children: runs(l.replace(/^\s*-\s+/, "")) }));
    else if (l.startsWith("$$")) cellChildren.push(para(l.slice(2).trim(), { ltr: true, align: AlignmentType.CENTER, bold: true }));
    else cellChildren.push(para(l, { after: 80, ...(isLatin(l) ? { size: 21 } : {}) }));
  }
  const border = { style: BorderStyle.SINGLE, size: 4, color: b.color };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W], visuallyRightToLeft: true,
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA }, shading: { fill: b.fill, type: ShadingType.CLEAR, color: "auto" },
      margins: { top: 120, bottom: 120, left: 180, right: 180 },
      borders: { top: border, bottom: border, left: border, right: { style: BorderStyle.SINGLE, size: 24, color: b.color } },
      children: cellChildren })] })],
  });
}

function table(rows) {
  const cells = rows.map(r => r.trim().replace(/^\||\|$/g, "").split("|").map(s => s.trim()));
  const n = Math.max(...cells.map(r => r.length));
  cells.forEach(r => { while (r.length < n) r.push(""); });
  const lens = Array.from({ length: n }, (_, i) => Math.max(6, Math.min(45, ...[Math.max(...cells.map(r => r[i].length))])));
  const tot = lens.reduce((a, b) => a + b, 0);
  let widths = lens.map(l => Math.floor(CONTENT_W * l / tot));
  widths[widths.length - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "A6A6A6" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, visuallyRightToLeft: true,
    rows: cells.map((r, ri) => new TableRow({
      tableHeader: ri === 0, cantSplit: true,
      children: r.map((c, ci) => new TableCell({
        width: { size: widths[ci], type: WidthType.DXA }, borders,
        shading: { fill: ri === 0 ? C.navy : (ri % 2 === 0 ? C.grey : "FFFFFF"), type: ShadingType.CLEAR, color: "auto" },
        margins: { top: 60, bottom: 60, left: 100, right: 100 },
        children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 0, line: 280 },
          children: runs(c, { size: 21, bold: ri === 0, color: ri === 0 ? "FFFFFF" : undefined }) })],
      })),
    })),
  });
}

function heading(text, level) {
  const map = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  return new Paragraph({ heading: map[level], bidirectional: true, pageBreakBefore: level === 1,
    children: [new TextRun({ text, rightToLeft: true, font: { ascii: FONT, hAnsi: FONT, cs: FONT } })] });
}

function partPage(title, sub) {
  return [
    new Paragraph({ children: [new PageBreak()] }),
    ...Array.from({ length: 8 }, () => new Paragraph({ children: [] })),
    new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER, spacing: { after: 300 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.gold, space: 8 } },
      children: runs(title, { size: 56, bold: true, color: C.navy }) }),
    para(sub || "", { align: AlignmentType.CENTER, size: 30, color: "595959" }),
  ];
}

let numInstance = 0;
function parse(src) {
  const out = []; const lines = src.split(/\r?\n/);
  let i = 0;
  while (i < lines.length) {
    let l = lines[i];
    if (!l.trim()) { i++; continue; }
    if (l.startsWith("@PART")) { const [t, s] = l.slice(5).split("|").map(x => x.trim()); out.push(...partPage(t, s)); i++; continue; }
    if (l.startsWith("### ")) { out.push(heading(l.slice(4).trim(), 3)); i++; continue; }
    if (l.startsWith("## ")) { out.push(heading(l.slice(3).trim(), 2)); i++; continue; }
    if (l.startsWith("# ")) { out.push(heading(l.slice(2).trim(), 1)); i++; continue; }
    const bk = l.match(/^(!>|\?>|>)\s?/);
    if (bk) {
      const kind = bk[1]; const buf = [];
      while (i < lines.length && lines[i].startsWith(kind) && !(kind === ">" && /^(!>|\?>)/.test(lines[i]))) {
        buf.push(lines[i].slice(kind.length).trim()); i++;
      }
      out.push(box(kind, buf.filter(Boolean))); out.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
      continue;
    }
    if (l.trim().startsWith("|")) {
      const buf = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) { if (!/^\s*\|[\s\-:|]+\|\s*$/.test(lines[i])) buf.push(lines[i]); i++; }
      out.push(table(buf)); out.push(new Paragraph({ spacing: { after: 60 }, children: [] }));
      continue;
    }
    if (l.startsWith("$$")) { out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80, after: 160 },
      shading: { fill: "FFF8E1", type: ShadingType.CLEAR, color: "auto" },
      children: runs(l.slice(2).trim(), { bold: true, color: C.navy, ltr: /^[\x00-\x7F]*$/.test(l) }) , bidirectional: !/^[\x00-\x7F]*$/.test(l)})); i++; continue; }
    const bm = l.match(/^(\s*)-\s+(.*)$/);
    if (bm) { out.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.JUSTIFIED,
      numbering: { reference: "bullets", level: bm[1].length >= 2 ? 1 : 0 }, spacing: { after: 60, line: 320 }, children: runs(bm[2]) })); i++; continue; }
    if (/^\d+\)\s+/.test(l)) {
      numInstance++;
      while (i < lines.length && /^\d+\)\s+/.test(lines[i])) {
        out.push(new Paragraph({ bidirectional: true, alignment: AlignmentType.JUSTIFIED,
          numbering: { reference: "numbers", level: 0, instance: numInstance }, spacing: { after: 60, line: 320 },
          children: runs(lines[i].replace(/^\d+\)\s+/, "")) })); i++;
      }
      continue;
    }
    out.push(para(l.trim())); i++;
  }
  return out;
}

function cover() {
  return [
    ...Array.from({ length: 6 }, () => new Paragraph({ children: [] })),
    para("الدليل الشامل لتاجر الحبوب والزيوت", { align: AlignmentType.CENTER, size: 64, bold: true, color: C.navy, after: 200 }),
    para("The Complete Grains & Oilseeds Trader Handbook", { align: AlignmentType.CENTER, size: 34, color: C.gold, ltr: true, after: 400 }),
    new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.gold, space: 4 } }, children: [] }),
    para("فول الصويا • القمح • الذرة • الزيوت النباتية", { align: AlignmentType.CENTER, size: 32, before: 300 }),
    para("من العرض الأول حتى التفريغ في المخزن: المواصفات، العقود، الإنكوترمز، الشحن، التفتيش، المستندات، الدفع، التخليص الجمركي، المخاطر", { align: AlignmentType.CENTER, size: 26, color: "595959" }),
    para("ومعه دليل البورصة: العقود الآجلة، الخيارات، بورصات السلع والأسهم", { align: AlignmentType.CENTER, size: 26, color: "595959", after: 600 }),
    para("مرجع مهني للتاجر الدولي في الاستيراد والتصدير والتجارة بين الأطراف", { align: AlignmentType.CENTER, size: 24, bold: true }),
    para("الإصدار الأول — نوفمبر 2026", { align: AlignmentType.CENTER, size: 24, color: "595959" }),
    new Paragraph({ children: [new PageBreak()] }),
    para("المحتويات", { align: AlignmentType.CENTER, size: 40, bold: true, color: C.navy, after: 240 }),
    new TableOfContents("المحتويات", { hyperlink: true, headingStyleRange: "1-2" }),
  ];
}

const files = fs.readdirSync(path.join(__dirname, "content")).filter(f => f.endsWith(".md")).sort();
const sources = files.map(f => fs.readFileSync(path.join(__dirname, "content", f), "utf8"));

// Renumber chapters sequentially (files may use provisional tokens like "4أ" or "ب1").
const chapMap = new Map(); let chapNo = 0;
for (const s of sources) for (const l of s.split(/\r?\n/)) {
  const m = l.match(/^# الفصل ([^:\s]+):/); if (m) chapMap.set(m[1], String(++chapNo));
}
function renumber(src) {
  let cur = null;
  return src.split(/\r?\n/).map(l => {
    const ch = l.match(/^# الفصل ([^:\s]+):/);
    if (ch) { cur = chapMap.get(ch[1]); return l.replace(/^# الفصل [^:\s]+:/, `# الفصل ${cur}:`); }
    const an = l.match(/^# ملحق (\d+):/); if (an) { cur = null; return l; }
    if (cur && /^#{2,3} [^\s.]+\.\d/.test(l)) l = l.replace(/^(#{2,3}) [^\s.]+\.(\d)/, `$1 ${cur}.$2`);
    return l.replace(/الفصل ([^\s:،.)]+)(?=[\s:،.)]|$)/g, (m0, tok) => chapMap.has(tok) ? `الفصل ${chapMap.get(tok)}` : m0);
  }).join("\n");
}
const body = [];
for (const s of sources) body.push(...parse(renumber(s)));

const hStyle = (size, color, before, after) => ({ run: { font: FONT, size, bold: true, color, sizeComplexScript: size, boldComplexScript: true },
  paragraph: { spacing: { before, after }, keepNext: true, keepLines: true } });

const doc = new Document({
  creator: "Grains Trading Handbook", title: "الدليل الشامل لتاجر الحبوب والزيوت",
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: FONT, size: BODY, sizeComplexScript: BODY } },
      heading1: hStyle(40, C.navy, 240, 240), heading2: hStyle(32, C.gold, 300, 120), heading3: hStyle(27, C.navy, 220, 100) },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 40, bold: true, color: C.navy, sizeComplexScript: 40, boldComplexScript: true },
        paragraph: { spacing: { before: 240, after: 240 }, keepNext: true, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.gold, space: 6 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 32, bold: true, color: C.gold, sizeComplexScript: 32, boldComplexScript: true },
        paragraph: { spacing: { before: 300, after: 120 }, keepNext: true, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT, size: 27, bold: true, color: C.navy, sizeComplexScript: 27, boldComplexScript: true },
        paragraph: { spacing: { before: 220, after: 100 }, keepNext: true, outlineLevel: 2 } },
    ],
  },
  numbering: { config: [
    { reference: "bullets", levels: [
      { level: 0, format: LevelFormat.BULLET, text: "●", alignment: AlignmentType.RIGHT, style: { paragraph: { indent: { right: 500, hanging: 280 } }, run: { color: C.gold } } },
      { level: 1, format: LevelFormat.BULLET, text: "○", alignment: AlignmentType.RIGHT, style: { paragraph: { indent: { right: 1000, hanging: 280 } } } } ] },
    { reference: "numbers", levels: [
      { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.RIGHT, style: { paragraph: { indent: { right: 560, hanging: 340 } }, run: { bold: true, color: C.navy } } } ] },
  ] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 1300, left: 1300, right: 1300 } } },
    headers: { default: new Header({ children: [new Paragraph({ bidirectional: true, alignment: AlignmentType.CENTER,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.gold, space: 2 } },
      children: runs("الدليل الشامل لتاجر الحبوب والزيوت", { size: 18, color: "7F7F7F" }) })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 20, color: "7F7F7F" })] })] }) },
    children: [...cover(), ...body],
  }],
});

const outName = process.argv[2] || path.join(__dirname, "Grains_Trading_Handbook.docx");
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(outName, buf); console.log("written", outName, files.length, "files"); });
