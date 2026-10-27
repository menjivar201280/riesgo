// Utilidades de maquetación para la guía (docx-js).
const {
  Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, HeadingLevel, PageBreak, VerticalAlign,
  TableLayoutType,
} = require("docx");

const C = {
  navy: "1F3864",
  blue: "2E75B6",
  light: "DEEAF6",
  pale: "F2F7FC",
  grey: "595959",
  border: "9DB7D5",
  white: "FFFFFF",
  amber: "FFF4E0",
  amberB: "C77C02",
  red: "FBE9E7",
  redB: "B23B2A",
  green: "E8F5E9",
  greenB: "2E7D32",
};
const FONT = "Calibri";
const PORTRAIT_W = 9360; // Carta, márgenes 1"
const LANDSCAPE_W = 12960;

// Convierte "texto con **negrita** y *cursiva*" en TextRuns.
function runs(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g;
  let last = 0, m;
  const s = String(text);
  while ((m = re.exec(s)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: s.slice(last, m.index), ...base }));
    const tok = m[0];
    if (tok.startsWith("**")) out.push(new TextRun({ text: tok.slice(2, -2), ...base, bold: true }));
    else out.push(new TextRun({ text: tok.slice(1, -1), ...base, italics: true }));
    last = m.index + tok.length;
  }
  if (last < s.length) out.push(new TextRun({ text: s.slice(last), ...base }));
  return out;
}

const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, pageBreakBefore: true, children: [new TextRun(t)] });
const H1nb = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, keepNext: true, children: [new TextRun(t)] });
const H3 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_3, keepNext: true, children: [new TextRun(t)] });

function P(text, o = {}) {
  return new Paragraph({
    alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { after: o.after ?? 120, before: o.before ?? 0 },
    keepNext: o.keepNext || false,
    indent: o.indent ? { left: o.indent } : undefined,
    children: runs(text, { size: o.size, color: o.color, bold: o.bold, italics: o.italics }),
  });
}

let listInstance = 0;
function B(items, level = 0) {
  const inst = ++listInstance;
  const out = [];
  for (const it of items) {
    if (Array.isArray(it)) { out.push(...B(it, level + 1)); continue; }
    out.push(new Paragraph({
      numbering: { reference: "bullets", level, instance: inst },
      spacing: { after: 60 },
      alignment: AlignmentType.LEFT,
      children: runs(it),
    }));
  }
  return out;
}
function N(items) {
  const inst = ++listInstance;
  return items.map((it) => new Paragraph({
    numbering: { reference: "numbers", level: 0, instance: inst },
    spacing: { after: 60 },
    alignment: AlignmentType.LEFT,
    children: runs(it),
  }));
}

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: C.border };
const borders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

function cell(content, width, o = {}) {
  const paras = (Array.isArray(content) ? content : [content]).map((c) =>
    c instanceof Paragraph ? c : new Paragraph({
      alignment: o.align || AlignmentType.LEFT,
      keepNext: o.keepNext || false,
      spacing: { after: 20, before: 20, line: 240 },
      children: runs(c, { bold: o.bold, color: o.color, size: o.size }),
    }));
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders,
    verticalAlign: o.vAlign || VerticalAlign.TOP,
    shading: o.fill ? { type: ShadingType.CLEAR, color: "auto", fill: o.fill } : undefined,
    margins: { top: 50, bottom: 50, left: 90, right: 90 },
    columnSpan: o.span,
    children: paras,
  });
}

// Tabla con encabezado repetible. widths en fracciones o DXA.
function T(headers, rows, widths, o = {}) {
  const total = o.total || PORTRAIT_W;
  const sum = widths.reduce((a, b) => a + b, 0);
  let w = widths.map((x) => Math.floor((x / sum) * total));
  w[w.length - 1] += total - w.reduce((a, b) => a + b, 0);
  const size = o.size || 18;
  const trs = [];
  if (headers) {
    trs.push(new TableRow({
      tableHeader: true, cantSplit: true,
      children: headers.map((h, i) => cell(h, w[i], { bold: true, color: C.white, fill: C.navy, size, keepNext: true })),
    }));
  }
  rows.forEach((r, ri) => {
    const isTotal = o.totalRows && o.totalRows.includes(ri);
    trs.push(new TableRow({
      cantSplit: true,
      children: r.map((c, i) => cell(c, w[i], {
        size,
        bold: isTotal || (o.firstBold && i === 0),
        fill: isTotal ? C.light : (o.zebra !== false && ri % 2 === 1 ? C.pale : undefined),
        align: o.alignCols && o.alignCols[i],
        keepNext: rows.length <= 8 && ri < rows.length - 1,
      })),
    }));
  });
  const tbl = new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: w, layout: TableLayoutType.FIXED, rows: trs });
  const out = [];
  if (o.caption) out.push(new Paragraph({ keepNext: true, spacing: { before: 120, after: 60 }, children: runs(o.caption, { bold: true, color: C.navy, size: 19 }) }));
  out.push(tbl);
  out.push(new Paragraph({ spacing: { after: 120 }, children: [] }));
  return out;
}

// Recuadro de una celda (fórmulas, avisos, ejemplos).
function box(title, lines, fill, accent, o = {}) {
  const total = o.total || PORTRAIT_W;
  const children = [];
  if (title) children.push(new Paragraph({ spacing: { after: 60 }, children: runs(title, { bold: true, color: accent, size: 20 }) }));
  for (const l of lines) {
    if (l instanceof Paragraph) { children.push(l); continue; }
    children.push(new Paragraph({
      alignment: o.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      spacing: { after: 50 },
      children: runs(l, { size: o.size || 20, font: o.mono ? "Consolas" : undefined }),
    }));
  }
  const side = { style: BorderStyle.SINGLE, size: 4, color: accent };
  const tbl = new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: [total], layout: TableLayoutType.FIXED,
    rows: [new TableRow({ cantSplit: !o.split, children: [new TableCell({
      width: { size: total, type: WidthType.DXA },
      borders: { top: side, bottom: side, right: side, left: { style: BorderStyle.SINGLE, size: 24, color: accent } },
      shading: { type: ShadingType.CLEAR, color: "auto", fill },
      margins: { top: 100, bottom: 100, left: 180, right: 180 },
      children,
    })] })],
  });
  return [tbl, new Paragraph({ spacing: { after: 120 }, children: [] })];
}
const F = (title, lines) => box("Fórmula — " + title, lines, C.light, C.navy, { center: true });
const EX = (title, lines, o = {}) => box("Ejemplo (datos ficticios) — " + title, lines, C.pale, C.blue, { split: true, ...o });
const LEGAL = (lines) => box("Referencia normativa", lines, C.green, C.greenB);
const GOOD = (lines) => box("Buena práctica recomendada (no constituye obligación legal)", lines, C.pale, C.blue);
const ALERT = (lines, t) => box(t || "Señal de alerta / error que debe evitarse", lines, C.red, C.redB);
const NOTE = (lines, t) => box(t || "Nota metodológica", lines, C.amber, C.amberB);

// Ficha estándar del concepto: definición, utilidad, aplicación, fórmula, ejemplo, interpretación, errores, decisión.
function ficha(rows) {
  const labels = ["Definición", "Para qué sirve", "Cómo se aplica", "Fórmula", "Ejemplo", "Interpretación", "Limitaciones / errores a evitar", "Influencia en la decisión"];
  const keys = ["def", "uso", "aplica", "formula", "ejemplo", "interp", "errores", "decision"];
  const data = keys.map((k, i) => rows[k] ? [labels[i], rows[k]] : null).filter(Boolean);
  return T(null, data, [26, 74], { firstBold: true, zebra: true, size: 19, caption: rows.titulo });
}

const spacer = () => new Paragraph({ spacing: { after: 80 }, children: [] });
const pageBreak = () => new Paragraph({ children: [new PageBreak()] });
const LANDSCAPE = { __section: "landscape" };
const PORTRAIT = { __section: "portrait" };

module.exports = {
  C, FONT, PORTRAIT_W, LANDSCAPE_W, runs, H1, H1nb, H2, H3, P, B, N, T, box, F, EX, LEGAL, GOOD, ALERT, NOTE,
  ficha, spacer, pageBreak, LANDSCAPE, PORTRAIT, cell,
};
