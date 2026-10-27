// Genera Guia_Integral_Evaluacion_Crediticia_El_Salvador_2026.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, LevelFormat,
  PageNumber, PageOrientation, TableOfContents, BorderStyle, ShadingType, SectionType,
  Table, TableRow, TableCell, WidthType, TableLayoutType,
} = require("docx");
const { C, FONT, H1nb, P } = require("./lib");

const content = [
  require("./p0_preliminar")(),
  ...require("./p1_cap01_04")(),
  ...require("./p2_cap05_08")(),
  ...require("./p3_cap09_10")(),
  ...require("./p4_cap11_13")(),
  ...require("./p5_cap14_16")(),
  ...require("./p6_cap17_anexos")(),
].flat(Infinity);

const TITLE = "Guía Integral de Evaluación Crediticia en El Salvador";

const header = () => new Header({
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.blue, space: 4 } },
    children: [new TextRun({ text: TITLE + " · Actualización 2026", size: 16, color: C.grey })],
  })],
});
const footer = () => new Footer({
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.border, space: 4 } },
    children: [
      new TextRun({ text: "Documento de referencia metodológica · Ejemplos con datos ficticios · Página ", size: 16, color: C.grey }),
      new TextRun({ children: [PageNumber.CURRENT], size: 16, color: C.grey }),
      new TextRun({ text: " de ", size: 16, color: C.grey }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: C.grey }),
    ],
  })],
});

const PAGE_P = { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1300, left: 1440, header: 600, footer: 600 } };
const PAGE_L = { size: { width: 12240, height: 15840, orientation: PageOrientation.LANDSCAPE }, margin: { top: 1080, right: 1440, bottom: 1000, left: 1440, header: 500, footer: 500 } };

// Divide el contenido en secciones según los marcadores de orientación.
const sections = [];
let cur = { orient: "portrait", children: [] };
for (const el of content) {
  if (el && el.__section) {
    if (cur.children.length) sections.push(cur);
    cur = { orient: el.__section, children: [] };
  } else cur.children.push(el);
}
if (cur.children.length) sections.push(cur);

const coverBox = (lines, fill) => new Table({
  width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], layout: TableLayoutType.FIXED,
  rows: [new TableRow({ children: [new TableCell({
    width: { size: 9360, type: WidthType.DXA },
    borders: { top: NB, bottom: NB, left: NB, right: NB },
    shading: { type: ShadingType.CLEAR, color: "auto", fill },
    margins: { top: 300, bottom: 300, left: 360, right: 360 },
    children: lines.map(([text, size, color, bold]) => new Paragraph({ spacing: { after: 80, line: 380, lineRule: "atLeast" }, children: [new TextRun({ text, size, color, bold, font: FONT })] })),
  })] })],
});
const NB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };

const cover = [
  new Paragraph({ spacing: { before: 1400 }, children: [] }),
  coverBox([["GUÍA INTEGRAL DE", 40, C.white, true], ["EVALUACIÓN CREDITICIA", 56, C.white, true], ["EFECTIVA EN EL SALVADOR", 40, C.white, true]], C.navy),
  coverBox([["Actualización 2026", 30, C.navy, true], ["Metodología, normativa de referencia, casos prácticos y herramientas", 24, C.navy, false]], C.light),
  new Paragraph({ spacing: { before: 900 }, children: [] }),
  P("Desde la recepción del expediente hasta el seguimiento posterior al desembolso", { align: AlignmentType.LEFT, color: C.navy, size: 24, bold: true }),
  P("Capacidad de pago · Historial crediticio · Garantías · Gestión de riesgos · Calificación interna · Informe · Trazabilidad · Seguimiento", { align: AlignmentType.LEFT, color: C.grey, size: 20 }),
  new Paragraph({ spacing: { before: 1600 }, children: [] }),
  P("**Versión:** 1.0 · **Fecha de elaboración:** octubre de 2026 · **Fecha de consulta normativa:** 27/10/2026", { align: AlignmentType.LEFT, size: 19 }),
  P("**Moneda de los ejemplos:** dólares de los Estados Unidos (US$). Todos los nombres, casos y cifras de ejemplo son ficticios.", { align: AlignmentType.LEFT, size: 19 }),
  P("Documento de referencia metodológica. No sustituye la normativa vigente ni las políticas aprobadas por cada institución.", { align: AlignmentType.LEFT, size: 19, italics: true, color: C.grey }),
];

const tocSection = [
  P("Índice", { size: 34, bold: true, color: C.navy, align: AlignmentType.LEFT, after: 200 }),
  P("*Para actualizar el índice en Microsoft Word: clic derecho sobre el índice > «Actualizar campos» > «Actualizar toda la tabla».*", { size: 18, color: C.grey }),
  new TableOfContents("Índice", { hyperlink: true, headingStyleRange: "1-2" }),
];

const docSections = [
  { properties: { page: PAGE_P, titlePage: true }, children: cover },
  { properties: { page: PAGE_P, type: SectionType.NEXT_PAGE }, headers: { default: header() }, footers: { default: footer() }, children: tocSection },
  ...sections.map((s) => ({
    properties: { page: s.orient === "landscape" ? PAGE_L : PAGE_P, type: SectionType.NEXT_PAGE },
    headers: { default: header() }, footers: { default: footer() },
    children: s.children,
  })),
];

const doc = new Document({
  creator: "Equipo de metodología de riesgo crediticio",
  title: "Guía Integral de Evaluación Crediticia en El Salvador — Actualización 2026",
  description: "Metodología integral de evaluación crediticia adaptada a El Salvador",
  features: { updateFields: true },
  styles: {
    default: { document: { run: { font: FONT, size: 21 }, paragraph: { spacing: { line: 276 } } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 34, bold: true, color: C.navy, font: FONT },
        paragraph: { spacing: { before: 120, after: 240 }, outlineLevel: 0, keepNext: true,
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.blue, space: 6 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, color: C.blue, font: FONT },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1, keepNext: true } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, color: C.navy, font: FONT },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2, keepNext: true } },
      { id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal", run: { bold: true, color: C.navy, size: 21 },
        paragraph: { spacing: { before: 100, after: 20 }, tabStops: [{ type: "right", position: 9350, leader: "dot" }] } },
      { id: "TOC2", name: "toc 2", basedOn: "Normal", next: "Normal", run: { size: 19 },
        paragraph: { spacing: { before: 0, after: 0 }, indent: { left: 360 }, tabStops: [{ type: "right", position: 9350, leader: "dot" }] } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1000, hanging: 270 } } } },
        { level: 2, format: LevelFormat.BULLET, text: "▪", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1440, hanging: 270 } } } },
      ] },
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 340 } } } },
      ] },
    ],
  },
  sections: docSections,
});

const out = path.resolve(__dirname, "..", "Guia_Integral_Evaluacion_Crediticia_El_Salvador_2026.docx");
Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(out, buf); console.log("OK", out, buf.length); });
