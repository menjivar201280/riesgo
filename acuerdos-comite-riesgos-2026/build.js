const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const THEME = {
  name: "Comite de Riesgos",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2A30", lt1: "FFFFFF", dk2: "0B3C49", lt2: "EEF3F4",
    accent1: "C8553D", accent2: "3A86A8", accent3: "6A4C93",
    accent4: "2A9D8F", accent5: "E0A100", accent6: "8D6E63",
    hlink: "3A86A8", folHlink: "6A4C93",
  },
};
const K = THEME.colors;
const GOLD = K.accent5, GOLD_TINT = "FFF1C7", MUTED = "5C6B73", CARD = "F3F6F7";

// ---------- Datos ----------
const TYPES = {
  INT: { name: "Gestión integral de riesgos", short: "Gestión integral", color: K.dk2 },
  CRE: { name: "Riesgo de crédito", short: "Crédito", color: K.accent1 },
  LIQ: { name: "Riesgo de liquidez y mercado", short: "Liquidez y mercado", color: K.accent2 },
  TEC: { name: "Seguridad de la información y ciberseguridad", short: "Seguridad de la información", color: K.accent3 },
  OPE: { name: "Riesgo operacional y continuidad", short: "Operacional y continuidad", color: K.accent4 },
  LEG: { name: "Riesgo legal y cumplimiento", short: "Legal y cumplimiento", color: K.accent6 },
  EST: { name: "Riesgo estratégico y gobierno corporativo", short: "Estratégico y gobierno", color: MUTED },
};

// [n, tipo, tema corto, tiene plan de acción]
const SESSIONS = [
  { no: 144, date: "04 feb 2026", items: [
    [1, "TEC", "Ciberseguridad y SGSI", true], [2, "INT", "Gestión integral", true],
    [3, "EST", "Informes semestrales", false], [4, "CRE", "Avalúos dic 2025", false],
    [5, "LIQ", "Manual de liquidez", false], [6, "CRE", "Estrés crediticio", false],
    [7, "LIQ", "Modelos predictivos", false], [8, "EST", "Segmentación", false],
    [9, "CRE", "Récord relacionados", false]] },
  { no: 145, date: "18 mar 2026", items: [
    [10, "LEG", "Litigios Soluc. Fin.", false], [11, "LEG", "Litigios G. Legal", true],
    [12, "INT", "Gestión integral", true], [13, "EST", "Gob. Corporativo 2025", false],
    [14, "LEG", "Ley FATCA", false], [15, "OPE", "Política R. Operac.", false],
    [16, "CRE", "Avalúos ene–feb", false], [17, "EST", "Segmentación", false],
    [18, "CRE", "Récord relacionados", false], [19, "TEC", "Ciberseg. y DBFD", true]] },
  { no: 146, date: "22 abr 2026", items: [
    [20, "INT", "Gestión integral", true], [21, "EST", "Informes anuales", false],
    [22, "CRE", "Avalúos marzo", false], [23, "EST", "Segmentación", false],
    [24, "CRE", "Récord relacionados", false], [25, "TEC", "Remediación ciber", true]] },
  { no: 147, date: "29 may 2026", items: [
    [26, "INT", "Gestión integral", true], [27, "CRE", "Manual de crédito", false],
    [28, "CRE", "Avalúos abril", true], [29, "EST", "Segmentación", false],
    [30, "CRE", "Récord relacionados", true], [31, "EST", "Gerencia de Riesgos", false],
    [32, "TEC", "Ciberseg. y política IA", false]] },
  { no: 148, date: "25 jun 2026", items: [
    [33, "INT", "Gestión integral", true], [34, "OPE", "Continuidad negocio", false],
    [35, "CRE", "Avalúos mayo", false], [36, "EST", "Segmentación", true],
    [37, "CRE", "Récord relacionados", true], [38, "TEC", "Ciberseguridad", false]] },
];
const ALL = SESSIONS.flatMap(s => s.items.map(i => ({ n: i[0], t: i[1], topic: i[2], plan: i[3], s: s.no })));
const TOTAL = ALL.length, PLANS = ALL.filter(a => a.plan).length;
const pad = n => String(n).padStart(2, "0");
if (TOTAL !== 38) throw new Error("conteo " + TOTAL);

// ---------- Iconos ----------
async function icon(Comp, color) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Acuerdos del Comité de Riesgos – Enero a junio 2026";
  pres.author = "Departamento de Administración de Riesgos";
  pres.subject = "DAR-2026-55";
  const C = pres.SchemeColor;

  const footer = { text: "Comité de Riesgos  ·  Informe DAR-2026-55  ·  Enero – junio 2026",
    options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } };

  pres.defineSlideMaster({
    title: "PORTADA", background: { color: K.dk2 },
    objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.35, w: 7.4, h: 1.7,
      fontFace: THEME.headFontFace, fontSize: 44, bold: true, color: "FFFFFF", valign: "top", align: "left", margin: 0 },
      text: "" } }],
  });
  pres.defineSlideMaster({
    title: "CONTENIDO", background: { color: "FFFFFF" },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.75,
        fontFace: THEME.headFontFace, fontSize: 32, bold: true, color: K.dk2, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: footer },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
  });

  const sub = (slide, text) => slide.addText(text, { x: 0.6, y: 1.1, w: 12.1, h: 0.4, fontSize: 15,
    color: MUTED, margin: 0, isTextBox: true, objectName: "Subtitulo" });

  // ======== 1. Portada + resumen ========
  pres.addSection({ title: "Resumen" });
  let s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Resumen" });
  s.addText("INFORME DAR-2026-55  ·  JUNTA DE VIGILANCIA", { x: 0.7, y: 0.75, w: 7.4, h: 0.4, fontSize: 13,
    bold: true, color: GOLD, charSpacing: 2, margin: 0, isTextBox: true, objectName: "Antetitulo" });
  s.addText("Acuerdos del Comité de Riesgos", { placeholder: "title" });
  s.addText("Primer semestre 2026  ·  Sesiones No. 144 a 148", { x: 0.7, y: 3.15, w: 7.4, h: 0.45, fontSize: 20,
    color: "CFE3E8", margin: 0, isTextBox: true, objectName: "Periodo" });
  s.addText([
    { text: "Resumen de los asuntos conocidos y decisiones adoptadas por el Comité, clasificados por tipo de riesgo, con los acuerdos que generan planes de acción resaltados.", options: { breakLine: true } },
    { text: "Base normativa: NRP-17, art. 20, inciso tercero.", options: { italic: true, color: "9FC3CC" } },
  ], { x: 0.7, y: 4.05, w: 6.6, h: 1.3, fontSize: 15, color: "FFFFFF", margin: 0, paraSpaceAfter: 8,
    valign: "top", isTextBox: true, objectName: "Proposito" });
  s.addText("Lic. Henrry Daniel Escamilla  ·  Jefe de Administración de Riesgos", { x: 0.7, y: 6.6, w: 7.4, h: 0.35,
    fontSize: 12, color: "9FC3CC", margin: 0, isTextBox: true, objectName: "Firma" });

  const kpis = [
    { v: "5", l: "sesiones ordinarias", ic: fa.FaCalendarAlt, hi: false },
    { v: String(TOTAL), l: "acuerdos adoptados", ic: fa.FaFileSignature, hi: false },
    { v: String(Object.keys(TYPES).length), l: "tipos de riesgo", ic: fa.FaLayerGroup, hi: false },
    { v: String(PLANS), l: "acuerdos con plan de acción", ic: fa.FaTasks, hi: true },
  ];
  for (let i = 0; i < kpis.length; i++) {
    const k = kpis[i], col = i % 2, row = Math.floor(i / 2);
    const x = 8.35 + col * 2.3, y = 1.35 + row * 2.6;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.05, h: 2.35, rectRadius: 0.12,
      fill: { color: k.hi ? GOLD : "134B5A" }, line: { type: "none" }, objectName: "KPI " + (i + 1) });
    s.addImage({ data: await icon(k.ic, k.hi ? K.dk2 : GOLD), x: x + 0.25, y: y + 0.25, w: 0.42, h: 0.42,
      altText: k.l, objectName: "Icono KPI " + (i + 1) });
    s.addText(k.v, { x: x + 0.25, y: y + 0.75, w: 1.6, h: 0.85, fontFace: THEME.headFontFace, fontSize: 54, bold: true,
      color: k.hi ? K.dk2 : "FFFFFF", margin: 0, valign: "middle", isTextBox: true, objectName: "Valor KPI " + (i + 1) });
    s.addText(k.l, { x: x + 0.25, y: y + 1.62, w: 1.65, h: 0.6, fontSize: 13, bold: k.hi,
      color: k.hi ? K.dk2 : "CFE3E8", margin: 0, valign: "top", isTextBox: true, objectName: "Etiqueta KPI " + (i + 1) });
  }
  s.addNotes(`El Comité de Riesgos celebró 5 sesiones (144 a 148) entre febrero y junio de 2026 y adoptó ${TOTAL} acuerdos (CR-2026-01 a CR-2026-38). ${PLANS} de ellos instruyen planes de acción, solicitudes o recomendaciones de seguimiento a áreas específicas.`);

  // ======== 2. Acuerdos por sesión ========
  pres.addSection({ title: "Acuerdos por sesión" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Acuerdos por sesión" });
  s.addText(`${TOTAL} acuerdos en 5 sesiones del Comité`, { placeholder: "title" });
  sub(s, "Número de acuerdo CR-2026-## por sesión, coloreado por tipo de riesgo; en dorado, los que tienen plan de acción");
  const colW = 2.26, gap = 0.2, x0 = 0.6, rowH = 0.36, rowGap = 0.05;
  SESSIONS.forEach((ss, ci) => {
    const x = x0 + ci * (colW + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.62, w: colW, h: 0.74, rectRadius: 0.08,
      fill: { color: K.dk2 }, line: { type: "none" }, objectName: `Sesion ${ss.no}` });
    s.addText([
      { text: `Sesión ${ss.no}`, options: { bold: true, fontSize: 15, color: "FFFFFF", breakLine: true } },
      { text: ss.date, options: { fontSize: 12, color: "CFE3E8" } },
    ], { x: x + 0.15, y: 1.62, w: 1.45, h: 0.74, margin: 0, valign: "middle", isTextBox: true, objectName: `Sesion ${ss.no} texto` });
    s.addText(String(ss.items.length), { x: x + 1.55, y: 1.62, w: 0.58, h: 0.74, fontFace: THEME.headFontFace,
      fontSize: 30, bold: true, color: GOLD, align: "right", valign: "middle", margin: 0, isTextBox: true,
      objectName: `Sesion ${ss.no} total` });
    ss.items.forEach(([n, t, topic, plan], ri) => {
      const y = 2.48 + ri * (rowH + rowGap);
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: colW, h: rowH, rectRadius: 0.05,
        fill: { color: plan ? GOLD_TINT : CARD }, line: plan ? { color: GOLD, width: 1.75 } : { type: "none" },
        objectName: `CR-${pad(n)} fondo` });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.05, y: y + 0.05, w: 0.66, h: rowH - 0.1, rectRadius: 0.04,
        fill: { color: TYPES[t].color }, line: { type: "none" }, objectName: `CR-${pad(n)} tipo` });
      s.addText(`CR-${pad(n)}`, { x: x + 0.05, y: y + 0.05, w: 0.66, h: rowH - 0.1, fontSize: 11, bold: true,
        color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: `CR-${pad(n)} numero` });
      s.addText(topic, { x: x + 0.8, y, w: colW - 0.85, h: rowH, fontSize: 11, bold: plan, color: K.dk1,
        valign: "middle", margin: 0, isTextBox: true, objectName: `CR-${pad(n)} tema` });
    });
  });
  // leyenda
  const leg = Object.values(TYPES);
  let lx = 0.6;
  const legW = [1.4, 0.95, 1.55, 1.9, 1.75, 1.5, 1.6];
  leg.forEach((t, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: 6.67, w: 0.18, h: 0.18, rectRadius: 0.03,
      fill: { color: t.color }, line: { type: "none" }, objectName: "Leyenda " + t.short });
    s.addText(t.short, { x: lx + 0.24, y: 6.6, w: legW[i] - 0.24, h: 0.32, fontSize: 10, color: MUTED, margin: 0,
      valign: "middle", isTextBox: true, objectName: "Leyenda texto " + t.short });
    lx += legW[i];
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: 6.67, w: 0.18, h: 0.18, rectRadius: 0.03,
    fill: { color: GOLD_TINT }, line: { color: GOLD, width: 1.5 }, objectName: "Leyenda plan" });
  s.addText("Con plan de acción", { x: lx + 0.24, y: 6.6, w: 1.5, h: 0.32, fontSize: 10, bold: true, color: K.dk1,
    margin: 0, valign: "middle", isTextBox: true, objectName: "Leyenda plan texto" });
  s.addNotes("Cada sesión ordinaria incluyó el Informe de Gestión Integral de Riesgos, la gestión de avalúos, la segmentación de asociados/clientes, el monitoreo del récord crediticio de personas relacionadas y los informes de Seguridad de la Información. La sesión 145 fue la más extensa con 10 acuerdos.");

  // ======== 3. Clasificación por tipo de riesgo ========
  pres.addSection({ title: "Tipos de riesgo" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Tipos de riesgo" });
  s.addText("Crédito concentra un tercio de los acuerdos", { placeholder: "title" });
  sub(s, "Clasificación de los 38 acuerdos por tipo de riesgo; resaltados en dorado los que tienen plan de acción");
  const order = ["CRE", "EST", "INT", "TEC", "LEG", "LIQ", "OPE"];
  const counts = order.map(k => ALL.filter(a => a.t === k).length);
  s.addChart(pres.charts.BAR, [{ name: "Acuerdos", labels: order.map(k => TYPES[k].short), values: counts }], {
    x: 0.5, y: 1.65, w: 5.6, h: 5.1, barDir: "bar", catAxisOrientation: "maxMin",
    chartColors: order.map(k => TYPES[k].color), barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 13, dataLabelFontBold: true,
    dataLabelColor: K.dk1, dataLabelFontFace: "+mn-lt",
    catAxisLabelFontSize: 12, catAxisLabelColor: K.dk1, catAxisLabelFontFace: "+mn-lt",
    catAxisLineShow: false, valAxisHidden: true, valGridLine: { style: "none" }, catGridLine: { style: "none" },
    showLegend: false, showTitle: true, title: "Número de acuerdos por tipo de riesgo", titleFontSize: 13,
    titleColor: MUTED, titleFontFace: "+mn-lt", objectName: "Grafico tipos de riesgo",
  });
  // listado
  order.forEach((k, i) => {
    const y = 1.7 + i * 0.73, items = ALL.filter(a => a.t === k);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.45, y, w: 6.25, h: 0.62, rectRadius: 0.06,
      fill: { color: CARD }, line: { type: "none" }, objectName: "Tipo " + TYPES[k].short });
    s.addShape(pres.shapes.OVAL, { x: 6.6, y: y + 0.21, w: 0.2, h: 0.2, fill: { color: TYPES[k].color },
      line: { type: "none" }, objectName: "Marca " + TYPES[k].short });
    s.addText(TYPES[k].short, { x: 6.92, y, w: 2.35, h: 0.62, fontSize: 13, bold: true, color: K.dk1, margin: 0,
      valign: "middle", isTextBox: true, objectName: "Nombre " + TYPES[k].short });
    const runs = [];
    items.forEach((a, j) => {
      runs.push({ text: pad(a.n), options: a.plan ? { bold: true, highlight: "FFD966", color: K.dk1 } : { color: K.dk1 } });
      if (j < items.length - 1) runs.push({ text: "  ", options: {} });
    });
    s.addText(runs, { x: 9.3, y, w: 3.3, h: 0.62, fontSize: 13, margin: 0, valign: "middle",
      isTextBox: true, objectName: "Acuerdos " + TYPES[k].short });
  });
  s.addNotes("Crédito (12 acuerdos) incluye avalúos, récord crediticio de personas relacionadas, prueba de estrés y el Manual de Riesgo de Crédito. Estratégico y gobierno (9) incluye segmentación de asociados, informes al Consejo, Gobierno Corporativo y la propuesta de Gerencia de Riesgos. Gestión integral (5) agrupa los informes mensuales que cubren crédito, liquidez, mercado, operacional, reputacional y legal.");

  // ======== 4. Gestión integral: semáforo ========
  pres.addSection({ title: "Gestión integral" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Gestión integral" });
  s.addText("El riesgo reputacional sube a Medio desde mayo", { placeholder: "title" });
  sub(s, "Nivel de riesgo aprobado en cada Informe de Gestión Integral de Riesgos");
  const levels = [
    ["Crédito", "M", "M", "M", "M", "M"], ["Liquidez", "M", "M", "M", "M", "M"],
    ["Mercado", "B", "B", "B", "B", "B"], ["Operacional", "M", "M", "M", "M", "M"],
    ["Reputacional", "B", "B", "B", "M", "M"], ["Legal", "M", "M", "M", "M", "M"],
  ];
  const LV = { M: { t: "Medio", fill: "F6C55B", c: K.dk1 }, B: { t: "Bajo", fill: "9BD3C4", c: K.dk1 } };
  const hdr = (t, sub2) => ({ text: [{ text: t, options: { bold: true, breakLine: true } }, { text: sub2, options: { fontSize: 10, color: "CFE3E8" } }],
    options: { fill: { color: K.dk2 }, color: "FFFFFF", align: "center", valign: "middle", fontSize: 12 } });
  const rows = [[{ text: "Tipo de riesgo", options: { fill: { color: K.dk2 }, color: "FFFFFF", bold: true, fontSize: 12, valign: "middle" } },
    hdr("S-144", "CR-02"), hdr("S-145", "CR-12"), hdr("S-146", "CR-20"), hdr("S-147", "CR-26"), hdr("S-148", "CR-33")]];
  levels.forEach(r => rows.push([
    { text: r[0], options: { bold: true, fontSize: 13, color: K.dk1, fill: { color: CARD }, valign: "middle" } },
    ...r.slice(1).map((v, i) => ({ text: LV[v].t, options: { fill: { color: (r[0] === "Reputacional" && i >= 3) ? "E8833A" : LV[v].fill },
      color: (r[0] === "Reputacional" && i >= 3) ? "FFFFFF" : LV[v].c, bold: true, align: "center", valign: "middle", fontSize: 13 } })),
  ]));
  s.addTable(rows, { x: 0.6, y: 1.7, w: 7.4, colW: [1.8, 1.12, 1.12, 1.12, 1.12, 1.12], rowH: [0.62, 0.6, 0.6, 0.6, 0.6, 0.6, 0.6],
    border: { type: "solid", pt: 2, color: "FFFFFF" }, fontFace: THEME.bodyFontFace, margin: 0.06, objectName: "Semaforo de riesgos" });
  s.addText([
    { text: "Recomendación recurrente: ", options: { bold: true } },
    { text: "estrategia integral de revisión de cartera, diversificación de depósitos y seguimiento a planes de acción derivados de quejas." },
  ], { x: 0.6, y: 6.0, w: 7.4, h: 0.75, fontSize: 13, color: K.dk1, margin: 10, valign: "middle",
    fill: { color: GOLD_TINT }, line: { color: GOLD, width: 1.25 }, isTextBox: true, objectName: "Recomendacion" });

  // indicadores de crédito
  s.addText("Indicadores de riesgo de crédito", { x: 8.5, y: 1.7, w: 4.2, h: 0.4, fontSize: 15, bold: true,
    color: K.dk2, margin: 0, isTextBox: true, objectName: "Indicadores titulo" });
  const stats = [
    { v: "17 → 19", l: "créditos/mes migrados a categoría E (S-144 a S-147)" },
    { v: "$10.78M → $8.59M", l: "cancelaciones anticipadas en 6 meses" },
    { v: "1.53% → 1.48%", l: "crecimiento interanual de cartera; tendencia decreciente en S-148" },
  ];
  stats.forEach((st, i) => {
    const y = 2.2 + i * 1.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.5, y, w: 4.2, h: 1.32, rectRadius: 0.08,
      fill: { color: CARD }, line: { type: "none" }, objectName: "Indicador " + (i + 1) });
    s.addText(st.v, { x: 8.7, y: y + 0.12, w: 3.9, h: 0.6, fontFace: THEME.headFontFace, fontSize: 26, bold: true,
      color: K.accent1, margin: 0, valign: "middle", isTextBox: true, objectName: "Indicador valor " + (i + 1) });
    s.addText(st.l, { x: 8.7, y: y + 0.72, w: 3.9, h: 0.5, fontSize: 12, color: MUTED, margin: 0, valign: "top",
      isTextBox: true, objectName: "Indicador etiqueta " + (i + 1) });
  });
  s.addNotes("Los cinco informes de gestión integral mantuvieron crédito, liquidez, operacional y legal en riesgo Medio y mercado en Bajo. El riesgo reputacional pasó de Bajo a Medio en las sesiones 147 y 148. En crédito persisten traslados de deuda a otros bancos, migración de créditos a categoría E y crecimiento de cartera vencida y reservas. El indicador de cartera vencida se reportó en 1.34% en las sesiones 146 y 147.");

  // ======== 5. Planes de acción ========
  pres.addSection({ title: "Planes de acción" });
  s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Planes de acción" });
  s.addText(`${PLANS} acuerdos con planes de acción en curso`, { placeholder: "title" });
  sub(s, "Instrucciones, solicitudes y recomendaciones de seguimiento derivadas de los acuerdos");
  const plans = [
    ["CR-01", "144", "TEC", "Actualizar todos los planes de acción por hallazgos de análisis de vulnerabilidad", "Seguridad de la Información / Gerencia de Sistemas"],
    ["CR-02 · 12 · 20 · 26", "144–147", "INT", "Revisar cartera (traslados, cancelaciones, líneas con más reservas); diversificar depósitos; verificar beneficiarios de depositantes > 65 años", "Administración de Riesgos y áreas de negocio"],
    ["CR-11", "145", "LEG", "Seguimiento oportuno de litigios y procesos sancionatorios hasta su resolución", "Gerencia Legal y de Cumplimiento Normativo"],
    ["CR-19", "145", "TEC", "Atender recomendaciones de alertas de fraude digital (DBFD)", "Seguridad de la Información"],
    ["CR-25", "146", "TEC", "Priorizar remediación por nivel de riesgo y asignar recursos a los planes de acción", "Gerencia de Sistemas"],
    ["CR-28", "147", "CRE", "Presentar al Consejo propuesta de ajuste del pago por informe de avalúo", "Administración de Riesgos"],
    ["CR-30", "147", "CRE", "Monitorear y comunicar a empleados con categoría de riesgo distinta de A1", "Administración de Riesgos"],
    ["CR-33", "148", "INT", "Analizar causas del traslado de créditos; plan de comunicación y educación financiera", "Gerencia de Negocios, Mercadeo, Comité de Educación"],
    ["CR-36", "148", "EST", "Compartir segmentación con Mercadeo; punto de equilibrio de agencias Usulután y Sonsonate", "Mercadeo / Gerencia de Finanzas"],
    ["CR-37", "148", "CRE", "Programas de educación financiera para empleados", "Educación y Responsabilidad Social"],
  ];
  const th = t => ({ text: t, options: { bold: true, color: "FFFFFF", fill: { color: K.dk2 }, fontSize: 12, valign: "middle" } });
  const prow = [[th("Acuerdo"), th("Sesión"), th("Tipo de riesgo"), th("Plan de acción / instrucción"), th("Responsable")]];
  plans.forEach((p, i) => {
    const bg = i % 2 ? "FFFFFF" : GOLD_TINT;
    prow.push([
      { text: p[0], options: { bold: true, color: K.dk2, fill: { color: bg } } },
      { text: p[1], options: { color: K.dk1, fill: { color: bg }, align: "center" } },
      { text: TYPES[p[2]].short, options: { bold: true, color: TYPES[p[2]].color === MUTED ? "4A575E" : TYPES[p[2]].color, fill: { color: bg } } },
      { text: p[3], options: { color: K.dk1, fill: { color: bg } } },
      { text: p[4], options: { color: MUTED, fill: { color: bg } } },
    ]);
  });
  s.addTable(prow, { x: 0.6, y: 1.6, w: 12.1, colW: [1.55, 0.85, 1.85, 5.0, 2.85], fontSize: 11,
    fontFace: THEME.bodyFontFace, valign: "middle", margin: [0.04, 0.08, 0.04, 0.08],
    rowH: [0.38, 0.42, 0.62, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42, 0.42],
    border: { type: "solid", pt: 0.75, color: "E3D7A8" }, objectName: "Tabla planes de accion" });
  s.addNotes("Los acuerdos CR-02, 12, 20 y 26 repiten la misma recomendación de gestión integral en cada sesión. CR-33 la concreta con instrucciones a la Gerencia de Negocios, Mercadeo y Comité de Educación, además de seguimiento a planes de acción operacionales, reputacionales y legales. CR-25 es el único que compromete recursos humanos, tecnológicos y presupuestarios para la remediación de hallazgos.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("ok", OUT, TOTAL, PLANS);
})().catch(e => { console.error(e); process.exit(1); });
