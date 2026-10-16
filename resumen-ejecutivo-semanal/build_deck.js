// Genera Resumen_Ejecutivo_Riesgos_COMEDICA_2_Slides.pptx (2 diapositivas, editable).
// Uso: node build_deck.js [salida.pptx]
// Requiere: pptxgenjs, react, react-dom, react-icons, sharp
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa6");

const OUT = process.argv[2] || "Resumen_Ejecutivo_Riesgos_COMEDICA_2_Slides.pptx";

// Paleta: azul marino, azul institucional, turquesa, blanco; ámbar solo para plazos.
const NAVY = "0B2545";
const BLUE = "13507F";
const TEAL = "0E9AA7";
const TEAL_LT = "D7F0F2";
const TEAL_DK = "0A7480";
const SLATE = "2E4A73";
const CARD = "F1F5F9";
const INK = "1B2733";
const AMBER = "B86E00";
const AMBER_LT = "FDF3E1";
const WHITE = "FFFFFF";
const FONT = "Calibri";

async function icon(Comp, color, px = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: px })
  );
  const buf = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.5, y: 0.32, w: 8.8, h: 0.6, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 26, bold: true, color: NAVY, valign: "middle",
  });
  slide.addText(subtitle, {
    x: 0.5, y: 0.92, w: 9.6, h: 0.32, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 13, color: BLUE, valign: "middle",
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 7.13, w: 12.33, h: 0.24, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 9, color: "4A5A6A", valign: "middle",
  });
}

// Ítem de una línea: viñeta cuadrada + etiqueta en negrita + texto.
function addItem(slide, x, y, w, label, text, color) {
  slide.addShape("rect", { x, y: y + 0.105, w: 0.08, h: 0.08, fill: { color }, line: { color } });
  const runs = label ? [{ text: label + " ", options: { bold: true, color: NAVY } }] : [];
  runs.push({ text, options: { color: INK } });
  slide.addText(runs, { x: x + 0.2, y, w: w - 0.2, h: 0.29, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 13, valign: "middle" });
}

async function card(slide, { x, y, w, h, ic, color, title, items }) {
  const HH = 0.62;
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.12, fill: { color: CARD }, line: { color: CARD } });
  slide.addShape("roundRect", { x, y, w, h: HH + 0.12, rectRadius: 0.12, fill: { color }, line: { color } });
  slide.addShape("rect", { x, y: y + HH, w, h: 0.2, fill: { color: CARD }, line: { color: CARD } });
  slide.addShape("ellipse", { x: x + 0.2, y: y + 0.1, w: 0.42, h: 0.42, fill: { color: WHITE }, line: { color: WHITE } });
  slide.addImage({ data: await icon(ic, color), x: x + 0.3, y: y + 0.2, w: 0.22, h: 0.22 });
  slide.addText(title, {
    x: x + 0.78, y: y + 0.1, w: w - 0.95, h: 0.42, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 16, bold: true, color: WHITE, valign: "middle",
  });
  const step = 0.34;
  items.forEach(([label, text], i) => addItem(slide, x + 0.28, y + HH + 0.22 + i * step, w - 0.5, label, text, color));
}

// Indicador único: avance del examen GIR.
function examBadge(slide, x, y, w, pct) {
  const h = 0.78;
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.12, fill: { color: TEAL_LT }, line: { color: TEAL_LT } });
  slide.addText([
    { text: "Examen GIR 2026", options: { bold: true, fontSize: 13, color: NAVY, breakLine: true } },
    { text: "Avance reportado", options: { fontSize: 10.5, color: BLUE } },
  ], { x: x + 0.22, y: y + 0.08, w: 1.7, h: 0.48, margin: 0, isTextBox: true, fontFace: FONT, valign: "middle" });
  const bx = x + 0.22, bw = w - 1.5;
  slide.addShape("roundRect", { x: bx, y: y + 0.6, w: bw, h: 0.09, rectRadius: 0.045, fill: { color: "B9DDE1" }, line: { color: "B9DDE1" } });
  slide.addShape("roundRect", { x: bx, y: y + 0.6, w: bw * pct / 100, h: 0.09, rectRadius: 0.045, fill: { color: TEAL_DK }, line: { color: TEAL_DK } });
  slide.addText(pct + "%", {
    x: x + w - 1.2, y, w: 1.0, h, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 30, bold: true, color: TEAL_DK, align: "right", valign: "middle",
  });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in, 16:9
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "Resumen Ejecutivo de Riesgos – COMÉDICA";
  pres.author = "Departamento de Administración de Riesgos";

  // ---------------- DIAPOSITIVA 1 ----------------
  // Matriz por frente (filas) x [vigente del informe anterior | nuevo de la semana].
  const s1 = pres.addSlide();
  s1.background = { color: WHITE };
  addTitle(s1, "Departamento de Administración de Riesgos", "Informe semanal · Resumen ejecutivo");

  const LX = 0.5, LW = 2.95, C1 = 3.6, CW = 4.54, C2 = C1 + CW + 0.15;
  const colHead = (x, t) => {
    s1.addShape("roundRect", { x, y: 1.32, w: CW, h: 0.34, rectRadius: 0.17, fill: { color: NAVY }, line: { color: NAVY } });
    s1.addText(t, { x, y: 1.32, w: CW, h: 0.34, margin: 0, isTextBox: true, align: "center", valign: "middle", fontFace: FONT, fontSize: 12, bold: true, color: WHITE });
  };
  colHead(C1, "En seguimiento");
  colHead(C2, "Actividades de la semana");

  const bullets = (items) => items.map((t, i) => ({
    text: t, options: { bullet: { indent: 12 }, color: INK, fontSize: 11, paraSpaceAfter: 2, breakLine: i < items.length - 1 },
  }));
  const cell = (x, y, h, content) => {
    s1.addShape("roundRect", { x, y, w: CW, h, rectRadius: 0.1, fill: { color: "EEF2F7" }, line: { color: "EEF2F7" } });
    s1.addText(content, { x: x + 0.18, y: y + 0.08, w: CW - 0.32, h: h - 0.16, margin: 0, isTextBox: true, fontFace: FONT, valign: "middle" });
  };
  const label = async (y, h, ic, title, sub, pct) => {
    s1.addShape("roundRect", { x: LX, y, w: LW, h, rectRadius: 0.1, fill: { color: NAVY }, line: { color: NAVY } });
    const cy = pct ? y + h / 2 - 0.42 : y + h / 2 - 0.2;
    s1.addShape("ellipse", { x: LX + 0.18, y: cy, w: 0.42, h: 0.42, fill: { color: "2B6CB0" }, line: { color: "2B6CB0" } });
    s1.addImage({ data: await icon(ic, WHITE), x: LX + 0.29, y: cy + 0.11, w: 0.2, h: 0.2 });
    s1.addText([
      { text: title, options: { bold: true, fontSize: 13.5, color: WHITE, breakLine: true } },
      { text: sub, options: { fontSize: 10.5, color: "C9D6E6" } },
    ], { x: LX + 0.75, y: cy - 0.1, w: LW - 0.85, h: 0.62, margin: 0, isTextBox: true, fontFace: FONT, valign: "middle" });
    if (pct) {
      const by = y + h / 2 + 0.3, bw = LW - 0.95;
      s1.addText("Examen GIR 2026", { x: LX + 0.18, y: by - 0.24, w: 1.6, h: 0.2, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 9.5, bold: true, color: "C9D6E6" });
      s1.addShape("roundRect", { x: LX + 0.18, y: by, w: bw, h: 0.11, rectRadius: 0.05, fill: { color: "3C5A80" }, line: { color: "3C5A80" } });
      s1.addShape("roundRect", { x: LX + 0.18, y: by, w: bw * pct / 100, h: 0.11, rectRadius: 0.05, fill: { color: "F2B544" }, line: { color: "F2B544" } });
      s1.addText(pct + "%", { x: LX + LW - 0.72, y: by - 0.12, w: 0.6, h: 0.34, margin: 0, isTextBox: true, align: "right", valign: "middle", fontFace: FONT, fontSize: 14, bold: true, color: WHITE });
    }
  };

  // Fila 1: requerimientos SSF (plazos vigentes, presentación base)
  const reg = (t, c, d, last) => [
    { text: t + "  ·  ", options: { bold: true, color: NAVY, fontSize: 11.5 } },
    { text: c, options: { color: BLUE, fontSize: 11 } },
    { text: "   " + d, options: { bold: true, color: AMBER, fontSize: 11.5, paraSpaceAfter: 3, breakLine: !last } },
  ];
  const rows = [
    { h: 1.15, ic: fa.FaBuildingColumns, t: "Cumplimiento Regulatorio", s: "Requerimientos SSF",
      a: [
        ...reg("PLD/FT", "Circ. 23486", "Plazo 16/10/2026"),
        ...reg("Gobierno Corporativo", "Circ. 22741", "Plazos 23/10 y 30/10"),
        ...reg("Plan de Recuperación", "Circ. 22737", "Plazo 03/11/2026"),
        ...reg("Riesgos Específicos", "Circ. 22742", "Inicio visita 09/10", true),
      ],
      b: bullets(["Matriz y requerimiento a áreas – nota 22741", "Pruebas de descargo – notas 22741 y 23486", "Presentación visita de supervisión Gob. Corporativo", "Resultados preliminares de Riesgo Específico"]) },
    { h: 1.28, ic: fa.FaShieldHalved, t: "Riesgo Operacional", s: "Continuidad y control", pct: 90,
      a: bullets(["Actualización Matriz BIA", "Indicadores de Gob. Corporativo y Reputacional", "Segunda prueba de continuidad de negocio"]),
      b: bullets(["Capacitación GIR 2026 e informe de seguimiento del examen", "Informe 2.ª prueba de continuidad (01/10/2026)", "Planes de acción Matriz 2026 y eventos TC / Orion", "Opiniones: Fábrica de Créditos y caso de asociado", "Información a Auditoría Interna (NRP-24) y Externa"]) },
    { h: 1.28, ic: fa.FaScaleBalanced, t: "Riesgo Normativo", s: "Avalúos y políticas",
      a: bullets(["Revisión de Política y Procedimientos de Avalúos", "Revisión de Perfil de Revisor", "Envíos regulatorios (DAR)"]),
      b: bullets(["Avalúos: perito, gestión electrónica y aviso a asociados", "Descargo de avalúo anterior y mutuo hipotecario", "Comparativos de políticas y acta de sesión 29/08/2026", "Acuerdos Semestrales del Comité de Riesgos", "Proyectos: formularios, BIA, Procesos y página web"]) },
    { h: 1.28, ic: fa.FaChartLine, t: "Riesgo Financiero", s: "Crédito, liquidez y TC",
      a: bullets(["Informe de Análisis de Cartera de TC", "Simulación de cartera vencida", "Actualización de Política de TC"]),
      b: bullets(["Informes al Comité: crédito, liquidez y mercado", "Envíos SSF: NRP-05, NCB-022, NRP-41 y CTRI-RECR", "Cartera TC y reserva de saneamiento", "Traslados de deuda y segmentación de productos", "Sistema de créditos NRP-41 y TC de alto riesgo"]) },
  ];
  let rowY = 1.78;
  for (const r of rows) {
    await label(rowY, r.h, r.ic, r.t, r.s, r.pct);
    cell(C1, rowY, r.h, r.a);
    cell(C2, rowY, r.h, r.b);
    rowY += r.h + 0.09;
  }

  addFooter(s1, "«En seguimiento»: asuntos del informe anterior que siguen vigentes. «Actividades de la semana»: documento de actividades semanales.");
  s1.addNotes(
    "Matriz por frente. Columna izquierda: asuntos del informe anterior que siguen vigentes (plazos SSF). " +
    "Columna derecha: actividades de la semana según el documento de actividades. Examen GIR 2026: avance 90%."
  );

  // ---------------- DIAPOSITIVA 2 ----------------
  const s2 = pres.addSlide();
  s2.background = { color: WHITE };
  addTitle(s2, "Compromisos Regulatorios y Próximos Hitos",
    "Requerimientos de la Superintendencia del Sistema Financiero (SSF) · Seguimiento semanal");

  // Leyenda
  const lg = [["Recepción", "rec"], ["Inicio de visita", "vis"], ["Plazo", "plazo"]];
  let lx = 9.95;
  for (const [t, k] of lg) {
    marker(s2, k, lx + 0.09, 0.62, 0.18);
    s2.addText(t, { x: lx + 0.24, y: 0.5, w: 1.0, h: 0.24, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 10, color: INK, valign: "middle" });
    lx += t.length > 10 ? 1.35 : 1.0;
  }

  function marker(slide, kind, cx, cy, d) {
    const o = { x: cx - d / 2, y: cy - d / 2, w: d, h: d };
    if (kind === "rec") slide.addShape("ellipse", { ...o, fill: { color: WHITE }, line: { color: NAVY, width: 2 } });
    else if (kind === "vis") slide.addShape("ellipse", { ...o, fill: { color: TEAL }, line: { color: TEAL, width: 1 } });
    else slide.addShape("diamond", { ...o, fill: { color: AMBER }, line: { color: AMBER, width: 1 } });
  }

  const LINE_Y = 2.74;
  const events = [
    { d: "22/09", k: "rec", t: "Recepción resultado preliminar Plan de Recuperación", c: "Circ. 22737" },
    { d: "24/09", k: "rec", t: "Recepción visita inicial Riesgos Específicos", c: "Circ. 22742" },
    { d: "02/10", k: "rec", t: "Recepción resultados preliminares PLD/FT", c: "Circ. 23486" },
    { d: "09/10", k: "vis", t: "Inicio visita Riesgos Específicos y recepción resultados Gob. Corp.", c: "Circ. 22742 · 22741" },
    { d: "16/10", k: "plazo", t: "Respuesta a resultados preliminares PLD/FT", c: "Circ. 23486" },
    { d: "23/10", k: "plazo", t: "Envío de nuevas observaciones Gob. Corporativo", c: "Circ. 22741" },
    { d: "30/10", k: "plazo", t: "Seguimiento de observaciones Gob. Corporativo", c: "Circ. 22741" },
    { d: "03/11", k: "plazo", t: "Envío del Plan de Recuperación", c: "Circ. 22737" },
  ];
  const x0 = 1.25, x1 = 12.08, step = (x1 - x0) / (events.length - 1);

  // Zona de plazos (fondo ámbar suave)
  const zx = x0 + 4 * step - 0.8;
  s2.addShape("roundRect", { x: zx, y: 1.38, w: 12.83 - zx, h: 2.66, rectRadius: 0.12, fill: { color: AMBER_LT }, line: { color: AMBER_LT } });
  s2.addText("FECHAS DE RESPUESTA SSF", {
    x: zx + 0.15, y: 1.44, w: 3, h: 0.24, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 9, bold: true, color: AMBER, charSpacing: 1,
  });

  s2.addShape("line", { x: 0.6, y: LINE_Y, w: 12.13, h: 0, line: { color: "AFC0D0", width: 2 } });

  events.forEach((e, i) => {
    const cx = x0 + i * step;
    const above = i % 2 === 0;
    const isP = e.k === "plazo";
    const bw = 1.62;
    s2.addShape("line", {
      x: cx, y: above ? LINE_Y - 0.28 : LINE_Y, w: 0, h: 0.28,
      line: { color: isP ? AMBER : "8FA3B8", width: 1 },
    });
    marker(s2, e.k, cx, LINE_Y, isP ? 0.26 : 0.22);
    const runs = [
      { text: e.d, options: { bold: true, fontSize: 14, color: isP ? AMBER : NAVY, breakLine: true } },
      { text: e.t, options: { fontSize: 10, color: INK, breakLine: true } },
      { text: e.c, options: { fontSize: 9.5, color: BLUE, bold: true } },
    ];
    s2.addText(runs, {
      x: cx - bw / 2, y: above ? 1.62 : 3.08, w: bw, h: 0.82, margin: 0, isTextBox: true,
      fontFace: FONT, align: "center", valign: above ? "bottom" : "top", paraSpaceAfter: 1,
    });
  });

  // Tarjetas de plazos
  const dl = [
    { d: "16/10/2026", tag: "Fecha límite", t: "Respuesta resultados preliminares PLD/FT", c: "Circular 23486", n: "Pruebas de descargo en seguimiento" },
    { d: "23/10/2026", tag: "Fecha límite", t: "Nuevas observaciones – Gob. Corporativo", c: "Circular 22741", n: "Requerimiento remitido a las áreas" },
    { d: "30/10/2026", tag: "Seguimiento", t: "Seguimiento observaciones – Gob. Corporativo", c: "Circular 22741", n: "Plan de solución, visitas anteriores" },
    { d: "03/11/2026", tag: "Fecha límite", t: "Envío del Plan de Recuperación", c: "Circular 22737", n: "Indicadores Gob. Corp. y Reputacional" },
  ];
  const cw = 1.94, cg = 0.16, cy = 4.24, ch = 1.98;
  const iconCal = await icon(fa.FaRegCalendarCheck, AMBER);
  dl.forEach((k, i) => {
    const x = 0.5 + i * (cw + cg);
    const first = i === 0;
    s2.addShape("roundRect", { x, y: cy, w: cw, h: ch, rectRadius: 0.1, fill: { color: first ? NAVY : CARD }, line: { color: first ? NAVY : CARD } });
    s2.addText(k.tag.toUpperCase(), {
      x: x + 0.16, y: cy + 0.12, w: cw - 0.3, h: 0.2, margin: 0, isTextBox: true,
      fontFace: FONT, fontSize: 8.5, bold: true, charSpacing: 1, color: first ? "F6C26B" : AMBER,
    });
    s2.addText(k.d, {
      x: x + 0.16, y: cy + 0.32, w: cw - 0.3, h: 0.38, margin: 0, isTextBox: true,
      fontFace: FONT, fontSize: 19, bold: true, color: first ? WHITE : NAVY, valign: "middle",
    });
    s2.addText([
      { text: k.t, options: { fontSize: 10.5, bold: true, color: first ? WHITE : INK, breakLine: true } },
      { text: k.c, options: { fontSize: 9.5, color: first ? "A9D8DE" : BLUE, breakLine: true } },
      { text: k.n, options: { fontSize: 9.5, italic: true, color: first ? "DCE6F0" : "3E4F60" } },
    ], {
      x: x + 0.16, y: cy + 0.76, w: cw - 0.28, h: ch - 0.86, margin: 0, isTextBox: true,
      fontFace: FONT, valign: "top", paraSpaceAfter: 2,
    });
  });

  // Visita Riesgos Específicos
  const ry = 6.36, rw = 4 * cw + 3 * cg;
  s2.addShape("roundRect", { x: 0.5, y: ry, w: rw, h: 0.68, rectRadius: 0.1, fill: { color: TEAL_LT }, line: { color: TEAL_LT } });
  s2.addShape("ellipse", { x: 0.64, y: ry + 0.12, w: 0.44, h: 0.44, fill: { color: TEAL }, line: { color: TEAL } });
  s2.addImage({ data: await icon(fa.FaMagnifyingGlassChart, WHITE), x: 0.75, y: ry + 0.23, w: 0.22, h: 0.22 });
  s2.addText([
    { text: "Visita inicial Riesgos Específicos · Circ. 22742  ", options: { bold: true, color: NAVY, fontSize: 11 } },
    { text: "Crédito, Liquidez y Mercado · recepción 24/09/2026, inicio 09/10/2026. Presentación de resultados preliminares registrada; plazo de respuesta no documentado.", options: { color: INK, fontSize: 10.5 } },
  ], { x: 1.22, y: ry + 0.06, w: rw - 0.85, h: 0.56, margin: 0, isTextBox: true, fontFace: FONT, valign: "middle" });

  // Puntos de control gerencial
  const px = 0.5 + rw + 0.22, pw = 12.83 - px, py = 4.24, ph = 2.8;
  s2.addShape("roundRect", { x: px, y: py, w: pw, h: ph, rectRadius: 0.12, fill: { color: NAVY }, line: { color: NAVY } });
  s2.addText("Puntos de control gerencial", {
    x: px + 0.25, y: py + 0.14, w: pw - 0.4, h: 0.34, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 15, bold: true, color: WHITE, valign: "middle",
  });
  const ctrl = [
    [fa.FaPeopleArrows, "Coordinación entre áreas", "Nota 22741: Gerencia Financiera, Cumplimiento, Talento Humano, Procesos y otras áreas"],
    [fa.FaFileCircleCheck, "Evidencias y pruebas de descargo", "Revisión y remisión para notas 22741 y 23486"],
    [fa.FaListCheck, "Entregables y planes de acción", "Matriz de Riesgos 2026, informes al Comité de Riesgos, Plan de Recuperación"],
    [fa.FaCircleQuestion, "Por confirmar", "Estado de envío de cada requerimiento y responsable asignado"],
  ];
  let iy = py + 0.58;
  for (const [ic, h, t] of ctrl) {
    s2.addShape("ellipse", { x: px + 0.25, y: iy + 0.02, w: 0.36, h: 0.36, fill: { color: TEAL }, line: { color: TEAL } });
    s2.addImage({ data: await icon(ic, WHITE), x: px + 0.34, y: iy + 0.11, w: 0.18, h: 0.18 });
    s2.addText([
      { text: h, options: { bold: true, fontSize: 11, color: WHITE, breakLine: true } },
      { text: t, options: { fontSize: 9.5, color: "D3DEEA" } },
    ], { x: px + 0.74, y: iy - 0.02, w: pw - 0.92, h: 0.58, margin: 0, isTextBox: true, fontFace: FONT, valign: "top" });
    iy += 0.55;
  }

  addFooter(s2, "Fechas según la presentación base (referencia 10/10/2026). Fecha límite, inicio de visita y recepción se distinguen en la leyenda; el estado de cada envío está por confirmar.");
  s2.addNotes(
    "Línea de tiempo de requerimientos SSF. Plazos: 16/10 PLD/FT (Circ. 23486), 23/10 nuevas observaciones y 30/10 seguimiento de Gobierno Corporativo (Circ. 22741), 03/11 Plan de Recuperación (Circ. 22737). " +
    "Visita de Riesgos Específicos (Circ. 22742): inicio 09/10, sin plazo de respuesta documentado. No se afirma si los envíos están cumplidos o pendientes."
  );

  await pres.writeFile({ fileName: OUT });
  console.log("OK", OUT);
})();
