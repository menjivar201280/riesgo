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
const GREEN = "2E7D4F";
const GREEN_LT = "E5F2EA";
const CARD = "F1F5F9";
const INK = "1B2733";
const AMBER = "B86E00";
const PLAZO = "1F6FB2"; // plazos SSF en la diapositiva 2
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
  addTitle(s1, "Departamento de Administración de Riesgos", "Informe semanal");

  const LX = 0.5, LW = 2.95, C1 = 3.6, CW = 4.54, C2 = C1 + CW + 0.15;
  const colHead = (x, t) => {
    s1.addShape("roundRect", { x, y: 1.32, w: CW, h: 0.34, rectRadius: 0.17, fill: { color: NAVY }, line: { color: NAVY } });
    s1.addText(t, { x, y: 1.32, w: CW, h: 0.34, margin: 0, isTextBox: true, align: "center", valign: "middle", fontFace: FONT, fontSize: 12, bold: true, color: WHITE });
  };
  colHead(C1, "En seguimiento");
  colHead(C2, "Actividades de la semana");

  const bullets = (items) => items.map((t, i) => ({
    text: t, options: { bullet: { indent: 12 }, color: INK, fontSize: 10.5, paraSpaceAfter: 2, breakLine: i < items.length - 1 },
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
      { text: title, options: { bold: true, fontSize: 13.5, color: WHITE, breakLine: !!sub } },
      ...(sub ? [{ text: sub, options: { fontSize: 10.5, color: "C9D6E6" } }] : []),
    ], { x: LX + 0.75, y: cy - 0.1, w: LW - 0.85, h: 0.62, margin: 0, isTextBox: true, fontFace: FONT, valign: "middle" });
    if (pct) {
      const by = y + h / 2 + 0.3, bw = LW - 0.95;
      s1.addText("Capacitación y Examen GIR 2026", { x: LX + 0.18, y: by - 0.24, w: 2.3, h: 0.2, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 9.5, bold: true, color: "C9D6E6" });
      s1.addShape("roundRect", { x: LX + 0.18, y: by, w: bw, h: 0.11, rectRadius: 0.05, fill: { color: "3C5A80" }, line: { color: "3C5A80" } });
      s1.addShape("roundRect", { x: LX + 0.18, y: by, w: bw * pct / 100, h: 0.11, rectRadius: 0.05, fill: { color: "5FB3E8" }, line: { color: "5FB3E8" } });
      s1.addText(pct + "%", { x: LX + LW - 0.72, y: by - 0.12, w: 0.6, h: 0.34, margin: 0, isTextBox: true, align: "right", valign: "middle", fontFace: FONT, fontSize: 14, bold: true, color: WHITE });
    }
  };

  // Texto literal: «En seguimiento» = presentación anterior; «Actividades de la semana» = documento de actividades.
  const reg = (t, c, d, last) => [
    { text: t + "  ·  ", options: { bold: true, color: NAVY, fontSize: 10.5 } },
    { text: c, options: { color: BLUE, fontSize: 10.5, breakLine: true } },
    ...[].concat(d).map((x, i, arr) => ({ text: x, options: { bold: true, color: AMBER, fontSize: 10.5, paraSpaceAfter: i === arr.length - 1 ? 3 : 0, breakLine: !(last && i === arr.length - 1) } })),
  ];
  const rows = [
    { h: 1.5, ic: fa.FaBuildingColumns, t: "Cumplimiento Regulatorio", s: "Requerimientos SSF",
      a: [
        ...reg("Resultados preliminares – PLD/FT", "Circ. 23486", "Plazo 16/10/2026"),
        ...reg("Resultados preliminares – Gobierno Corporativo", "Circ. 22741", ["Plazo 23/10/2026 · Nuevas observaciones", "Plazo 30/10/2026 · Seguimiento de observaciones"]),
        ...reg("Resultado preliminar Plan de Recuperación", "Circ. 22737", "Plazo 03/11/2026", true),
      ],
      b: bullets(["Requerimiento según nota DS-SABAO-BCS-BC-22741", "Requerimiento según nota DS-SABAO-BCS-BC-23486", "Presentación resultados preliminares de Gestión de Riesgo Específico", "Presentación visita de supervisión gobierno corporativo"]) },
    { h: 1.26, ic: fa.FaShieldHalved, t: "Riesgo Operacional", s: "", pct: 93,
      a: bullets(["Capacitación y Examen GIR 2026", "Actualización Matriz BIA", "Segunda prueba de continuidad de negocio"]),
      b: bullets(["Capacitación de GIR 2026", "Informe de segunda prueba de Continuidad del Negocio realizada el 01.10.2026", "Seguimiento planes de acción Matriz de Riesgos Institucional 2026", "Seguimiento eventos de riesgos por caídas de TC 25.09.2026 y Orion 26.09.2026"]) },
    { h: 1.3, ic: fa.FaScaleBalanced, t: "Riesgo Normativo", s: "",
      a: bullets(["Revisión de Política de Avalúos", "Revisión de Procedimientos de Administración de Avalúos", "Revisión de Perfil de Revisor"]),
      b: bullets(["Asignación a perito valuador", "Descargo de avalúo anterior y documento mutuo hipotecario", "Elaboración de acta de sesión de fecha 29/08/2026", "Actualización de política de servicios tercerizados", "Actualización de política de gobierno corporativo"]) },
    { h: 0.98, ic: fa.FaChartLine, t: "Riesgo Financiero", s: "",
      a: bullets(["Informe de Análisis de Cartera de TC", "Simulación de cartera vencida", "Actualización de Política de TC"]),
      b: bullets(["Envío de información a la Central de Riesgo (NRP-41)", "Monitoreo semanal de la reserva de saneamiento", "Registro permanente de los traslados de deuda"]) },
  ];
  let rowY = 1.78;
  for (const r of rows) {
    await label(rowY, r.h, r.ic, r.t, r.s, r.pct);
    cell(C1, rowY, r.h, r.a);
    cell(C2, rowY, r.h, r.b);
    rowY += r.h + 0.09;
  }

  s1.addNotes(
    "Matriz por frente. Columna izquierda: asuntos del informe anterior que siguen vigentes (plazos SSF). " +
    "Columna derecha: actividades de la semana según el documento de actividades. Examen GIR 2026: avance 93%."
  );

  // ---------------- DIAPOSITIVA 2 ----------------
  // Fecha de corte: cambiar aquí para recalcular "HOY" y los días restantes.
  const CORTE = new Date(2026, 9, 16);
  const fmt = (d) => String(d.getDate()).padStart(2, "0") + "/" + String(d.getMonth() + 1).padStart(2, "0");
  const dias = (d) => Math.round((d - CORTE) / 86400000);

  const s2 = pres.addSlide();
  s2.background = { color: WHITE };
  addTitle(s2, "Calendario de requerimientos SSF",
    "Compromisos regulatorios y próximos hitos · Corte " + fmt(CORTE) + "/" + CORTE.getFullYear());

  function marker(slide, kind, cx, cy, d, past) {
    const o = { x: cx - d / 2, y: cy - d / 2, w: d, h: d };
    if (kind === "plazo") slide.addShape("diamond", { ...o, fill: { color: PLAZO }, line: { color: PLAZO, width: 1 } });
    else if (kind === "vis") slide.addShape("ellipse", { ...o, fill: { color: TEAL }, line: { color: TEAL, width: 1 } });
    else slide.addShape("ellipse", { ...o, fill: { color: past ? NAVY : WHITE }, line: { color: NAVY, width: 2 } });
  }

  // Leyenda
  const checkIcon = await icon(fa.FaCheck, WHITE);
  let lx = 10.0;
  for (const [t, k] of [["Transcurrido", "done"], ["Plazo SSF", "plazo"]]) {
    if (k === "done") {
      s2.addShape("ellipse", { x: lx, y: 0.53, w: 0.18, h: 0.18, fill: { color: GREEN }, line: { color: GREEN } });
      s2.addImage({ data: checkIcon, x: lx + 0.04, y: 0.57, w: 0.1, h: 0.1 });
    } else marker(s2, k, lx + 0.09, 0.62, 0.18, true);
    s2.addText(t, { x: lx + 0.24, y: 0.5, w: 1.2, h: 0.24, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 10, color: INK, valign: "middle" });
    lx += 1.45;
  }

  const LINE_Y = 2.6;
  const D = (dd, mm) => new Date(2026, mm - 1, dd);
  const events = [
    { d: D(22, 9), k: "rec", t: "Resultado preliminar Plan de Recuperación", c: "Circular 22737" },
    { d: D(24, 9), k: "rec", t: "Visita inicial Riesgos Específicos", c: "Circular 22742" },
    { d: D(2, 10), k: "rec", t: "Resultados preliminares PLD/FT", c: "Circular 23486" },
    { d: D(9, 10), k: "vis", t: "Inicio visita Riesgos Específicos · Resultados Gobierno Corporativo", c: "Circular 22742 · 22741" },
    { d: D(16, 10), k: "plazo", t: "Respuesta resultados preliminares PLD/FT", c: "Circular 23486" },
    { d: D(23, 10), k: "plazo", t: "Nuevas observaciones Gobierno Corporativo", c: "Circular 22741" },
    { d: D(30, 10), k: "plazo", t: "Seguimiento de observaciones Gobierno Corporativo", c: "Circular 22741" },
    { d: D(3, 11), k: "plazo", t: "Envío del Plan de Recuperación", c: "Circular 22737" },
  ];
  // Posición proporcional a la fecha: la línea muestra el avance real del calendario.
  const x0 = 1.4, x1 = 12.0, T0 = events[0].d, TN = events[events.length - 1].d;
  const span = (TN - T0) / 86400000;
  const xAt = (d) => x0 + ((d - T0) / 86400000) / span * (x1 - x0);
  const xs = events.map((e) => xAt(e.d));
  const hoyX = xAt(CORTE);
  const transcurridos = Math.round((CORTE - T0) / 86400000);

  // Línea: tramo transcurrido en azul marino, tramo pendiente en gris
  s2.addShape("roundRect", { x: 0.5, y: 1.38, w: hoyX - 0.86 - 0.5, h: 2.48, rectRadius: 0.12, fill: { color: GREEN_LT }, line: { color: GREEN_LT } });
  s2.addText("TRANSCURRIDO", { x: 0.65, y: 1.43, w: 2, h: 0.22, margin: 0, isTextBox: true, fontFace: FONT, fontSize: 9, bold: true, charSpacing: 1, color: GREEN });
  s2.addShape("roundRect", { x: x0 - 0.15, y: LINE_Y - 0.06, w: x1 - x0 + 0.3, h: 0.12, rectRadius: 0.06, fill: { color: "D5DEE8" }, line: { color: "D5DEE8" } });
  s2.addShape("roundRect", { x: x0 - 0.15, y: LINE_Y - 0.06, w: hoyX - x0 + 0.15, h: 0.12, rectRadius: 0.06, fill: { color: GREEN }, line: { color: GREEN } });

  events.forEach((e, i) => {
    const cx = xs[i], above = i % 2 === 0, isP = e.k === "plazo", past = dias(e.d) < 0;
    s2.addShape("line", { x: cx, y: above ? LINE_Y - 0.26 : LINE_Y, w: 0, h: 0.26, line: { color: past ? GREEN : (isP ? PLAZO : "8FA3B8"), width: 1 } });
    if (past) {
      s2.addShape("ellipse", { x: cx - 0.14, y: LINE_Y - 0.14, w: 0.28, h: 0.28, fill: { color: GREEN }, line: { color: WHITE, width: 1.5 } });
      s2.addImage({ data: checkIcon, x: cx - 0.07, y: LINE_Y - 0.07, w: 0.14, h: 0.14 });
    } else marker(s2, e.k, cx, LINE_Y, isP ? 0.26 : 0.22, past);
    const bw = 1.55;
    s2.addText([
      { text: (isP ? "Plazo " : "") + fmt(e.d), options: { bold: true, fontSize: 13, color: past ? GREEN : (isP ? PLAZO : NAVY), breakLine: true } },
      { text: e.t, options: { fontSize: 10, color: INK, breakLine: true } },
      { text: e.c, options: { fontSize: 9.5, color: BLUE, bold: true } },
    ], {
      x: cx - bw / 2, y: above ? 1.5 : 2.94, w: bw, h: 0.84, margin: 0, isTextBox: true,
      fontFace: FONT, align: "center", valign: above ? "bottom" : "top",
    });
  });

  // Marca HOY
  s2.addShape("line", { x: hoyX, y: LINE_Y + 0.18, w: 0, h: 0.62, line: { color: PLAZO, width: 1.5, dashType: "dash" } });
  s2.addShape("roundRect", { x: hoyX - 0.55, y: LINE_Y + 0.8, w: 1.1, h: 0.3, rectRadius: 0.15, fill: { color: PLAZO }, line: { color: PLAZO } });
  s2.addText("HOY " + fmt(CORTE), { x: hoyX - 0.55, y: LINE_Y + 0.8, w: 1.1, h: 0.3, margin: 0, isTextBox: true, align: "center", valign: "middle", fontFace: FONT, fontSize: 10.5, bold: true, color: WHITE });

  // Próximos vencimientos
  s2.addText("Próximos vencimientos", {
    x: 0.5, y: 4.0, w: 6, h: 0.36, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 16, bold: true, color: NAVY, valign: "middle",
  });
  const venc = [
    { d: D(16, 10), t: "Envío respuesta resultados preliminares PLD/FT", c: "Circular 23486", n: "" },
    { d: D(23, 10), t: "Envío de nuevas observaciones – Gobierno Corporativo", c: "Circular 22741", n: "" },
    { d: D(30, 10), t: "Seguimiento de observaciones – Gobierno Corporativo", c: "Circular 22741", n: "" },
    { d: D(3, 11), t: "Envío del Plan de Recuperación", c: "Circular 22737", n: "" },
  ].filter((v) => dias(v.d) >= 0);
  const cw = (12.33 - 0.2 * (venc.length - 1)) / venc.length, cy = 4.46, ch = 2.5;
  venc.forEach((v, i) => {
    const x = 0.5 + i * (cw + 0.2), n = dias(v.d), first = i === 0;
    s2.addShape("roundRect", { x, y: cy, w: cw, h: ch, rectRadius: 0.12, fill: { color: first ? NAVY : CARD }, line: { color: first ? NAVY : CARD } });
    s2.addText(n === 0 ? [
      { text: fmt(v.d), options: { fontSize: 40, bold: true, color: "9CCBF2" } },
    ] : [
      { text: String(n), options: { fontSize: 40, bold: true, color: first ? "9CCBF2" : PLAZO } },
      { text: n === 1 ? "  día" : "  días", options: { fontSize: 15, bold: true, color: first ? WHITE : NAVY } },
    ], { x: x + 0.22, y: cy + 0.14, w: cw - 0.4, h: 0.72, margin: 0, isTextBox: true, fontFace: FONT, valign: "middle" });
    s2.addText([
      ...(n === 0 ? [] : [{ text: fmt(v.d) + "/" + v.d.getFullYear(), options: { fontSize: 13, bold: true, color: first ? WHITE : NAVY, breakLine: true } }]),
      { text: v.t, options: { fontSize: 11.5, color: first ? WHITE : INK, breakLine: true } },
      { text: v.c, options: { fontSize: 10, color: first ? "A9D8DE" : BLUE, breakLine: !!v.n } },
      ...(v.n ? [{ text: v.n, options: { fontSize: 10, italic: true, color: first ? "DCE6F0" : "3E4F60" } }] : []),
    ], { x: x + 0.22, y: cy + 0.92, w: cw - 0.4, h: ch - 1.04, margin: 0, isTextBox: true, fontFace: FONT, valign: "top", paraSpaceAfter: 3 });
  });

  addFooter(s2, "Riesgos Específicos (Circular 22742): inicio de visita 09/10/2026.");
  s2.addNotes(
    "Corte " + fmt(CORTE) + ". Plazo de respuesta PLD/FT: 16/10 (Circular 23486). Siguen Gobierno Corporativo el 23/10 y 30/10 (Circular 22741) " +
    "y el Plan de Recuperación el 03/11 (Circular 22737)."
  );

  await pres.writeFile({ fileName: OUT });
  console.log("OK", OUT);
})();
