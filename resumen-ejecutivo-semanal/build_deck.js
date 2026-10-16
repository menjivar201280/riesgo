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
    x: 0.5, y: 0.32, w: 9.6, h: 0.6, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 26, bold: true, color: NAVY, valign: "middle",
  });
  slide.addText(subtitle, {
    x: 0.5, y: 0.92, w: 9.6, h: 0.32, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 13, color: BLUE, valign: "middle",
  });
}

function addPill(slide, text, x, y, w) {
  slide.addShape("roundRect", {
    x, y, w, h: 0.36, rectRadius: 0.18, fill: { color: NAVY }, line: { color: NAVY },
  });
  slide.addText(text, {
    x, y, w, h: 0.36, margin: 0, isTextBox: true, align: "center", valign: "middle",
    fontFace: FONT, fontSize: 11, bold: true, color: WHITE,
  });
}

function addFooter(slide, text) {
  slide.addText(text, {
    x: 0.5, y: 7.13, w: 12.33, h: 0.24, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 9, color: "4A5A6A", valign: "middle",
  });
}

// Lista con etiqueta en negrita + texto.
function itemRuns(items, size) {
  return items.map(([label, text], i) => ({
    text: "",
    options: {},
    _label: label, _text: text, _last: i === items.length - 1,
  })).flatMap((it) => [
    { text: it._label + "  ", options: { bold: true, color: NAVY, bullet: { code: "25A0", indent: 12 }, paraSpaceAfter: 4, fontSize: size } },
    { text: it._text, options: { color: INK, fontSize: size, breakLine: !it._last } },
  ]);
}

function progress(slide, x, y, label, pct) {
  slide.addText(label, {
    x, y, w: 1.55, h: 0.22, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 9.5, bold: true, color: NAVY, align: "right",
  });
  const bw = 1.05;
  slide.addShape("roundRect", { x, y: y + 0.29, w: bw, h: 0.11, rectRadius: 0.05, fill: { color: "D5DDE6" }, line: { color: "D5DDE6" } });
  slide.addShape("roundRect", { x, y: y + 0.29, w: bw * pct / 100, h: 0.11, rectRadius: 0.05, fill: { color: TEAL }, line: { color: TEAL } });
  slide.addText(pct + "%", {
    x: x + bw + 0.05, y: y + 0.2, w: 0.45, h: 0.28, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 12, bold: true, color: TEAL, align: "right", valign: "middle",
  });
}

async function card(slide, { x, y, w, h, ic, title, synth, items, status }) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.12, fill: { color: CARD }, line: { color: CARD } });
  slide.addShape("ellipse", { x: x + 0.22, y: y + 0.2, w: 0.56, h: 0.56, fill: { color: NAVY }, line: { color: NAVY } });
  slide.addImage({ data: await icon(ic, WHITE), x: x + 0.355, y: y + 0.335, w: 0.29, h: 0.29 });
  const tw = status ? w - 0.98 - 1.85 : w - 1.2;
  slide.addText(title, {
    x: x + 0.94, y: y + 0.17, w: tw, h: 0.34, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 15, bold: true, color: NAVY, valign: "middle",
  });
  slide.addText(synth, {
    x: x + 0.94, y: y + 0.5, w: w - 1.2, h: 0.28, margin: 0, isTextBox: true,
    fontFace: FONT, fontSize: 11, italic: true, color: BLUE, valign: "middle",
  });
  if (status) progress(slide, x + w - 1.78, y + 0.12, status.label, status.pct);
  slide.addText(itemRuns(items, 11), {
    x: x + 0.24, y: y + 0.92, w: w - 0.46, h: h - 1.04, margin: 0, isTextBox: true,
    fontFace: FONT, valign: "top",
  });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in, 16:9
  pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
  pres.title = "Resumen Ejecutivo de Riesgos – COMÉDICA";
  pres.author = "Departamento de Administración de Riesgos";

  // ---------------- DIAPOSITIVA 1 ----------------
  const s1 = pres.addSlide();
  s1.background = { color: WHITE };
  addTitle(s1, "Gestión Integral de Riesgos | Resumen Ejecutivo Semanal",
    "Departamento de Administración de Riesgos · COMÉDICA de R.L.");
  addPill(s1, "Seguimiento semanal · S1 – S2", 10.43, 0.45, 2.4);

  const W = 6.04, H = 2.72, X1 = 0.5, X2 = 6.79, Y1 = 1.48, Y2 = 4.36;

  await card(s1, {
    x: X1, y: Y1, w: W, h: H, ic: fa.FaShieldHalved,
    title: "Riesgo Operacional y Continuidad",
    synth: "Cultura de riesgos, continuidad del negocio y control de eventos",
    status: { label: "Examen GIR 2026", pct: 90 },
    items: [
      ["Capacitación GIR 2026:", "capacitación, examen e informe de seguimiento del examen"],
      ["Continuidad:", "informe de la 2.ª prueba (01/10/2026) y actualización de Matriz BIA"],
      ["Matriz de Riesgos Institucional 2026:", "seguimiento de planes de acción"],
      ["Eventos:", "caídas de TC (25/09) y Orion (26/09); ventana de mantenimiento"],
      ["Opiniones e informes:", "Fábrica de Créditos; caso asociado J. Jacobo Marroquín; Acuerdos Semestrales del Comité de Riesgos (Junta de Vigilancia)"],
      ["Auditorías:", "información a Auditoría Interna (NRP-24) y Auditoría Externa"],
    ],
  });

  await card(s1, {
    x: X2, y: Y1, w: W, h: H, ic: fa.FaScaleBalanced,
    title: "Riesgo Normativo y SSF",
    synth: "Requerimientos SSF coordinados con las áreas y pruebas de descargo",
    status: { label: "Avalúos (último reporte)", pct: 60 },
    items: [
      ["Gobierno Corporativo · 22741:", "matriz de requerimiento, solicitudes a áreas y seguimiento de descargos (visita 18432)"],
      ["PLD/FT · 23486:", "seguimiento al envío de pruebas de descargo"],
      ["Riesgo Específico · 22742:", "presentación de resultados preliminares"],
      ["Plan de Recuperación · 22737:", "indicadores de Gobierno Corporativo y Reputacional"],
      ["Avalúos:", "política, procedimientos y perfil de revisor; asignación a perito, gestión electrónica y aviso a asociados"],
      ["Otros:", "envíos regulatorios (DAR); acta de sesión del 29/08/2026"],
    ],
  });

  await card(s1, {
    x: X1, y: Y2, w: W, h: H, ic: fa.FaChartLine,
    title: "Riesgo Financiero y Crediticio",
    synth: "Medición, reportes regulatorios y control de cartera y reservas",
    items: [
      ["Crédito:", "informe al Comité – pérdidas esperadas, matrices de transición, cosechas"],
      ["Liquidez y mercado:", "informe al Comité, escenarios regulatorios y anexo 03 NRP-05"],
      ["Activos de riesgo:", "NCB-022, Central de Riesgo (NRP-41), CTRI-RECR y créditos relacionados"],
      ["Tarjeta y reservas:", "análisis de cartera TC, simulación de vencida; reserva de saneamiento"],
      ["Seguimiento:", "traslados de deuda, segmentación de productos, consultas de aspirantes"],
      ["Largo plazo:", "Política de TC, sistema de créditos (NRP-41), expedientes TC de alto riesgo"],
    ],
  });

  await card(s1, {
    x: X2, y: Y2, w: W, h: H, ic: fa.FaDiagramProject,
    title: "Proyectos Transversales",
    synth: "Mejora institucional en coordinación con otras áreas",
    items: [
      ["BIA:", "elaboración del cuestionario para entrevistas"],
      ["Formularios:", "unificación de Actualización de Datos (Atención al Asociado) y Conozca a tu Cliente (Oficialía de Cumplimiento)"],
      ["Comparativos de políticas:", "servicios tercerizados y gobierno corporativo"],
      ["Procesos:", "Matriz de Administración de Sistemas Informáticos y solicitud de cambio de estructura del Departamento"],
      ["Atención al Asociado:", "actualización de página web"],
    ],
  });

  addFooter(s1, "Fuentes: presentación base «Resumen Ejecutivo Riesgos» y documento de actividades semanales. Sin estado documentado: se presenta solo la actividad.");
  s1.addNotes(
    "Resumen de los cuatro frentes de gestión, integrando las actividades de S1 y S2. " +
    "Examen GIR 2026: avance 90% (reportado). Avalúos: 60% según la presentación anterior, por confirmar. " +
    "Los demás frentes se presentan sin estado porque los documentos no lo indican."
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
