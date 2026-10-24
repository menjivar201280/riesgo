// Genera: Seguimiento_Acuerdos_Comite_Riesgo_Sesion_150.pptx
// Uso: node build.js <carpeta_con_imagenes_de_notas>
const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.PPTX_SKILL + "/scripts/apply_theme.js");

const IMG = process.argv[2] || path.join(__dirname, "evidencias");
const OUT = path.join(__dirname, "Seguimiento_Acuerdos_Comite_Riesgo_Sesion_150.pptx");

const THEME = {
  name: "Comite de Riesgos",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2433", lt1: "FFFFFF", dk2: "0E2A47", lt2: "EEF3F7",
    accent1: "0F7C8C", accent2: "2E8B57", accent3: "D99A1E",
    accent4: "B5473A", accent5: "5B6B7F", accent6: "8CC6CF",
    hlink: "0F7C8C", folHlink: "5B6B7F",
  },
};
const H = THEME.colors;

async function icon(Comp, color = "#FFFFFF") {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "Seguimiento Acuerdos Comité de Riesgo - Sesión 150";
  pres.company = "COMEDICA de R.L.";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;

  const I = {};
  for (const [k, comp] of Object.entries({
    users: fa.FaUsers, coins: fa.FaCoins, sync: fa.FaSyncAlt, pension: fa.FaHandHoldingUsd,
    user: fa.FaUserTie, gavel: fa.FaGavel, file: fa.FaFileSignature, check: fa.FaCheckCircle,
    clock: fa.FaHourglassHalf, info: fa.FaInfoCircle, flag: fa.FaFlag, bank: fa.FaUniversity,
    tasks: fa.FaTasks, cal: fa.FaCalendarAlt,
  })) I[k] = await icon(comp);

  // ---------- Layouts ----------
  pres.defineSlideMaster({
    title: "PORTADA",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.0, w: 7.6, h: 2.0,
        fontSize: 44, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.35, w: 7.6, h: 1.0,
        fontSize: 18, color: C.accent6, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENIDO",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.8, 0.6],
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 0.35, w: 10, h: 0.35,
        fontSize: 12, bold: true, color: C.accent1, charSpacing: 2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.68, w: 12.1, h: 0.72,
        fontSize: 34, bold: true, color: C.text2, valign: "top", align: "left", margin: 0 }, text: "" } },
      { text: { text: "COMEDICA de R.L.  ·  Comité de Riesgos  ·  Seguimiento de acuerdos Sesión 150",
        options: { x: 0.6, y: 7.0, w: 9, h: 0.3, fontSize: 10, color: C.accent5, margin: 0 } } },
    ],
    slideNumber: { x: 12.3, y: 7.0, w: 0.45, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  });
  pres.defineSlideMaster({
    title: "CIERRE",
    background: { color: C.text2 },
    objects: [
      { placeholder: { options: { name: "kicker", type: "body", x: 0.8, y: 0.55, w: 10, h: 0.35,
        fontSize: 12, bold: true, color: C.accent6, charSpacing: 2, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 0.9, w: 11.7, h: 0.8,
        fontSize: 36, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });

  // ---------- Helpers ----------
  const tb = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, valign: "top", fontSize: 14, color: C.text1, ...o });
  const card = (s, x, y, w, h, name, fill = C.background2) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, rectRadius: 0.12, objectName: name });
  const iconCircle = (s, img, x, y, d, fill = C.accent1, name = "Icono") => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: name + " fondo" });
    const p = d * 0.24;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: name });
  };
  const cardHeader = (s, x, y, img, label, fill = C.accent1) => {
    iconCircle(s, img, x + 0.25, y + 0.16, 0.4, fill, "Icono " + label);
    tb(s, label, { x: x + 0.78, y: y + 0.21, w: 5, h: 0.3, fontSize: 13, bold: true, color: C.accent1, charSpacing: 1.5 });
  };
  const chip = (s, x, y, w, h, text, fill, fontSize = 12) =>
    s.addText(text, { isTextBox: true, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.08, x, y, w, h,
      fill: { color: fill }, color: C.background1, bold: true, fontSize, align: "center", valign: "middle", margin: 0, objectName: "Estado " + text });

  const STATUS = {
    ok: { label: "CUMPLIDO", color: C.accent2, img: "check", note: "Acuerdo atendido con evidencia presentada al Comité." },
    prog: { label: "EN PROCESO", color: C.accent3, img: "clock", note: "Nota recibida; análisis pendiente de presentación." },
    nf: { label: "ATENDIDO · NO FAVORABLE", color: C.accent4, img: "info", note: "Gestión realizada; instituciones no aceptan aún a COMEDICA." },
  };

  const NOTES = [
    {
      num: "01", topicImg: "users", short: "Concentración de mayores depositantes",
      title: "Concentración de mayores depositantes",
      name: "Lic. José Humberto Barillas", cargo: "Gerente de Negocios",
      status: "ok", thumb: "1ba1bc7c-1.png", thumbTitle: "Nota del Comité",
      thumbText: "10/09/2026 · Recibida en Depto. de Créditos el 10/09/2026",
      acuerdoH: 1.55,
      acuerdo: [{ text: "Realizar visitas personalizadas a los asociados de mayor concentración en depósitos, para fortalecer la relación y promover la incorporación de sus hijos como asociados, como medida para gestionar el riesgo de concentración. Compartir las gestiones con el Comité." }],
    },
    {
      num: "02", topicImg: "coins", short: "Liberación de reservas y recuperación de cartera castigada",
      title: "Liberación de reservas e ingresos por recuperación",
      name: "Ing. José Augusto De la O", cargo: "Gerente de Finanzas y Planificación",
      status: "ok", thumb: "287acb7a-1.png", thumbTitle: "Memorando 04_2026",
      thumbText: "30/09/2026 · Recibido por el Consejo de Administración el 01/10/2026",
      acuerdoH: 2.35,
      acuerdo: [
        { text: "Presentar un análisis explicativo sobre la liberación de reservas reflejada como ingresos en los estados financieros, que incluya:", options: { breakLine: true, paraSpaceAfter: 4 } },
        { text: "Origen, composición y criterios contables de reconocimiento.", options: { bullet: { type: "number" }, breakLine: true } },
        { text: "Monto de recuperación de cartera castigada y su participación.", options: { bullet: { type: "number" }, breakLine: true } },
        { text: "Conciliación: liberación de reservas vs. recuperaciones vs. otros conceptos.", options: { bullet: { type: "number" }, breakLine: true, paraSpaceAfter: 4 } },
        { text: "Remitir a la Asistente del Consejo de Administración.", options: { italic: true, color: C.accent5 } },
      ],
    },
    {
      num: "03", topicImg: "sync", short: "Cartera en proceso de reestructuración",
      title: "Cartera en proceso de reestructuración",
      name: "Lic. Carlos Roberto Aguilar", cargo: "Jefe de Soluciones Financieras",
      status: "prog", thumb: "c3944e2d-1.png", thumbTitle: "Nota del Comité",
      thumbText: "10/09/2026 · Recibida en Depto. de Soluciones el 10/09/2026",
      acuerdoH: 2.1,
      acuerdo: [
        { text: "Presentar ante el Comité un análisis integral de la cartera en proceso de reestructuración, que incluya:", options: { breakLine: true, paraSpaceAfter: 4 } },
        { text: "Comportamiento y características de la cartera reestructurada.", options: { bullet: true, breakLine: true } },
        { text: "Principales causas, factores o circunstancias que la originan.", options: { bullet: true, breakLine: true } },
        { text: "Tendencias observadas y riesgos asociados.", options: { bullet: true } },
      ],
      timeline: [["27/08/2026", "Acuerdo en Sesión 150"], ["10/09/2026", "Nota recibida por Depto. de Soluciones (Griselda Sigüenza)"], ["Pendiente", "Presentación del análisis al Comité de Riesgos"]],
    },
    {
      num: "04", topicImg: "pension", short: "Servicio de pago de pensiones en la Cooperativa",
      title: "Pago de pensiones en la Cooperativa",
      name: "Lic. Luis Marcelo Cruz Vigil", cargo: "Consulta remitida el 02/10/2026",
      status: "nf", thumb: "b4460910-1.png", thumbTitle: "Constancias a AFP",
      thumbText: "Solicitudes del 05/11/2025, con sello de recibido 06/11/2025",
      acuerdoH: 1.3,
      acuerdo: [{ text: "Informar si se han realizado gestiones, análisis o acercamientos para implementar el servicio de pago de pensiones en la Cooperativa, así como su estado y avances." }],
    },
  ];

  // ---------- 1. Portada ----------
  {
    const s = pres.addSlide({ masterName: "PORTADA" });
    s.addText("Seguimiento Acuerdos Comité de Riesgo", { placeholder: "title" });
    s.addText("Sesión 150 del 27 de agosto de 2026  ·  Estado de notas, responsables, acuerdos y evidencias", { placeholder: "body" });
    tb(s, "COMEDICA DE R.L.  ·  ADMINISTRACIÓN DE RIESGOS", { x: 0.8, y: 1.45, w: 7.6, h: 0.35, fontSize: 13, bold: true, color: C.accent6, charSpacing: 2 });
    s.addShape(pres.shapes.OVAL, { x: 8.95, y: 1.45, w: 3.6, h: 3.6, fill: { color: C.accent1 }, line: { type: "none" }, objectName: "Círculo sesión" });
    s.addShape(pres.shapes.OVAL, { x: 8.75, y: 1.25, w: 4.0, h: 4.0, fill: { type: "none" }, line: { color: C.accent6, width: 1.5 }, objectName: "Anillo sesión" });
    tb(s, "SESIÓN", { x: 8.95, y: 2.2, w: 3.6, h: 0.4, fontSize: 18, bold: true, color: C.background1, align: "center", charSpacing: 4 });
    tb(s, "150", { x: 8.95, y: 2.55, w: 3.6, h: 1.3, fontSize: 80, bold: true, color: C.background1, align: "center" });
    tb(s, "27 · 08 · 2026", { x: 8.95, y: 3.9, w: 3.6, h: 0.4, fontSize: 16, color: C.background1, align: "center" });
    tb(s, "Seguimiento al 24 de octubre de 2026", { x: 0.8, y: 6.55, w: 7, h: 0.35, fontSize: 12, color: C.accent6 });
  }

  // ---------- 2. Tablero resumen ----------
  {
    const s = pres.addSlide({ masterName: "CONTENIDO" });
    s.addText("TABLERO DE SEGUIMIENTO", { placeholder: "kicker" });
    s.addText("4 acuerdos en seguimiento: 3 atendidos y 1 en proceso", { placeholder: "title" });
    const stats = [["4", "Notas emitidas por el Comité", C.text2], ["3", "Acuerdos atendidos con respuesta", C.accent2], ["1", "Análisis pendiente de presentar", C.accent3]];
    stats.forEach(([n, l, col], i) => {
      const x = 0.6 + i * 4.13;
      card(s, x, 1.65, 3.87, 1.35, "Indicador " + (i + 1));
      tb(s, n, { x: x + 0.3, y: 1.75, w: 1.1, h: 1.15, fontSize: 60, bold: true, color: col, valign: "middle" });
      tb(s, l, { x: x + 1.45, y: 1.75, w: 2.2, h: 1.15, fontSize: 15, color: C.text1, valign: "middle" });
    });
    NOTES.forEach((d, i) => {
      const x = 0.6 + i * 3.1, y = 3.3, w = 2.85, h = 3.45;
      const st = STATUS[d.status];
      card(s, x, y, w, h, "Nota " + d.num);
      iconCircle(s, I[d.topicImg], x + 0.25, y + 0.25, 0.7, C.accent1, "Icono nota " + d.num);
      tb(s, "NOTA " + d.num, { x: x + 1.1, y: y + 0.42, w: 1.6, h: 0.35, fontSize: 14, bold: true, color: C.accent1, charSpacing: 1.5 });
      tb(s, d.short, { x: x + 0.25, y: y + 1.15, w: w - 0.5, h: 0.95, fontSize: 16, bold: true, color: C.text2 });
      tb(s, [{ text: d.name, options: { breakLine: true } }, { text: d.cargo, options: { color: C.accent5 } }],
        { x: x + 0.25, y: y + 2.1, w: w - 0.5, h: 0.75, fontSize: 12 });
      chip(s, x + 0.25, y + h - 0.6, w - 0.5, 0.38, st.label, st.color, st.label.length > 12 ? 11 : 12);
    });
  }

  // ---------- 3. Matriz consolidada ----------
  {
    const s = pres.addSlide({ masterName: "CONTENIDO" });
    s.addText("MATRIZ DE SEGUIMIENTO", { placeholder: "kicker" });
    s.addText("Notas, responsables, acuerdos y evidencias", { placeholder: "title" });
    const hd = (t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text2 }, fontSize: 12, valign: "middle" } });
    const st = (k) => ({ text: STATUS[k].label, options: { bold: true, color: C.background1, fill: { color: STATUS[k].color }, align: "center", valign: "middle", fontSize: 11 } });
    const resp = (n, c) => [{ text: n, options: { bold: true, breakLine: true } }, { text: c, options: { color: C.accent5 } }];
    const num = (n) => ({ text: n, options: { bold: true, fontSize: 20, color: C.accent1, align: "center", valign: "middle" } });
    const rows = [
      [hd("N° nota"), hd("Responsable"), hd("Acuerdo"), hd("Evidencia / respuesta"), { text: "Estado", options: { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center", fontSize: 12, valign: "middle" } }],
      [num("01"), { text: resp("Lic. José Humberto Barillas", "Gerente de Negocios") },
        "Visitas personalizadas a los mayores depositantes para gestionar el riesgo de concentración y promover la afiliación de sus hijos",
        "Informe de visitas del 16/10/2026: 10 depositantes visitados ($27.15 M); 69.40% con beneficiarios afiliados", st("ok")],
      [num("02"), { text: resp("Ing. José Augusto De la O", "Gerente de Finanzas y Planificación") },
        "Análisis de liberación de reservas e ingresos por recuperación de cartera castigada (CR-2026-49, lit. j)",
        "Memorando 04_2026 (30/09/2026): liberación $1,073,528.39 (4.33% de ingresos); recuperación castigada $121,441.73 (0.49%)", st("ok")],
      [num("03"), { text: resp("Lic. Carlos Roberto Aguilar", "Jefe de Soluciones Financieras") },
        "Análisis integral de la cartera en reestructuración: características, causas, tendencias y riesgos",
        "Nota recibida el 10/09/2026. Análisis pendiente de presentación", st("prog")],
      [num("04"), { text: resp("Lic. Luis Marcelo Cruz Vigil", "Consulta del 02/10/2026") },
        "Informar gestiones para prestar el servicio de pago de pensiones en la Cooperativa",
        "Solicitudes a AFP CONFIA, AFP CRECER, IPSFA e ISP. Respuesta 03/10/2026: COMEDICA aún no califica como entidad receptora", st("nf")],
    ];
    s.addTable(rows, {
      x: 0.6, y: 1.65, w: 12.13, colW: [0.95, 2.45, 3.75, 3.48, 1.5],
      rowH: [0.45, 1.15, 1.15, 1.15, 1.15],
      fontSize: 12, color: C.text1, valign: "middle", margin: [0.06, 0.1, 0.06, 0.1],
      border: { type: "solid", pt: 0.75, color: "D5DEE6" }, fill: { color: C.background1 },
    });
  }

  // ---------- 4-7. Detalle por nota ----------
  for (const d of NOTES) {
    const s = pres.addSlide({ masterName: "CONTENIDO" });
    const st = STATUS[d.status];
    s.addText("NOTA " + d.num + "  ·  ACUERDO SESIÓN 150", { placeholder: "kicker" });
    s.addText(d.title, { placeholder: "title" });

    // Columna izquierda: responsable, estado, documento soporte
    card(s, 0.6, 1.6, 3.7, 1.85, "Tarjeta responsable");
    cardHeader(s, 0.6, 1.6, I.user, "RESPONSABLE");
    tb(s, d.name, { x: 0.85, y: 2.25, w: 3.25, h: 0.4, fontSize: 17, bold: true, color: C.text2 });
    tb(s, d.cargo, { x: 0.85, y: 2.68, w: 3.25, h: 0.6, fontSize: 14, color: C.accent5 });

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 3.65, w: 3.7, h: 1.0, fill: { color: st.color }, line: { type: "none" }, rectRadius: 0.12, objectName: "Tarjeta estado" });
    s.addImage({ data: I[st.img], x: 0.85, y: 3.9, w: 0.5, h: 0.5, objectName: "Icono estado" });
    tb(s, [{ text: "ESTADO", options: { fontSize: 11, charSpacing: 1.5, breakLine: true } }, { text: st.label, options: { fontSize: st.label.length > 12 ? 15 : 20, bold: true } }],
      { x: 1.55, y: 3.78, w: 2.65, h: 0.75, color: C.background1, valign: "middle" });

    const th = 1.95, tw = th * 0.773;
    s.addImage({ path: path.join(IMG, d.thumb), x: 0.6, y: 4.85, w: tw, h: th, objectName: "Documento soporte",
      altText: "Imagen del documento soporte de la nota " + d.num,
      shadow: { type: "outer", color: "1B2433", opacity: 0.25, blur: 6, offset: 2, angle: 90 } });
    tb(s, d.thumbTitle.toUpperCase(), { x: 2.3, y: 4.95, w: 2.0, h: 0.3, fontSize: 11, bold: true, color: C.accent1, charSpacing: 1.5 });
    tb(s, d.thumbText, { x: 2.3, y: 5.28, w: 2.0, h: 1.3, fontSize: 12, color: C.accent5 });

    // Columna derecha: acuerdo y evidencia
    const X = 4.6, W = 8.13;
    card(s, X, 1.6, W, d.acuerdoH, "Tarjeta acuerdo");
    cardHeader(s, X, 1.6, I.gavel, "ACUERDO");
    tb(s, d.acuerdo, { x: X + 0.3, y: 2.28, w: W - 0.6, h: d.acuerdoH - 0.75, fontSize: 14, paraSpaceAfter: 2 });

    const ey = 1.6 + d.acuerdoH + 0.2, eh = 6.75 - ey;
    card(s, X, ey, W, eh, "Tarjeta evidencia");
    cardHeader(s, X, ey, I.file, "EVIDENCIA / RESPUESTA");
    const by = ey + 0.7;

    if (d.num === "01") {
      tb(s, [{ text: "Informe de visitas del 16/10/2026: ", options: { bold: true } },
        { text: "se visitó a los 10 mayores depositantes ($27,152,269.99) por medio de los jefes de agencia, en coordinación con la Subgerencia de Canales." }],
        { x: X + 0.3, y: by, w: 5.1, h: 0.8 });
      const segs = [
        ["69.40%", "6 de 10 · beneficiarios afiliados", "$18,862,867.16 · mantener fidelización", C.accent2],
        ["24.20%", "3 de 10 · depositantes jurídicos", "$6,557,464.37 · negociación anticipada", C.accent1],
        ["6.40%", "1 de 10 · beneficiarios no afiliados", "$1,731,938.46 · plan de afiliación", C.accent4],
      ];
      segs.forEach(([p, a, b, col], i) => {
        const y = by + 0.88 + i * 0.6;
        s.addShape(pres.shapes.OVAL, { x: X + 0.32, y: y + 0.12, w: 0.22, h: 0.22, fill: { color: col }, line: { type: "none" }, objectName: "Marcador " + p });
        tb(s, p, { x: X + 0.7, y, w: 1.15, h: 0.5, fontSize: 22, bold: true, color: col, valign: "middle" });
        tb(s, [{ text: a, options: { bold: true, breakLine: true } }, { text: b, options: { color: C.accent5, fontSize: 12 } }],
          { x: X + 1.9, y: y - 0.02, w: 3.4, h: 0.6, fontSize: 14 });
      });
      s.addChart(pres.charts.DOUGHNUT, [{ name: "Participación", labels: ["Beneficiarios afiliados", "Jurídicos", "Beneficiarios no afiliados"], values: [69.4, 24.2, 6.4] }], {
        x: X + 5.55, y: ey + 0.45, w: 2.4, h: 2.6, holeSize: 62, chartColors: [H.accent2, H.accent1, H.accent4],
        showLegend: false, showValue: false, showPercent: false, showTitle: false, dataBorder: { pt: 1.5, color: "FFFFFF" }, objectName: "Gráfico concentración",
      });
      tb(s, [{ text: "$27.15 M", options: { bold: true, fontSize: 18, color: C.text2, breakLine: true } }, { text: "10 depositantes", options: { fontSize: 11, color: C.accent5 } }],
        { x: X + 5.55, y: ey + 1.4, w: 2.4, h: 0.7, align: "center", valign: "middle" });
    } else if (d.num === "04") {
      tb(s, [{ text: "Respuesta del 03/10/2026: ", options: { bold: true } },
        { text: "se enviaron solicitudes el 05/11/2025 a cuatro instituciones pagadoras de pensiones para que los asociados reciban su pensión en COMEDICA (recibidas el 06/11/2025)." }],
        { x: X + 0.3, y: by, w: W - 0.6, h: 0.8 });
      const inst = ["AFP CONFIA", "AFP CRECER", "IPSFA", "ISP"];
      inst.forEach((n, i) => {
        const x = X + 0.3 + i * 1.9, y = by + 0.95;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.7, h: 0.95, fill: { color: C.background1 }, line: { color: "D5DEE6", width: 1 }, rectRadius: 0.1, objectName: "Institución " + n });
        iconCircle(s, I.bank, x + 0.6, y + 0.1, 0.45, C.accent5, "Icono " + n);
        tb(s, n, { x, y: y + 0.6, w: 1.7, h: 0.3, fontSize: 13, bold: true, color: C.text2, align: "center" });
      });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X + 0.3, y: by + 2.1, w: W - 0.6, h: 0.65, fill: { color: C.background1 }, line: { color: H.accent4, width: 1.25 }, rectRadius: 0.1, objectName: "Resultado" });
      s.addImage({ data: await icon(fa.FaExclamationTriangle, "#" + H.accent4), x: X + 0.5, y: by + 2.25, w: 0.35, h: 0.35, objectName: "Icono resultado" });
      tb(s, [{ text: "Resultado: ", options: { bold: true, color: C.accent4 } }, { text: "las instituciones indicaron por teléfono que COMEDICA actualmente no califica como entidad receptora de pensiones." }],
        { x: X + 1.0, y: by + 2.15, w: W - 1.45, h: 0.55, valign: "middle" });
    } else if (d.num === "02") {
      const stats = [
        ["$1,073,528.39", "Liberación de reservas", "4.33% de los ingresos totales · 100% reversión de estimaciones por incobrabilidad (cta. 1149)", C.accent1],
        ["$121,441.73", "Recuperación de cartera castigada", "0.49% de los ingresos totales · 81.06% corresponde a intereses refinanciados", C.accent2],
      ];
      stats.forEach(([n, a, b, col], i) => {
        const x = X + 0.3 + i * 3.84;
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: by, w: 3.69, h: 1.15, fill: { color: C.background1 }, line: { color: "D5DEE6", width: 1 }, rectRadius: 0.1, objectName: "Cifra " + a });
        tb(s, n, { x: x + 0.2, y: by + 0.08, w: 3.3, h: 0.42, fontSize: 22, bold: true, color: col });
        tb(s, [{ text: a, options: { bold: true, color: C.text2, breakLine: true } }, { text: b, options: { color: C.accent5 } }],
          { x: x + 0.2, y: by + 0.5, w: 3.35, h: 0.6, fontSize: 11 });
      });
      tb(s, [{ text: "Conclusión: ", options: { bold: true } }, { text: "mismo ámbito (riesgo crediticio), origen distinto: la liberación revierte provisiones por mejora crediticia; la recuperación son cobros de créditos ya castigados." }],
        { x: X + 0.3, y: by + 1.25, w: W - 0.6, h: 0.6, fontSize: 13 });
    } else {
      // Línea de tiempo: acuerdo -> nota recibida -> pendiente
      tb(s, [{ text: "A la fecha " }, { text: "no se ha recibido el análisis solicitado", options: { bold: true } }, { text: "; se mantiene en seguimiento." }],
        { x: X + 0.3, y: by, w: W - 0.6, h: 0.35 });
      const ty = by + 0.5, d0 = 0.36, lw = 2.6;
      const nodes = [X + 0.3, X + W / 2 - d0 / 2, X + W - 0.3 - d0];
      s.addShape(pres.shapes.LINE, { x: nodes[0] + d0, y: ty + d0 / 2, w: nodes[2] - nodes[0] - d0, h: 0, line: { color: H.accent6, width: 2 }, objectName: "Línea de tiempo" });
      d.timeline.forEach(([f, t], i) => {
        const cx = nodes[i], pend = f === "Pendiente";
        s.addShape(pres.shapes.OVAL, { x: cx, y: ty, w: d0, h: d0, fill: { color: pend ? C.accent3 : C.accent1 }, line: { color: H.lt2, width: 3 }, objectName: "Hito " + (i + 1) });
        const lx = i === 0 ? X + 0.3 : i === 1 ? X + W / 2 - lw / 2 : X + W - 0.3 - lw;
        const align = ["left", "center", "right"][i];
        tb(s, f, { x: lx, y: ty + 0.45, w: lw, h: 0.3, fontSize: 15, bold: true, color: pend ? C.accent3 : C.text2, align });
        tb(s, t, { x: lx, y: ty + 0.77, w: lw, h: 0.55, fontSize: 12, color: C.accent5, align });
      });
    }
  }

  // ---------- 8. Próximos pasos ----------
  {
    const s = pres.addSlide({ masterName: "CIERRE" });
    s.addText("PRÓXIMOS PASOS", { placeholder: "kicker" });
    s.addText("Acciones para cerrar los acuerdos de la Sesión 150", { placeholder: "title" });
    const steps = [
      ["03", I.clock, C.accent3, "Requerir la presentación del análisis de cartera reestructurada", "Jefatura de Soluciones Financieras · nota recibida el 10/09/2026"],
      ["01", I.users, C.accent2, "Monitorear el plan de afiliación y los vencimientos", "Depositante con beneficiarios no afiliados (6.40%) y depositantes jurídicos (24.20%)"],
      ["04", I.pension, C.accent4, "Solicitar respuesta formal por escrito de las instituciones", "Documentar los requisitos para que COMEDICA califique como entidad receptora de pensiones"],
    ];
    steps.forEach(([n, img, col, t, sub], i) => {
      const y = 2.15 + i * 1.5;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.8, y, w: 11.73, h: 1.25, fill: { color: "16365A" }, line: { type: "none" }, rectRadius: 0.12, objectName: "Paso " + (i + 1) });
      iconCircle(s, img, 1.1, y + 0.25, 0.75, col, "Icono paso " + (i + 1));
      tb(s, "NOTA " + n, { x: 2.15, y: y + 0.2, w: 3, h: 0.3, fontSize: 12, bold: true, color: C.accent6, charSpacing: 1.5 });
      tb(s, t, { x: 2.15, y: y + 0.45, w: 10, h: 0.4, fontSize: 20, bold: true, color: C.background1 });
      tb(s, sub, { x: 2.15, y: y + 0.85, w: 10, h: 0.3, fontSize: 13, color: C.accent6 });
    });
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("OK", OUT);
})();
