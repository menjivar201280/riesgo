const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/eb134167-ea2a-4bd7-8c34-83283d88af55_b18ad4a5-0fa1-4688-82cb-4bf483ec9783/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";
const THEME = {
  name: "SSF Gobierno Corporativo",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "142033", lt1: "FFFFFF", dk2: "173A6E", lt2: "EEF2F8",
    accent1: "173A6E", // azul marino 2025
    accent2: "1D7A4B", // verde 2026
    accent3: "B42318", // rojo no superada
    accent4: "D99A1E", // ámbar parcial
    accent5: "586579", // gris texto secundario
    accent6: "A86410", // ámbar oscuro (texto)
    hlink: "173A6E", folHlink: "586579",
  },
};
const LINE_HEX = "D9E0EA";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Supervisiones de Gobierno Corporativo SSF a COMEDICA de R.L.";
pres.subject = "Resultados de visitas de supervisión 2025 y 2026";
pres.author = "COMEDICA de R.L.";
const C = pres.SchemeColor;
const NAVY = C.accent1, GREEN = C.accent2, RED = C.accent3, AMBER = C.accent4, MUTED = C.accent5, AMBER_INK = C.accent6;
const WHITE = C.background1, TINT = C.background2, INK = C.text1;
const W = 13.333, M = 0.6, PANEL = 4.5;
const FOOT = "SSF · Gobierno Corporativo (NRP-17) · COMEDICA de R.L.";

// ---------- Diseños ----------
pres.defineSlideMaster({
  title: "RESUMEN PANEL",
  background: { color: WHITE },
  objects: [
    { rect: { x: 0, y: 0, w: PANEL, h: 7.5, fill: { color: NAVY }, line: { color: NAVY } } },
    { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.55, w: PANEL - 2 * M + 0.2, h: 0.6, fontSize: 11, bold: true, color: TINT, charSpacing: 2, margin: 0, valign: "top" }, text: "SECCIÓN" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 1.25, w: PANEL - 2 * M + 0.2, h: 1.6, fontFace: THEME.headFontFace, fontSize: 30, bold: true, color: WHITE, valign: "top", align: "left", margin: 0 }, text: "Título" } },
    { text: { text: FOOT, options: { x: PANEL + 0.5, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: WHITE },
  objects: [
    { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.35, w: 9, h: 0.3, fontSize: 11, bold: true, color: MUTED, charSpacing: 2, margin: 0 }, text: "SECCIÓN" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.65, w: 8.6, h: 0.7, fontFace: THEME.headFontFace, fontSize: 30, bold: true, color: INK, valign: "top", align: "left", margin: 0 }, text: "Título" } },
    { text: { text: FOOT, options: { x: M, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
});

// ---------- Ayudas ----------
const addSlide = (master, section, eyebrow, title) => {
  pres.addSection({ title: section });
  const s = pres.addSlide({ masterName: master, sectionTitle: section });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(title, { placeholder: "title" });
  return s;
};
const card = (s, name, x, y, w, h, fill = TINT, extra = {}) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name, x, y, w, h, rectRadius: 0.1, fill: { color: fill, ...(extra.transparency != null ? { transparency: extra.transparency } : {}) }, line: extra.line || { color: fill, transparency: 100 } });
const txt = (s, name, text, o) => s.addText(text, { objectName: name, isTextBox: true, margin: 0, fontSize: 14, color: INK, valign: "top", ...o });
const pill = (s, name, text, x, y, w, color, fs = 10.5, h = 0.3) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name + " fondo", x, y, w, h, rectRadius: h / 2, fill: { color }, line: { color, transparency: 100 } });
  s.addText(text, { objectName: name, isTextBox: true, x, y, w, h, margin: 0, align: "center", valign: "middle", fontSize: fs, bold: true, color: WHITE, charSpacing: 1 });
};
const badge = (s, name, n, x, y, d, color) => {
  s.addShape(pres.shapes.OVAL, { objectName: name + " círculo", x, y, w: d, h: d, fill: { color }, line: { color, transparency: 100 } });
  s.addText(String(n), { objectName: name, isTextBox: true, x, y, w: d, h: d, margin: 0, align: "center", valign: "middle", fontSize: d > 0.42 ? 15 : 12, bold: true, color: WHITE, fontFace: THEME.headFontFace });
};
const legendDot = (s, name, text, x, y, color, w = 2.2) => {
  s.addShape(pres.shapes.OVAL, { objectName: name + " punto", x, y: y + 0.07, w: 0.16, h: 0.16, fill: { color }, line: { color, transparency: 100 } });
  txt(s, name, text, { x: x + 0.24, y, w, h: 0.3, fontSize: 11, color: MUTED, valign: "middle" });
};

// ============ SLIDE 1 · RESUMEN ============
{
  const s = addSlide("RESUMEN PANEL", "1 · Resumen de notas SSF", "SUPERVISIONES DE GOBIERNO CORPORATIVO · COMEDICA DE R.L.", "Notas de notificación SSF 2025 – 2026");
  // KPIs en el panel
  const k = [["4", "notas oficiales recibidas"], ["9", "observaciones en 2025"], ["7 + 2", "nuevas y de seguimiento en 2026"], ["1 · 1", "no superada · parcialmente superada"]];
  k.forEach(([n, l], i) => {
    const x = M + (i % 2) * 1.75, y = 3.45 + Math.floor(i / 2) * 1.55;
    txt(s, "KPI valor " + (i + 1), n, { x, y, w: 1.6, h: 0.65, fontSize: 34, bold: true, color: WHITE, fontFace: THEME.headFontFace, valign: "bottom" });
    txt(s, "KPI texto " + (i + 1), l, { x, y: y + 0.7, w: 1.6, h: 0.6, fontSize: 12, color: TINT });
  });
  // Filas por año: nombramiento → informe
  const RX = PANEL + 0.5, RW = W - M - RX;
  const rows = [
    { y: 0.55, yr: "2025", sub: "Visita 1 · corte 30/04/2025", color: NAVY, a: ["09/04/2025", "Nombramiento de visita ampliada", "DS-SABAO-BCS-BC-9337", "Corte 30/04/2025 · inicio 19/05/2025"], b: ["05/11/2025", "Informe de resultados preliminares", "DS-SABAO-BCS-BC-17003", "9 observaciones"] },
    { y: 3.75, yr: "2026", sub: "Visita 2 · corte 31/07/2026", color: GREEN, a: ["31/07/2026", "Nombramiento de visita ampliada", "SABAO-BCS-BC-18432", "Corte 31/07/2026 · inicio 21/08/2026"], b: ["09/10/2026", "Informe de resultados preliminares", "DS-SABAO-BCS-BC-22741", "7 nuevas + 2 de seguimiento"] },
  ];
  const cw = (RW - 0.6) / 2, ch = 2.15;
  rows.forEach(r => {
    txt(s, "Año " + r.yr, r.yr, { x: RX, y: r.y, w: 1.4, h: 0.55, fontSize: 30, bold: true, color: r.color, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Visita " + r.yr, r.sub, { x: RX + 1.45, y: r.y, w: 4, h: 0.55, fontSize: 14, color: MUTED, valign: "middle" });
    [r.a, r.b].forEach(([fecha, tipo, no, det], j) => {
      const x = RX + j * (cw + 0.6), y = r.y + 0.75;
      card(s, `Tarjeta ${r.yr}-${j + 1}`, x, y, cw, ch, WHITE, { line: { color: LINE_HEX, width: 1 } });
      badge(s, `Paso ${r.yr}-${j + 1}`, j + 1, x + 0.25, y + 0.25, 0.4, r.color);
      txt(s, `Fecha ${r.yr}-${j + 1}`, fecha, { x: x + 0.8, y: y + 0.25, w: cw - 1.0, h: 0.4, fontSize: 13, bold: true, color: r.color, fontFace: "Courier New", valign: "middle" });
      txt(s, `Tipo ${r.yr}-${j + 1}`, tipo, { x: x + 0.25, y: y + 0.8, w: cw - 0.5, h: 0.35, fontSize: 15, bold: true });
      txt(s, `Nota ${r.yr}-${j + 1}`, "Nota No. " + no, { x: x + 0.25, y: y + 1.17, w: cw - 0.5, h: 0.3, fontSize: 12, color: MUTED });
      pill(s, `Detalle ${r.yr}-${j + 1}`, det, x + 0.25, y + 1.6, cw - 0.5, r.color, 11, 0.34);
    });
    s.addShape(pres.shapes.RIGHT_ARROW, { objectName: "Flecha " + r.yr, x: RX + cw + 0.14, y: r.y + 0.75 + ch / 2 - 0.16, w: 0.32, h: 0.32, fill: { color: r.color }, line: { color: r.color, transparency: 100 } });
  });
  s.addNotes("Cuatro comunicados oficiales. Cada año la SSF primero notifica el nombramiento de la visita ampliada y luego remite el informe de resultados preliminares. 2025: 9 observaciones. 2026: 7 nuevas (Anexo A) y 2 de seguimiento (Anexo B), de las cuales 1 no está superada y 1 está parcialmente superada.");
}

// ============ SLIDE 2 · OBSERVACIONES 2025 ============
{
  const s = addSlide("CONTENIDO", "2 · Observaciones 2025", "NOTA DS-SABAO-BCS-BC-17003 · VISITA CON CORTE AL 30/04/2025", "Observaciones 2025");
  legendDot(s, "Leyenda 2025", "Observación 2025", 9.35, 0.85, NAVY, 1.4);
  legendDot(s, "Leyenda abierta", "Sigue abierta en 2026", 10.9, 0.85, RED, 1.9);
  const O = [
    ["Falta de Política de Operaciones Vinculadas", "Consejo de Administración / Gerencia General", "Requerida por Art. 12 NRP-17"],
    ["Sin evidencia de firma de confidencialidad e inducción en actas", "Consejo de Administración / Secretaría", "Incumplimiento del Manual de Inducción"],
    ["Comités de Apoyo no integran personal ejecutivo", "Consejo de Administración / RRHH", "Incumplimiento Art. 3 literal h) NRP-17", "2026: NO SUPERADA", RED],
    ["Inconsistencia en pago de dietas (Reglamento vs. Asamblea General)", "Consejo de Administración / Asamblea", "Actualizar Reglamento y autorización de Asamblea"],
    ["Normativa interna desactualizada (más de 2 años sin revisión)", "Departamento de Procesos / Líderes de Área", "9 documentos vencidos (RIT, Auditoría, Riesgos, etc.)", "2026: PARCIALMENTE SUPERADA", AMBER_INK],
    ["Falta de soporte en expedientes de créditos de Ferias de Crédito", "Gerencia de Créditos / Riesgos", "Sin acuerdos de CA ni soporte de tasas"],
    ["Falta de información obligatoria en declaración jurada de personas relacionadas", "Directivos y Equipo Gerencial", "Incumplimiento Art. 9 NRP-75"],
    ["Omisiones en procedimientos del Manual de Compras de Bienes y Servicios", "Servicios Generales / Área Contratante", "Cotizaciones sin sellar y falta de cláusula penal"],
    ["Uso del Fondo de Educación en fines distintos a los legales", "Comité de Educación / Gerencia Financiera", "Fuera del Art. 59 Ley de Cooperativas", "2026: REINCIDE", AMBER_INK],
  ];
  const gap = 0.15, cw = (W - 2 * M - 2 * gap) / 3, ch = 1.75, y0 = 1.45;
  O.forEach(([t, r, c, st, col], i) => {
    const x = M + (i % 3) * (cw + gap), y = y0 + Math.floor(i / 3) * (ch + gap);
    card(s, "Observación 2025-" + (i + 1), x, y, cw, ch, TINT);
    badge(s, "Número 2025-" + (i + 1), i + 1, x + 0.2, y + 0.2, 0.42, col || NAVY);
    txt(s, "Título 2025-" + (i + 1), t, { x: x + 0.75, y: y + 0.13, w: cw - 0.95, h: 0.62, fontSize: 12.5, bold: true, valign: "middle" });
    const lines = [
      { text: "Responsable  ", options: { bold: true, color: MUTED } }, { text: r, options: { breakLine: true } },
      { text: "Criterio SSF  ", options: { bold: true, color: MUTED } }, { text: c, options: st ? { breakLine: true } : {} },
    ];
    if (st) lines.push({ text: "↳ " + st, options: { bold: true, color: col } });
    txt(s, "Detalle 2025-" + (i + 1), lines, { x: x + 0.2, y: y + 0.84, w: cw - 0.4, h: ch - 0.92, fontSize: 10.5, paraSpaceAfter: 1 });
  });
  s.addNotes("Nueve observaciones de la Nota 17003. En rojo y ámbar, las que siguen abiertas en la visita 2026: comités sin ejecutivos (no superada), normativa desactualizada (parcialmente superada) y Fondo de Educación (reincide).");
}

// ============ SLIDE 3 · OBSERVACIONES 2026 ============
{
  const s = addSlide("CONTENIDO", "3 · Observaciones 2026", "NOTA DS-SABAO-BCS-BC-22741 · VISITA CON CORTE AL 31/07/2026", "Observaciones 2026");
  legendDot(s, "Leyenda A", "Anexo A · nuevas", 9.0, 0.85, GREEN, 1.5);
  legendDot(s, "Leyenda B", "Anexo B · seguimiento", 10.75, 0.85, AMBER, 1.9);
  const A = [
    ["Falta de inducción sobre partidas relevantes de EEFF a nuevos miembros", "Gerencia Financiera / RRHH", "Art. 11 NRP-17"],
    ["Falta de Certificación de Lectura del Código de Conducta", "Junta de Vigilancia / Comités de Apoyo", "Código de Conducta"],
    ["Falta de formalización y aprobación de Políticas por el Consejo", "Consejo de Administración", "Políticas de Retribución y Control Interno"],
    ["Información de comités desactualizada en sitio web institucional", "Oficialía de Cumplimiento / Comunicaciones", "Arts. 27 y 28 NRP-17"],
    ["Incongruencia entre Estatutos y Código de Gobierno Corporativo", "Asesoría Legal / Consejo de Administración", "Restricción de cargos remunerados"],
    ["Reglamento del Fondo de Educación incluye conceptos no delimitados por Ley", "Comité de Educación / Gerencia Financiera", "Art. 5 del Reglamento"],
    ["Limitada transparencia a la Asamblea sobre el Fondo de Educación", "Consejo de Administración / Comité de Educación", "$100,774.28 no detallados"],
  ];
  const gap = 0.15, cw = (W - 2 * M - 3 * gap) / 4, ch = 1.6;
  txt(s, "Etiqueta Anexo A", "ANEXO A · OBSERVACIONES DETERMINADAS EN LA VISITA 2026", { x: M, y: 1.42, w: 8, h: 0.28, fontSize: 11, bold: true, color: GREEN, charSpacing: 2 });
  const cell = (i) => ({ x: M + (i % 4) * (cw + gap), y: 1.75 + Math.floor(i / 4) * (ch + gap) });
  A.forEach(([t, r, c], i) => {
    const { x, y } = cell(i);
    card(s, "Anexo A-" + (i + 1), x, y, cw, ch, TINT);
    badge(s, "Número A-" + (i + 1), i + 1, x + 0.18, y + 0.16, 0.36, GREEN);
    txt(s, "Responsable A-" + (i + 1), r, { x: x + 0.65, y: y + 0.13, w: cw - 0.8, h: 0.42, fontSize: 10.5, color: MUTED, valign: "middle" });
    txt(s, "Título A-" + (i + 1), t, { x: x + 0.18, y: y + 0.6, w: cw - 0.36, h: 0.65, fontSize: 11.5, bold: true });
    txt(s, "Criterio A-" + (i + 1), c, { x: x + 0.18, y: y + ch - 0.33, w: cw - 0.36, h: 0.24, fontSize: 10.5, bold: true, color: GREEN });
  });
  // Tarjeta de plazo Anexo A (ocupa la 8.ª celda)
  {
    const { x, y } = cell(7);
    card(s, "Plazo Anexo A", x, y, cw, ch, GREEN);
    txt(s, "Plazo A número", "10", { x: x + 0.2, y: y + 0.12, w: 0.9, h: 0.7, fontSize: 40, bold: true, color: WHITE, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Plazo A título", "días hábiles para descargos", { x: x + 1.1, y: y + 0.17, w: cw - 1.25, h: 0.6, fontSize: 12, bold: true, color: WHITE, valign: "middle" });
    txt(s, "Plazo A detalle", "Buzón DESCARGOS_OBSERVACIONES\nVence aprox. 23/10/2026", { x: x + 0.2, y: y + 0.88, w: cw - 0.35, h: 0.5, fontSize: 10.5, color: WHITE });
  }
  // Anexo B
  const yB = 5.6;
  txt(s, "Etiqueta Anexo B", "ANEXO B · SEGUIMIENTO A OBSERVACIONES DE VISITAS ANTERIORES", { x: M, y: yB - 0.32, w: 8, h: 0.26, fontSize: 11, bold: true, color: AMBER_INK, charSpacing: 2 });
  const B = [
    ["Comités de Apoyo no integran personal ejecutivo", "Responsable: Consejo de Administración", "NO SUPERADA", RED],
    ["Normativa interna desactualizada", "Depto. de Procesos · RIT, Comité de Auditoría, Compras", "PARCIALMENTE SUPERADA", AMBER],
  ];
  const bw = (W - 2 * M - cw - 2 * gap) / 2, bh = 1.25;
  B.forEach(([t, r, st, col], i) => {
    const x = M + i * (bw + gap);
    card(s, "Anexo B-" + (i + 8), x, yB, bw, bh, col, { transparency: 88 });
    badge(s, "Número B-" + (i + 8), i + 8, x + 0.2, yB + 0.14, 0.4, col);
    pill(s, "Estado B-" + (i + 8), st, x + 0.75, yB + 0.19, st.length > 12 ? 2.4 : 1.45, col, 10.5);
    txt(s, "Título B-" + (i + 8), t, { x: x + 0.2, y: yB + 0.62, w: bw - 0.4, h: 0.3, fontSize: 12.5, bold: true });
    txt(s, "Responsable B-" + (i + 8), r, { x: x + 0.2, y: yB + 0.93, w: bw - 0.4, h: 0.25, fontSize: 10.5, color: MUTED });
  });
  {
    const x = M + 2 * (bw + gap);
    card(s, "Plazo Anexo B", x, yB, cw, bh, AMBER_INK);
    txt(s, "Plazo B número", "15", { x: x + 0.2, y: yB + 0.02, w: 0.9, h: 0.7, fontSize: 40, bold: true, color: WHITE, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Plazo B título", "días hábiles para plan de solución", { x: x + 1.1, y: yB + 0.07, w: cw - 1.25, h: 0.6, fontSize: 12, bold: true, color: WHITE, valign: "middle" });
    txt(s, "Plazo B detalle", "Buzón PLANES_SOLUC_VISITASUPERVIS\nVence aprox. 30/10/2026", { x: x + 0.2, y: yB + 0.75, w: cw - 0.3, h: 0.45, fontSize: 10.5, color: WHITE });
  }
  s.addNotes("Anexo A: siete observaciones nuevas; plazo de 10 días hábiles desde el día hábil siguiente a la recepción para remitir pruebas de descargo, previo conocimiento del Consejo de Administración. Anexo B: dos observaciones anteriores siguen abiertas; plan de solución en 15 días hábiles con acciones, responsable, plazo por acción y certificación del punto de acta del Consejo. Fechas aproximadas sin asuetos; validar con el Sistema de Control de Envíos.");
}

// ============ SLIDE 4 · LÍNEA DE TIEMPO Y EVOLUCIÓN ============
{
  const s = addSlide("CONTENIDO", "4 · Línea de tiempo y evolución", "LÍNEA DE TIEMPO COMPARATIVA Y EVOLUCIÓN DEL CUMPLIMIENTO", "De 2025 a 2026");
  legendDot(s, "Leyenda 2025", "2025 · Visita 1", 9.6, 0.85, NAVY, 1.3);
  legendDot(s, "Leyenda 2026", "2026 · Visita 2", 11.1, 0.85, GREEN, 1.3);
  const H = [
    ["09 ABR 2025", "Nombramiento SSF", "Nota 9337 · corte 30/04", NAVY],
    ["19 MAY 2025", "Inicio de visita en sitio", "Supervisión ampliada", NAVY],
    ["05 NOV 2025", "Informe preliminar", "Nota 17003 · 9 observaciones", NAVY],
    ["31 JUL 2026", "Nombramiento SSF", "Nota 18432 · corte 31/07", GREEN],
    ["21 AGO 2026", "Inicio de visita ampliada", "Supervisión ampliada", GREEN],
    ["09 OCT 2026", "Informe preliminar", "Nota 22741 · 7 nuevas + 2", GREEN],
    ["HOY", "Plazos de respuesta", "Descargos 10 d.h. · Plan 15 d.h.", GREEN, true],
  ];
  const axisY = 3.0, x0 = M + 0.9, x1 = W - M - 0.9, step = (x1 - x0) / (H.length - 1);
  const xs = H.map((_, i) => x0 + i * step);
  const mid = (xs[2] + xs[3]) / 2;
  // Eje: tramo 2025, pausa de 8 meses, tramo 2026
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Eje 2025", x: M, y: axisY - 0.06, w: xs[2] - M + 0.15, h: 0.12, rectRadius: 0.06, fill: { color: NAVY }, line: { color: NAVY, transparency: 100 } });
  s.addShape(pres.shapes.LINE, { objectName: "Pausa entre visitas", x: xs[2] + 0.25, y: axisY, w: xs[3] - xs[2] - 0.5, h: 0, line: { color: MUTED, width: 1.5, dashType: "dash" } });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Eje 2026", x: xs[3] - 0.15, y: axisY - 0.06, w: W - M - xs[3] + 0.15, h: 0.12, rectRadius: 0.06, fill: { color: GREEN }, line: { color: GREEN, transparency: 100 } });
  txt(s, "Pausa texto", "8 meses", { x: mid - 0.6, y: axisY + 0.09, w: 1.2, h: 0.2, fontSize: 10, italic: true, color: MUTED, align: "center" });
  const cw = 2.75, chh = 1.12;
  H.forEach(([d, t, sub, col, now], i) => {
    const cx = xs[i], above = i % 2 === 0;
    const y = above ? axisY - 0.32 - chh : axisY + 0.32;
    const x = Math.min(Math.max(cx - cw / 2, M), W - M - cw);
    s.addShape(pres.shapes.LINE, { objectName: "Conector " + (i + 1), x: cx, y: above ? y + chh : axisY, w: 0, h: above ? axisY - (y + chh) : y - axisY, line: { color: col, width: 1.25 } });
    if (now) s.addShape(pres.shapes.OVAL, { objectName: "Halo hoy", x: cx - 0.24, y: axisY - 0.24, w: 0.48, h: 0.48, fill: { color: col, transparency: 75 }, line: { color: col, transparency: 100 } });
    s.addShape(pres.shapes.OVAL, { objectName: "Hito " + (i + 1), x: cx - 0.13, y: axisY - 0.13, w: 0.26, h: 0.26, fill: { color: now ? col : WHITE }, line: { color: col, width: 3 } });
    card(s, "Tarjeta hito " + (i + 1), x, y, cw, chh, now ? col : TINT);
    const fg = now ? WHITE : INK;
    txt(s, "Fecha hito " + (i + 1), d, { x: x + 0.18, y: y + 0.12, w: cw - 0.36, h: 0.26, fontSize: 12, bold: true, color: now ? WHITE : col, fontFace: "Courier New" });
    txt(s, "Texto hito " + (i + 1), t, { x: x + 0.18, y: y + 0.4, w: cw - 0.36, h: 0.3, fontSize: 13.5, bold: true, color: fg });
    txt(s, "Detalle hito " + (i + 1), sub, { x: x + 0.18, y: y + 0.73, w: cw - 0.36, h: 0.28, fontSize: 11, color: now ? WHITE : MUTED });
  });
  // Evolución
  const yE = 4.95;
  txt(s, "Evolución título", "EVOLUCIÓN · LO QUE PERSISTE DE 2025 A 2026", { x: M, y: yE - 0.02, w: 8, h: 0.28, fontSize: 11, bold: true, color: MUTED, charSpacing: 2 });
  const gap = 0.2, ew = (W - 2 * M - 3 * gap) / 4, eh = 1.6, ey = yE + 0.35;
  // Hallazgos críticos 2025
  card(s, "Hallazgos críticos 2025", M, ey, ew, eh, NAVY);
  txt(s, "Hallazgos título", "Hallazgos críticos 2025", { x: M + 0.2, y: ey + 0.14, w: ew - 0.4, h: 0.3, fontSize: 13, bold: true, color: WHITE });
  s.addText(["Sin Política de Operaciones Vinculadas", "Normativa interna desactualizada", "Uso inapropiado del Fondo de Educación"].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
    { objectName: "Hallazgos lista", isTextBox: true, x: M + 0.2, y: ey + 0.5, w: ew - 0.35, h: 1.0, fontSize: 11, color: WHITE, margin: 0, valign: "top", paraSpaceAfter: 2 });
  const E = [
    ["NO SUPERADA", RED, "Comités de Apoyo sin personal ejecutivo", "2025 obs. 3 → 2026 Anexo B, obs. 8"],
    ["PARCIALMENTE SUPERADA", AMBER, "Normativa interna desactualizada", "2025 obs. 5 → 2026 Anexo B, obs. 9 (RIT, Auditoría, Compras)"],
    ["REINCIDE", AMBER, "Gobernanza del Fondo de Educación", "2025 obs. 9 → 2026 Anexo A, obs. 6 y 7 ($100,774.28 sin detalle)"],
  ];
  E.forEach(([st, col, t, d], i) => {
    const x = M + (i + 1) * (ew + gap);
    card(s, "Evolución " + (i + 1), x, ey, ew, eh, col, { transparency: 88 });
    pill(s, "Evolución estado " + (i + 1), st, x + 0.2, ey + 0.18, st.length > 12 ? 2.3 : 1.3, col, 10.5);
    txt(s, "Evolución tema " + (i + 1), t, { x: x + 0.2, y: ey + 0.6, w: ew - 0.4, h: 0.45, fontSize: 13, bold: true });
    txt(s, "Evolución detalle " + (i + 1), d, { x: x + 0.2, y: ey + 1.05, w: ew - 0.4, h: 0.48, fontSize: 10.5, color: MUTED });
  });
  s.addNotes("Ambas visitas siguieron el mismo ciclo: nombramiento, inicio de visita e informe preliminar, con 8 meses entre el informe de 2025 y el nuevo nombramiento. Hoy corren los plazos: 10 días hábiles para descargos del Anexo A (aprox. 23/10/2026) y 15 días hábiles para el plan de solución del Anexo B (aprox. 30/10/2026). Persisten tres temas: comités sin ejecutivos (no superada), normativa interna (parcialmente superada) y Fondo de Educación (reincide). Paleta: azul marino #173A6E 2025, verde #1D7A4B 2026, rojo #B42318 no superada, ámbar #D99A1E parcial.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
