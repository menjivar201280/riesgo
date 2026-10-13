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
  const k = [["4", "notas oficiales recibidas"], ["9", "observaciones en 2025"], ["7", "observaciones nuevas en 2026"], ["2", "de 2025 en seguimiento: 1 no superada, 1 parcial"]];
  k.forEach(([n, l], i) => {
    const x = M + (i % 2) * 1.75, y = 3.45 + Math.floor(i / 2) * 1.55;
    txt(s, "KPI valor " + (i + 1), n, { x, y, w: 1.6, h: 0.65, fontSize: 34, bold: true, color: WHITE, fontFace: THEME.headFontFace, valign: "bottom" });
    txt(s, "KPI texto " + (i + 1), l, { x, y: y + 0.7, w: 1.6, h: 0.6, fontSize: 12, color: TINT });
  });
  // Filas por año: nombramiento → informe
  const RX = PANEL + 0.5, RW = W - M - RX;
  const rows = [
    { y: 0.55, yr: "2025", sub: "Visita 1 · corte 30/04/2025", color: NAVY, a: ["09/04/2025", "Nombramiento de visita ampliada", "DS-SABAO-BCS-BC-9337", "Corte 30/04/2025 · inicio 19/05/2025"], b: ["05/11/2025", "Informe de resultados preliminares", "DS-SABAO-BCS-BC-17003", "9 observaciones"] },
    { y: 3.75, yr: "2026", sub: "Visita 2 · corte 31/07/2026", color: GREEN, a: ["31/07/2026", "Nombramiento de visita ampliada", "SABAO-BCS-BC-18432", "Corte 31/07/2026 · inicio 21/08/2026"], b: ["09/10/2026", "Informe de resultados preliminares", "DS-SABAO-BCS-BC-22741", "7 nuevas + seguimiento a 2 de 2025"] },
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
  s.addNotes("Cuatro comunicados oficiales. Cada año la SSF primero notifica el nombramiento de la visita ampliada y luego remite el informe de resultados preliminares. 2025: 9 observaciones. 2026: 7 observaciones nuevas (Anexo A) y seguimiento a 2 de las 9 observaciones de 2025 (Anexo B): la obs. 3 no superada y la obs. 5 parcialmente superada.");
}

// ============ SLIDE 2 · OBSERVACIONES 2025 ============
{
  const s = addSlide("CONTENIDO", "2 · Observaciones 2025", "NOTA DS-SABAO-BCS-BC-17003 · VISITA CON CORTE AL 30/04/2025", "Observaciones 2025");
  legendDot(s, "Leyenda 2025", "Observación 2025", 9.35, 0.85, NAVY, 1.4);
  legendDot(s, "Leyenda abierta", "En seguimiento SSF 2026", 10.85, 0.85, RED, 1.95);
  const O = [
    ["Falta de Política de Operaciones Vinculadas", "Consejo de Administración / Gerencia General", "Requerida por Art. 12 NRP-17"],
    ["Sin evidencia de firma de confidencialidad e inducción en actas", "Consejo de Administración / Secretaría", "Incumplimiento del Manual de Inducción"],
    ["Comités de Apoyo no integran personal ejecutivo", "Consejo de Administración / RRHH", "Incumplimiento Art. 3 literal h) NRP-17", "Seguimiento 2026: NO SUPERADA", RED],
    ["Inconsistencia en pago de dietas (Reglamento vs. Asamblea General)", "Consejo de Administración / Asamblea", "Actualizar Reglamento y autorización de Asamblea"],
    ["Normativa interna desactualizada (más de 2 años sin revisión)", "Departamento de Procesos / Líderes de Área", "9 documentos vencidos (RIT, Auditoría, Riesgos, etc.)", "Seguimiento 2026: PARCIALMENTE SUPERADA", AMBER_INK],
    ["Falta de soporte en expedientes de créditos de Ferias de Crédito", "Gerencia de Créditos / Riesgos", "Sin acuerdos de CA ni soporte de tasas"],
    ["Falta de información obligatoria en declaración jurada de personas relacionadas", "Directivos y Equipo Gerencial", "Incumplimiento Art. 9 NRP-75"],
    ["Omisiones en procedimientos del Manual de Compras de Bienes y Servicios", "Servicios Generales / Área Contratante", "Cotizaciones sin sellar y falta de cláusula penal"],
    ["Uso del Fondo de Educación en fines distintos a los legales", "Comité de Educación / Gerencia Financiera", "Fuera del Art. 59 Ley de Cooperativas"],
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
  s.addNotes("Nueve observaciones de la Nota 17003. En rojo y ámbar, las 2 observaciones a las que la SSF dio seguimiento en la visita 2026 (Anexo B): obs. 3 comités sin ejecutivos (no superada) y obs. 5 normativa desactualizada (parcialmente superada).");
}

// ============ SLIDE 3 · OBSERVACIONES 2026 ============
{
  const s = addSlide("CONTENIDO", "3 · Observaciones 2026", "NOTA DS-SABAO-BCS-BC-22741 · VISITA CON CORTE AL 31/07/2026", "Observaciones 2026");
  legendDot(s, "Leyenda A", "Anexo A · nuevas", 8.8, 0.85, GREEN, 1.5);
  legendDot(s, "Leyenda B", "Anexo B · seguimiento 2025", 10.55, 0.85, NAVY, 2.1);
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
  txt(s, "Etiqueta Anexo B", "ANEXO B · SEGUIMIENTO A 2 DE LAS 9 OBSERVACIONES DE 2025 (NO SON NUEVAS)", { x: M, y: yB - 0.32, w: 8, h: 0.26, fontSize: 11, bold: true, color: AMBER_INK, charSpacing: 2 });
  const B = [
    ["Comités de Apoyo no integran personal ejecutivo", "Responsable: Consejo de Administración", "NO SUPERADA", RED, 3],
    ["Normativa interna desactualizada", "Depto. de Procesos · RIT, Comité de Auditoría, Compras", "PARCIALMENTE SUPERADA", AMBER, 5],
  ];
  const bw = (W - 2 * M - cw - 2 * gap) / 2, bh = 1.25;
  B.forEach(([t, r, st, col, n25], i) => {
    const x = M + i * (bw + gap);
    card(s, "Anexo B-" + (i + 8), x, yB, bw, bh, col, { transparency: 88 });
    pill(s, "Origen B-" + n25, "OBS. " + n25 + " DE 2025", x + 0.2, yB + 0.19, 1.45, NAVY, 10.5);
    pill(s, "Estado B-" + n25, st, x + 1.8, yB + 0.19, st.length > 12 ? 2.4 : 1.45, col, 10.5);
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
  s.addNotes("Anexo A: siete observaciones nuevas; plazo de 10 días hábiles desde el día hábil siguiente a la recepción para remitir pruebas de descargo, previo conocimiento del Consejo de Administración. Anexo B: no son observaciones nuevas, es el seguimiento a las obs. 3 y 5 de 2025, que siguen abiertas; plan de solución en 15 días hábiles con acciones, responsable, plazo por acción y certificación del punto de acta del Consejo. Fechas aproximadas sin asuetos; validar con el Sistema de Control de Envíos.");
}

// ============ SLIDE 4 · LÍNEA DE TIEMPO Y EVOLUCIÓN ============
{
  const s = addSlide("CONTENIDO", "4 · Línea de tiempo y evolución", "LÍNEA DE TIEMPO COMPARATIVA Y EVOLUCIÓN DEL CUMPLIMIENTO", "Línea de tiempo: de hoy a 2025");
  // Hitos de lo más reciente (arriba) a lo más antiguo (abajo)
  const H = [
    ["HOY", "Plazos de respuesta en curso", "Descargos 10 días hábiles (≈23/10) · Plan de solución 15 días hábiles (≈30/10)", GREEN, true],
    ["09 OCT 2026", "Informe preliminar · Nota 22741", "7 nuevas (Anexo A) + seguimiento a obs. 3 y 5 de 2025 (Anexo B)", GREEN],
    ["21 AGO 2026", "Inicio de la visita de supervisión ampliada", "", GREEN],
    ["31 JUL 2026", "Nombramiento SSF · Nota 18432", "Visita con corte al 31/07/2026", GREEN],
    ["05 NOV 2025", "Informe preliminar · Nota 17003", "9 observaciones", NAVY],
    ["19 MAY 2025", "Inicio formal de la visita en sitio", "", NAVY],
    ["09 ABR 2025", "Nombramiento SSF · Nota 9337", "Visita con corte al 30/04/2025", NAVY],
  ];
  const ax = M + 1.2, y0 = 1.6, step = 0.74, gapYear = 0.22, tx = ax + 0.35, tw = 5.0;
  const ys = H.map((_, i) => y0 + i * step + (i >= 4 ? gapYear : 0));
  // Eje: 2026 verde, pausa punteada, 2025 azul
  s.addShape(pres.shapes.LINE, { objectName: "Eje 2026", x: ax, y: ys[0], w: 0, h: ys[3] - ys[0], line: { color: GREEN, width: 3 } });
  s.addShape(pres.shapes.LINE, { objectName: "Pausa entre visitas", x: ax, y: ys[3] + 0.2, w: 0, h: ys[4] - ys[3] - 0.4, line: { color: MUTED, width: 1.5, dashType: "dash" } });
  s.addShape(pres.shapes.LINE, { objectName: "Eje 2025", x: ax, y: ys[4], w: 0, h: ys[6] - ys[4], line: { color: NAVY, width: 3 } });
  pill(s, "Año 2026", "2026", M, ys[0] - 0.04, 0.85, GREEN, 11);
  pill(s, "Año 2025", "2025", M, ys[4] - 0.04, 0.85, NAVY, 11);
  txt(s, "Pausa texto", "8 meses", { x: M - 0.05, y: (ys[3] + ys[4]) / 2 - 0.1, w: 1.0, h: 0.22, fontSize: 10, italic: true, color: MUTED, align: "center" });
  H.forEach(([d, t, sub, col, now], i) => {
    const y = ys[i];
    if (now) {
      card(s, "Fondo hoy", tx - 0.12, y - 0.1, tw + 0.2, step - 0.06, col, { transparency: 85 });
      s.addShape(pres.shapes.OVAL, { objectName: "Halo hoy", x: ax - 0.2, y: y - 0.04, w: 0.4, h: 0.4, fill: { color: col, transparency: 70 }, line: { color: col, transparency: 100 } });
    }
    s.addShape(pres.shapes.OVAL, { objectName: "Hito " + (i + 1), x: ax - 0.12, y: y + 0.04, w: 0.24, h: 0.24, fill: { color: now ? col : WHITE }, line: { color: col, width: 2.5 } });
    txt(s, "Hito fecha " + (i + 1), [{ text: d + "   ", options: { bold: true, color: col, fontFace: "Courier New", fontSize: 11.5 } }, { text: t, options: { bold: true, fontSize: 12.5 } }], { x: tx, y: y, w: tw, h: 0.3, valign: "middle" });
    if (sub) txt(s, "Hito detalle " + (i + 1), sub, { x: tx, y: y + 0.3, w: tw, h: 0.25, fontSize: 10.5, color: MUTED });
  });
  // Columna derecha: detalle por año (2026 arriba) y seguimiento
  const rx = 7.35, rw = W - M - rx;
  card(s, "Detalle 2026", rx, 1.45, rw, 1.42, GREEN, { transparency: 88 });
  txt(s, "Detalle 2026 título", "Resultado 2026 · Nota 22741", { x: rx + 0.2, y: 1.57, w: rw - 0.4, h: 0.28, fontSize: 13, bold: true, color: GREEN });
  txt(s, "Detalle 2026 texto", [
    { text: "Anexo A: ", options: { bold: true } }, { text: "7 observaciones nuevas · descargos en 10 días hábiles", options: { breakLine: true } },
    { text: "Anexo B: ", options: { bold: true } }, { text: "seguimiento a obs. 3 y 5 de 2025 · plan de solución en 15 días hábiles" },
  ], { x: rx + 0.2, y: 1.92, w: rw - 0.4, h: 0.88, fontSize: 11.5, paraSpaceAfter: 3 });
  card(s, "Detalle 2025", rx, 3.02, rw, 1.25, NAVY, { transparency: 88 });
  txt(s, "Detalle 2025 título", "Hallazgos críticos 2025 · Nota 17003", { x: rx + 0.2, y: 3.14, w: rw - 0.4, h: 0.28, fontSize: 13, bold: true, color: NAVY });
  txt(s, "Detalle 2025 texto", [
    { text: "Ausencia de Política de Operaciones Vinculadas", options: { bullet: true, breakLine: true } },
    { text: "Normativa interna desactualizada (9 documentos)", options: { bullet: true, breakLine: true } },
    { text: "Uso inapropiado del Fondo de Educación", options: { bullet: true } },
  ], { x: rx + 0.2, y: 3.48, w: rw - 0.4, h: 0.75, fontSize: 11.5, paraSpaceAfter: 1 });
  txt(s, "Evolución título", "SEGUIMIENTO A OBSERVACIONES 2025", { x: rx, y: 4.42, w: rw, h: 0.26, fontSize: 11, bold: true, color: MUTED, charSpacing: 2 });
  const E = [
    ["NO SUPERADA", RED, "Comités de Apoyo sin ejecutivos", "Obs. 3 de 2025 · Anexo B"],
    ["PARCIALMENTE SUPERADA", AMBER, "Normativa interna desactualizada", "Obs. 5 de 2025 · Anexo B"],
    ["MISMO TEMA · NUEVAS", GREEN, "Fondo de Educación", "Obs. 6 y 7 nuevas · no es seguimiento"],
  ];
  E.forEach(([st, col, t, d], i) => {
    const y = 4.75 + i * 0.72;
    card(s, "Evolución " + (i + 1), rx, y, rw, 0.64, col, { transparency: 88 });
    pill(s, "Evolución estado " + (i + 1), st, rx + 0.15, y + 0.17, 2.2, col, 10);
    txt(s, "Evolución tema " + (i + 1), t, { x: rx + 2.5, y: y + 0.08, w: rw - 2.6, h: 0.26, fontSize: 12, bold: true });
    txt(s, "Evolución detalle " + (i + 1), d, { x: rx + 2.5, y: y + 0.34, w: rw - 2.6, h: 0.24, fontSize: 10.5, color: MUTED });
  });
  s.addNotes("La línea de tiempo se lee de arriba hacia abajo, de lo más reciente a lo más antiguo. Hoy corren los plazos: 10 días hábiles para descargos del Anexo A (aprox. 23/10/2026) y 15 días hábiles para el plan de solución del Anexo B (aprox. 30/10/2026). 2026: nombramiento 31/07, inicio 21/08 e informe preliminar 09/10 con 7 observaciones nuevas y seguimiento a las obs. 3 y 5 de 2025. 2025: nombramiento 09/04, inicio 19/05 e informe preliminar 05/11 con 9 observaciones. Seguimiento: obs. 3 no superada; obs. 5 parcialmente superada. El Fondo de Educación vuelve como tema, pero con observaciones nuevas.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
