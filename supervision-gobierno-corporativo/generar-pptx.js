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
  const s = addSlide("CONTENIDO", "4 · Línea de tiempo y evolución", "LÍNEA DE TIEMPO COMPARATIVA Y EVOLUCIÓN DEL CUMPLIMIENTO", "Línea de tiempo 2025 vs. 2026");
  const gap = 0.3, bw = (W - 2 * M - gap) / 2, by = 1.45, bh = 3.35;
  const blocks = [
    { yr: "2025", sub: "Visita 1 · Nota 17003 · 9 observaciones", color: NAVY, hitos: [
        ["09 ABR 2025", "Nombramiento SSF · Nota 9337"],
        ["19 MAY 2025", "Inicio formal de la visita en sitio"],
        ["05 NOV 2025", "Informe preliminar · Nota 17003 · 9 obs."]],
      side: { t: "Hallazgos críticos 2025", items: [
        [["Ausencia de Política de ", {}], ["Operaciones Vinculadas", { bold: true }]],
        [["Normativa interna ", {}], ["desactualizada", { bold: true }], [" (9 documentos)", {}]],
        [["Uso inapropiado del ", {}], ["Fondo de Educación", { bold: true }]]] } },
    { yr: "2026", sub: "Visita 2 · Nota 22741 · 7 nuevas + 2 de 2025", color: GREEN, hitos: [
        ["31 JUL 2026", "Nombramiento SSF · Nota 18432"],
        ["21 AGO 2026", "Inicio de la visita ampliada"],
        ["09 OCT 2026", "Informe preliminar · Nota 22741 · 7 + 2"],
        ["HOY", "Plazos de respuesta en curso", true]],
      side: { t: "Resultado y plazos 2026", items: [
        [["Anexo A: ", { bold: true }], ["7 observaciones nuevas · descargos en 10 días hábiles (aprox. 23/10/2026)", {}]],
        [["Anexo B: ", { bold: true }], ["seguimiento a obs. 3 y 5 de 2025 · plan de solución en 15 días hábiles (aprox. 30/10/2026)", {}]]] } },
  ];
  blocks.forEach((b, k) => {
    const bx = M + k * (bw + gap);
    card(s, "Bloque " + b.yr, bx, by, bw, bh, TINT);
    txt(s, "Año bloque " + b.yr, b.yr, { x: bx + 0.25, y: by + 0.15, w: 1.2, h: 0.5, fontSize: 26, bold: true, color: b.color, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Subtítulo bloque " + b.yr, b.sub, { x: bx + 1.4, y: by + 0.15, w: bw - 1.6, h: 0.5, fontSize: 12, bold: true, color: MUTED, valign: "middle" });
    // Hitos verticales
    const tx = bx + 0.25, tw = 2.85, step = b.hitos.length === 3 ? 0.8 : 0.6, hy = by + 0.88;
    s.addShape(pres.shapes.LINE, { objectName: "Eje " + b.yr, x: tx + 0.11, y: hy + 0.11, w: 0, h: step * (b.hitos.length - 1), line: { color: b.color, width: 2.5 } });
    b.hitos.forEach(([d, t, now], i) => {
      const y = hy + i * step;
      if (now) s.addShape(pres.shapes.OVAL, { objectName: "Halo hoy", x: tx - 0.07, y: y - 0.07, w: 0.36, h: 0.36, fill: { color: b.color, transparency: 70 }, line: { color: b.color, transparency: 100 } });
      s.addShape(pres.shapes.OVAL, { objectName: `Hito ${b.yr}-${i + 1}`, x: tx, y, w: 0.22, h: 0.22, fill: { color: now ? b.color : WHITE }, line: { color: b.color, width: 2.5 } });
      txt(s, `Fecha ${b.yr}-${i + 1}`, d, { x: tx + 0.4, y: y - 0.03, w: tw, h: 0.24, fontSize: 11, bold: true, color: b.color, fontFace: "Courier New" });
      txt(s, `Hito ${b.yr}-${i + 1} texto`, t, { x: tx + 0.4, y: y + 0.21, w: tw, h: 0.4, fontSize: 11.5, bold: !!now });
    });
    // Detalle del año
    const sx = bx + 3.6, sw = bw - 3.6 - 0.2;
    card(s, "Detalle " + b.yr, sx, by + 0.85, sw, bh - 1.05, WHITE);
    txt(s, "Detalle título " + b.yr, b.side.t, { x: sx + 0.18, y: by + 1.0, w: sw - 0.36, h: 0.3, fontSize: 12.5, bold: true, color: b.color });
    const runs = [];
    b.side.items.forEach((it, i) => it.forEach(([t, o], j) => runs.push({ text: t, options: { ...o, ...(j === 0 ? { bullet: true } : {}), ...(j === it.length - 1 && i < b.side.items.length - 1 ? { breakLine: true } : {}) } })));
    s.addText(runs, { objectName: "Detalle lista " + b.yr, isTextBox: true, x: sx + 0.18, y: by + 1.38, w: sw - 0.3, h: bh - 1.65, fontSize: 11, color: INK, margin: 0, valign: "top", paraSpaceAfter: 4 });
  });
  // Seguimiento / evolución
  const yE = 5.0;
  txt(s, "Evolución título", "EVOLUCIÓN · SEGUIMIENTO A OBSERVACIONES 2025", { x: M, y: yE, w: 8, h: 0.26, fontSize: 11, bold: true, color: MUTED, charSpacing: 2 });
  const E = [
    ["NO SUPERADA", RED, "Comités de Apoyo sin personal ejecutivo", "Obs. 3 de 2025 · seguimiento en Anexo B 2026"],
    ["PARCIALMENTE SUPERADA", AMBER, "Normativa interna desactualizada", "Obs. 5 de 2025 · seguimiento en Anexo B 2026 (RIT, Auditoría, Compras)"],
    ["MISMO TEMA · NUEVAS", GREEN, "Fondo de Educación", "No es seguimiento: obs. 6 y 7 de 2026 son nuevas (Anexo A), mismo tema de la obs. 9 de 2025"],
  ];
  const eg = 0.2, ew = (W - 2 * M - 2 * eg) / 3, eh = 1.5, ey = yE + 0.35;
  E.forEach(([st, col, t, d], i) => {
    const x = M + i * (ew + eg);
    card(s, "Evolución " + (i + 1), x, ey, ew, eh, col, { transparency: 88 });
    pill(s, "Evolución estado " + (i + 1), st, x + 0.2, ey + 0.17, st.length > 12 ? 2.3 : 1.3, col, 10.5);
    txt(s, "Evolución tema " + (i + 1), t, { x: x + 0.2, y: ey + 0.6, w: ew - 0.4, h: 0.3, fontSize: 13, bold: true });
    txt(s, "Evolución detalle " + (i + 1), d, { x: x + 0.2, y: ey + 0.93, w: ew - 0.4, h: 0.48, fontSize: 10.5, color: MUTED });
  });
  s.addNotes("Bloque 2025: nombramiento (09/04/2025), inicio de visita (19/05/2025) e informe preliminar con 9 observaciones (05/11/2025); hallazgos críticos: Política de Operaciones Vinculadas, normativa desactualizada y Fondo de Educación. Bloque 2026: nombramiento (31/07/2026), inicio (21/08/2026) e informe preliminar (09/10/2026) con 7 observaciones nuevas (Anexo A) y seguimiento a 2 observaciones de 2025 (Anexo B). Plazos: 10 días hábiles para descargos y 15 días hábiles para plan de solución. Seguimiento: obs. 3 de 2025 no superada; obs. 5 de 2025 parcialmente superada. El Fondo de Educación vuelve a aparecer como tema, pero con observaciones nuevas, no como seguimiento.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
