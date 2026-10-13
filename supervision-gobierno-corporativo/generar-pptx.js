const path = require("path");
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

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Supervisiones de Gobierno Corporativo SSF a COMEDICA de R.L.";
pres.subject = "Resultados de visitas de supervisión 2025 y 2026";
pres.author = "COMEDICA de R.L.";
const C = pres.SchemeColor;
const NAVY = C.accent1, GREEN = C.accent2, RED = C.accent3, AMBER = C.accent4, MUTED = C.accent5, AMBER_INK = C.accent6;
const W = 13.333, M = 0.6;

// ---------- Layouts ----------
pres.defineSlideMaster({
  title: "PORTADA",
  background: { color: C.text2 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.1, w: 11.5, h: 1.9, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "Título" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.15, w: 11.5, h: 1.2, fontSize: 20, color: C.background2, valign: "top" }, text: "Subtítulo" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: C.background1 },
  margin: [0.5, M, 0.6, M],
  objects: [
    { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.35, w: 9, h: 0.35, fontSize: 12, bold: true, color: MUTED, charSpacing: 2, margin: 0 }, text: "SECCIÓN" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.7, w: W - 2 * M, h: 0.75, fontFace: THEME.headFontFace, fontSize: 30, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "Título" } },
    { text: { text: "SSF · Gobierno Corporativo (NRP-17) · COMEDICA de R.L.", options: { x: M, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: MUTED, margin: 0 } } },
  ],
  slideNumber: { x: W - M - 0.6, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: MUTED, align: "right" },
});

const content = (section, eyebrow, title) => {
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: section });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(title, { placeholder: "title" });
  return s;
};
const box = (s, name, x, y, w, h, fill, opts = {}) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name, x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, ...opts });
const txt = (s, name, text, o) => s.addText(text, { objectName: name, isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top", ...o });
const pill = (s, name, text, x, y, w, color) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name + " fondo", x, y, w, h: 0.34, rectRadius: 0.17, fill: { color }, line: { color } });
  s.addText(text, { objectName: name, isTextBox: true, x, y, w, h: 0.34, margin: 0, align: "center", valign: "middle", fontSize: 11, bold: true, color: C.background1, charSpacing: 1 });
};

// ---------- Tabla de observaciones ----------
const HEAD = ["#", "Nombre de la observación", "Responsable / Área implicada", "Estado actual / Criterio SSF"];
const COLW = [0.55, 5.2, 3.2, 3.18];
function obsTable(s, name, rows, startNo, headColor, y, opts = {}) {
  const head = HEAD.map(t => ({ text: t, options: { bold: true, color: C.background1, fill: { color: headColor }, fontSize: 12 } }));
  const body = rows.map((r, i) => {
    const zebra = i % 2 ? C.background1 : C.background2;
    const st = r.status;
    const estado = st
      ? [{ text: st.label, options: { bold: true, color: st.color, breakLine: true } }, { text: r.estado, options: { color: C.text1 } }]
      : r.estado;
    return [
      { text: String(startNo + i), options: { bold: true, color: MUTED, fill: { color: zebra }, align: "center" } },
      { text: r.link ? [{ text: r.obs, options: { bold: true, breakLine: true } }, { text: "↳ " + r.link.text, options: { bold: true, color: r.link.color, fontSize: 11 } }] : r.obs, options: { bold: !r.link, fill: { color: zebra } } },
      { text: r.resp, options: { fill: { color: zebra } } },
      { text: estado, options: { fill: { color: zebra } } },
    ];
  });
  s.addTable([head, ...body], {
    objectName: name, x: M, y, w: W - 2 * M, colW: COLW, fontFace: THEME.bodyFontFace, fontSize: 14, color: C.text1,
    valign: "middle", margin: [0.1, 0.12, 0.1, 0.12], border: { type: "solid", pt: 0.75, color: "D9E0EA" }, autoPage: false, ...opts,
  });
}

// ============ 1. PORTADA ============
pres.addSection({ title: "Portada" });
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Portada" });
  s.addText("Supervisiones de Gobierno Corporativo SSF a COMEDICA de R.L.", { placeholder: "title" });
  s.addText("Resultados de las visitas de supervisión ampliada 2025 y 2026 · Notas oficiales, observaciones y evolución del cumplimiento", { placeholder: "body" });
  txt(s, "Etiqueta", "SUPERINTENDENCIA DEL SISTEMA FINANCIERO · NRP-17", { x: M, y: 1.45, w: 9, h: 0.4, fontSize: 13, bold: true, color: C.background2, charSpacing: 3 });
  pill(s, "Etiqueta 2025", "2025 · NOTA 17003", M, 5.75, 2.3, NAVY);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Contorno 2025", x: M, y: 5.75, w: 2.3, h: 0.34, rectRadius: 0.17, fill: { type: "none" }, line: { color: C.background2, width: 1 } });
  pill(s, "Etiqueta 2026", "2026 · NOTA 22741", M + 2.5, 5.75, 2.3, GREEN);
  s.addNotes("Presentación de resultados de las supervisiones de Gobierno Corporativo realizadas por la SSF a COMEDICA de R.L. en 2025 y 2026. Estructura: 1) notas oficiales, 2) observaciones 2025, 3) observaciones 2026, 4) línea de tiempo y evolución.");
}

// ============ 2. SLIDER 1 · RESUMEN ============
pres.addSection({ title: "1 · Resumen de notas SSF" });
{
  const s = content("1 · Resumen de notas SSF", "1 · RESUMEN EJECUTIVO", "Notas de notificación SSF recibidas 2025 - 2026");
  // KPIs
  const kpis = [["4", "notas oficiales SSF", C.text1], ["9", "observaciones 2025", NAVY], ["7 + 2", "nuevas + seguimiento 2026", GREEN], ["1 / 1", "no superada / parcial", RED]];
  const kw = (W - 2 * M - 3 * 0.3) / 4;
  kpis.forEach(([n, l, c], i) => {
    const x = M + i * (kw + 0.3);
    box(s, "KPI " + l, x, 1.6, kw, 1.0, C.background2);
    txt(s, "KPI valor " + l, n, { x: x + 0.2, y: 1.68, w: kw - 0.4, h: 0.5, fontSize: 28, bold: true, color: c, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "KPI texto " + l, l, { x: x + 0.2, y: 2.18, w: kw - 0.4, h: 0.32, fontSize: 13, color: MUTED });
  });
  const cols = [
    { y: "2025", sub: "Visita 1 · corte 30/04/2025", color: NAVY, notes: [
      ["Nombramiento de visita ampliada", "Nota No. DS-SABAO-BCS-BC-9337", "09/04/2025", "Corte: 30/04/2025 · Inicio: 19/05/2025"],
      ["Informe de resultados preliminares", "Nota No. DS-SABAO-BCS-BC-17003", "05/11/2025", "9 observaciones"]] },
    { y: "2026", sub: "Visita 2 · corte 31/07/2026", color: GREEN, notes: [
      ["Nombramiento de visita ampliada", "Nota No. SABAO-BCS-BC-18432", "31/07/2026", "Corte: 31/07/2026 · Inicio: 21/08/2026"],
      ["Informe de resultados preliminares", "Nota No. DS-SABAO-BCS-BC-22741", "09/10/2026", "7 nuevas (Anexo A) · 2 de seguimiento (Anexo B)"]] },
  ];
  const cw = (W - 2 * M - 0.4) / 2;
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.4);
    box(s, "Encabezado " + c.y, x, 2.9, cw, 0.6, c.color);
    txt(s, "Año " + c.y, c.y, { x: x + 0.25, y: 2.9, w: 1.5, h: 0.6, fontSize: 22, bold: true, color: C.background1, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Subtítulo " + c.y, c.sub, { x: x + 1.6, y: 2.9, w: cw - 1.85, h: 0.6, fontSize: 13, bold: true, color: C.background1, align: "right", valign: "middle" });
    c.notes.forEach(([tipo, no, fecha, det], j) => {
      const y = 3.65 + j * 1.6;
      box(s, `Tarjeta ${c.y}-${j + 1}`, x, y, cw, 1.45, C.background1, { line: { color: "D9E0EA", width: 1 } });
      txt(s, `Tipo ${c.y}-${j + 1}`, tipo, { x: x + 0.25, y: y + 0.15, w: cw - 1.9, h: 0.4, fontSize: 16, bold: true });
      pill(s, `Fecha ${c.y}-${j + 1}`, fecha, x + cw - 1.55, y + 0.17, 1.3, c.color);
      txt(s, `Nota ${c.y}-${j + 1}`, no, { x: x + 0.25, y: y + 0.58, w: cw - 0.5, h: 0.3, fontSize: 13, color: MUTED, fontFace: "Courier New" });
      txt(s, `Detalle ${c.y}-${j + 1}`, det, { x: x + 0.25, y: y + 0.92, w: cw - 0.5, h: 0.35, fontSize: 14, bold: true, color: c.color });
    });
  });
  s.addNotes("Cuatro comunicados oficiales: en cada año la SSF primero notifica el nombramiento de la visita ampliada y luego remite el informe de resultados preliminares. 2025: 9 observaciones. 2026: 7 nuevas y 2 de seguimiento.");
}

// ============ 3-4. SLIDER 2 · OBS 2025 ============
const OBS25 = [
  { obs: "Falta de Política de Operaciones Vinculadas", resp: "Consejo de Administración / Gerencia General", estado: "Requerida por Art. 12 NRP-17" },
  { obs: "Sin evidencia de firma de confidencialidad e inducción en actas", resp: "Consejo de Administración / Secretaría", estado: "Incumplimiento del Manual de Inducción" },
  { obs: "Comités de Apoyo no integran personal ejecutivo", resp: "Consejo de Administración / RRHH", estado: "Incumplimiento Art. 3 literal h) NRP-17", link: { text: "2026: NO SUPERADA", color: RED } },
  { obs: "Inconsistencia en pago de dietas (Reglamento vs. Asamblea General)", resp: "Consejo de Administración / Asamblea", estado: "Requiere actualización del Reglamento y autorización de Asamblea" },
  { obs: "Normativa interna desactualizada (más de 2 años sin revisión)", resp: "Departamento de Procesos / Líderes de Área", estado: "9 documentos vencidos (RIT, Auditoría, Riesgos, etc.)", link: { text: "2026: PARCIALMENTE SUPERADA", color: AMBER_INK } },
  { obs: "Falta de soporte en expedientes de créditos de Ferias de Crédito", resp: "Gerencia de Créditos / Riesgos", estado: "Sin acuerdos de CA ni soporte de tasas" },
  { obs: "Falta de información obligatoria en declaración jurada de personas relacionadas", resp: "Directivos y Equipo Gerencial", estado: "Incumplimiento Art. 9 NRP-75" },
  { obs: "Omisiones en procedimientos del Manual de Compras de Bienes y Servicios", resp: "Servicios Generales / Área Contratante", estado: "Cotizaciones sin sellar y falta de cláusula penal" },
  { obs: "Uso del Fondo de Educación en fines distintos a los legales", resp: "Comité de Educación / Gerencia Financiera", estado: "Gastos socio-culturales/operativos (Art. 59 Ley de Cooperativas)", link: { text: "2026: REINCIDE (Anexo A, obs. 6 y 7)", color: AMBER_INK } },
];
pres.addSection({ title: "2 · Observaciones 2025" });
{
  const s = content("2 · Observaciones 2025", "2 · NOTA DS-SABAO-BCS-BC-17003 · AÑO 2025", "Detalle de observaciones 2025 (1 de 2)");
  obsTable(s, "Tabla observaciones 2025 1-5", OBS25.slice(0, 5), 1, NAVY, 1.65);
  s.addNotes("Observaciones 1 a 5 de la Nota 17003. Las observaciones 3 y 5 se volvieron a señalar en la visita 2026 (Anexo B).");
}
{
  const s = content("2 · Observaciones 2025", "2 · NOTA DS-SABAO-BCS-BC-17003 · AÑO 2025", "Detalle de observaciones 2025 (2 de 2)");
  obsTable(s, "Tabla observaciones 2025 6-9", OBS25.slice(5), 6, NAVY, 1.65);
  // Callout de persistencia
  const y = 5.55, cw = (W - 2 * M - 0.6) / 3;
  [["Obs. 3 · Comités sin ejecutivos", "NO SUPERADA en 2026", RED], ["Obs. 5 · Normativa desactualizada", "PARCIALMENTE SUPERADA en 2026", AMBER_INK], ["Obs. 9 · Fondo de Educación", "REINCIDE en 2026", AMBER_INK]].forEach(([a, b, c], i) => {
    const x = M + i * (cw + 0.3);
    box(s, "Persistencia " + (i + 1), x, y, cw, 1.05, C.background2);
    txt(s, "Persistencia título " + (i + 1), a, { x: x + 0.2, y: y + 0.15, w: cw - 0.4, h: 0.35, fontSize: 14, bold: true });
    txt(s, "Persistencia estado " + (i + 1), b, { x: x + 0.2, y: y + 0.55, w: cw - 0.4, h: 0.35, fontSize: 14, bold: true, color: c });
  });
  s.addNotes("Observaciones 6 a 9. Abajo, las tres observaciones de 2025 que siguen abiertas en la visita 2026.");
}

// ============ 5-6. SLIDER 3 · OBS 2026 ============
const PLAZO = { label: "DESCARGO · 10 DÍAS HÁBILES", color: GREEN };
const OBS26A = [
  { obs: "Falta de inducción sobre partidas relevantes de EEFF a nuevos miembros", resp: "Gerencia Financiera / RRHH", estado: "Art. 11 NRP-17" },
  { obs: "Falta de Certificación de Lectura del Código de Conducta", resp: "Junta de Vigilancia / Comités de Apoyo", estado: "Código de Conducta" },
  { obs: "Falta de formalización y aprobación de Políticas por el Consejo de Administración", resp: "Consejo de Administración", estado: "Políticas de Retribución y Control Interno" },
  { obs: "Información de comités desactualizada en sitio web institucional", resp: "Oficialía de Cumplimiento / Comunicaciones", estado: "Arts. 27 y 28 NRP-17" },
  { obs: "Incongruencia entre Estatutos y Código de Gobierno Corporativo", resp: "Asesoría Legal / Consejo de Administración", estado: "Restricción de cargos remunerados" },
  { obs: "Reglamento del Fondo de Educación incluye conceptos no delimitados por Ley", resp: "Comité de Educación / Gerencia Financiera", estado: "Art. 5 del Reglamento" },
  { obs: "Limitada transparencia en información presentada a Asamblea sobre Fondo de Educación", resp: "Consejo de Administración / Comité de Educación", estado: "$100,774.28 no detallados" },
].map(r => ({ ...r, status: PLAZO }));
pres.addSection({ title: "3 · Observaciones 2026" });
{
  const s = content("3 · Observaciones 2026", "3 · NOTA DS-SABAO-BCS-BC-22741 · ANEXO A", "Anexo A: observaciones nuevas de la visita 2026");
  obsTable(s, "Tabla Anexo A", OBS26A, 1, GREEN, 1.6, { fontSize: 13, margin: [0.06, 0.1, 0.06, 0.1] });
  s.addNotes("Siete observaciones nuevas. Para todas, el plazo es de 10 días hábiles contados desde el día hábil siguiente a la recepción de la nota, para remitir pruebas de descargo documentadas por el buzón DESCARGOS_OBSERVACIONES, previo conocimiento del Consejo de Administración.");
}
{
  const s = content("3 · Observaciones 2026", "3 · NOTA DS-SABAO-BCS-BC-22741 · ANEXO B", "Anexo B: seguimiento a observaciones anteriores");
  obsTable(s, "Tabla Anexo B", [
    { obs: "Comités de Apoyo no integran personal ejecutivo", resp: "Consejo de Administración", estado: "Plan de solución en 15 días hábiles", status: { label: "NO SUPERADA", color: RED } },
    { obs: "Normativa interna desactualizada (RIT, Comité de Auditoría, Compras)", resp: "Departamento de Procesos", estado: "Plan de solución en 15 días hábiles", status: { label: "PARCIALMENTE SUPERADA", color: AMBER_INK } },
  ], 8, GREEN, 1.65);
  const y = 3.85, cw = (W - 2 * M - 0.4) / 2;
  [
    { n: "10", t: "días hábiles · Anexo A", d: "Pruebas de descargo documentadas\nBuzón: DESCARGOS_OBSERVACIONES\nVence aprox. 23/10/2026", c: GREEN },
    { n: "15", t: "días hábiles · Anexo B", d: "Plan de solución: acciones, responsable y plazo por acción, con certificación del punto de acta del Consejo\nBuzón: PLANES_SOLUC_VISITASUPERVIS\nVence aprox. 30/10/2026", c: AMBER_INK },
  ].forEach((k, i) => {
    const x = M + i * (cw + 0.4);
    box(s, "Plazo " + k.n, x, y, cw, 2.65, C.background2);
    txt(s, "Plazo número " + k.n, k.n, { x: x + 0.3, y: y + 0.2, w: 1.4, h: 0.95, fontSize: 54, bold: true, color: k.c, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Plazo título " + k.n, k.t, { x: x + 1.75, y: y + 0.35, w: cw - 2.0, h: 0.65, fontSize: 18, bold: true, valign: "middle" });
    txt(s, "Plazo detalle " + k.n, k.d, { x: x + 0.3, y: y + 1.25, w: cw - 0.6, h: 1.3, fontSize: 13, color: C.text1, paraSpaceAfter: 4 });
  });
  txt(s, "Nota de fechas", "Fechas estimadas en días hábiles desde el 12/10/2026, sin asuetos. Validar con la fecha de recepción del Sistema de Control de Envíos.", { x: M, y: 6.6, w: W - 2 * M, h: 0.3, fontSize: 10, color: MUTED, italic: true });
  s.addNotes("Dos observaciones de visitas anteriores siguen abiertas: la integración de comités con personal ejecutivo (no superada) y la normativa interna desactualizada (parcialmente superada). Para ambas se exige un plan de solución en 15 días hábiles.");
}

// ============ 7-8. SLIDER 4 · LÍNEA DE TIEMPO ============
pres.addSection({ title: "4 · Línea de tiempo y evolución" });
{
  const s = content("4 · Línea de tiempo y evolución", "4 · LÍNEA DE TIEMPO COMPARATIVA", "Línea de tiempo 2025 vs. 2026");
  const lineW = 7.3;
  const blocks = [
    { y: 1.75, year: "2025", color: NAVY, hitos: [["09 ABR 2025", "Nombramiento SSF", "Nota 9337 · corte 30/04"], ["19 MAY 2025", "Inicio de visita en sitio", ""], ["05 NOV 2025", "Informe preliminar", "Nota 17003 · 9 obs."]],
      side: { t: "Hallazgos críticos 2025", items: ["Sin Política de Operaciones Vinculadas", "Normativa interna desactualizada", "Uso inapropiado del Fondo de Educación"] } },
    { y: 4.35, year: "2026", color: GREEN, hitos: [["31 JUL 2026", "Nombramiento SSF", "Nota 18432 · corte 31/07"], ["21 AGO 2026", "Inicio de visita ampliada", ""], ["09 OCT 2026", "Informe preliminar", "Nota 22741 · 7 + 2 obs."], ["ACTUAL", "Plazos en curso", "10 y 15 días hábiles"]],
      side: { t: "Plazos de respuesta", items: ["Descargos (Anexo A): 10 días hábiles, aprox. 23/10/2026", "Plan de solución (Anexo B): 15 días hábiles, aprox. 30/10/2026"] } },
  ];
  blocks.forEach(b => {
    pill(s, "Año línea " + b.year, b.year, M, b.y, 0.95, b.color);
    const x0 = M + 1.25, lineY = b.y + 0.17;
    s.addShape(pres.shapes.LINE, { objectName: "Eje " + b.year, x: x0, y: lineY, w: lineW, h: 0, line: { color: b.color, width: 2.5 } });
    const n = b.hitos.length, step = lineW / n;
    b.hitos.forEach(([d, t, sub], i) => {
      const cx = x0 + step * i + 0.1;
      const last = d === "ACTUAL";
      s.addShape(pres.shapes.OVAL, { objectName: `Hito ${b.year}-${i + 1}`, x: cx, y: lineY - 0.14, w: 0.28, h: 0.28, fill: { color: last ? b.color : C.background1 }, line: { color: b.color, width: 3 } });
      txt(s, `Hito fecha ${b.year}-${i + 1}`, d, { x: cx, y: lineY + 0.32, w: step - 0.25, h: 0.3, fontSize: 12, bold: true, color: b.color, fontFace: "Courier New" });
      txt(s, `Hito texto ${b.year}-${i + 1}`, t, { x: cx, y: lineY + 0.65, w: step - 0.25, h: 0.6, fontSize: 14, bold: true });
      if (sub) txt(s, `Hito detalle ${b.year}-${i + 1}`, sub, { x: cx, y: lineY + 1.25, w: step - 0.25, h: 0.5, fontSize: 12, color: MUTED });
    });
    const sx = x0 + lineW + 0.3, sw = W - M - sx;
    box(s, "Recuadro " + b.year, sx, b.y - 0.1, sw, 2.25, C.background2);
    txt(s, "Recuadro título " + b.year, b.side.t, { x: sx + 0.2, y: b.y + 0.05, w: sw - 0.4, h: 0.35, fontSize: 14, bold: true, color: b.color });
    s.addText(b.side.items.map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
      { objectName: "Recuadro lista " + b.year, isTextBox: true, x: sx + 0.2, y: b.y + 0.45, w: sw - 0.4, h: 1.6, fontSize: 13, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 5 });
  });
  s.addNotes("Ambas visitas siguieron el mismo ciclo: nombramiento, inicio de visita e informe preliminar. En 2026 la entidad está ahora dentro de los plazos de respuesta: 10 días hábiles para descargos del Anexo A y 15 días hábiles para el plan de solución del Anexo B.");
}
{
  const s = content("4 · Línea de tiempo y evolución", "4 · ANÁLISIS DE EVOLUCIÓN", "Observaciones reincidentes y persistentes");
  const rows = [
    ["Integración de Comités de Apoyo con personal ejecutivo", "Obs. 3 · Art. 3 lit. h) NRP-17", "NO SUPERADA", "Anexo B, obs. 8", RED],
    ["Actualización de normativa interna", "Obs. 5 · 9 documentos vencidos", "PARCIALMENTE SUPERADA", "Anexo B, obs. 9: RIT, Comité de Auditoría, Compras", AMBER_INK],
    ["Gobernanza del Fondo de Educación", "Obs. 9 · Art. 59 Ley de Cooperativas", "REINCIDE", "Anexo A, obs. 6 y 7: Reglamento sin criterios legales y $100,774.28 sin detalle a la Asamblea", AMBER_INK],
  ];
  const xT = M, wT = 3.9, x25 = M + 4.15, w25 = 3.1, xA = x25 + w25 + 0.15, x26 = xA + 0.6, w26 = W - M - x26;
  txt(s, "Encabezado tema", "TEMA", { x: xT, y: 1.6, w: wT, h: 0.3, fontSize: 12, bold: true, color: MUTED, charSpacing: 2 });
  txt(s, "Encabezado 2025", "2025 · NOTA 17003", { x: x25, y: 1.6, w: w25, h: 0.3, fontSize: 12, bold: true, color: NAVY, charSpacing: 2 });
  txt(s, "Encabezado 2026", "2026 · NOTA 22741", { x: x26, y: 1.6, w: w26, h: 0.3, fontSize: 12, bold: true, color: GREEN, charSpacing: 2 });
  rows.forEach(([tema, a, st, b, c], i) => {
    const y = 2.05 + i * 1.6, h = 1.4;
    box(s, "Fila evolución " + (i + 1), M, y, W - 2 * M, h, C.background2);
    txt(s, "Tema " + (i + 1), tema, { x: xT + 0.25, y, w: wT - 0.3, h, fontSize: 17, bold: true, fontFace: THEME.headFontFace, valign: "middle" });
    pill(s, "Estado 2025 " + (i + 1), "OBSERVADA", x25, y + 0.3, 1.6, NAVY);
    txt(s, "Detalle 2025 " + (i + 1), a, { x: x25, y: y + 0.75, w: w25, h: 0.45, fontSize: 13, color: C.text1 });
    s.addShape(pres.shapes.RIGHT_ARROW, { objectName: "Flecha " + (i + 1), x: xA, y: y + 0.5, w: 0.45, h: 0.4, fill: { color: MUTED }, line: { color: MUTED } });
    pill(s, "Estado 2026 " + (i + 1), st, x26, y + 0.3, st.length > 12 ? 2.9 : 1.9, c);
    txt(s, "Detalle 2026 " + (i + 1), b, { x: x26, y: y + 0.75, w: w26 - 0.2, h: 0.6, fontSize: 13, color: C.text1 });
  });
  s.addNotes("Tres temas persisten entre visitas. Comités con ejecutivos: no superada. Normativa interna: parcialmente superada. Fondo de Educación: reincide con hallazgos de falta de transparencia ($100,774.28 no detallados a la Asamblea) y conceptos no delimitados por ley en su Reglamento.");
}

// ============ 9. ANEXO · KIT DE DISEÑO ============
pres.addSection({ title: "Anexo · Kit de diseño" });
{
  const s = content("Anexo · Kit de diseño", "ANEXO · GUÍA VISUAL", "Paleta, tipografía y uso de la plantilla");
  const sw = [["173A6E", NAVY, "Azul marino corporativo", "Hitos y tablas 2025"], ["1D7A4B", GREEN, "Verde", "Hitos y tablas 2026"], ["B42318", RED, "Rojo", "Estado NO SUPERADA"], ["D99A1E", AMBER, "Amarillo / ámbar", "PARCIALMENTE SUPERADA y reincidencias"], ["EEF2F8", C.background2, "Gris azulado claro", "Fondo de tarjetas"]];
  sw.forEach(([hex, c, n, u], i) => {
    const y = 1.65 + i * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Muestra " + n, x: M, y, w: 0.75, h: 0.75, rectRadius: 0.08, fill: { color: c }, line: { color: "D9E0EA", width: 1 } });
    txt(s, "Color " + n, [{ text: n + "  ", options: { bold: true } }, { text: "#" + hex, options: { color: MUTED, fontFace: "Courier New" } }], { x: M + 0.95, y: y + 0.05, w: 4.6, h: 0.32, fontSize: 14 });
    txt(s, "Uso " + n, u, { x: M + 0.95, y: y + 0.4, w: 4.6, h: 0.3, fontSize: 13, color: MUTED });
  });
  const x = 6.6, w = W - M - x;
  box(s, "Tipografía", x, 1.65, w, 2.1, C.background2);
  txt(s, "Tipografía título", "Tipografía", { x: x + 0.25, y: 1.8, w: w - 0.5, h: 0.35, fontSize: 16, bold: true });
  s.addText([
    { text: "Títulos: Cambria Bold, 30-40 pt", options: { bullet: true, breakLine: true } },
    { text: "Texto y tablas: Calibri, 13-16 pt", options: { bullet: true, breakLine: true } },
    { text: "Números de nota y fechas: Courier New", options: { bullet: true, breakLine: true } },
    { text: "Etiquetas: 11-12 pt en MAYÚSCULAS con espaciado", options: { bullet: true } },
  ], { objectName: "Tipografía lista", isTextBox: true, x: x + 0.25, y: 2.2, w: w - 0.5, h: 1.45, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 4 });
  box(s, "Uso plantilla", x, 3.95, w, 2.2, C.background2);
  txt(s, "Uso plantilla título", "Cómo editar en PowerPoint", { x: x + 0.25, y: 4.1, w: w - 0.5, h: 0.35, fontSize: 16, bold: true });
  s.addText([
    { text: "Diseño > Variantes > Colores: el tema “SSF Gobierno Corporativo” controla toda la paleta", options: { bullet: true, breakLine: true } },
    { text: "Vista > Patrón de diapositivas: diseños PORTADA y CONTENIDO", options: { bullet: true, breakLine: true } },
    { text: "Las secciones 1 a 4 corresponden a las 4 pestañas del dashboard", options: { bullet: true, breakLine: true } },
    { text: "Para Canva o Gamma: importar este .pptx directamente", options: { bullet: true } },
  ], { objectName: "Uso plantilla lista", isTextBox: true, x: x + 0.25, y: 4.5, w: w - 0.5, h: 1.55, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 4 });
  s.addNotes("Guía de colores y fuentes para mantener la consistencia si se agregan diapositivas o se migra la presentación a Canva o Gamma.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
