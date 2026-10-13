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
const pill = (s, name, text, x, y, w, color, fs = 11) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name + " fondo", x, y, w, h: 0.34, rectRadius: 0.17, fill: { color }, line: { color } });
  s.addText(text, { objectName: name, isTextBox: true, x, y, w, h: 0.34, margin: 0, align: "center", valign: "middle", fontSize: fs, bold: true, color: C.background1, charSpacing: 1 });
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

// ============ SLIDER 1 · RESUMEN ============
pres.addSection({ title: "1 · Resumen de notas SSF" });
{
  const s = content("1 · Resumen de notas SSF", "1 · SUPERVISIONES DE GOBIERNO CORPORATIVO SSF A COMEDICA DE R.L.", "Notas de notificación SSF 2025 - 2026");
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

// ============ SLIDER 2 · OBS 2025 ============
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
  const s = content("2 · Observaciones 2025", "2 · NOTA DS-SABAO-BCS-BC-17003 · AÑO 2025", "Detalle de observaciones 2025");
  obsTable(s, "Tabla observaciones 2025", OBS25, 1, NAVY, 1.6, { fontSize: 12, margin: [0.05, 0.1, 0.05, 0.1] });
  s.addNotes("Nueve observaciones de la Nota 17003 (visita con corte al 30/04/2025). Las observaciones 3, 5 y 9 siguen abiertas en la visita 2026: comités sin ejecutivos (no superada), normativa desactualizada (parcialmente superada) y Fondo de Educación (reincide).");
}

// ============ SLIDER 3 · OBS 2026 (Anexo A + B) ============
const OBS26A = [
  ["Falta de inducción sobre partidas relevantes de EEFF a nuevos miembros", "Gerencia Financiera / RRHH", "Art. 11 NRP-17"],
  ["Falta de Certificación de Lectura del Código de Conducta", "Junta de Vigilancia / Comités de Apoyo", "Código de Conducta"],
  ["Falta de formalización y aprobación de Políticas por el Consejo de Administración", "Consejo de Administración", "Políticas de Retribución y Control Interno"],
  ["Información de comités desactualizada en sitio web institucional", "Oficialía de Cumplimiento / Comunicaciones", "Arts. 27 y 28 NRP-17"],
  ["Incongruencia entre Estatutos y Código de Gobierno Corporativo", "Asesoría Legal / Consejo de Administración", "Restricción de cargos remunerados"],
  ["Reglamento del Fondo de Educación incluye conceptos no delimitados por Ley", "Comité de Educación / Gerencia Financiera", "Art. 5 del Reglamento"],
  ["Limitada transparencia en información presentada a Asamblea sobre Fondo de Educación", "Consejo de Administración / Comité de Educación", "$100,774.28 no detallados"],
];
const OBS26B = [
  ["Comités de Apoyo no integran personal ejecutivo", "Consejo de Administración", "NO SUPERADA", RED],
  ["Normativa interna desactualizada (RIT, Comité de Auditoría, Compras)", "Departamento de Procesos", "PARCIALMENTE SUPERADA", AMBER_INK],
];
pres.addSection({ title: "3 · Observaciones 2026" });
{
  const s = content("3 · Observaciones 2026", "3 · NOTA DS-SABAO-BCS-BC-22741 · AÑO 2026", "Detalle de observaciones 2026");
  const cell = (text, i, o = {}) => ({ text, options: { fill: { color: i % 2 ? C.background1 : C.background2 }, ...o } });
  const sec = (text, color) => [{ text, options: { colspan: 4, bold: true, color: C.background1, fill: { color }, fontSize: 11 } }];
  const rows = [
    HEAD.map(t => ({ text: t, options: { bold: true, color: C.text1, fill: { color: "D9E0EA" }, fontSize: 11 } })),
    sec("ANEXO A · OBSERVACIONES DETERMINADAS EN LA VISITA 2026 · Descargo en 10 días hábiles (buzón DESCARGOS_OBSERVACIONES, vence aprox. 23/10/2026)", GREEN),
    ...OBS26A.map(([o, r, c], i) => [
      cell(String(i + 1), i, { bold: true, color: MUTED, align: "center" }), cell(o, i, { bold: true }), cell(r, i),
      cell([{ text: "Descargo 10 días hábiles", options: { bold: true, color: GREEN, breakLine: true } }, { text: c }], i),
    ]),
    sec("ANEXO B · SEGUIMIENTO A OBSERVACIONES DE VISITAS ANTERIORES · Plan de solución en 15 días hábiles (buzón PLANES_SOLUC_VISITASUPERVIS, vence aprox. 30/10/2026)", NAVY),
    ...OBS26B.map(([o, r, st, col], i) => [
      cell(String(i + 8), i, { bold: true, color: MUTED, align: "center" }), cell(o, i, { bold: true }), cell(r, i),
      cell([{ text: st, options: { bold: true, color: col, breakLine: true } }, { text: "Plan de solución en 15 días hábiles" }], i),
    ]),
  ];
  s.addTable(rows, {
    objectName: "Tabla observaciones 2026", x: M, y: 1.55, w: W - 2 * M, colW: COLW, fontFace: THEME.bodyFontFace, fontSize: 11.5, color: C.text1,
    valign: "middle", margin: [0.04, 0.1, 0.04, 0.1], border: { type: "solid", pt: 0.75, color: "D9E0EA" }, autoPage: false,
  });
  s.addNotes("Anexo A: siete observaciones nuevas; plazo de 10 días hábiles desde el día hábil siguiente a la recepción para remitir pruebas de descargo documentadas, previo conocimiento del Consejo de Administración. Anexo B: dos observaciones de visitas anteriores siguen abiertas; plan de solución en 15 días hábiles con acciones, responsable, plazo por acción y certificación del punto de acta del Consejo. Fechas aproximadas sin asuetos; validar con el Sistema de Control de Envíos.");
}

// ============ SLIDER 4 · LÍNEA DE TIEMPO Y EVOLUCIÓN ============
pres.addSection({ title: "4 · Línea de tiempo y evolución" });
{
  const s = content("4 · Línea de tiempo y evolución", "4 · LÍNEA DE TIEMPO COMPARATIVA Y EVOLUCIÓN DEL CUMPLIMIENTO", "Línea de tiempo 2025 vs. 2026");
  const x0 = M + 1.2, lineW = 6.9, sx = x0 + lineW + 0.35, sw = W - M - sx;
  const blocks = [
    { y: 1.75, year: "2025", color: NAVY, hitos: [["09 ABR 2025", "Nombramiento SSF", "Nota 9337 · corte 30/04"], ["19 MAY 2025", "Inicio de visita en sitio", ""], ["05 NOV 2025", "Informe preliminar", "Nota 17003 · 9 obs."]],
      side: { t: "Hallazgos críticos 2025", items: ["Sin Política de Operaciones Vinculadas", "Normativa interna desactualizada", "Uso inapropiado del Fondo de Educación"] } },
    { y: 3.3, year: "2026", color: GREEN, hitos: [["31 JUL 2026", "Nombramiento SSF", "Nota 18432 · corte 31/07"], ["21 AGO 2026", "Inicio de visita ampliada", ""], ["09 OCT 2026", "Informe preliminar", "Nota 22741 · 7 + 2 obs."], ["ACTUAL", "Plazos en curso", "Vencen 23/10 y 30/10"]],
      side: { t: "Plazos de respuesta", items: ["Descargos Anexo A: 10 días hábiles", "Plan de solución Anexo B: 15 días hábiles"] } },
  ];
  blocks.forEach(b => {
    pill(s, "Año línea " + b.year, b.year, M, b.y - 0.17, 0.95, b.color);
    s.addShape(pres.shapes.LINE, { objectName: "Eje " + b.year, x: x0, y: b.y, w: lineW, h: 0, line: { color: b.color, width: 2.5 } });
    const step = lineW / b.hitos.length;
    b.hitos.forEach(([d, t, sub], i) => {
      const cx = x0 + step * i + 0.05, tw = step - 0.15;
      s.addShape(pres.shapes.OVAL, { objectName: `Hito ${b.year}-${i + 1}`, x: cx, y: b.y - 0.13, w: 0.26, h: 0.26, fill: { color: d === "ACTUAL" ? b.color : C.background1 }, line: { color: b.color, width: 3 } });
      txt(s, `Hito fecha ${b.year}-${i + 1}`, d, { x: cx, y: b.y + 0.25, w: tw, h: 0.25, fontSize: 11, bold: true, color: b.color, fontFace: "Courier New" });
      txt(s, `Hito texto ${b.year}-${i + 1}`, t, { x: cx, y: b.y + 0.52, w: tw, h: 0.45, fontSize: 13, bold: true });
      if (sub) txt(s, `Hito detalle ${b.year}-${i + 1}`, sub, { x: cx, y: b.y + 0.97, w: tw, h: 0.3, fontSize: 10.5, color: MUTED });
    });
    box(s, "Recuadro " + b.year, sx, b.y - 0.25, sw, 1.45, C.background2);
    txt(s, "Recuadro título " + b.year, b.side.t, { x: sx + 0.2, y: b.y - 0.12, w: sw - 0.4, h: 0.3, fontSize: 13, bold: true, color: b.color });
    s.addText(b.side.items.map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
      { objectName: "Recuadro lista " + b.year, isTextBox: true, x: sx + 0.2, y: b.y + 0.22, w: sw - 0.4, h: 0.9, fontSize: 12, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 2 });
  });
  // Evolución
  txt(s, "Evolución título", "ANÁLISIS DE EVOLUCIÓN · OBSERVACIONES REINCIDENTES Y PERSISTENTES", { x: M, y: 4.85, w: W - 2 * M, h: 0.3, fontSize: 12, bold: true, color: MUTED, charSpacing: 2 });
  const evo = [
    ["Integración de comités con ejecutivos", "Obs. 3", "NO SUPERADA", RED, "Anexo B, obs. 8"],
    ["Actualización de normativa interna", "Obs. 5", "PARCIALMENTE SUPERADA", AMBER_INK, "Anexo B, obs. 9: RIT, Comité de Auditoría, Compras"],
    ["Gobernanza del Fondo de Educación", "Obs. 9", "REINCIDE", AMBER_INK, "Reglamento sin criterios legales y $100,774.28 sin detalle a la Asamblea"],
  ];
  const cw = (W - 2 * M - 0.6) / 3, cy = 5.22;
  evo.forEach(([t, a, st, col, d], i) => {
    const x = M + i * (cw + 0.3);
    box(s, "Evolución " + (i + 1), x, cy, cw, 1.6, C.background2);
    txt(s, "Evolución tema " + (i + 1), t, { x: x + 0.2, y: cy + 0.12, w: cw - 0.4, h: 0.3, fontSize: 14, bold: true, fontFace: THEME.headFontFace });
    pill(s, "Evolución 2025 " + (i + 1), a.toUpperCase(), x + 0.2, cy + 0.55, 0.9, NAVY);
    s.addShape(pres.shapes.RIGHT_ARROW, { objectName: "Evolución flecha " + (i + 1), x: x + 1.22, y: cy + 0.6, w: 0.32, h: 0.24, fill: { color: MUTED }, line: { color: MUTED } });
    pill(s, "Evolución 2026 " + (i + 1), st, x + 1.66, cy + 0.55, st.length > 12 ? cw - 1.86 : 1.3, col, st.length > 12 ? 10 : 11);
    txt(s, "Evolución detalle " + (i + 1), d, { x: x + 0.2, y: cy + 1.0, w: cw - 0.4, h: 0.5, fontSize: 11.5, color: C.text1 });
  });
  s.addNotes("Ambas visitas siguieron el mismo ciclo: nombramiento, inicio de visita e informe preliminar. En 2026 la entidad está dentro de los plazos de respuesta. Tres temas persisten: comités con ejecutivos (no superada), normativa interna (parcialmente superada) y Fondo de Educación (reincide). Paleta: azul marino #173A6E para 2025, verde #1D7A4B para 2026, rojo #B42318 no superada, ámbar #D99A1E parcial.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
