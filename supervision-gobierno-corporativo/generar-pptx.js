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
const W = 13.333, M = 0.6, PANEL = 2.75;
const FOOT = "SSF · Gobierno Corporativo (NRP-17) · COMEDICA de R.L.";

// ---------- Diseños ----------
pres.defineSlideMaster({
  title: "RESUMEN PANEL",
  background: { color: WHITE },
  objects: [
    { rect: { x: 0, y: 0, w: PANEL, h: 7.5, fill: { color: NAVY }, line: { color: NAVY } } },
    { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.55, w: PANEL - 2 * M + 0.2, h: 1.4, fontSize: 11, bold: true, color: TINT, charSpacing: 2, margin: 0, valign: "top" }, text: "SECCIÓN" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.2, w: PANEL - 2 * M + 0.2, h: 1.6, fontFace: THEME.headFontFace, fontSize: 40, bold: true, color: WHITE, valign: "top", align: "left", margin: 0 }, text: "Título" } },
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
  const s = addSlide("RESUMEN PANEL", "1 · Resumen de circulares SSF", "SUPERVISIONES DE GOBIERNO CORPORATIVO · COMEDICA DE R.L.", "2025 –\n2026");
  const RX = PANEL + 0.5, RW = W - M - RX;
  const rows = [
    { y: 0.35, yr: "2025", sub: "Visita 1 · supervisión ampliada", color: NAVY, cards: [
      { fecha: "09/04/2025", tipo: "Nombramiento de visita ampliada", no: "DS-SABAO-BCS-BC-9337", ini: ["INICIO DE VISITA", "19", "MAY 2025"], fin: ["ACTA DE CIERRE", "04", "JUL 2025"], facts: [["Corte", "30/04/2025"], ["Acta de cierre", "04/07/2025"]] },
      { fecha: "05/11/2025", tipo: "Resultados preliminares", no: "DS-SABAO-BCS-BC-17003", ini: ["RECEPCIÓN", "05", "NOV 2025"], fin: ["DESCARGOS", "18", "NOV 2025"],
        groups: [{ t: "9 OBSERVACIONES", col: NAVY, cols: [
          [[1, "Política de Operaciones Vinculadas"], [2, "Confidencialidad e inducción"], [3, "Comités sin personal ejecutivo", "seguimiento 2026"], [4, "Pago de dietas"], [5, "Normativa interna desactualizada", "seguimiento 2026"]],
          [[6, "Soporte de Ferias de Crédito"], [7, "Declaración jurada (NRP-75)"], [8, "Manual de Compras"], [9, "Uso del Fondo de Educación"]]] }] }] },
    { y: 3.68, yr: "2026", sub: "Visita 2 · supervisión ampliada", color: GREEN, cards: [
      { fecha: "31/07/2026", tipo: "Nombramiento de visita ampliada", no: "SABAO-BCS-BC-18432", ini: ["INICIO DE VISITA", "21", "AGO 2026"], fin: ["INFORME", "09", "OCT 2026"], facts: [["Corte", "31/07/2026"], ["Norma", "NRP-17"]] },
      { fecha: "09/10/2026", tipo: "Resultados preliminares", no: "DS-SABAO-BCS-BC-22741", ini: ["RECEPCIÓN", "09", "OCT 2026"], fin: ["VENCE APROX.", "23", "OCT 2026"],
        groups: [
          { t: "7 NUEVAS · ANEXO A", col: GREEN, cols: [[[1, "Inducción sobre estados financieros"], [2, "Certificación del Código de Conducta"], [3, "Políticas sin aprobar por el Consejo"], [4, "Sitio web de comités desactualizado"], [5, "Integración de Comités de Apoyo"], [6, "Reglamento del Fondo de Educación"], [7, "Transparencia ante la Asamblea"]]] },
          { t: "2 EN SEGUIMIENTO", col: RED, cols: [[["Obs. 3", "Comités sin ejecutivos", "No superada"], ["Obs. 5", "Normativa interna", "Parcialmente superada"]]] }] }] },
  ];
  const gapC = 0.42, cwA = 3.05, cwB = RW - cwA - gapC, ch = 2.62, dw = 1.08;
  rows.forEach(r => {
    txt(s, "Año " + r.yr, r.yr, { x: RX, y: r.y, w: 1.3, h: 0.5, fontSize: 28, bold: true, color: r.color, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Visita " + r.yr, r.sub, { x: RX + 1.3, y: r.y, w: 4, h: 0.5, fontSize: 14, color: MUTED, valign: "middle" });
    r.cards.forEach((c, j) => {
      const cw = j ? cwB : cwA, x = RX + (j ? cwA + gapC : 0), y = r.y + 0.6, n = `${r.yr}-${j + 1}`;
      card(s, "Tarjeta " + n, x, y, cw, ch, WHITE, { line: { color: LINE_HEX, width: 1 } });
      // Bloque de fechas: inicio y final
      const bx = x + 0.13, by = y + 0.13, bh = ch - 0.26;
      card(s, "Fechas fondo " + n, bx, by, dw, bh, r.color);
      [c.ini, c.fin].forEach(([lab, d, my], k) => {
        const fy = by + 0.2 + k * (bh / 2);
        txt(s, `Fecha etiqueta ${n}-${k}`, lab, { x: bx, y: fy, w: dw, h: 0.2, fontSize: 8, bold: true, color: WHITE, align: "center", charSpacing: 1 });
        txt(s, `Fecha día ${n}-${k}`, d, { x: bx, y: fy + 0.22, w: dw, h: 0.45, fontSize: 26, bold: true, color: WHITE, fontFace: THEME.headFontFace, align: "center", valign: "middle" });
        txt(s, `Fecha mes ${n}-${k}`, my, { x: bx, y: fy + 0.68, w: dw, h: 0.22, fontSize: 10, bold: true, color: WHITE, align: "center", charSpacing: 1 });
      });
      s.addShape(pres.shapes.LINE, { objectName: "Fechas separador " + n, x: bx + 0.2, y: by + bh / 2, w: dw - 0.4, h: 0, line: { color: WHITE, width: 0.75, transparency: 40 } });
      const cx = x + dw + 0.3, cww = cw - dw - 0.45;
      if (!c.groups) {
        txt(s, "Tipo " + n, c.tipo, { x: cx, y: y + 0.15, w: cww, h: 0.62, fontSize: 13.5, bold: true, valign: "middle" });
        txt(s, "Circular " + n, [{ text: "Circular No.  ", options: { color: MUTED, fontSize: 10 } }, { text: c.fecha, options: { bold: true, color: r.color, fontSize: 10.5, breakLine: true } }, { text: c.no, options: { bold: true, color: r.color } }], { x: cx, y: y + 0.9, w: cww, h: 0.5, fontSize: 10.5 });
        c.facts.forEach(([l, v], q) => {
          const fy = y + 1.66 + q * 0.42;
          s.addShape(pres.shapes.LINE, { objectName: `Separador ${n}-${q}`, x: cx, y: fy - 0.04, w: cww, h: 0, line: { color: LINE_HEX, width: 0.75 } });
          txt(s, `Dato ${n}-${q}`, [{ text: l + "  ", options: { color: MUTED } }, { text: v, options: { bold: true } }], { x: cx, y: fy, w: cww, h: 0.34, fontSize: 10.5, valign: "middle" });
        });
        return;
      }
      // Resultados preliminares con el listado de observaciones
      txt(s, "Tipo " + n, c.tipo, { x: cx, y: y + 0.14, w: 3, h: 0.32, fontSize: 13.5, bold: true, valign: "middle" });
      txt(s, "Circular " + n, [{ text: "Circular No.  ", options: { color: MUTED } }, { text: c.fecha + "  ·  " + c.no, options: { bold: true, color: r.color } }], { x: cx, y: y + 0.47, w: cww, h: 0.24, fontSize: 10 });
      s.addShape(pres.shapes.LINE, { objectName: "Separador lista " + n, x: cx, y: y + 0.8, w: cww, h: 0, line: { color: LINE_HEX, width: 0.75 } });
      const ncol = c.groups.reduce((a, g) => a + g.cols.length, 0);
      const widths = r.yr === "2026" ? [2.5, cww - 2.5] : Array(ncol).fill(cww / ncol);
      const xAt = i => cx + widths.slice(0, i).reduce((a, b) => a + b, 0);
      const colw = widths[1];
      let ci = 0;
      c.groups.forEach((g, gi) => {
        txt(s, `Grupo ${n}-${gi}`, g.t, { x: xAt(ci), y: y + 0.9, w: widths.slice(ci, ci + g.cols.length).reduce((a, b) => a + b, 0) - 0.1, h: 0.22, fontSize: 9, bold: true, color: g.col, charSpacing: 1 });
        g.cols.forEach(col => {
          const runs = [];
          col.forEach(([num, t, tag], k) => {
            const last = k === col.length - 1, isB = typeof num === "string";
            runs.push({ text: num + "  ", options: { bold: true, color: tag ? RED : g.col } });
            if (isB) {
              runs.push({ text: t, options: { breakLine: true } });
              runs.push({ text: tag, options: { bold: true, color: RED, fontSize: 9.5, breakLine: !last } });
            } else {
              runs.push({ text: t, options: { color: tag ? RED : INK, bold: !!tag, breakLine: !last } });
            }
          });
          txt(s, `Lista ${n}-${ci}`, runs, { x: xAt(ci), y: y + 1.17, w: widths[ci] - 0.12, h: ch - 1.27, fontSize: col.length > 5 ? 9.5 : 10, paraSpaceAfter: col.length > 5 ? 1 : 3 });
          ci++;
        });
      });
      if (r.yr === "2025") txt(s, "Nota rojo", "En rojo: con seguimiento de la SSF en 2026", { x: xAt(1), y: y + ch - 0.36, w: colw - 0.12, h: 0.22, fontSize: 9, italic: true, color: RED });
    });
    s.addShape(pres.shapes.RIGHT_ARROW, { objectName: "Flecha " + r.yr, x: RX + cwA + 0.07, y: r.y + 0.6 + ch / 2 - 0.14, w: 0.28, h: 0.28, fill: { color: r.color }, line: { color: r.color, transparency: 100 } });
  });
  s.addNotes("Cuatro circulares de la SSF. Cada año la SSF primero notifica el nombramiento de la visita ampliada y luego remite los resultados preliminares. 2025: visita del 19/05/2025 al 04/07/2025 (Acta de Cierre de Visita); Circular 17003 con 9 observaciones; descargos recibidos por la SSF el 18/11/2025 a las 5:19 p. m. (constancia de recepción, archivo RESPUESTA SSF NOTA DS-SABAO-BCS-BC-17003). 2026: Circular 22741 con 7 observaciones nuevas (Anexo A) y seguimiento a 2 de las 9 observaciones de 2025 (Anexo B): la obs. 3 no superada y la obs. 5 parcialmente superada.");
}

// ============ SLIDE 2 · OBSERVACIONES 2025 ============
{
  const s = addSlide("CONTENIDO", "2 · Observaciones 2025", "CIRCULAR DS-SABAO-BCS-BC-17003 · VISITA CON CORTE AL 30/04/2025", "Observaciones 2025");
  legendDot(s, "Leyenda 2025", "Observación 2025", 9.35, 0.85, NAVY, 1.4);
  legendDot(s, "Leyenda abierta", "En seguimiento SSF 2026", 10.85, 0.85, RED, 1.95);
  // Responsable y respuesta: Matriz de requerimientos (descargos) Circular 17003
  const O = [
    ["Falta de Política de Operaciones Vinculadas", "Gerencia General / Jefe de Administración de Riesgos", "Se creará la política para aprobación del Consejo de Administración"],
    ["Sin evidencia de firma de confidencialidad ni fechas de inducción en actas", "Gerencia General / Jefe de Administración de Riesgos", "Acta 2579 del Consejo de Administración: fechas de inducción aprobadas"],
    ["Comités de Apoyo no integran personal ejecutivo", "Gerencia General / Jefe de Administración de Riesgos", "Integración sustentada en Estatutos (arts. 61 y 73)", "Seguimiento 2026: NO SUPERADA", RED],
    ["Inconsistencia en pago de dietas (Reglamento vs. Asamblea General)", "Gerencia General / Jefe de Administración de Riesgos", "Se actualizará el Reglamento de Compensaciones Económicas"],
    ["Normativa interna desactualizada (más de 2 años sin revisión)", "Gerencia General / Jefe de Administración de Riesgos, Auditoría Interna, Gerencia de Finanzas, Seguridad de la Información y Gerencia Administrativa", "3 de 9 documentos actualizados; el resto en proceso", "Seguimiento 2026: PARCIALMENTE SUPERADA", RED],
    ["Falta de soporte en expedientes de créditos de Ferias de Crédito", "Gerencia de Negocios / Administración de Cartera", "Se enviaron resoluciones y acuerdos del Consejo de cada feria"],
    ["Falta de información obligatoria en declaración jurada de personas relacionadas", "Gerencia General / Jefe de Administración de Riesgos", "Se pidió a la SSF el detalle de las personas señaladas"],
    ["Omisiones en procedimientos del Manual de Compras de Bienes y Servicios", "Gerencia Administrativa", "Firmas de Mercadeo y Gerencia General; contrato terminado"],
    ["Uso del Fondo de Educación en fines distintos a los legales", "Gerencia General / Jefe de Administración de Riesgos", "El Consejo de Administración suspendió el Comité Sociocultural (Acta 2576)"],
  ];
  const gap = 0.15, cw = (W - 2 * M - 2 * gap) / 3, ch = 1.75, y0 = 1.45;
  O.forEach(([t, r, c, st, col], i) => {
    const x = M + (i % 3) * (cw + gap), y = y0 + Math.floor(i / 3) * (ch + gap);
    card(s, "Observación 2025-" + (i + 1), x, y, cw, ch, TINT);
    badge(s, "Número 2025-" + (i + 1), i + 1, x + 0.2, y + 0.2, 0.42, col || NAVY);
    txt(s, "Título 2025-" + (i + 1), t, { x: x + 0.75, y: y + 0.13, w: cw - 0.95, h: 0.62, fontSize: 12.5, bold: true, valign: "middle" });
    const lines = [
      { text: "Responsable  ", options: { bold: true, color: MUTED } }, { text: r, options: { breakLine: true } },
      { text: "Descargo  ", options: { bold: true, color: MUTED } }, { text: c, options: st ? { breakLine: true } : {} },
    ];
    if (st) lines.push({ text: "↳ " + st, options: { bold: true, color: col } });
    txt(s, "Detalle 2025-" + (i + 1), lines, { x: x + 0.2, y: y + 0.82, w: cw - 0.4, h: ch - 0.88, fontSize: 10, paraSpaceAfter: 1 });
  });
  s.addNotes("Nueve observaciones de la Circular 17003. Responsable y descargo según la Matriz de requerimientos enviada a la SSF el 18/11/2025 (ref. NB07-2025-974, buzón DESCARGOS_OBSERVACIONES). Obs. 5: actualizados el Reglamento del Comité de Finanzas (21/07/2025), la Política de Seguridad de la Información (02/06/2025) y el Manual de Inducción (25/08/2025). En rojo, las 2 observaciones con seguimiento en 2026 (Anexo B): obs. 3 no superada y obs. 5 parcialmente superada.");
}

// ============ SLIDE 3 · OBSERVACIONES 2026 ============
{
  const s = addSlide("CONTENIDO", "3 · Observaciones 2026", "CIRCULAR DS-SABAO-BCS-BC-22741 · VISITA CON CORTE AL 31/07/2026", "Observaciones 2026");
  legendDot(s, "Leyenda A", "Anexo A · nuevas", 8.8, 0.85, GREEN, 1.5);
  legendDot(s, "Leyenda B", "Anexo B · seguimiento 2025", 10.55, 0.85, NAVY, 2.1);
  // Responsable y respuesta: Matriz de requerimientos y Plan de solución, Circular 22741
  const A = [
    ["Falta de inducción sobre estados financieros a nuevos miembros", "Gerencia Financiera", ""],
    ["Falta de Certificación de Lectura del Código de Conducta", "Gerencia de Oficialía de Cumplimiento", "6 de 7 certificaciones localizadas; se pide reconsiderar"],
    ["Falta de formalización y aprobación de políticas", "Talento Humano · Departamento de Procesos", "Se elaborará la Política de Retribución para aprobación del Consejo"],
    ["Información de comités desactualizada en el sitio web", "Administración de Riesgos", ""],
    ["Incongruencia entre Estatutos y Código de Gobierno Corporativo", "Gerencia General · Administración de Riesgos", ""],
    ["Reglamento del Fondo de Educación con conceptos no delimitados por Ley", "Gerencia Administrativa · Jefe de Educación y Responsabilidad Social", "Revisión integral del Reglamento del Fondo"],
    ["Limitada transparencia a la Asamblea sobre el Fondo de Educación", "Gerencia Administrativa · Jefatura de Educación y Responsabilidad Social", "Memoria de Labores 2026 con conceptos y montos"],
  ];
  const gap = 0.12, cw = (W - 2 * M - 3 * gap) / 4, ch = 1.58;
  txt(s, "Etiqueta Anexo A", "ANEXO A · OBSERVACIONES NUEVAS DETERMINADAS EN LA VISITA 2026", { x: M, y: 1.4, w: 9, h: 0.26, fontSize: 11, bold: true, color: GREEN, charSpacing: 2 });
  const cell = (i) => ({ x: M + (i % 4) * (cw + gap), y: 1.72 + Math.floor(i / 4) * (ch + gap) });
  A.forEach(([t, r, c], i) => {
    const { x, y } = cell(i);
    card(s, "Anexo A-" + (i + 1), x, y, cw, ch, TINT);
    badge(s, "Número A-" + (i + 1), i + 1, x + 0.16, y + 0.14, 0.36, GREEN);
    txt(s, "Responsable A-" + (i + 1), [{ text: "Responsable", options: { bold: true, color: GREEN, breakLine: true, fontSize: 9.5 } }, { text: r }], { x: x + 0.62, y: y + 0.08, w: cw - 0.72, h: 0.5, fontSize: 10, color: INK, valign: "middle" });
    txt(s, "Título A-" + (i + 1), t, { x: x + 0.16, y: y + 0.64, w: cw - 0.32, h: 0.55, fontSize: 11.5, bold: true });
    txt(s, "Respuesta A-" + (i + 1), c ? "Respuesta: " + c : "Respuesta: sin comentario en la matriz", { x: x + 0.16, y: y + 1.18, w: cw - 0.32, h: 0.36, fontSize: 9.5, italic: true, color: MUTED });
  });
  {
    const { x, y } = cell(7);
    card(s, "Plazo Anexo A", x, y, cw, ch, GREEN);
    txt(s, "Plazo A número", "10", { x: x + 0.2, y: y + 0.1, w: 0.9, h: 0.7, fontSize: 40, bold: true, color: WHITE, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Plazo A título", "días hábiles para descargos", { x: x + 1.1, y: y + 0.15, w: cw - 1.25, h: 0.6, fontSize: 12, bold: true, color: WHITE, valign: "middle" });
    txt(s, "Plazo A detalle", "Buzón DESCARGOS_OBSERVACIONES\nVence aprox. 23/10/2026", { x: x + 0.2, y: y + 0.88, w: cw - 0.35, h: 0.55, fontSize: 10.5, color: WHITE });
  }
  const yB = 5.4;
  txt(s, "Etiqueta Anexo B", "ANEXO B · SEGUIMIENTO A 2 DE LAS 9 OBSERVACIONES DE 2025 (NO SON NUEVAS)", { x: M, y: yB - 0.3, w: 9, h: 0.26, fontSize: 11, bold: true, color: AMBER_INK, charSpacing: 2 });
  const B = [
    ["Comités de Apoyo no integran personal ejecutivo", "Gerencia General · Jefe de Administración de Riesgos", "Plan de solución: acciones aún sin registrar", "NO SUPERADA", RED, 3],
    ["Normativa interna desactualizada", "Talento Humano · Auditor Interno · Servicios Generales", "RIT y Manual de Compras en proceso · Reglamento del Comité de Auditoría aprobado (28/08/2026) y enviado al Consejo (05/10/2026)", "PARCIALMENTE SUPERADA", AMBER, 5],
  ];
  const bw = (W - 2 * M - cw - 2 * gap) / 2, bh = 1.45;
  B.forEach(([t, r, av, st, col, n25], i) => {
    const x = M + i * (bw + gap);
    card(s, "Anexo B-" + n25, x, yB, bw, bh, col, { transparency: 88 });
    pill(s, "Origen B-" + n25, "OBS. " + n25 + " DE 2025", x + 0.2, yB + 0.14, 1.45, NAVY, 10.5);
    pill(s, "Estado B-" + n25, st, x + 1.8, yB + 0.14, st.length > 12 ? 2.4 : 1.45, col, 10.5);
    txt(s, "Título B-" + n25, t, { x: x + 0.2, y: yB + 0.5, w: bw - 0.4, h: 0.28, fontSize: 12.5, bold: true });
    txt(s, "Responsable B-" + n25, [{ text: "Responsable: ", options: { bold: true, color: MUTED } }, { text: r }], { x: x + 0.2, y: yB + 0.78, w: bw - 0.4, h: 0.24, fontSize: 10, color: INK });
    txt(s, "Avance B-" + n25, [{ text: "Avance: ", options: { bold: true, color: MUTED } }, { text: av }], { x: x + 0.2, y: yB + 1.02, w: bw - 0.4, h: 0.4, fontSize: 9.5, color: INK });
  });
  {
    const x = M + 2 * (bw + gap);
    card(s, "Plazo Anexo B", x, yB, cw, bh, AMBER_INK);
    txt(s, "Plazo B número", "15", { x: x + 0.2, y: yB + 0.08, w: 0.9, h: 0.7, fontSize: 40, bold: true, color: WHITE, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Plazo B título", "días hábiles para plan de solución", { x: x + 1.1, y: yB + 0.13, w: cw - 1.25, h: 0.6, fontSize: 12, bold: true, color: WHITE, valign: "middle" });
    txt(s, "Plazo B detalle", "Buzón PLANES_SOLUC_VISITASUPERVIS\nVence aprox. 30/10/2026", { x: x + 0.2, y: yB + 0.85, w: cw - 0.3, h: 0.5, fontSize: 10.5, color: WHITE });
  }
  s.addNotes("Anexo A: siete observaciones nuevas. Responsable y respuesta según la Matriz de requerimientos de la Circular 22741 (las obs. 1, 4 y 5 aún no tienen comentario en la matriz). Anexo B: seguimiento a las obs. 3 y 5 de 2025; responsable y avance según el Plan de solución. Obs. 5: Reglamento Interno de Trabajo en proceso; Reglamento del Comité de Auditoría aprobado por el Comité el 28/08/2026 y remitido el 05/10/2026 al Consejo de Administración para ratificación; Manual de Compras en proceso. Plazos: 10 días hábiles para descargos y 15 días hábiles para el plan de solución (fechas aproximadas sin asuetos).");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("written", OUT);
})();
