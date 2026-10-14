const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/eb134167-ea2a-4bd7-8c34-83283d88af55_b18ad4a5-0fa1-4688-82cb-4bf483ec9783/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "matriz23486.pptx";
const THEME = {
  name: "SSF Gobierno Corporativo", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "142033", lt1: "FFFFFF", dk2: "173A6E", lt2: "EEF2F8", accent1: "173A6E", accent2: "1D7A4B", accent3: "B42318", accent4: "D99A1E", accent5: "586579", accent6: "A86410", hlink: "173A6E", folHlink: "586579" },
};
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Matriz de requerimientos · Circular DS-SABAO-BCS-BC-23486";
pres.author = "COMEDICA de R.L.";
const C = pres.SchemeColor;
const NAVY = C.accent1, GREEN = C.accent2, RED = C.accent3, MUTED = C.accent5, AMBER_INK = C.accent6, WHITE = C.background1, TINT = C.background2, INK = C.text1;
const W = 13.333, M = 0.45;

pres.defineSlideMaster({
  title: "MATRIZ",
  background: { color: WHITE },
  objects: [
    { placeholder: { options: { name: "eyebrow", type: "body", x: M, y: 0.25, w: 8, h: 0.26, fontSize: 10, bold: true, color: MUTED, charSpacing: 2, margin: 0 }, text: "SECCIÓN" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.5, w: 7.6, h: 0.55, fontFace: THEME.headFontFace, fontSize: 24, bold: true, color: INK, align: "left", valign: "top", margin: 0 }, text: "Título" } },
    { text: { text: "SSF · Matriz de requerimientos · COMEDICA de R.L.", options: { x: M, y: 7.18, w: 6, h: 0.22, fontSize: 8.5, color: MUTED, margin: 0 } } },
  ],
});

const s = pres.addSlide({ masterName: "MATRIZ" });
s.addText("CIRCULAR DS-SABAO-BCS-BC-23486 · MATRIZ DE REQUERIMIENTOS", { placeholder: "eyebrow" });
s.addText("Observaciones de la SSF y respuesta de COMEDICA", { placeholder: "title" });

// Cifras clave
const kpis = [["17", "nuevas", NAVY], ["9", "en seguimiento", MUTED], ["5", "no superadas", RED], ["4", "parcialmente superadas", AMBER_INK]];
let kx = 8.2;
kpis.forEach(([n, l, col], i) => {
  const w = [0.85, 1.1, 1.05, 1.45][i];
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Cifra " + l, x: kx, y: 0.3, w, h: 0.72, rectRadius: 0.08, fill: { color: TINT }, line: { color: TINT, transparency: 100 } });
  s.addText([{ text: n, options: { bold: true, fontSize: 20, color: col, fontFace: THEME.headFontFace, breakLine: true } }, { text: l, options: { fontSize: 8.5, color: MUTED } }],
    { objectName: "Cifra texto " + l, isTextBox: true, x: kx, y: 0.3, w, h: 0.72, margin: 0, align: "center", valign: "middle" });
  kx += w + 0.08;
});

const BORDER = [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: "D9E0EA" }, { type: "none" }];
const head = (col) => ["Observación", "Comentarios", "Responsable"].map(t => ({ text: t.toUpperCase(), options: { bold: true, color: WHITE, fill: { color: col }, fontSize: 8, charSpacing: 1 } }));
const sec = (t, col) => [{ text: t, options: { colspan: 3, bold: true, color: col, fill: { color: TINT }, fontSize: 7.5, charSpacing: 1 } }];
const row = (i, obs, com, resp, estado) => {
  const z = i % 2 ? TINT : WHITE;
  const o = [{ text: obs, options: { bold: true, color: INK } }];
  if (estado) o.push({ text: "\n" + estado.toUpperCase(), options: { bold: true, fontSize: 7, color: /no superada/i.test(estado) ? RED : AMBER_INK } });
  const tint = estado ? (/no superada/i.test(estado) ? "FBEAE8" : "FDF4E3") : "FFFFFF";
  return [
    { text: o, options: { fill: { color: tint } } },
    { text: com, options: { color: INK, fill: { color: "FFFFFF" } } },
    { text: resp, options: { color: NAVY, bold: true, fill: { color: "FFFFFF" } } },
  ];
};
const COMP = "Compromiso al 14/10/2026";
const GC = "Gerencia de Cumplimiento";
const GCSI = "Gerencia de Cumplimiento / Seguridad de la Información";

// Panel izquierdo: observaciones nuevas
const left = [
  head(NAVY),
  sec("A. GESTIÓN DE RIESGO DE LDA/FT/FPADM", NAVY),
  row(0, "1. Expedientes de clientes incompletos o sin actualizar", COMP, GC),
  row(1, "2. No se identifica bien la actividad económica de los clientes", COMP, GC),
  sec("B. IMPLEMENTACIÓN DE LA HERRAMIENTA DE MONITOREO", NAVY),
  row(0, "3. Metodología y matriz de riesgo LA/FT/FPADM de la entidad sin actualizar", COMP, GC),
  row(1, "4. Metodología y matriz de riesgo de clientes o asociados sin actualizar", COMP, GC),
  row(0, "5. El envío de reportes regulados a la UIF (Web Service) no está automatizado", COMP, GC),
  row(1, "6. La validación de Listas de Cautela no está automatizada", COMP, GC),
  row(0, "7. No hay evidencia de los avances reportados en el cronograma", COMP, GC),
  row(1, "8. Revisión de documentos según el artículo 22 de la NRP-23", "Se adjuntará evidencia de la fecha de la matriz de riesgos", "Jefe de Riesgos"),
  row(0, "9. Debilidades en el contrato del servicio de software", "Gerencia Legal envió propuesta para modificar el contrato (correo del 09/10/2026)", "Gerencia de Cumplimiento / Gerencia Legal"),
  row(1, "10. Las pruebas de continuidad del negocio tienen un alcance limitado", "Plan de acción con el proveedor crítico", "Seguridad de la Información"),
  row(0, "11. Fallas en los procedimientos de respaldos", "Plan de acción", "Gerencia de Cumplimiento / Gerencia de Sistemas"),
  row(1, "12. La hora del sistema no cuadra con la herramienta de monitoreo", "Descargo con correos de gestión", GC),
  row(0, "13. Verificación de las condiciones en ACRM", "Descargo", GC),
  row(1, "14. Acceso limitado a los datos de COMEDICA en la herramienta de monitoreo", "Descargo", GCSI),
  row(0, "15. No se logra identificar la base de datos del CRM", COMP, GCSI),
  row(1, "16. Falta documentación del artículo 22 de la NRP-23 por la compra del CRM", "Descargo", GCSI),
  row(0, "17. Servicios tercerizados", "Seguridad de la Información reevaluará riesgos y controles del servicio ACRM con el proveedor · vence 30/11/2026", GCSI),
];
// Panel derecho: seguimiento
const right = [
  head(RED),
  sec("C.1 CRÉDITO, OPERACIONAL Y TECNOLÓGICO · CIRCULAR 16572 (10/09/2025)", NAVY),
  row(0, "Obs. 2 · Auditoría Interna no logra que se corrijan sus observaciones", "Préstamos y tarjetas de crédito: ya solventadas. Informes AI 32, 37 y 52 de 2024 siguen en seguimiento · vence 30/11/2026", "Auditoría Interna", "Parcialmente superada"),
  row(1, "Obs. 3 · Auditoría Interna da seguimiento fuera de tiempo", "Se creó una matriz de seguimiento de informes que se actualiza cada trimestre", "Auditoría Interna", "No superada"),
  row(0, "Obs. 2 · Seguimiento a informes del Comité de Auditoría y de Auditoría Interna 2023", "Se retomará el seguimiento de lo pendiente en las Actas 173 y 175 · vence 30/11/2026", "Auditoría Interna", "Parcialmente superada"),
  sec("C.2 NRP-36 (LDA/FT/FPADM) · CIRCULAR 18445 (30/10/2025)", NAVY),
  row(1, "Obs. 1 · No se aplica el factor de riesgo «empleados»", "Sin comentarios ni documentación", "Gerencia de Oficialía de Cumplimiento", "No superada"),
  row(0, "Obs. 4 · Observaciones de Auditoría Interna de años anteriores sin corregir", "Sin comentarios ni documentación", "Auditoría Interna", "Parcialmente superada"),
  sec("C.3 GUÍAS DE CALIDAD · CIRCULAR 29295 (03/12/2025)", NAVY),
  row(1, "Liquidez · Obs. 1 · Al plan de contingencia le faltan elementos", "Se ajustará el Plan de Contingencia de Liquidez · vence 16/11/2026", "Gerencia Financiera", "No superada"),
  row(0, "Liquidez · Obs. 2 · Roles y responsabilidades del plan de contingencia no cuadran", "Se ajustará el Plan de Contingencia de Liquidez · vence 16/11/2026", "Gerencia Financiera", "No superada"),
  row(1, "Operacional y tecnológico · Obs. 1 · Normativa interna desactualizada", "Ya actualizadas: Política de Riesgo Operacional y Manual de Riesgo Tecnológico. Falta llevar al Comité de Riesgos la Política de servicios críticos tercerizados", "Administración de Riesgos", "Parcialmente superada"),
  row(0, "Operacional y tecnológico · Obs. 2 · El BIA no incluye el objetivo mínimo de continuidad (MBCO)", "Actualización del BIA programada para noviembre de 2026", "Administración de Riesgos", "No superada"),
];

const PWL = 6.75, PWR = W - 2 * M - 0.25 - PWL, TY = 1.42, PW = PWL;
s.addText([{ text: "OBSERVACIONES NUEVAS DE LA VISITA", options: { bold: true, color: NAVY } }], { objectName: "Etiqueta nuevas", isTextBox: true, x: M, y: 1.13, w: PW, h: 0.24, fontSize: 9.5, charSpacing: 2, margin: 0 });
s.addText([{ text: "SEGUIMIENTO A OBSERVACIONES DE VISITAS ANTERIORES", options: { bold: true, color: RED } }], { objectName: "Etiqueta seguimiento", isTextBox: true, x: M + PWL + 0.25, y: 1.13, w: PWR, h: 0.24, fontSize: 9.5, charSpacing: 2, margin: 0 });
const opts = (x, colW) => ({ x, y: TY, w: colW.reduce((a, b) => a + b, 0), colW, fontFace: THEME.bodyFontFace, fontSize: 7.5, valign: "middle", margin: [0.025, 0.05, 0.025, 0.05], border: BORDER, autoPage: false });
s.addTable(left, { objectName: "Tabla observaciones nuevas", ...opts(M, [2.5, 2.0, PWL - 4.5]) });
s.addTable(right, { objectName: "Tabla seguimiento", ...opts(M + PWL + 0.25, [2.05, 2.25, PWR - 4.3]) });

s.addText("Todas con compromiso al 14/10/2026", { objectName: "Nota compromiso", isTextBox: true, x: M + PWL - 2.6, y: 1.13, w: 2.6, h: 0.24, fontSize: 8.5, italic: true, color: MUTED, align: "right", margin: 0 });
// Próximos vencimientos
const vx = M + PWL + 0.25, vy = 6.32, vw = PWR;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Vencimientos fondo", x: vx, y: vy, w: vw, h: 0.78, rectRadius: 0.08, fill: { color: TINT }, line: { color: TINT, transparency: 100 } });
s.addText("PRÓXIMOS VENCIMIENTOS", { objectName: "Vencimientos título", isTextBox: true, x: vx + 0.15, y: vy + 0.08, w: 3, h: 0.2, fontSize: 8, bold: true, color: NAVY, charSpacing: 2, margin: 0 });
[["16 NOV 2026", "Plan de Contingencia de Liquidez"], ["30 NOV 2026", "Auditoría Interna y servicio ACRM"], ["NOV 2026", "Actualización del BIA"]].forEach(([d, t], i) => {
  const cw = (vw - 0.3 - 0.2) / 3, cx = vx + 0.15 + i * (cw + 0.1);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Vence " + i, x: cx, y: vy + 0.32, w: cw, h: 0.38, rectRadius: 0.06, fill: { color: WHITE }, line: { color: "D9E0EA", width: 0.75 } });
  s.addText([{ text: d, options: { bold: true, color: RED, fontSize: 8.5, breakLine: true } }, { text: t, options: { color: INK, fontSize: 7.5 } }], { objectName: "Vence texto " + i, isTextBox: true, x: cx + 0.08, y: vy + 0.33, w: cw - 0.16, h: 0.36, margin: 0, valign: "middle" });
});

s.addNotes("Matriz de requerimientos de la Circular DS-SABAO-BCS-BC-23486 (versión corta). Izquierda: 17 observaciones nuevas (A. Gestión de riesgo de LDA/FT/FPADM y B. Implementación de la herramienta de monitoreo); en la matriz todas tienen 'Comentario CA: 14/10/2026 Compromiso'. Derecha: 9 observaciones de visitas anteriores en seguimiento: 5 no superadas y 4 parcialmente superadas. La redacción se adaptó a un lenguaje sencillo sin cambiar el contenido.");

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("written", OUT); })();
