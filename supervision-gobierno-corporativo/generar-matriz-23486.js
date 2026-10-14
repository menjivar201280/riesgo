const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/eb134167-ea2a-4bd7-8c34-83283d88af55_b18ad4a5-0fa1-4688-82cb-4bf483ec9783/pptx/scripts/apply_theme.js");
const ROWS = require("./m23486_rows.json");
const TITLE = "Matriz de requerimientos DS-SABAO-BCS-BC-23486 · Gestión de Riesgo Específico";
const SECH = { A: +(process.env.HA || 0.6), B: 0 };
const FS1 = +(process.env.FS1 || 7.6), FS2 = +(process.env.FS2 || 7.2);
const OUT = process.argv[2] || "m23486v.pptx";
const THEME = { name: "SSF Gobierno Corporativo", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "142033", lt1: "FFFFFF", dk2: "173A6E", lt2: "EEF2F8", accent1: "173A6E", accent2: "1D7A4B", accent3: "B42318", accent4: "D99A1E", accent5: "586579", accent6: "A86410", hlink: "173A6E", folHlink: "586579" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Matriz de requerimientos DS-SABAO-BCS-BC-23486";
const C = pres.SchemeColor, NAVY = C.accent1, RED = C.accent3, MUTED = C.accent5, WHITE = C.background1, TINT = C.background2, INK = C.text1;
const W = 13.333, M = 0.45;
pres.defineSlideMaster({ title: "MATRIZ", background: { color: WHITE }, objects: [
  { placeholder: { options: { name: "title", type: "title", x: M, y: 0.25, w: W - 2 * M - 2.4, h: 0.45, fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: INK, align: "left", valign: "top", margin: 0 }, text: "Título" } },
  { image: { path: "logo_comedica.png", x: W - M - 2.15, y: 0.12, w: 2.15, h: 0.66 } },
  ] 
});

const PEND = t => t.replace(/Comentario CA: 14\/10\/2026 Compromiso/g, "Pendiente");
const [HEAD, ...BODY] = ROWS.map(r => r.length === 3 ? [r[0], PEND(r[1]), r[2]] : r);
const splitAt = BODY.findIndex(r => r.length === 1 && r[0].startsWith("Dar seguimiento"));
const BORDER = [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: "D9E0EA" }, { type: "none" }];
const head = (col) => HEAD.map(t => ({ text: t.replace(/\n/g, " "), options: { bold: true, color: WHITE, fill: { color: col }, fontSize: 9 } }));
const lines = t => t.split("\n").map((p, i, a) => ({ text: p, options: { breakLine: i < a.length - 1 } }));
function toRow(r, num) {
  if (r.length === 1) return [{ text: lines(r[0]), options: { colspan: 3, bold: true, color: NAVY, fill: { color: TINT } } }];
  const [o, c, f] = r;
  const est = /Estado[^:]*:\s*(.*)/i.exec(o);
  const tint = est ? (/no superada/i.test(est[1]) ? "FBEAE8" : "FDF4E3") : "FFFFFF";
  const oRuns = o.split("\n").map((p, i, a) => ({ text: (i === 0 && num ? num + "  " : "") + p, options: { breakLine: i < a.length - 1, bold: i === 0, color: /^Estado/i.test(p) ? (/no superada/i.test(p) ? RED : C.accent6) : INK } }));
  return [
    { text: oRuns, options: { fill: { color: tint } } },
    { text: lines(c), options: { color: INK } },
    { text: lines(f), options: { color: NAVY, bold: true, align: "left" } },
  ];
}
function slide(title, rows, col, colW, fs) {
  const s = pres.addSlide({ masterName: "MATRIZ" });
  s.addText(title, { placeholder: "title" });
  s.addTable([head(col), ...rows.map(r => toRow(r))], { objectName: "Tabla matriz", x: M, y: 0.8, w: W - 2 * M, colW, fontFace: THEME.bodyFontFace, fontSize: fs, valign: "middle", align: "justify", margin: [0.02, 0.06, 0.02, 0.06], border: BORDER, autoPage: false });
  return s;
}
const split3 = BODY.findIndex(r => r.length === 1 && r[0].startsWith("C.3. Evaluación"));
// Diapositiva 1: secciones A y B en cuadros separados y numerados
const s1 = pres.addSlide({ masterName: "MATRIZ" });
s1.addText(TITLE, { placeholder: "title" });
const secIdx = BODY.slice(0, splitAt).map((r, i) => r.length === 1 ? i : -1).filter(i => i >= 0);
const tableOpts = { x: M, w: W - 2 * M, colW: [4.35, 6.33, 1.75], fontFace: THEME.bodyFontFace, fontSize: FS1, valign: "middle", align: "justify", margin: [0.015, 0.06, 0.015, 0.06], border: BORDER, autoPage: false };
let y = 0.76;
secIdx.forEach((si, k) => {
  const end = k + 1 < secIdx.length ? secIdx[k + 1] : splitAt;
  const letter = BODY[si][0].trim().charAt(0);
  const rows = BODY.slice(si + 1, end).map((r, n) => toRow(r, letter + "." + (n + 1)));
  s1.addText(BODY[si][0], { objectName: "Sección " + letter, isTextBox: true, x: M, y, w: W - 2 * M, h: 0.26, fontSize: 11, bold: true, color: NAVY, margin: 0, valign: "middle" });
  s1.addTable([head(NAVY), ...rows], { objectName: "Tabla sección " + letter, y: y + 0.3, ...tableOpts });
  y += 0.3 + SECH[letter] + 0.12;
});
s1.addNotes("Matriz de requerimientos DS-SABAO-BCS-BC-23486 (versión corta). Secciones A y B. 'Comentario CA: 14/10/2026 Compromiso' se muestra como 'Pendiente'.");
const s2 = slide(TITLE, BODY.slice(splitAt), RED, [5.4, 5.53, 1.5], FS2);
s2.addNotes("Texto tal como aparece en la Matriz de requerimientos DS-SABAO-BCS-BC-23486 (versión corta). Seguimiento a observaciones de visitas anteriores.");
(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("written", OUT); })();
