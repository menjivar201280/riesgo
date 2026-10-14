const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/eb134167-ea2a-4bd7-8c34-83283d88af55_b18ad4a5-0fa1-4688-82cb-4bf483ec9783/pptx/scripts/apply_theme.js");
const ROWS = require("./m23486_rows.json");
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
  { placeholder: { options: { name: "title", type: "title", x: M, y: 0.25, w: W - 2 * M, h: 0.45, fontFace: THEME.headFontFace, fontSize: 20, bold: true, color: INK, align: "left", valign: "top", margin: 0 }, text: "Título" } },
  { text: { text: "COMEDICA de R.L.", options: { x: M, y: 7.2, w: 6, h: 0.2, fontSize: 8.5, color: MUTED, margin: 0 } } } ],
  slideNumber: { x: W - M - 0.5, y: 7.2, w: 0.5, h: 0.2, fontSize: 8.5, color: MUTED, align: "right" } });

const [HEAD, ...BODY] = ROWS;
const splitAt = BODY.findIndex(r => r.length === 1 && r[0].startsWith("Dar seguimiento"));
const BORDER = [{ type: "none" }, { type: "none" }, { type: "solid", pt: 0.75, color: "D9E0EA" }, { type: "none" }];
const head = (col) => HEAD.map(t => ({ text: lines(t), options: { bold: true, color: WHITE, fill: { color: col }, fontSize: 9 } }));
const lines = t => t.split("\n").map((p, i, a) => ({ text: p, options: { breakLine: i < a.length - 1 } }));
function toRow(r) {
  if (r.length === 1) return [{ text: lines(r[0]), options: { colspan: 3, bold: true, color: NAVY, fill: { color: TINT }, fontSize: 7 } }];
  const [o, c, f] = r;
  const est = /Estado[^:]*:\s*(.*)/i.exec(o);
  const tint = est ? (/no superada/i.test(est[1]) ? "FBEAE8" : "FDF4E3") : "FFFFFF";
  const oRuns = o.split("\n").map((p, i, a) => ({ text: p, options: { breakLine: i < a.length - 1, bold: i === 0, color: /^Estado/i.test(p) ? (/no superada/i.test(p) ? RED : C.accent6) : INK } }));
  return [
    { text: oRuns, options: { fill: { color: tint } } },
    { text: lines(c), options: { color: INK } },
    { text: lines(f), options: { color: NAVY, bold: true } },
  ];
}
function slide(title, rows, col, colW, fs) {
  const s = pres.addSlide({ masterName: "MATRIZ" });
  s.addText(title, { placeholder: "title" });
  s.addTable([head(col), ...rows.map(toRow)], { objectName: "Tabla matriz", x: M, y: 0.8, w: W - 2 * M, colW, fontFace: THEME.bodyFontFace, fontSize: fs, valign: "middle", margin: [0.02, 0.06, 0.02, 0.06], border: BORDER, autoPage: false });
  return s;
}
const s1 = slide("Matriz de requerimientos DS-SABAO-BCS-BC-23486", BODY.slice(0, splitAt), NAVY, [4.9, 5.6, 1.93], 7.5);
s1.addNotes("Texto tal como aparece en la Matriz de requerimientos DS-SABAO-BCS-BC-23486 (versión corta). Secciones A y B.");
const s2 = slide("Matriz de requerimientos DS-SABAO-BCS-BC-23486", BODY.slice(splitAt), RED, [5.4, 5.43, 1.6], 7);
s2.addNotes("Texto tal como aparece en la Matriz de requerimientos DS-SABAO-BCS-BC-23486 (versión corta). Seguimiento a observaciones de visitas anteriores.");
(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("written", OUT); })();
