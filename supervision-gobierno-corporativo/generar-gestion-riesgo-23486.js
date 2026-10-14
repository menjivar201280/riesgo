// Presentación Matriz 23486 con el mismo estilo que "Supervisión GC SSF COMEDICA"
const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/eb134167-ea2a-4bd7-8c34-83283d88af55_b18ad4a5-0fa1-4688-82cb-4bf483ec9783/pptx/scripts/apply_theme.js");
const D = require("./mz.json");
const OUT = process.argv[2] || "deck.pptx";
const THEME = { name: "SSF Gobierno Corporativo", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "142033", lt1: "FFFFFF", dk2: "173A6E", lt2: "EEF2F8", accent1: "173A6E", accent2: "1D7A4B", accent3: "B42318", accent4: "D99A1E", accent5: "586579", accent6: "A86410", hlink: "173A6E", folHlink: "586579" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Matriz de requerimientos DS-SABAO-BCS-BC-23486";
const C = pres.SchemeColor;
const NAVY = C.accent1, GREEN = C.accent2, RED = C.accent3, AMBER = C.accent4, MUTED = C.accent5, AMBER_INK = C.accent6;
const WHITE = C.background1, TINT = C.background2, INK = C.text1, TEAL = "0E7490";
const W = 13.333, M = 0.5, PANEL = 2.75;
const CODE = "DS-SABAO-BCS-BC-23486";
const FOOT = "SSF · Gestión de Riesgo Específico · COMEDICA de R.L.";

pres.defineSlideMaster({ title: "PANEL", background: { color: WHITE },
  objects: [{ rect: { x: 0, y: 0, w: PANEL, h: 7.5, fill: { color: NAVY }, line: { color: NAVY } } },
    { text: { text: FOOT, options: { x: PANEL + 0.45, y: 7.08, w: 6, h: 0.25, fontSize: 9, color: MUTED, margin: 0 } } }],
  slideNumber: { x: W - M - 0.5, y: 7.08, w: 0.5, h: 0.25, fontSize: 9, color: MUTED, align: "right" } });
pres.defineSlideMaster({ title: "CONTENIDO", background: { color: WHITE },
  objects: [{ text: { text: FOOT, options: { x: M, y: 7.12, w: 6, h: 0.25, fontSize: 9, color: MUTED, margin: 0 } } }],
  slideNumber: { x: W - M - 0.5, y: 7.12, w: 0.5, h: 0.25, fontSize: 9, color: MUTED, align: "right" } });

const card = (s, name, x, y, w, h, fill = TINT, extra = {}) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name, x, y, w, h, rectRadius: 0.1, fill: { color: fill, ...(extra.transparency != null ? { transparency: extra.transparency } : {}) }, line: extra.line || { color: fill, transparency: 100 } });
const txt = (s, name, text, o) => s.addText(text, { objectName: name, isTextBox: true, margin: 0, fontSize: 12, color: INK, valign: "top", ...o });
const pill = (s, name, text, x, y, w, color, fs = 9.5, h = 0.28) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name + " fondo", x, y, w, h, rectRadius: h / 2, fill: { color }, line: { color, transparency: 100 } });
  s.addText(text, { objectName: name, isTextBox: true, x, y, w, h, margin: 0, align: "center", valign: "middle", fontSize: fs, bold: true, color: WHITE, charSpacing: 0.5 });
};
const badge = (s, name, n, x, y, d, color, fs) => {
  s.addShape(pres.shapes.OVAL, { objectName: name + " círculo", x, y, w: d, h: d, fill: { color }, line: { color, transparency: 100 } });
  s.addText(String(n), { objectName: name, isTextBox: true, x, y, w: d, h: d, margin: 0, align: "center", valign: "middle", fontSize: fs || 11, bold: true, color: WHITE, fontFace: THEME.headFontFace });
};
const logo = (s, x, y, w) => s.addImage({ path: "logo.png", x, y, w, h: w * 0.66 / 2.15, objectName: "Logo COMEDICA" });
function header(s, eyebrow, title) {
  txt(s, "Eyebrow", eyebrow, { x: M, y: 0.32, w: 9.5, h: 0.26, fontSize: 10.5, bold: true, color: MUTED, charSpacing: 2 });
  txt(s, "Título", title, { x: M, y: 0.6, w: 9.5, h: 0.6, fontSize: 28, bold: true, fontFace: THEME.headFontFace, valign: "middle" });
  logo(s, W - M - 2.0, 0.28, 2.0);
}

// ---------- Datos ----------
const nuevas = D.filter(r => r.slide === 0 && r.obs);
const seg = []; let secCtx = [];
D.filter(r => r.slide === 1).forEach(r => { if (r.sec) secCtx.push(r.sec); else seg.push({ ...r, ctx: [...secCtx] }); });
const estadoDe = o => { const m = /Estado[^:]*:\s*(.*)/i.exec(o); return m ? m[1].trim().replace(/\.$/, "") : ""; };
const circDe = ctx => { const t = ctx.join(" "); const m = t.match(/DS-SABAO-BCS-(?:BC-)?\d+/g); return m ? m[m.length - 1] : ""; };
const subDe = ctx => { const last = ctx[ctx.length - 1] || ""; if (/NRP-36/.test(last)) return "NRP-36"; const mm = last.match(/C\.\d(\.\d)?\./); if (mm) return mm[0].replace(/\.$/, ""); return subOld(ctx); };
const subOld = ctx => { for (let i = ctx.length - 1; i >= 0; i--) { const m = ctx[i].match(/C\.\d(\.\d)?\.[^\n]*$/m); if (m && m[0].length < 60) return m[0].trim(); } return /NRP-36/.test(ctx.join(" ")) ? "NRP-36 · LDA/FT/FPADM" : ""; };
const noSup = seg.filter(r => /^no superada/i.test(estadoDe(r.obs))).length;
const parcial = seg.length - noSup;

// ============ 1. RESUMEN ============
{
  const s = pres.addSlide({ masterName: "PANEL" });
  txt(s, "Panel título", "MATRIZ DE\nREQUERIMIENTOS", { x: 0.2, y: 0.55, w: PANEL - 0.4, h: 0.9, fontSize: 15, bold: true, color: WHITE, charSpacing: 2, align: "center", fontFace: "Arial" });
  txt(s, "Panel subtítulo", "Gestión de Riesgo Específico", { x: 0.3, y: 1.5, w: PANEL - 0.6, h: 0.8, fontSize: 20, bold: true, color: WHITE, fontFace: THEME.headFontFace, align: "center" });
  txt(s, "Panel código", CODE, { x: 0.25, y: 2.45, w: PANEL - 0.5, h: 0.3, fontSize: 11, bold: true, color: TINT, align: "center" });
  card(s, "Logo fondo", 0.3, 4.5, PANEL - 0.6, 0.85, WHITE);
  logo(s, 0.38, 4.58, PANEL - 0.76);

  const RX = PANEL + 0.45, RW = W - M - RX;
  // Bloque observaciones nuevas
  const rows = [
    { y: 0.4, title: "Observaciones nuevas", sub: "Visita de supervisión · Gestión de Riesgo Específico", color: GREEN, big: String(nuevas.length), bigLab: "OBSERVACIONES NUEVAS",
      items: [["A", "Gestión de Riesgo de LDA/FT/FPADM", "2 observaciones", TEAL], ["B", "Proceso de implementación de la Herramienta de Monitoreo", "15 observaciones", GREEN]] },
    { y: 3.72, title: "Seguimiento a visitas anteriores", sub: "", color: RED, big: String(seg.length), bigLab: "EN SEGUIMIENTO",
      items: [["16572", "Riesgos de Crédito, Operacional y Tecnológico · 10/09/2025", "3 observaciones", NAVY], ["18445", "NRP-36 · LDA/FT/FPADM · 30/10/2025", "2 observaciones", NAVY], ["29295", "Guías de Calidad · 03/12/2025", "4 observaciones", NAVY]] },
  ];
  rows.forEach((r, k) => {
    txt(s, "Bloque " + k, r.title, { x: RX, y: r.y, w: 6, h: 0.45, fontSize: 24, bold: true, color: r.color, fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Bloque sub " + k, r.sub, { x: RX + 5.2, y: r.y, w: RW - 5.2, h: 0.45, fontSize: 12, color: MUTED, valign: "middle", align: "right" });
    const cy = r.y + 0.6, ch = 2.5;
    card(s, "Tarjeta " + k, RX, cy, RW, ch, WHITE, { line: { color: "D9E0EA", width: 1 } });
    card(s, "Cifra fondo " + k, RX + 0.15, cy + 0.15, 1.6, ch - 0.3, r.color);
    txt(s, "Cifra " + k, r.big, { x: RX + 0.15, y: cy + 0.55, w: 1.6, h: 0.8, fontSize: 48, bold: true, color: WHITE, align: "center", fontFace: THEME.headFontFace, valign: "middle" });
    txt(s, "Cifra lab " + k, r.bigLab, { x: RX + 0.2, y: cy + 1.4, w: 1.5, h: 0.5, fontSize: 9, bold: true, color: WHITE, align: "center", charSpacing: 1 });
    const ix = RX + 2.05, iw = RW - 2.25, n = r.items.length, ih = (ch - 0.3) / n;
    r.items.forEach(([tag, t, c, col], i) => {
      const iy = cy + 0.15 + i * ih;
      if (i) s.addShape(pres.shapes.LINE, { x: ix, y: iy, w: iw, h: 0, line: { color: "D9E0EA", width: 0.75 } });
      const tw = tag.length === 1 ? 0.5 : 1.75;
      pill(s, "Etiqueta " + k + i, tag.length === 1 ? tag : "Circular " + tag, ix, iy + ih / 2 - 0.16, tw, col, tag.length === 1 ? 13 : 10, 0.32);
      txt(s, "Item " + k + i, t, { x: ix + tw + 0.2, y: iy, w: iw - tw - 2.1, h: ih, fontSize: 14, bold: true, valign: "middle" });
      txt(s, "Item cifra " + k + i, c, { x: ix + iw - 1.8, y: iy, w: 1.8, h: ih, fontSize: 13, color: MUTED, valign: "middle", align: "right" });
    });
  });
  // estados
  pill(s, "No superadas", `${noSup} NO SUPERADAS`, W - M - 4.35, 3.8, 1.7, RED, 9.5, 0.32);
  pill(s, "Parciales", `${parcial} PARCIALMENTE SUPERADAS`, W - M - 2.55, 3.8, 2.55, AMBER, 9.5, 0.32);
  s.addNotes("Matriz de requerimientos DS-SABAO-BCS-BC-23486. 17 observaciones nuevas (A: 2, B: 15) y 9 observaciones de visitas anteriores en seguimiento (Circulares 16572, 18445 y 29295): " + noSup + " no superadas y " + parcial + " parcialmente superadas.");
}

// ============ Utilidades de cuadrícula con altura variable ============
const CH = fs => fs * 0.0064;           // ancho medio de carácter (pulgadas) en Calibri
const linesOf = (t, fs, w) => t.split("\n").reduce((a, l) => a + Math.max(1, Math.ceil(l.length * CH(fs) / w)), 0);
function fitFs(items, cols, y0, y1, hOfFs, start) {
  const gap = 0.14, cw = (W - 2 * M - (cols - 1) * gap) / cols;
  for (let fs = start; fs > 7; fs -= 0.25) {
    const rows = []; for (let i = 0; i < items.length; i += cols) rows.push(items.slice(i, i + cols));
    const tot = rows.map(r => Math.max(...r.map(it => hOfFs(it, cw, fs)))).reduce((a, b) => a + b, 0);
    if (tot <= y1 - y0 - gap * (rows.length - 1)) return fs;
  }
  return 7;
}
function grid(items, cols, y0, y1, heightOf, draw) {
  const gap = 0.14, cw = (W - 2 * M - (cols - 1) * gap) / cols, rows = [];
  for (let i = 0; i < items.length; i += cols) rows.push(items.slice(i, i + cols));
  const need = rows.map(r => Math.max(...r.map(it => heightOf(it, cw))));
  const total = need.reduce((a, b) => a + b, 0), avail = y1 - y0 - gap * (rows.length - 1);
  const k = avail / total; let y = y0;
  rows.forEach((r, ri) => { const h = need[ri] * k; r.forEach((it, ci) => draw(it, M + ci * (cw + gap), y, cw, h)); y += h + gap; });
  return k;
}

// ============ 2–3. OBSERVACIONES NUEVAS (A y B separadas) ============
const parseN = r => { const m = /^([AB]\.\d+)\s+([\s\S]*)$/.exec(r.obs); return { num: m[1], t: m[2], r }; };
const SECT = { A: ["A. Gestión de Riesgo de LDA/FT/FPADM", TEAL], B: ["B. Proceso de implementación de la Herramienta de Monitoreo", GREEN] };
function nuevasSlide(groups, eyebrow, title, fs0) {
  const s = pres.addSlide({ masterName: "CONTENIDO" });
  header(s, eyebrow, title);
  const gap = 0.14, BAND = 0.36, top = 1.32, bottom = 7.0;
  const cwOf = cols => (W - 2 * M - (cols - 1) * gap) / cols;
  const hdrH = (it, cw) => Math.max(0.56, 0.2 + linesOf(it.r.resp, 10, cw - 0.84) * 0.17);
  const hF = (it, cw, fs) => hdrH(it, cw) + 0.16 + linesOf(it.t, fs, cw - 0.32) * fs * 0.0175 + 0.12 + linesOf("Comentarios  " + it.r.com.replace(/\n/g, " "), fs - 1, cw - 0.32) * (fs - 1) * 0.0172 + 0.18;
  const rowsOf = g => { const out = []; for (let i = 0; i < g.items.length; i += g.cols) out.push(g.items.slice(i, i + g.cols)); return out; };
  const fixed = groups.length * (BAND + 0.08) + (groups.length - 1) * 0.16;
  let fs = fs0;
  for (; fs > 7; fs -= 0.25) {
    let need = 0, nRows = 0;
    groups.forEach(g => rowsOf(g).forEach(r => { need += Math.max(...r.map(it => hF(it, cwOf(g.cols), fs))); nRows++; }));
    if (need <= bottom - top - fixed - gap * (nRows - groups.length)) break;
  }
  let need = 0, nRows = 0;
  groups.forEach(g => rowsOf(g).forEach(r => { need += Math.max(...r.map(it => hF(it, cwOf(g.cols), fs))); nRows++; }));
  const k = (bottom - top - fixed - gap * (nRows - groups.length)) / need;
  let y = top;
  groups.forEach((g, gi) => {
    const [label, col] = SECT[g.key];
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: "Banda " + g.key, x: M, y, w: W - 2 * M, h: BAND, rectRadius: 0.06, fill: { color: col }, line: { color: col, transparency: 100 } });
    txt(s, "Banda texto " + g.key + gi, label + (g.cont ? " (continuación)" : ""), { x: M + 0.18, y, w: W - 2 * M - 0.36, h: BAND, fontSize: 13, bold: true, color: WHITE, valign: "middle" });
    y += BAND + 0.08;
    const cw = cwOf(g.cols);
    rowsOf(g).forEach(row => {
      const h = Math.max(...row.map(it => hF(it, cw, fs))) * k;
      row.forEach((it, ci) => {
        const x = M + ci * (cw + gap), { num, t, r } = it;
        card(s, "Obs " + num, x, y, cw, h, TINT);
        badge(s, "Número " + num, num, x + 0.12, y + 0.12, 0.5, col, 11);
        const hh = hdrH(it, cw);
        txt(s, "Responsable " + num, [{ text: "RESPONSABLE", options: { bold: true, color: MUTED, fontSize: 8.5, breakLine: true } }, { text: r.resp, options: { fontSize: 10 } }], { x: x + 0.72, y: y + 0.08, w: cw - 0.84, h: hh, valign: "middle" });
        const th = linesOf(t, fs, cw - 0.32) * fs * 0.0175 + 0.04, ty = y + 0.08 + hh + 0.1;
        txt(s, "Título " + num, t, { x: x + 0.16, y: ty, w: cw - 0.32, h: th, fontSize: fs, bold: true, color: NAVY, align: "justify" });
        const pend = r.com.trim() === "Pendiente";
        txt(s, "Comentario " + num, [{ text: "Comentarios  ", options: { bold: true, color: MUTED } }, { text: r.com.replace(/\n/g, " "), options: { color: pend ? AMBER_INK : INK, bold: pend } }],
          { x: x + 0.16, y: ty + th + 0.08, w: cw - 0.32, h: y + h - (ty + th + 0.08) - 0.1, fontSize: fs - 1, align: "justify" });
      });
      y += h + gap;
    });
    y += 0.16 - gap;
  });
  console.log("nuevas", title, "fs", fs, "k", k.toFixed(2));
  return s;
}
const NA = nuevas.map(parseN).filter(x => x.num[0] === "A"), NB = nuevas.map(parseN).filter(x => x.num[0] === "B");
const s2 = nuevasSlide([{ key: "A", items: NA, cols: 2 }, { key: "B", items: NB.slice(0, 6), cols: 3 }], "CIRCULAR " + CODE + " · OBSERVACIONES NUEVAS", "Observaciones nuevas", 13);
s2.addNotes("A. Gestión de Riesgo de LDA/FT/FPADM (A.1–A.2) y B. Proceso de implementación de la Herramienta de Monitoreo (B.1–B.6). Texto de la matriz.");
const s3 = nuevasSlide([{ key: "B", items: NB.slice(6, 11), cols: 3, cont: true }], "CIRCULAR " + CODE + " · OBSERVACIONES NUEVAS", "Observaciones nuevas", 13);
s3.addNotes("B. Proceso de implementación de la Herramienta de Monitoreo (B.7–B.11). Texto de la matriz.");
const s3b = nuevasSlide([{ key: "B", items: NB.slice(11), cols: 2, cont: true }], "CIRCULAR " + CODE + " · OBSERVACIONES NUEVAS", "Observaciones nuevas", 13);
s3b.addNotes("B. Proceso de implementación de la Herramienta de Monitoreo (B.12–B.15). Texto de la matriz.");

// ============ 4–5. SEGUIMIENTO ============
function segSlide(list, title, cols, fs0) {
  const s = pres.addSlide({ masterName: "CONTENIDO" });
  header(s, "CIRCULAR " + CODE + " · SEGUIMIENTO A OBSERVACIONES DE VISITAS ANTERIORES", title);
  const hF = (r, cw, fs) => { const L = r.obs.split("\n").filter(l => !/^Estado/i.test(l));
    return 0.5 + linesOf(L[0], fs, cw - 0.3) * fs * 0.0175 + linesOf(L.slice(1).join("\n") || "", fs - 1, cw - 0.3) * (fs - 1) * 0.017 + 0.1
      + linesOf("Comentarios  " + r.com.replace(/\n/g, " "), fs - 1, cw - 0.3) * (fs - 1) * 0.0172 + 0.42; };
  const fs = fitFs(list, cols, 1.35, 7.0, hF, fs0); const hOf = (r, cw) => hF(r, cw, fs);
  const k = grid(list, cols, 1.35, 7.0, hOf, (r, x, y, cw, h) => {
    const i = seg.indexOf(r), est = estadoDe(r.obs), no = /^no superada/i.test(est), col = no ? RED : AMBER;
    card(s, "Seg " + i, x, y, cw, h, col, { transparency: 88 });
    const circ = circDe(r.ctx).replace("DS-SABAO-BCS-", "").replace("BC-", "");
    pill(s, "Circular " + i, "Circular " + circ + " · " + subDe(r.ctx), x + 0.12, y + 0.1, 1.75, NAVY, 8.5, 0.26);
    pill(s, "Estado " + i, est.toUpperCase(), x + 1.95, y + 0.1, no ? 1.15 : 1.85, col, 8.5, 0.26);
    const L = r.obs.split("\n").filter(l => !/^Estado/i.test(l));
    const oh = linesOf(L[0], fs, cw - 0.3) * fs * 0.0175 + linesOf(L.slice(1).join("\n") || "", fs - 1, cw - 0.3) * (fs - 1) * 0.017 * (L.length > 1 ? 1 : 0) + 0.06;
    txt(s, "Obs " + i, L.map((l, kk) => ({ text: l, options: { bold: kk === 0, color: kk === 0 ? NAVY : INK, fontSize: kk === 0 ? fs : fs - 1, breakLine: kk < L.length - 1 } })), { x: x + 0.15, y: y + 0.46, w: cw - 0.3, h: oh, align: "justify" });
    txt(s, "Com " + i, [{ text: "Comentarios  ", options: { bold: true, color: MUTED } }, { text: r.com.replace(/\n/g, " ") }], { x: x + 0.15, y: y + 0.46 + oh + 0.06, w: cw - 0.3, h: h - 0.46 - oh - 0.42, fontSize: fs - 1, align: "justify" });
    txt(s, "Resp " + i, [{ text: "Responsable  ", options: { bold: true, color: MUTED } }, { text: r.resp, options: { bold: true, color: NAVY } }], { x: x + 0.15, y: y + h - 0.32, w: cw - 0.3, h: 0.24, fontSize: fs - 0.5, valign: "middle" });
  });
  console.log("seg", title, "fs", fs, "k", k.toFixed(2));
  return s;
}
const g1 = seg.filter(r => !/29295/.test(circDe(r.ctx))), g2 = seg.filter(r => /29295/.test(circDe(r.ctx)));
const s4 = segSlide(g1, "Seguimiento · Circulares 16572 y 18445", 3, 12);
s4.addNotes(D.filter(r => r.sec && !/29295|C\.3/.test(r.sec)).map(r => r.sec).join("\n\n"));
const s5 = segSlide(g2, "Seguimiento · Circular 29295", 2, 13);
s5.addNotes(D.filter(r => r.sec && /29295|C\.3/.test(r.sec)).map(r => r.sec).join("\n\n"));

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("written", OUT); })();
