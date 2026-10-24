const pptxgen = require('pptxgenjs');
const React = require('react');
const RDS = require('react-dom/server');
const sharp = require('sharp');
const fa = require('react-icons/fa');
const { applyTheme } = require(process.argv[2] + '/scripts/apply_theme.js');
const OUT = process.argv[3];

const NAVY = '0F1B3D', INK = '1F2937', MUTED = '6B7280', TRACK = 'E5E9F0', BG = 'F3F5F9', GREEN = '16A34A';
const THEME = { name: 'Riesgos', headFontFace: 'Calibri', bodyFontFace: 'Calibri',
  colors: { dk1: '1F2937', lt1: 'FFFFFF', dk2: NAVY, lt2: BG, accent1: '0E9F8E', accent2: '2F6FDE',
    accent3: '7A4FD3', accent4: 'E07B12', accent5: GREEN, accent6: '6B7280', hlink: '2F6FDE', folHlink: '7A4FD3' } };

async function icon(Comp, color, size = 256) {
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: '#' + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

const CATS = [
  { n: 1, t: 'Cumplimiento Regulatorio', s: 'Requerimientos SSF y acuerdos', c: '0E9F8E', tint: 'E3F5F2', ic: fa.FaBalanceScale,
    items: [['Req. adicional SSF · 22742', 90], ['Req. DS-SABAO-BCS-BC-23486', 100], ['Gobierno Corporativo · 22741', 100],
            ['Recalificación · Asociado 22494', 100], ['Servicios Tercerizados', 100], ['Período de gracia · Asociado 25572', 100]] },
  { n: 2, t: 'Riesgo Operacional', s: 'Gestión y continuidad', c: '2F6FDE', tint: 'E6EEFC', ic: fa.FaCogs,
    items: [['Plan de Riesgo Operacional 2026', 45], ['3.ª prueba de continuidad', 45], ['Eventos de riesgo informados', 67],
            ['Caídas TC y Orion', 83], ['Comité de Riesgos', 100], ['Boletín de riesgos', 100],
            ['Capacitación / Examen GIR', 100], ['Piloto Contact Center', 100]] },
  { n: 3, t: 'Riesgo Normativo', s: 'Avalúos', c: '7A4FD3', tint: 'EFE9FB', ic: fa.FaFileSignature,
    items: [['Control de Avalúos', 90], ['Acta de sesión 29/08/2026', 90], ['Informe Comité para Consejo', 95],
            ['Actualización de avalúos', 98], ['Manuales Internos', 100], ['Junta de Vigilancia', 100]] },
  { n: 4, t: 'Riesgo Financiero', s: 'Activos y modelos de riesgo', c: 'E07B12', tint: 'FDF0E2', ic: fa.FaChartLine,
    items: [['Proyectos a largo plazo', 63], ['Otros / Monitoreo', 76], ['Riesgo Crediticio', 90],
            ['Liquidez y Mercado', 97], ['Activos de Riesgo', 100], ['Opiniones y análisis crediticios', 100]] },
];

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.theme = { headFontFace: 'Calibri', bodyFontFace: 'Calibri' };
  pres.title = 'Informe semanal · Administración de Riesgos';
  const s = pres.addSlide();
  s.background = { color: BG };

  const all = CATS.flatMap(c => c.items);
  const total = all.length, done = all.filter(i => i[1] === 100).length;
  const avg = all.reduce((a, i) => a + i[1], 0) / total;
  const pending = total - done;

  // Title block
  s.addText('INFORME SEMANAL · AVANCE DE ACTIVIDADES', { x: 0.45, y: 0.38, w: 7, h: 0.3, margin: 0, fontSize: 12, bold: true,
    color: '2F6FDE', charSpacing: 2, isTextBox: true });
  s.addText('Departamento de Administración de Riesgos', { x: 0.45, y: 0.74, w: 7.35, h: 0.5, margin: 0, fontSize: 24, bold: true,
    color: NAVY, isTextBox: true, fit: 'none' });

  // KPI tiles
  const kpis = [
    { v: String(total), l: 'Actividades', ic: fa.FaTasks, c: NAVY },
    { v: avg.toFixed(1).replace('.', ',') + '%', l: 'Avance prom.', ic: fa.FaChartPie, c: '2F6FDE' },
    { v: String(done), l: 'Completadas', ic: fa.FaCheckCircle, c: GREEN },
    { v: String(pending), l: 'En proceso', ic: fa.FaHourglassHalf, c: 'E07B12' },
  ];
  const kw = 1.15, kg = 0.13, kx0 = 13.333 - 0.45 - (4 * kw + 3 * kg);
  for (let i = 0; i < kpis.length; i++) {
    const k = kpis[i], x = kx0 + i * (kw + kg), y = 0.36;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: kw, h: 0.98, rectRadius: 0.1, fill: { color: 'FFFFFF' },
      line: { color: 'E3E7EE', width: 0.75 }, shadow: { type: 'outer', color: '000000', opacity: 0.08, blur: 6, offset: 2, angle: 90 } });
    s.addImage({ data: await icon(k.ic, k.c), x: x + 0.14, y: y + 0.15, w: 0.24, h: 0.24 });
    s.addText(k.v, { x: x + 0.14, y: y + 0.38, w: kw - 0.2, h: 0.36, margin: 0, fontSize: 20, bold: true, color: k.c, isTextBox: true });
    s.addText(k.l, { x: x + 0.14, y: y + 0.72, w: kw - 0.2, h: 0.2, margin: 0, fontSize: 10, color: MUTED, isTextBox: true });
  }

  // Category cards
  const mx = 0.45, gap = 0.2, cw = (13.333 - 2 * mx - 3 * gap) / 4, cy = 1.62, ch = 5.45;
  for (let i = 0; i < CATS.length; i++) {
    const c = CATS[i], x = mx + i * (cw + gap);
    const cavg = c.items.reduce((a, it) => a + it[1], 0) / c.items.length;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: cy, w: cw, h: ch, rectRadius: 0.12, fill: { color: 'FFFFFF' },
      line: { color: 'E3E7EE', width: 0.75 }, shadow: { type: 'outer', color: '000000', opacity: 0.08, blur: 8, offset: 2, angle: 90 } });
    // header tint
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.1, y: cy + 0.1, w: cw - 0.2, h: 1.12, rectRadius: 0.1, fill: { color: c.tint }, line: { color: c.tint } });
    s.addShape(pres.shapes.OVAL, { x: x + 0.22, y: cy + 0.22, w: 0.5, h: 0.5, fill: { color: c.c }, line: { color: c.c } });
    s.addImage({ data: await icon(c.ic, 'FFFFFF'), x: x + 0.345, y: cy + 0.345, w: 0.25, h: 0.25 });
    s.addText(c.t.toUpperCase(), { x: x + 0.84, y: cy + 0.22, w: cw - 1.0, h: 0.5, margin: 0, fontSize: 13, bold: true,
      color: NAVY, valign: 'middle', isTextBox: true });
    s.addText(c.s, { x: x + 0.22, y: cy + 0.8, w: cw - 1.3, h: 0.24, margin: 0, fontSize: 10, color: c.c, bold: true, isTextBox: true });
    s.addText(Math.round(cavg) + '%', { x: x + cw - 1.05, y: cy + 0.62, w: 0.85, h: 0.36, margin: 0, fontSize: 22, bold: true,
      color: c.c, align: 'right', valign: 'middle', isTextBox: true });
    s.addText(c.items.length + ' actividades', { x: x + cw - 1.25, y: cy + 0.97, w: 1.05, h: 0.18, margin: 0, fontSize: 9,
      color: MUTED, align: 'right', isTextBox: true });

    // rows
    const ix = x + 0.25, iw = cw - 0.5, ry0 = cy + 1.42, avail = ch - 1.42 - 0.12;
    const rh = Math.min(0.68, avail / c.items.length);
    for (let j = 0; j < c.items.length; j++) {
      const [name, p] = c.items[j], y = ry0 + j * rh, ok = p === 100, col = ok ? GREEN : c.c;
      s.addText(name, { x: ix, y, w: iw, h: 0.26, margin: 0, fontSize: 11, bold: true, color: INK, isTextBox: true, valign: 'middle' });
      s.addText(p + '%', { x: ix + iw - 0.6, y: y + 0.24, w: ok ? 0.42 : 0.6, h: 0.22, margin: 0, fontSize: 11, bold: true, color: col,
        align: 'right', valign: 'middle', isTextBox: true });
      if (ok) s.addImage({ data: await icon(fa.FaCheckCircle, GREEN), x: ix + iw - 0.14, y: y + 0.28, w: 0.14, h: 0.14 });
      const by = y + 0.31, bw = iw - 0.62;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix, y: by, w: bw, h: 0.08, rectRadius: 0.04, fill: { color: TRACK }, line: { color: TRACK } });
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ix, y: by, w: bw * p / 100, h: 0.08, rectRadius: 0.04, fill: { color: col }, line: { color: col } });
    }
  }

  // footer legend
  s.addText([
    { text: '● ', options: { color: GREEN } }, { text: 'Completada (100%)     ', options: { color: MUTED } },
    { text: '● ', options: { color: MUTED } }, { text: 'En proceso: color de la categoría     ', options: { color: MUTED } },
    { text: 'Porcentaje de cabecera = avance promedio de la categoría', options: { color: MUTED, italic: true } },
  ], { x: 0.45, y: 7.12, w: 12.4, h: 0.24, margin: 0, fontSize: 10, isTextBox: true });

  s.addNotes(`Resumen: ${total} actividades, avance promedio ${avg.toFixed(1)}%, ${done} completadas y ${pending} en proceso. ` +
    'Prioridades: Plan de Riesgo Operacional 2026 y 3.ª prueba de continuidad (45%), Proyectos a largo plazo (63%), Eventos de riesgo informados (67%).');

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
})();
