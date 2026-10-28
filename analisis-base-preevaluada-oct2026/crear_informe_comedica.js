// Genera Informe_Auditoria_Base_PreEvaluada_COMEDICA.docx
// Uso: node crear_informe_comedica.js <hallazgos.json> <controles.json> <salida.docx>
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, HeadingLevel,
  AlignmentType, LevelFormat, BorderStyle, Footer, PageNumber,
} = require('docx');

const [hjson, cjson, out] = process.argv.slice(2);
const { hallazgos, cambios } = JSON.parse(fs.readFileSync(hjson, 'utf8'));
const controles = JSON.parse(fs.readFileSync(cjson, 'utf8'));

const NAVY = '1F3864';
const FONT = 'Arial';
const W = 9360; // ancho útil carta con márgenes de 1"
const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };

const p = (text, opts = {}) => new Paragraph({
  spacing: { after: 120 }, ...opts.para,
  children: [new TextRun({ text, font: FONT, size: opts.size || 21, bold: opts.bold, italics: opts.italics, color: opts.color })],
});
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 140 },
  children: [new TextRun({ text: t, font: FONT, size: 28, bold: true, color: NAVY })] });
const bullet = (t) => new Paragraph({ numbering: { reference: 'vi', level: 0 }, spacing: { after: 60 },
  children: [new TextRun({ text: t, font: FONT, size: 21 })] });

function table(headers, rows, widths, fillCol) {
  const cell = (t, w, head, shade) => new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: head ? { fill: NAVY, type: ShadingType.CLEAR } : (shade ? { fill: shade, type: ShadingType.CLEAR } : undefined),
    margins: { top: 60, bottom: 60, left: 90, right: 90 },
    children: [new Paragraph({ children: [new TextRun({ text: String(t ?? ''), font: FONT, size: head ? 18 : 17, bold: head,
      color: head ? 'FFFFFF' : '000000' })] })],
  });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true)) }),
      ...rows.map((r) => new TableRow({ children: r.map((v, i) => cell(v, widths[i], false, fillCol ? fillCol(r, i) : null)) }))],
  });
}
const CRIT = { 'Crítica': 'F4B183', 'Alta': 'F8CBAD', 'Media': 'FFE699', 'Baja': 'C6E0B4' };
const RES = { OK: 'C6E0B4', 'OK (por diseño)': 'C6E0B4', ALERTA: 'FFE699', REVISAR: 'F8CBAD' };

const children = [
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'COMÉDICA', font: FONT, size: 22, bold: true, color: '2E75B6' })] }),
  new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: 'Informe de auditoría de la base preevaluada', font: FONT, size: 40, bold: true, color: NAVY })] }),
  p('Archivo auditado: Recomendacion_Base_PreEvaluada (1).xlsx · Base al 30/09/2026 · Memorando del 27/10/2026 · Fecha del informe: 28/10/2026',
    { italics: true, color: '595959', size: 19 }),
  p('Este informe explica en lenguaje sencillo qué se encontró en el archivo, qué se corrigió y qué decisiones quedan pendientes. '
    + 'La preevaluación es preliminar: no es una aprobación de crédito ni una autorización de desembolso.', { bold: true }),

  h1('1. Resumen'),
  bullet('El archivo auditado mostraba 492 asociados "recomendados", pero 57 de ellos tenían riesgo ALTO. Esa contradicción ya está resuelta: '
    + 'ningún asociado con riesgo ALTO queda como recomendado; todos pasan a "Requiere revisión adicional".'),
  bullet('El archivo no tenía ninguna fórmula: todo estaba pegado como valor. El archivo corregido tiene 239,405 fórmulas; si se pega una base nueva, '
    + 'todo se recalcula solo (resumen, niveles, listas, controles y gráficos).'),
  bullet('Resultado corregido sobre 2,262 registros: 528 cumplen los criterios iniciales del memorando; 425 recomendados preliminarmente; '
    + '103 requieren revisión adicional; 1,536 están pendientes por información insuficiente; 198 no son elegibles según el memorando.'),
  bullet('Todas las pruebas de control pasaron, salvo una ALERTA que se deja visible a propósito: 123 recomendados tienen un monto de vehículo '
    + 'fuera de la tabla oficial. Es una decisión que corresponde a la institución.'),

  h1('2. Debilidades detectadas'),
  p('La matriz completa está en la hoja 8_Hallazgos del Excel. Criticidad: Crítica = puede llevar a ofrecer crédito a quien no corresponde '
    + 'o a decidir con cifras equivocadas; Alta = afecta la clasificación de muchos registros; Media = afecta la interpretación o la '
    + 'trazabilidad; Baja = forma.', { size: 19, italics: true }),
  table(['N.º', 'Hoja', 'Hallazgo', 'Evidencia', 'Criticidad', 'Estado'],
    hallazgos.map((h, i) => [i + 1, h[0], h[1], h[2], h[3], h[6]]), [450, 1200, 2250, 2960, 1100, 1400],
    (r, i) => (i === 4 ? CRIT[r[4]] : null)),

  h1('3. Cifras verificadas: archivo original frente al corregido'),
  table(['Concepto', 'Original', 'Corregido', 'Explicación'], [
    ['Registros en la base', '2,262', '2,262', 'Se conservan todos; el memorando dice 2,263 (diferencia de 1, pendiente de la fuente).'],
    ['Cumplen criterios iniciales', '528', '528', 'Mismo criterio del memorando, ahora con fórmula.'],
    ['Recomendados', '492 (57 con riesgo ALTO)', '425 (0 con riesgo ALTO)', '67 pasan a revisión: 55 por riesgo ALTO y 12 por tarjeta con deuda alta.'],
    ['Con observaciones / revisión', '36', '103', 'Los 36 anteriores + 67 que salen de recomendados.'],
    ['Pendientes por información', 'No se distinguían', '1,536', 'Sobre todo ingreso no informado (1,283).'],
    ['No elegibles', 'No se distinguían', '198', 'Ingreso 0, fondo 0, categoría fuera de A1/A2 o edad.'],
    ['Hipotecario recomendado', '279 · $37,280,000', '264 · $35,265,000', 'Montos de la base, solo a recomendados.'],
    ['Vehículo recomendado', '364 · $8,785,000', '347 · $8,365,000', '123 fuera de su tabla (alerta).'],
    ['Tarjeta recomendada', '152 · $528,000', '105 · $414,500', '47 tarjetas más quedan "a evaluar" en revisión ($113,500).'],
  ], [2100, 1900, 1900, 3460]),
  p('Detalle de los 492 recomendados originales: 425 siguen recomendados y 67 pasan a revisión. De esos 67, 55 tienen riesgo ALTO '
    + '(29 que "podían pagar", 20 de "solo tarjeta" y 6 que "pagaban su deuda con el hipotecario") y 12 son de "solo tarjeta" con riesgo '
    + 'MEDIO pero deuda alta para su ingreso. El nivel de riesgo se recalculó con una regla más clara (punto 5).',
    { size: 19, italics: true, para: { spacing: { before: 120, after: 120 } } }),

  h1('4. Riesgos de usar el archivo sin corregir'),
  bullet('Ofrecer crédito a asociados con riesgo ALTO pensando que ya estaban "recomendados".'),
  bullet('Tomar decisiones con números que no se actualizan cuando cambia la base.'),
  bullet('Leer "Monto a ofrecer" como monto aprobado.'),
  bullet('Ofrecer vehículos por montos que la propia tabla de la Cooperativa no permite.'),
  bullet('Descartar a asociados a los que solo les falta un dato, igual que a quienes no cumplen.'),

  h1('5. Cambios efectuados'),
  table(['N.º', 'Cambio', 'Motivo', 'Impacto', 'Estado'], cambios.map((c, i) => [i + 1, c[0], c[1], c[3], c[5]]),
    [450, 2900, 2400, 2210, 1400]),
  p('Nivel de riesgo (propuesto): se cuentan los datos en nivel Alto entre antigüedad, ingreso, deuda SSF, cuotas de aportación vencidas '
    + 'y categoría SSF. Ninguno = BAJO; 1 = MEDIO; 2 o más = ALTO. El puntaje con los pesos 13%, 30%, 20%, 12% y 25% (suman 100%) se '
    + 'conserva como piloto: solo ordena la lista, no decide.', { para: { spacing: { before: 120, after: 120 } } }),

  h1('6. Criterios pendientes de aprobación institucional'),
  bullet('Rangos de ingreso ($1,500 y $3,000), deuda SSF ($1,000 y $10,000), cuotas vencidas (1 y 3) y deuda ÷ ingreso anual (25% y 100%). '
    + 'Son propuestas de Riesgo. La hoja 5 muestra el efecto de tres alternativas para la deuda ÷ ingreso anual (75%, 100% y 150%).'),
  bullet('La regla del nivel de riesgo por datos en nivel Alto, y los pesos del puntaje piloto.'),
  bullet('Si se aceptan montos de vehículo fuera de la tabla RANGOS VEHICULO (123 recomendados).'),
  bullet('Si el hipotecario puede usarse para pagar deudas de otros bancos (17 recomendados condicionados).'),
  bullet('Una tabla oficial para el hipotecario (hoy se toma el monto de la base).'),
  bullet('Solvencia en Hacienda para vehículos mayores de $25,000 (propuesta; el memorando solo la pide para hipotecario y tarjeta).'),

  h1('7. Resultados de las pruebas'),
  p('Las pruebas están en la hoja 5_Distribucion_Controles y se recalculan solas. El archivo se recalculó con LibreOffice (0 errores en '
    + '239,405 fórmulas) y cada estado, nivel, puntaje y monto se comparó con un cálculo independiente en Python: coinciden en los 2,262 registros. '
    + 'También se probó la automatización: al agregar un registro y borrar el ingreso de un recomendado, el resumen, las listas y los controles '
    + 'se actualizaron solos.', { size: 20 }),
  table(['Prueba', 'Resultado', 'Dato'], controles.map((c) => [c[0], c[1], c[2] ?? '']), [6060, 1700, 1600], (r, i) => (i === 1 ? RES[r[1]] : null)),

  h1('8. Limitaciones de la metodología'),
  bullet('La base trae el saldo de las deudas, no la cuota mensual: la relación cuota/ingreso y la capacidad de pago real no se pueden calcular.'),
  bullet('Lista negra, créditos castigados, personas relacionadas/empleados y solvencia en Hacienda no están en la base.'),
  bullet('1,283 registros no tienen ingreso; no se pueden evaluar hasta completar el dato.'),
  bullet('Los rangos se calcularon con la distribución de esta base; no son límites normativos.'),

  h1('9. Recomendaciones para su uso institucional'),
  bullet('Usar la hoja 3 solo como lista de preevaluación: cada caso pasa a la evaluación crediticia formal (paso 6).'),
  bullet('Revisar primero los casos de la hoja 4 que tienen "Tarjeta a evaluar".'),
  bullet('Aprobar o ajustar los parámetros en la hoja 2 (celdas azules); el libro se recalcula solo.'),
  bullet('Pedir a la fuente el ingreso faltante y la cuota mensual de las deudas, para medir la capacidad de pago real.'),
  bullet('Antes de usar el libro, revisar que todas las pruebas de la hoja 5 digan OK.'),
];

const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: 'vi', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      new TextRun({ text: 'Informe de auditoría – COMÉDICA · Página ', font: FONT, size: 16, color: '808080' }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '808080' })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(out, b); console.log('ok', out); });
