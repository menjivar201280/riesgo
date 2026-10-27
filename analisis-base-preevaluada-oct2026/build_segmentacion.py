# -*- coding: utf-8 -*-
"""Construye el libro de segmentacion de riesgos a partir de la base pre evaluada.

Uso: python3 -I build.py <base_original.xlsx> <salida.xlsx>
La base original se conserva intacta (hojas BD FINAL, Variables, RANGOS ...);
todas las clasificaciones se escriben como formulas que leen la hoja Parametros.
"""
import sys, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

SRC, OUT = sys.argv[1], sys.argv[2]
wb = openpyxl.load_workbook(SRC)
src = wb['BD FINAL']
N = src.max_row - 1            # registros de datos
R1, RN = 2, src.max_row        # filas de datos
NCOL_ORIG = 31                 # columnas con encabezado (A:AE)

F = 'Arial'
fn = Font(name=F, size=10)
fb = Font(name=F, size=10, bold=True)
fh = Font(name=F, size=10, bold=True, color='FFFFFF')
ftitle = Font(name=F, size=14, bold=True, color='1F3864')
fsub = Font(name=F, size=11, bold=True, color='1F3864')
finput = Font(name=F, size=10, color='0000FF', bold=True)
fill_h = PatternFill('solid', fgColor='1F3864')
fill_calc = PatternFill('solid', fgColor='2E75B6')
fill_in = PatternFill('solid', fgColor='FFFF00')
fill_note = PatternFill('solid', fgColor='F2F2F2')
fill_warn = PatternFill('solid', fgColor='FCE4D6')
thin = Side(style='thin', color='BFBFBF')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical='top')
USD = '"$"#,##0.00'
MULT = '0.00"x"'
PCT = '0.0%'


def hdr(ws, row, values, fill=fill_h, widths=None):
    for j, v in enumerate(values, 1):
        c = ws.cell(row, j, v)
        c.font, c.fill, c.alignment, c.border = fh, fill, Alignment(wrap_text=True, vertical='center'), box
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[L(j)].width = w


def put(ws, row, col, v, font=fn, fmt=None, fill=None, border=True, al=wrap):
    c = ws.cell(row, col, v)
    c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = box
    if al: c.alignment = al
    return c


# ---------------------------------------------------------------- Parametros
P = {}
wp = wb.create_sheet('Parametros')
wp['A1'] = 'PARÁMETROS DE SEGMENTACIÓN (celdas amarillas = editables; todo el libro se recalcula)'
wp['A1'].font = ftitle
wp['A2'] = ('Estado: PROPUESTA PENDIENTE DE VALIDACIÓN POR EL ÁREA DE RIESGOS. Ningún límite de esta hoja '
            'es normativo ni está aprobado institucionalmente.')
wp['A2'].font = Font(name=F, size=10, bold=True, color='C00000')
hdr(wp, 4, ['ID', 'Parámetro', 'Valor', 'Unidad', 'Origen del valor / justificación', 'Estado'],
    widths=[18, 46, 16, 14, 90, 30])
params = [
    # id, descripcion, valor, unidad, fmt, origen, estado
    ('FECHA_CORTE', 'Fecha de corte de la base (para recalcular edad)', datetime.date(2026, 9, 30), 'Fecha', 'yyyy-mm-dd',
     'MEMORANDO: universo al 30 de septiembre de 2026. Nota: la columna EDAD coincide mejor con el 02/09/2026 '
     '(2,250 de 2,262), por lo que parece calculada antes del corte.', 'Criterio del memorando'),
    ('EDAD_MIN', 'Edad mínima del universo (años cumplidos)', 30, 'Años', '0',
     'Memorando, criterio 7: edad desde 30 hasta 55 años cumplidos.', 'Criterio del memorando'),
    ('EDAD_MAX', 'Edad máxima del universo (años cumplidos)', 55, 'Años', '0',
     'Memorando, criterio 7.', 'Criterio del memorando'),
    ('EDAD_ADULTA', 'Edad mínima supuesta de afiliación (solo para alerta de consistencia)', 18, 'Años', '0',
     'SUPUESTO: si ANTIGUEDAD_ANIOS > EDAD - este valor, se marca alerta de verificación (no se reclasifica).',
     'Supuesto pendiente'),
    ('ANT_ALTO_MAX', 'Antigüedad: límite superior nivel ALTO', 1, 'Años completos', '0',
     'Solicitado en el prompt (0 a 1 año = Alto). Coincide con la hoja "Variables" del archivo.', 'Criterio solicitado'),
    ('ANT_MEDIO_MAX', 'Antigüedad: límite superior nivel MEDIO', 4, 'Años completos', '0',
     'Solicitado en el prompt (2 a 4 años = Medio; 5 o más = Bajo).', 'Criterio solicitado'),
    ('ING_L1', 'Ingreso mensual: inicio rango II', 1500, 'USD/mes', USD,
     'Aprox. percentil 25 de los ingresos > 0 (P25 = $1,484.40). Coincide con el corte $1,500 de la hoja RANGOS TC.',
     'Pendiente de aprobación'),
    ('ING_L2', 'Ingreso mensual: inicio rango III', 3000, 'USD/mes', USD,
     'Cerca del percentil 60 (P60 = $3,129.98). Es el ingreso mínimo de la tabla RANGOS VEHICULO y el mínimo '
     'observado en la pre evaluación hipotecaria.', 'Pendiente de aprobación'),
    ('ING_L3', 'Ingreso mensual: inicio rango IV', 5000, 'USD/mes', USD,
     'Aprox. percentil 83. Corte de RANGOS VEHICULO ($5,000) y de la pre evaluación hipotecaria observada (150,000).',
     'Pendiente de aprobación'),
    ('ING_ATIPICO', 'Ingreso atípico: umbral de verificación (P75 + 3×RIC)', None, 'USD/mes', USD,
     'FÓRMULA VIVA sobre ingresos > 0: percentil 75 + 3 × (P75 − P25). Valores mayores se marcan para verificar, '
     'no se excluyen.', 'Método pendiente de aprobación'),
    ('DEU_L1', 'Saldo SSF: inicio deuda intermedia', 1000, 'USD', USD,
     'Aprox. percentil 33 de saldos > 0 (P33 = $961.95).', 'Pendiente de aprobación'),
    ('DEU_L2', 'Saldo SSF: inicio deuda elevada', 10000, 'USD', USD,
     'Aprox. percentil 63 de saldos > 0 (P60 = $7,539.71; P67 = $11,920.18).', 'Pendiente de aprobación'),
    ('DEU_L3', 'Saldo SSF: inicio deuda muy elevada', 50000, 'USD', USD,
     'Aprox. percentil 87 de saldos > 0 (P80 = $27,138.76; P90 = $73,940.01). Nivel típico de saldos hipotecarios.',
     'Pendiente de aprobación'),
    ('DTI_L1', 'Deuda/ingreso: inicio nivel medio', 3, 'Veces el ingreso mensual', MULT,
     'Aprox. percentil 60 de la relación saldo/ingreso (P60 = 2.50x). Umbral ESTADÍSTICO, no regulatorio.',
     'Pendiente de aprobación'),
    ('DTI_L2', 'Deuda/ingreso: inicio nivel alto', 12, 'Veces el ingreso mensual', MULT,
     'Equivale a un año de ingreso; percentil 82 aprox. (P80 = 10.67x). Umbral estadístico.', 'Pendiente de aprobación'),
    ('DTI_L3', 'Deuda/ingreso: inicio nivel muy alto', 24, 'Veces el ingreso mensual', MULT,
     'Dos años de ingreso; percentil 92 aprox. (P90 = 20.88x). Umbral estadístico.', 'Pendiente de aprobación'),
    ('FON_L1', 'Fondo compensatorio: inicio rango medio', 3000, 'USD', USD,
     'Aprox. percentil 28 de fondos > 0 (P25 = $2,591.46).', 'Pendiente de aprobación'),
    ('FON_L2', 'Fondo compensatorio: inicio rango alto', 15000, 'USD', USD,
     'Aprox. percentil 53 (mediana = $12,692.09).', 'Pendiente de aprobación'),
    ('FON_L3', 'Fondo compensatorio: inicio rango muy alto', 50000, 'USD', USD,
     'Aprox. percentil 85 (P80 = $38,499.32; P90 = $69,807.93).', 'Pendiente de aprobación'),
    ('FOI_L1', 'Fondo/ingreso: inicio cobertura media', 1, 'Meses de ingreso', MULT,
     'Fondo equivalente a 1 mes de ingreso (P50 ≈ 1.22 meses).', 'Pendiente de aprobación'),
    ('FOI_L2', 'Fondo/ingreso: inicio cobertura alta', 6, 'Meses de ingreso', MULT,
     'Seis meses de ingreso (P70 aprox.).', 'Pendiente de aprobación'),
    ('FOI_L3', 'Fondo/ingreso: inicio cobertura muy alta', 12, 'Meses de ingreso', MULT,
     'Un año de ingreso (P80 aprox.).', 'Pendiente de aprobación'),
    ('TC_TOPE_SIN_SOLV', 'Tope TC si no está solvente en Hacienda', 4000, 'USD', USD,
     'Memorando, otros criterios 5: con límite > $4,000 y sin solvencia en Hacienda se otorga con límite $4,000.', 'Criterio del memorando'),
    ('HIPO_SOLV_MIN', 'Monto hipotecario desde el cual se exige solvencia en Hacienda', 30000, 'USD', USD,
     'Memorando, otros criterios 6: créditos hipotecarios mayores a $30,000 requieren verificar solvencia.', 'Criterio del memorando'),
    ('APO_L1', 'Aportaciones acumuladas: inicio nivel medio', 500, 'USD', USD,
     'Aprox. percentil 20 de MONTO_TOTAL_APORTACIONES (P20 = $517; P25 = $696).', 'Pendiente de aprobación'),
    ('APO_L2', 'Aportaciones acumuladas: inicio nivel alto', 2000, 'USD', USD,
     'Aprox. percentil 73 (P75 = $2,070).', 'Pendiente de aprobación'),
    ('MORA_L1', 'Cuotas de aportación vencidas: inicio nivel medio', 1, 'Cuotas', '0',
     'Cualquier atraso (memorando: no estar al día condiciona la TC).', 'Pendiente de aprobación'),
    ('MORA_L2', 'Cuotas de aportación vencidas: inicio nivel alto', 3, 'Cuotas', '0',
     '3 o más cuotas (atraso de un trimestre o más); P95 de la base = 2 cuotas.', 'Pendiente de aprobación'),
    ('EDAD_T1', 'Edad: inicio del tramo 2', 40, 'Años', '0', 'Tramos de 10 años dentro del rango 30-55 del memorando (descriptivo).',
     'Pendiente de aprobación'),
    ('EDAD_T2', 'Edad: inicio del tramo 3', 50, 'Años', '0', 'Tramo final: relevante para plazos largos (hipotecario).',
     'Pendiente de aprobación'),
    ('CAT_OK1', 'Categoría aceptada 1', 'A1', 'Texto', '@', 'Memorando, criterios 5 y 6 (A1 y A2 interna y externa).',
     'Criterio del memorando'),
    ('CAT_OK2', 'Categoría aceptada 2', 'A2', 'Texto', '@', 'Memorando, criterios 5 y 6 (A1 y A2 interna y externa).',
     'Criterio del memorando'),
]
r = 5
for pid, desc, val, unit, fmt, orig, est in params:
    put(wp, r, 1, pid, fb)
    put(wp, r, 2, desc)
    c = put(wp, r, 3, val, finput, fmt, fill_in)
    put(wp, r, 4, unit)
    put(wp, r, 5, orig)
    put(wp, r, 6, est, Font(name=F, size=10, color='C00000'))
    P[pid] = f'Parametros!$C${r}'
    r += 1

# Tablas de producto (transcripcion numerica de las hojas del archivo)
r += 1
wp.cell(r, 1, 'TABLAS DE PRODUCTO (límite inferior de ingreso mensual → monto). Fuente indicada en cada tabla.').font = fsub
r += 1
TBL = {}


def table(name, title, rows, source):
    global r
    wp.cell(r, 1, title).font = fb
    wp.cell(r, 5, source).font = Font(name=F, size=9, italic=True)
    r += 1
    hdr(wp, r, ['Desde ingreso (USD)', 'Monto / límite', 'Segmento'], fill=fill_calc)
    r += 1
    start = r
    for lo, amt, seg in rows:
        put(wp, r, 1, lo, finput, USD, fill_in)
        put(wp, r, 2, amt, finput, '#,##0' if isinstance(amt, (int, float)) else None, fill_in)
        put(wp, r, 3, seg)
        r += 1
    TBL[name] = (f'Parametros!$A${start}:$A${r-1}', f'Parametros!$B${start}:$B${r-1}', f'Parametros!$C${start}:$C${r-1}')
    r += 1


table('TC', 'Tarjeta de crédito', [
    (0, 'NO APLICA', 'Ingreso menor a $500'),
    (500, 500, 'CLASICA LOCAL (PRIMERA TARJETA)'), (750, 1500, 'CLASICA INTERNACIONAL'),
    (1000, 2000, 'CLASICA INTERNACIONAL'), (1500, 2500, 'DORADA'), (1700, 3000, 'DORADA'),
    (2000, 3500, 'PLATINUM'), (2500, 4500, 'PLATINUM'), (3000, 5000, 'PLATINUM'),
    (3500, 7000, 'BLACK'), (7500, 8000, 'BLACK'), (8000, 10000, 'BLACK')],
    'Transcripción numérica de la hoja "RANGOS TC" del archivo (los textos con "$" se convirtieron a número).')
table('VEH', 'Crédito de vehículo', [
    (0, 'NO APLICA', 'Menor a $3,000.00'), (3000, 20000, ''), (5000, 25000, ''), (10000, 30000, ''), (15000, 40000, '')],
    'Transcripción de la hoja "RANGOS VEHICULO". La hoja dice "INGRESOS"; los datos coinciden con INGRESO mensual.')
table('HIPO', 'Crédito hipotecario (REGLA INFERIDA, NO DOCUMENTADA)', [
    (0, 'NO APLICA', 'Sin regla observada'), (3000, 110000, 'Inferido'), (5000, 150000, 'Inferido'), (9000, 205000, 'Inferido')],
    'INFERIDA de los datos: con INGRESO de $3,000-$4,999.99 se asignó 110,000; $5,000-$8,900 → 150,000; '
    '≥ $9,000 → 205,000. El archivo NO incluye tabla hipotecaria: el corte exacto entre $8,900 y $9,000 debe confirmarse.')
wp.freeze_panes = 'A5'

# ---------------------------------------------------------------- BD_CLASIFICADA
ws = wb.create_sheet('BD_CLASIFICADA')
orig_hdr = [src.cell(1, j).value for j in range(1, NCOL_ORIG + 1)]
for j, h in enumerate(orig_hdr, 1):
    c = ws.cell(1, j, h)
    c.font, c.fill, c.border = fh, fill_h, box
for i in range(R1, RN + 1):
    for j in range(1, NCOL_ORIG + 1):
        s = src.cell(i, j)
        c = ws.cell(i, j, s.value)
        c.font = fn
        if s.number_format and s.number_format != 'General':
            c.number_format = s.number_format

# columnas de origen
C = dict(ASOC='A', EDAD='C', FNAC='E', ANT='O', ING='P', FON='Q', HIPO='R', VEH='S', TC='T', SEG='U', CVA='V',
         CRED='X', CATI='Y', APD='Z', APM='AA', TTC='AB', CATE='AC', SSF='AD', INST='AE', DUI='I')
calc = []  # (header, template, fmt, nota)


def add(h, t, fmt=None, nota=None):
    calc.append((h, t, fmt, nota))


def col_of(h):
    return L(NCOL_ORIG + 1 + [x[0] for x in calc].index(h))


V = lambda k: f'{C[k]}{{r}}'           # celda de origen en la fila
ING_OK = f'AND(ISNUMBER({V("ING")}),{V("ING")}>0)'
SSF_OK = f'ISNUMBER({V("SSF")})'
NUMV = lambda k: f'ISNUMBER({V(k)})'

# --- edad y antiguedad
add('EDAD_RECALCULADA',
    f'=YEAR({P["FECHA_CORTE"]})-VALUE(LEFT({V("FNAC")},4))-IF(OR(VALUE(MID({V("FNAC")},6,2))>MONTH({P["FECHA_CORTE"]}),'
    f'AND(VALUE(MID({V("FNAC")},6,2))=MONTH({P["FECHA_CORTE"]}),VALUE(RIGHT({V("FNAC")},2))>DAY({P["FECHA_CORTE"]}))),1,0)',
    '0', 'Edad en años cumplidos a la FECHA_CORTE a partir de CLI_FECHA_NACIMIENTO (texto AAAA-MM-DD).')
add('CHK_EDAD', f'=IF({V("EDAD")}={{EDAD_RECALCULADA}},"Coincide","Difiere - verificar")')
add('NIVEL_ANTIGUEDAD',
    f'=IF(NOT({NUMV("ANT")}),"Sin dato",IF(OR({V("ANT")}<0,{V("ANT")}>{V("EDAD")}),"Inconsistente - revisar",'
    f'IF({V("ANT")}<={P["ANT_ALTO_MAX"]},"Alto",IF({V("ANT")}<={P["ANT_MEDIO_MAX"]},"Medio","Bajo"))))',
    None, 'Años completos. Vacío ≠ 0. Negativos o mayores que la edad = inconsistentes.')
add('ALERTA_ANT_EDAD',
    f'=IF(AND({NUMV("ANT")},{V("ANT")}>=0,{V("ANT")}<={V("EDAD")},{V("ANT")}>{V("EDAD")}-{P["EDAD_ADULTA"]}),1,0)',
    '0', '1 = antigüedad mayor que los años transcurridos desde EDAD_ADULTA (verificar fecha de afiliación).')
# --- ingreso
add('ING_VALIDO', f'=IF({ING_OK},{V("ING")},"")', USD, 'Auxiliar: ingreso > 0; vacío si no es válido.')
add('RANGO_INGRESO',
    f'=IF(NOT({NUMV("ING")}),"Sin dato",IF({V("ING")}<=0,"Ingreso cero - revisar",'
    f'IF({V("ING")}<{P["ING_L1"]},"I - Ingreso bajo",IF({V("ING")}<{P["ING_L2"]},"II - Ingreso medio",'
    f'IF({V("ING")}<{P["ING_L3"]},"III - Ingreso medio alto","IV - Ingreso alto")))))')
add('NIVEL_INGRESO',
    '=IF({RANGO_INGRESO}="I - Ingreso bajo","Alto",IF({RANGO_INGRESO}="II - Ingreso medio","Medio",'
    'IF({RANGO_INGRESO}="III - Ingreso medio alto","Bajo",IF({RANGO_INGRESO}="IV - Ingreso alto","Muy bajo","Sin información"))))',
    None, 'Nivel de riesgo DESCRIPTIVO por capacidad de ingreso; no es riesgo crediticio integral.')
add('ALERTA_ING_ATIPICO', f'=IF(AND({ING_OK},{V("ING")}>{P["ING_ATIPICO"]}),1,0)', '0')
# --- deuda SSF
add('ESTADO_DEUDA_SSF',
    f'=IF(AND(NOT({SSF_OK}),LEN(TRIM({V("INST")}&""))=0),"Sin registro SSF en la base",'
    f'IF(NOT({SSF_OK}),"Dato faltante",IF({V("SSF")}<0,"Saldo negativo (a favor) - revisar",'
    f'IF({V("SSF")}=0,"Saldo cero","Con saldo"))))')
add('SSF_VALIDO', f'=IF(AND({SSF_OK},{V("SSF")}>0),{V("SSF")},"")', USD, 'Auxiliar: saldo SSF > 0.')
add('RANGO_DEUDA',
    f'=IF(NOT({SSF_OK}),IF(LEN(TRIM({V("INST")}&""))=0,"Sin registro SSF","Dato faltante"),'
    f'IF({V("SSF")}<=0,"Sin saldo (cero o a favor)",IF({V("SSF")}<{P["DEU_L1"]},"D1 - Deuda baja",'
    f'IF({V("SSF")}<{P["DEU_L2"]},"D2 - Deuda intermedia",IF({V("SSF")}<{P["DEU_L3"]},"D3 - Deuda elevada",'
    f'"D4 - Deuda muy elevada")))))')
add('NIVEL_DEUDA',
    '=IF({RANGO_DEUDA}="D1 - Deuda baja","Bajo",IF({RANGO_DEUDA}="D2 - Deuda intermedia","Medio",'
    'IF({RANGO_DEUDA}="D3 - Deuda elevada","Alto",IF({RANGO_DEUDA}="D4 - Deuda muy elevada","Muy alto",'
    'IF({RANGO_DEUDA}="Sin saldo (cero o a favor)","Sin deuda","Sin información")))))',
    None, 'Nivel por saldo SSF. Deuda alta no significa incumplimiento.')
add('MULT_DEUDA_INGRESO', f'=IF(AND({ING_OK},{SSF_OK}),MAX({V("SSF")},0)/{V("ING")},"")', MULT,
    'Saldo SSF / ingreso mensual, en VECES el ingreso mensual (no es relación cuota-ingreso).')
add('RANGO_DEUDA_INGRESO',
    f'=IF(NOT({ING_OK}),"No calculable - ingreso no válido",IF(NOT({SSF_OK}),"No calculable - sin dato SSF",'
    '{MULT_DEUDA_INGRESO_IF}))')
add('NIVEL_DEUDA_INGRESO',
    '=IF({RANGO_DEUDA_INGRESO}="R1 - Bajo","Bajo",IF({RANGO_DEUDA_INGRESO}="R2 - Medio","Medio",'
    'IF({RANGO_DEUDA_INGRESO}="R3 - Alto","Alto",IF({RANGO_DEUDA_INGRESO}="R4 - Muy alto","Muy alto",'
    'IF({RANGO_DEUDA_INGRESO}="Sin deuda","Sin deuda","Sin información")))))')
# --- fondos
add('RANGO_FONDO',
    f'=IF(NOT({NUMV("FON")}),"Sin dato",IF({V("FON")}<=0,"Sin fondo (cero)",IF({V("FON")}<{P["FON_L1"]},"F1 - Fondo bajo",'
    f'IF({V("FON")}<{P["FON_L2"]},"F2 - Fondo medio",IF({V("FON")}<{P["FON_L3"]},"F3 - Fondo alto","F4 - Fondo muy alto")))))',
    None, 'Descriptivo: no se supone que el fondo esté disponible para pagar deudas.')
add('NIVEL_RESPALDO_FONDO',
    '=IF({RANGO_FONDO}="F1 - Fondo bajo","Bajo",IF({RANGO_FONDO}="F2 - Fondo medio","Medio",'
    'IF({RANGO_FONDO}="F3 - Fondo alto","Alto",IF({RANGO_FONDO}="F4 - Fondo muy alto","Muy alto",'
    'IF({RANGO_FONDO}="Sin fondo (cero)","Sin fondo","Sin información")))))',
    None, 'Nivel de RESPALDO (monto del fondo), no de riesgo: falta validar si el fondo está disponible.')
add('FON_VALIDO', f'=IF(AND({NUMV("FON")},{V("FON")}>0),{V("FON")},"")', USD, 'Auxiliar: fondo > 0.')
add('MESES_FONDO_INGRESO', f'=IF(AND({ING_OK},{NUMV("FON")}),{V("FON")}/{V("ING")},"")', MULT,
    'Fondo compensatorio / ingreso mensual (meses de ingreso).')
add('RANGO_FONDO_INGRESO',
    f'=IF(NOT({ING_OK}),"No calculable - ingreso no válido",IF({V("FON")}<=0,"Sin fondo",'
    f'IF({{MESES_FONDO_INGRESO}}<{P["FOI_L1"]},"M1 - Cobertura baja",IF({{MESES_FONDO_INGRESO}}<{P["FOI_L2"]},"M2 - Cobertura media",'
    f'IF({{MESES_FONDO_INGRESO}}<{P["FOI_L3"]},"M3 - Cobertura alta","M4 - Cobertura muy alta")))))')
add('NIVEL_COBERTURA_FONDO',
    '=IF({RANGO_FONDO_INGRESO}="M1 - Cobertura baja","Bajo",IF({RANGO_FONDO_INGRESO}="M2 - Cobertura media","Medio",'
    'IF({RANGO_FONDO_INGRESO}="M3 - Cobertura alta","Alto",IF({RANGO_FONDO_INGRESO}="M4 - Cobertura muy alta","Muy alto",'
    'IF({RANGO_FONDO_INGRESO}="Sin fondo","Sin fondo","Sin información")))))',
    None, 'Nivel de respaldo del fondo en meses de ingreso (no es riesgo).')
# --- otros criterios
add('NIVEL_APORTACIONES',
    f'=IF(NOT({NUMV("APM")}),"Sin información",IF({V("APM")}<{P["APO_L1"]},"Bajo",IF({V("APM")}<{P["APO_L2"]},"Medio","Alto")))',
    None, 'Nivel de aportaciones acumuladas (vínculo/ahorro con la cooperativa); no es nivel de riesgo.')
add('NIVEL_MORA_APORTACIONES',
    f'=IF(NOT({NUMV("CVA")}),"Sin información",IF({V("CVA")}<{P["MORA_L1"]},"Bajo",IF({V("CVA")}<{P["MORA_L2"]},"Medio","Alto")))',
    None, 'Nivel de riesgo por cuotas de aportación vencidas.')
add('TRAMO_EDAD',
    f'=IF({{EDAD_RECALCULADA}}<{P["EDAD_T1"]},"30 a 39",IF({{EDAD_RECALCULADA}}<{P["EDAD_T2"]},"40 a 49",'
    f'IF({{EDAD_RECALCULADA}}<={P["EDAD_MAX"]},"50 a 55","Mayor de 55")))',
    None, 'Tramo descriptivo de edad al corte; no se asigna nivel de riesgo por edad.')
# --- historial
add('CAT_INTERNA',
    f'=IF(LEN(TRIM({V("CATI")}&""))>0,{V("CATI")},IF({V("CRED")}="SI","Dato faltante (con crédito vigente)",'
    '"Sin historial interno"))')
add('CAT_EXTERNA_SSF',
    f'=IF(LEN(TRIM({V("CATE")}&""))>0,{V("CATE")},IF({SSF_OK},"Dato faltante (con saldo SSF)","Sin registro SSF"))')
add('NIVEL_CAT_INTERNA', '=IF({CAT_INTERNA}="A1","Bajo",IF({CAT_INTERNA}="A2","Medio",IF(LEN({CAT_INTERNA})=2,"Alto","Sin información")))',
    None, 'A1 = Bajo, A2 = Medio, otra categoría = Alto; sin categoría = Sin información.')
add('NIVEL_CAT_EXTERNA', '=IF({CAT_EXTERNA_SSF}="A1","Bajo",IF({CAT_EXTERNA_SSF}="A2","Medio",IF(LEN({CAT_EXTERNA_SSF})=2,"Alto","Sin información")))')
# --- criterios de elegibilidad del memorando
crit_first = len(calc)
add('C_EDAD', f'=IF(AND({{EDAD_RECALCULADA}}>={P["EDAD_MIN"]},{{EDAD_RECALCULADA}}<={P["EDAD_MAX"]}),"Cumple","No cumple")',
    None, 'Edad a la FECHA_CORTE del memorando (30/09/2026), recalculada desde la fecha de nacimiento.')
add('C_INGRESO', f'=IF(NOT({NUMV("ING")}),"Sin información",IF({V("ING")}>0,"Cumple","No cumple"))')
add('C_FONDO', f'=IF(NOT({NUMV("FON")}),"Sin información",IF({V("FON")}>0,"Cumple","No cumple"))')
add('C_ANTIGUEDAD', '=IF(OR({NIVEL_ANTIGUEDAD}="Sin dato",{NIVEL_ANTIGUEDAD}="Inconsistente - revisar"),"Sin información","Cumple")')
add('C_CAT_INTERNA',
    f'=IF(OR({{CAT_INTERNA}}={P["CAT_OK1"]},{{CAT_INTERNA}}={P["CAT_OK2"]}),"Cumple",'
    f'IF(LEN(TRIM({V("CATI")}&""))=0,"Sin información","No cumple"))')
add('C_CAT_EXTERNA',
    f'=IF(OR({{CAT_EXTERNA_SSF}}={P["CAT_OK1"]},{{CAT_EXTERNA_SSF}}={P["CAT_OK2"]}),"Cumple",'
    f'IF(LEN(TRIM({V("CATE")}&""))=0,"Sin información","No cumple"))')
crit_last = len(calc) - 1
add('C_APORTACIONES', f'=IF({V("APD")}="SI","Cumple",IF({V("APD")}="NO","No cumple","Sin información"))',
    None, 'Memorando: no excluye; condiciona la TC a ponerse al día.')
add('C_SIN_TARJETA', f'=IF({V("TTC")}="NO","Cumple",IF({V("TTC")}="SI","No cumple","Sin información"))',
    None, 'Solo aplica al universo de tarjeta de crédito.')
add('C_CONSULTAS_EXTERNAS', '="Requiere consulta externa"', None,
    'Lista negra, créditos castigados, personas relacionadas/empleados y solvencia en Hacienda: no existen en la base.')
add('ELEGIBILIDAD_GENERAL', '=IF(COUNTIF({CRITRANGE},"No cumple")>0,"No cumple criterio verificable",'
    'IF(COUNTIF({CRITRANGE},"Sin información")>0,"Información incompleta","Cumple criterios verificables"))',
    None, 'Preliminar. Siempre queda pendiente la consulta externa. No es decisión crediticia.')
add('ELEGIBILIDAD_TC', '=IF({C_SIN_TARJETA}="No cumple","Excluido: ya tiene tarjeta",{ELEGIBILIDAD_GENERAL})')
add('CONDICIONES_MEMO',
    f'=IF(AND({NUMV("TC")},{V("APD")}="NO"),"TC: ponerse al día en aportaciones; ","")'
    f'&IF(AND({NUMV("TC")},{V("TC")}>{P["TC_TOPE_SIN_SOLV"]}),"TC > tope: verificar solvencia Hacienda o limitar; ","")'
    f'&IF(AND({NUMV("HIPO")},{V("HIPO")}>{P["HIPO_SOLV_MIN"]}),"Hipotecario: verificar solvencia Hacienda; ","")',
    None, 'Condiciones de otorgamiento del memorando (otros criterios 4, 5 y 6).')
add('TC_LIMITE_SIN_SOLVENCIA', f'=IF({NUMV("TC")},MIN({V("TC")},{P["TC_TOPE_SIN_SOLV"]}),"")', '#,##0',
    'Límite TC aplicable si el asociado NO está solvente en Hacienda (memorando, otros 5).')
add('MOTIVOS',
    '=IF({C_EDAD}<>"Cumple","Edad; ","")&IF({C_INGRESO}="No cumple","Ingreso cero; ",IF({C_INGRESO}="Sin información","Ingreso sin dato; ",""))'
    '&IF({C_FONDO}<>"Cumple","Fondo cero; ","")&IF({C_ANTIGUEDAD}<>"Cumple","Antigüedad sin dato o inconsistente; ","")'
    '&IF({C_CAT_INTERNA}="No cumple","Categoría interna fuera de A1/A2; ",IF({C_CAT_INTERNA}="Sin información","Sin categoría interna; ",""))'
    '&IF({C_CAT_EXTERNA}="No cumple","Categoría SSF fuera de A1/A2; ",IF({C_CAT_EXTERNA}="Sin información","Sin calificación SSF; ",""))'
    )
# --- productos
for k, nm in (('TC', 'TC'), ('VEH', 'VEHICULO'), ('HIPO', 'HIPO')):
    lo, amt, _ = TBL[k]
    add(f'ESPERADO_{nm}', f'=IF({ING_OK},LOOKUP({V("ING")},{lo},{amt}),"")', '#,##0',
        f'Monto según tabla de la hoja Parametros ({nm}).')
    add(f'CONSIST_{nm}',
        f'=IF(NOT({NUMV(k)}),"No pre evaluado",IF(NOT({ING_OK}),"Pre evaluado sin ingreso válido",'
        f'IF({V(k)}={{ESPERADO_{nm}}},"Coincide con tabla","No coincide con tabla")))')
add('LIMITE_TC_SOBRE_INGRESO', f'=IF(AND({ING_OK},{NUMV("TC")}),{V("TC")}/{V("ING")},"")', MULT,
    'Límite TC pre evaluado / ingreso mensual.')
add('VEHICULO_SOBRE_INGRESO', f'=IF(AND({ING_OK},{NUMV("VEH")}),{V("VEH")}/{V("ING")},"")', MULT)
add('HIPO_SOBRE_INGRESO', f'=IF(AND({ING_OK},{NUMV("HIPO")}),{V("HIPO")}/{V("ING")},"")', MULT)
# --- alertas
al_first = len(calc)
add('AL_ING_BAJO_DEUDA_ELEVADA',
    '=IF(AND({RANGO_INGRESO}="I - Ingreso bajo",OR({RANGO_DEUDA}="D3 - Deuda elevada",{RANGO_DEUDA}="D4 - Deuda muy elevada")),1,0)', '0')
add('AL_DEUDA_ALTA_VS_INGRESO',
    '=IF(OR({RANGO_DEUDA_INGRESO}="R3 - Alto",{RANGO_DEUDA_INGRESO}="R4 - Muy alto"),1,0)', '0')
add('AL_ANTIG_BAJA_SIN_HISTORIAL',
    '=IF(AND({NIVEL_ANTIGUEDAD}="Alto",{CAT_INTERNA}="Sin historial interno"),1,0)', '0')
add('AL_FONDO_BAJO_DEUDA_ELEVADA',
    '=IF(AND(OR({RANGO_FONDO}="Sin fondo (cero)",{RANGO_FONDO}="F1 - Fondo bajo"),'
    'OR({RANGO_DEUDA}="D3 - Deuda elevada",{RANGO_DEUDA}="D4 - Deuda muy elevada")),1,0)', '0')
add('AL_CATEGORIA_A2', '=IF(OR({CAT_INTERNA}="A2",{CAT_EXTERNA_SSF}="A2"),1,0)', '0')
add('AL_APORT_VENCIDAS', f'=IF(OR({V("APD")}="NO",{V("CVA")}>0),1,0)', '0')
al_last = len(calc) - 1
dq_first = len(calc)
add('DQ_PRECA_SIN_INGRESO',
    '=IF(OR({CONSIST_TC}="Pre evaluado sin ingreso válido",{CONSIST_VEHICULO}="Pre evaluado sin ingreso válido",'
    '{CONSIST_HIPO}="Pre evaluado sin ingreso válido"),1,0)', '0')
add('DQ_PRECA_FUERA_TABLA',
    '=IF(OR({CONSIST_TC}="No coincide con tabla",{CONSIST_VEHICULO}="No coincide con tabla",'
    '{CONSIST_HIPO}="No coincide con tabla"),1,0)', '0')
add('DQ_DATO_INCONSISTENTE',
    '=IF(OR({NIVEL_ANTIGUEDAD}="Inconsistente - revisar",{ALERTA_ANT_EDAD}=1,{ALERTA_ING_ATIPICO}=1,'
    '{ESTADO_DEUDA_SSF}="Saldo negativo (a favor) - revisar",{CHK_EDAD}<>"Coincide",' + f'LEN({V("DUI")})<>9),1,0)', '0')
dq_last = len(calc) - 1
add('TOTAL_ALERTAS_RIESGO', '=SUM({ALRANGE})', '0')
add('TOTAL_ALERTAS_DATOS', '=SUM({DQRANGE})', '0')
add('ESTADO_DESCRIPTIVO',
    '=IF({TOTAL_ALERTAS_RIESGO}>=2,"2 o más alertas de riesgo",IF({TOTAL_ALERTAS_RIESGO}=1,"1 alerta de riesgo",'
    '{SINALERTA}))', None,
    'Segmentación descriptiva por número de alertas. NO es una calificación ni una decisión crediticia.')

names = [x[0] for x in calc]
crit_rng = f'{col_of(names[crit_first])}{{r}}:{col_of(names[crit_last])}{{r}}'
al_rng = f'{col_of(names[al_first])}{{r}}:{col_of(names[al_last])}{{r}}'
dq_rng = f'{col_of(names[dq_first])}{{r}}:{col_of(names[dq_last])}{{r}}'
dti_if = (f'IF({{MULT_DEUDA_INGRESO}}=0,"Sin deuda",IF({{MULT_DEUDA_INGRESO}}<{P["DTI_L1"]},"R1 - Bajo",'
          f'IF({{MULT_DEUDA_INGRESO}}<{P["DTI_L2"]},"R2 - Medio",IF({{MULT_DEUDA_INGRESO}}<{P["DTI_L3"]},"R3 - Alto","R4 - Muy alto"))))')
sin_alerta = f'IF({ING_OK},"Sin alertas identificadas","Sin alertas, información insuficiente (ingreso no válido)")'


def render(t, r):
    t = t.replace('{MULT_DEUDA_INGRESO_IF}', dti_if).replace('{CRITRANGE}', crit_rng).replace('{ALRANGE}', al_rng)
    t = t.replace('{DQRANGE}', dq_rng).replace('{SINALERTA}', sin_alerta)
    for n in sorted(names, key=len, reverse=True):
        t = t.replace('{' + n + '}', f'{col_of(n)}{{r}}')
    return t.replace('{r}', str(r))


fill_cols = {}
for j, (h, t, fmt, nota) in enumerate(calc):
    col = NCOL_ORIG + 1 + j
    c = ws.cell(1, col, h)
    c.font, c.fill, c.border = fh, fill_calc, box
    c.alignment = Alignment(wrap_text=True, vertical='center')
    if nota:
        c.comment = Comment(nota, 'Metodología')
    for i in range(R1, RN + 1):
        cc = ws.cell(i, col, render(t, i))
        cc.font = fn
        if fmt: cc.number_format = fmt
    ws.column_dimensions[L(col)].width = 22
for j in range(1, NCOL_ORIG + 1):
    ws.column_dimensions[L(j)].width = 14
ws.row_dimensions[1].height = 45
ws.freeze_panes = 'C2'
ws.auto_filter.ref = f'A1:{L(NCOL_ORIG + len(calc))}{RN}'

COL = {n: col_of(n) for n in names}
RNG = lambda n: f'BD_CLASIFICADA!${COL[n]}${R1}:${COL[n]}${RN}'
ORNG = lambda k: f'BD_CLASIFICADA!${C[k]}${R1}:${C[k]}${RN}'
TOTAL = f'COUNTA(BD_CLASIFICADA!$A${R1}:$A${RN})'

# parametro vivo: umbral de ingreso atipico
for row in wp.iter_rows(min_row=5, max_row=40):
    if row[0].value == 'ING_ATIPICO':
        row[2].value = (f'=PERCENTILE({RNG("ING_VALIDO")},0.75)+3*(PERCENTILE({RNG("ING_VALIDO")},0.75)'
                        f'-PERCENTILE({RNG("ING_VALIDO")},0.25))')
        row[2].fill = PatternFill(None)
        row[2].font = fn

# ---------------------------------------------------------------- Consolidado
CATS = {
    'NIVEL_ANTIGUEDAD': ['Alto', 'Medio', 'Bajo', 'Sin dato', 'Inconsistente - revisar'],
    'RANGO_INGRESO': ['I - Ingreso bajo', 'II - Ingreso medio', 'III - Ingreso medio alto', 'IV - Ingreso alto',
                      'Ingreso cero - revisar', 'Sin dato'],
    'NIVEL_INGRESO': ['Alto', 'Medio', 'Bajo', 'Muy bajo', 'Sin información'],
    'NIVEL_DEUDA': ['Bajo', 'Medio', 'Alto', 'Muy alto', 'Sin deuda', 'Sin información'],
    'NIVEL_DEUDA_INGRESO': ['Bajo', 'Medio', 'Alto', 'Muy alto', 'Sin deuda', 'Sin información'],
    'NIVEL_COBERTURA_FONDO': ['Bajo', 'Medio', 'Alto', 'Muy alto', 'Sin fondo', 'Sin información'],
    'NIVEL_APORTACIONES': ['Bajo', 'Medio', 'Alto', 'Sin información'],
    'NIVEL_MORA_APORTACIONES': ['Bajo', 'Medio', 'Alto', 'Sin información'],
    'TRAMO_EDAD': ['30 a 39', '40 a 49', '50 a 55', 'Mayor de 55'],
    'NIVEL_CAT_INTERNA': ['Bajo', 'Medio', 'Alto', 'Sin información'],
    'NIVEL_CAT_EXTERNA': ['Bajo', 'Medio', 'Alto', 'Sin información'],
    'NIVEL_RESPALDO_FONDO': ['Bajo', 'Medio', 'Alto', 'Muy alto', 'Sin fondo', 'Sin información'],
    'ESTADO_DEUDA_SSF': ['Con saldo', 'Saldo cero', 'Saldo negativo (a favor) - revisar', 'Dato faltante',
                         'Sin registro SSF en la base'],
    'RANGO_DEUDA': ['D1 - Deuda baja', 'D2 - Deuda intermedia', 'D3 - Deuda elevada', 'D4 - Deuda muy elevada',
                    'Sin saldo (cero o a favor)', 'Dato faltante', 'Sin registro SSF'],
    'RANGO_DEUDA_INGRESO': ['Sin deuda', 'R1 - Bajo', 'R2 - Medio', 'R3 - Alto', 'R4 - Muy alto',
                            'No calculable - ingreso no válido', 'No calculable - sin dato SSF'],
    'RANGO_FONDO': ['Sin fondo (cero)', 'F1 - Fondo bajo', 'F2 - Fondo medio', 'F3 - Fondo alto', 'F4 - Fondo muy alto',
                    'Sin dato'],
    'RANGO_FONDO_INGRESO': ['Sin fondo', 'M1 - Cobertura baja', 'M2 - Cobertura media', 'M3 - Cobertura alta',
                            'M4 - Cobertura muy alta', 'No calculable - ingreso no válido'],
    'CAT_INTERNA': ['A1', 'A2', 'Sin historial interno', 'Dato faltante (con crédito vigente)'],
    'CAT_EXTERNA_SSF': ['A1', 'A2', 'Dato faltante (con saldo SSF)', 'Sin registro SSF'],
    'C_EDAD': ['Cumple', 'No cumple'],
    'C_INGRESO': ['Cumple', 'No cumple', 'Sin información'],
    'C_FONDO': ['Cumple', 'No cumple', 'Sin información'],
    'C_ANTIGUEDAD': ['Cumple', 'Sin información'],
    'C_CAT_INTERNA': ['Cumple', 'No cumple', 'Sin información'],
    'C_CAT_EXTERNA': ['Cumple', 'No cumple', 'Sin información'],
    'C_APORTACIONES': ['Cumple', 'No cumple', 'Sin información'],
    'C_SIN_TARJETA': ['Cumple', 'No cumple', 'Sin información'],
    'C_CONSULTAS_EXTERNAS': ['Requiere consulta externa'],
    'ELEGIBILIDAD_GENERAL': ['Cumple criterios verificables', 'Información incompleta', 'No cumple criterio verificable'],
    'ELEGIBILIDAD_TC': ['Cumple criterios verificables', 'Información incompleta', 'No cumple criterio verificable',
                        'Excluido: ya tiene tarjeta'],
    'CONSIST_TC': ['Coincide con tabla', 'No coincide con tabla', 'Pre evaluado sin ingreso válido', 'No pre evaluado'],
    'CONSIST_VEHICULO': ['Coincide con tabla', 'No coincide con tabla', 'Pre evaluado sin ingreso válido', 'No pre evaluado'],
    'CONSIST_HIPO': ['Coincide con tabla', 'No coincide con tabla', 'Pre evaluado sin ingreso válido', 'No pre evaluado'],
    'ESTADO_DESCRIPTIVO': ['Sin alertas identificadas', 'Sin alertas, información insuficiente (ingreso no válido)',
                           '1 alerta de riesgo', '2 o más alertas de riesgo'],
}
wc = wb.create_sheet('Consolidado')
wc['A1'] = 'CONSOLIDADO DE ASOCIADOS POR VARIABLE, RANGO Y NIVEL (fórmulas COUNTIF sobre BD_CLASIFICADA)'
wc['A1'].font = ftitle
wc['A2'] = 'Total de asociados procesados:'
wc['A2'].font = fb
wc['C2'] = f'={TOTAL}'
wc['C2'].font = fb
hdr(wc, 4, ['Variable (columna calculada)', 'Categoría / nivel', 'Cantidad', '% del total', 'Control'],
    widths=[30, 52, 12, 12, 34])
r = 5
CONS_ROW = {}
for var, cats in CATS.items():
    start = r
    for cat in cats:
        put(wc, r, 1, var, fb if r == start else fn)
        put(wc, r, 2, cat)
        put(wc, r, 3, f'=COUNTIF({RNG(var)},B{r})', fmt='#,##0')
        put(wc, r, 4, f'=C{r}/$C$2', fmt=PCT)
        CONS_ROW[(var, cat)] = r
        r += 1
    put(wc, r, 1, var, fb, fill=fill_note)
    put(wc, r, 2, 'Total variable', fb, fill=fill_note)
    put(wc, r, 3, f'=SUM(C{start}:C{r-1})', fb, '#,##0', fill_note)
    put(wc, r, 4, f'=SUM(D{start}:D{r-1})', fb, PCT, fill_note)
    put(wc, r, 5, f'=IF(C{r}=$C$2,"OK: cada asociado una sola vez","ERROR: no cuadra")', fb, fill=fill_note)
    r += 2
wc.freeze_panes = 'A5'

# ---------------------------------------------------------------- Estadisticas
we = wb.create_sheet('Estadisticas')
we['A1'] = 'ESTADÍSTICAS DESCRIPTIVAS (fórmulas vivas; base para proponer los rangos)'
we['A1'].font = ftitle
stats_vars = [
    ('Ingreso mensual (> 0)', RNG('ING_VALIDO'), USD, ORNG('ING')),
    ('Saldo SSF (> 0)', RNG('SSF_VALIDO'), USD, ORNG('SSF')),
    ('Fondo compensatorio (> 0)', RNG('FON_VALIDO'), USD, ORNG('FON')),
    ('Saldo SSF / ingreso (veces)', RNG('MULT_DEUDA_INGRESO'), MULT, None),
    ('Fondo / ingreso (meses)', RNG('MESES_FONDO_INGRESO'), MULT, None),
    ('Antigüedad (años, valor original)', ORNG('ANT'), '0', ORNG('ANT')),
    ('Monto total aportaciones', ORNG('APM'), USD, ORNG('APM')),
    ('Cuotas vencidas de aportación', ORNG('CVA'), '0', ORNG('CVA')),
    ('Edad', ORNG('EDAD'), '0', ORNG('EDAD')),
]
cols = ['Indicador', 'n (numéricos)', 'Mínimo', 'P10', 'P25', 'Mediana', 'Promedio', 'P75', 'P90', 'P95', 'P99',
        'Máximo', 'Vacíos en columna original', 'Ceros en original', 'Negativos en original']
hdr(we, 3, cols, widths=[34] + [14] * 14)
for i, (lab, rg, fmt, orng) in enumerate(stats_vars):
    rr = 4 + i
    put(we, rr, 1, lab, fb)
    put(we, rr, 2, f'=COUNT({rg})', fmt='#,##0')
    funcs = [f'MIN({rg})', f'PERCENTILE({rg},0.1)', f'PERCENTILE({rg},0.25)', f'MEDIAN({rg})', f'AVERAGE({rg})',
             f'PERCENTILE({rg},0.75)', f'PERCENTILE({rg},0.9)', f'PERCENTILE({rg},0.95)', f'PERCENTILE({rg},0.99)', f'MAX({rg})']
    for k, fx in enumerate(funcs):
        put(we, rr, 3 + k, '=' + fx, fmt=fmt)
    if orng:
        put(we, rr, 13, f'={N}-COUNT({orng})', fmt='#,##0')
        put(we, rr, 14, f'=COUNTIF({orng},0)', fmt='#,##0')
        put(we, rr, 15, f'=COUNTIF({orng},"<0")', fmt='#,##0')
    else:
        for k in (13, 14, 15):
            put(we, rr, k, 'n/a')
rr = 4 + len(stats_vars) + 1
notes = [
    'Notas:',
    '• Los percentiles usan PERCENTILE (método inclusivo de Excel). Ingreso, saldo y fondo se calculan solo sobre valores > 0; '
    'ceros, vacíos y negativos se reportan aparte para no distorsionar la distribución.',
    '• "Vacíos" en INGRESO incluye celdas sin dato; un vacío NO se interpreta como cero.',
    '• La relación saldo/ingreso se expresa en VECES el ingreso mensual (ej. 12x = el saldo equivale a 12 meses de ingreso).',
    '• No existen cuotas mensuales de préstamos en la base: NO es posible calcular la relación cuota-ingreso ni la carga financiera mensual.',
]
for t in notes:
    we.cell(rr, 1, t).font = fb if t == 'Notas:' else fn
    rr += 1

# ---------------------------------------------------------------- Matriz
wm = wb.create_sheet('Matriz_Segmentacion')
wm['A1'] = 'MATRIZ DE SEGMENTACIÓN — RANGOS Y NIVELES PROPUESTOS (pendiente de validación por Riesgos)'
wm['A1'].font = ftitle
mh = ['Variable evaluada', 'Columna de origen', 'Rango inferior', 'Rango superior', 'Unidad', 'Nivel / segmento asignado',
      'Justificación del rango', 'Cantidad de asociados', '% de asociados', 'Tratamiento de valores faltantes',
      'Limitaciones de interpretación']
hdr(wm, 3, mh, widths=[22, 22, 15, 15, 14, 30, 50, 12, 11, 36, 44])


def prm(pid, delta=None):
    return f'={P[pid]}' if delta is None else f'={P[pid]}-{delta}'


M = []  # (var, origen, lo, hi, unidad, nivel, justif, countvar, countcat, faltantes, limit, fmt)
ant_lim = 'Antigüedad mide permanencia, no capacidad ni comportamiento de pago. Valores enteros (años completos).'
ant_falt = 'Vacío = "Sin dato" (no se trata como 0). Negativos o > edad = "Inconsistente - revisar".'
M += [
    ('Antigüedad como asociado', 'ANTIGUEDAD_ANIOS', 0, prm('ANT_ALTO_MAX'), 'Años completos', 'Alto',
     'Criterio solicitado (0 a 1 año).', 'NIVEL_ANTIGUEDAD', 'Alto', ant_falt, ant_lim, '0'),
    ('Antigüedad como asociado', 'ANTIGUEDAD_ANIOS', prm('ANT_ALTO_MAX', '-1'), prm('ANT_MEDIO_MAX'), 'Años completos', 'Medio',
     'Criterio solicitado (2 a 4 años).', 'NIVEL_ANTIGUEDAD', 'Medio', ant_falt, ant_lim, '0'),
    ('Antigüedad como asociado', 'ANTIGUEDAD_ANIOS', prm('ANT_MEDIO_MAX', '-1'), 'Sin límite', 'Años completos', 'Bajo',
     'Criterio solicitado (5 años o más).', 'NIVEL_ANTIGUEDAD', 'Bajo', ant_falt, ant_lim, '0'),
    ('Antigüedad como asociado', 'ANTIGUEDAD_ANIOS', '—', '—', '—', 'Sin dato', 'Celda vacía en la base.',
     'NIVEL_ANTIGUEDAD', 'Sin dato', ant_falt, '37% de la base no tiene antigüedad registrada.', None),
    ('Antigüedad como asociado', 'ANTIGUEDAD_ANIOS', '—', '—', '—', 'Inconsistente - revisar',
     'Valores imposibles: -3992, -1 y 1776 años.', 'NIVEL_ANTIGUEDAD', 'Inconsistente - revisar', ant_falt, 'Corregir en el sistema de origen.', None),
]
ing_lim = 'INGRESO es el ingreso registrado en el sistema; la base no indica si es bruto o disponible (neto de descuentos). No es ingreso disponible.'
ing_falt = 'Vacío = "Sin dato"; 0 = "Ingreso cero - revisar". No se calculan razones con ingreso 0 o vacío.'
M += [
    ('Ingreso mensual', 'INGRESO', 0.01, prm('ING_L1', '0.01'), 'USD/mes', 'I - Ingreso bajo (nivel Alto)',
     'Hasta aprox. P25 de ingresos > 0; corte coincide con RANGOS TC.', 'RANGO_INGRESO', 'I - Ingreso bajo', ing_falt, ing_lim, USD),
    ('Ingreso mensual', 'INGRESO', prm('ING_L1'), prm('ING_L2', '0.01'), 'USD/mes', 'II - Ingreso medio (nivel Medio)',
     'P25 a P60 aprox.; $3,000 es el ingreso mínimo de RANGOS VEHICULO.', 'RANGO_INGRESO', 'II - Ingreso medio', ing_falt, ing_lim, USD),
    ('Ingreso mensual', 'INGRESO', prm('ING_L2'), prm('ING_L3', '0.01'), 'USD/mes', 'III - Ingreso medio alto (nivel Bajo)',
     'P60 a P83 aprox.; coincide con tramo de vehículo 20,000 e hipotecario 110,000.', 'RANGO_INGRESO', 'III - Ingreso medio alto', ing_falt, ing_lim, USD),
    ('Ingreso mensual', 'INGRESO', prm('ING_L3'), 'Sin límite', 'USD/mes', 'IV - Ingreso alto (nivel Muy bajo)',
     'Sobre P83 aprox. Valores > umbral atípico se marcan para verificación.', 'RANGO_INGRESO', 'IV - Ingreso alto', ing_falt, ing_lim, USD),
    ('Ingreso mensual', 'INGRESO', 0, 0, 'USD/mes', 'Ingreso cero - revisar', 'Ingreso registrado en 0.',
     'RANGO_INGRESO', 'Ingreso cero - revisar', ing_falt, 'Puede ser dato no capturado; requiere verificación.', USD),
    ('Ingreso mensual', 'INGRESO', '—', '—', '—', 'Sin dato', 'Celda vacía.', 'RANGO_INGRESO', 'Sin dato', ing_falt,
     'Más de la mitad de la base no tiene ingreso: limita toda medición de capacidad de pago.', None),
]
deu_lim = ('SALDO_SSF trae UNA institución por asociado; no se puede confirmar que sea el saldo consolidado de todas sus '
           'obligaciones. Deuda elevada NO implica incumplimiento.')
deu_falt = 'Sin saldo ni institución = "Sin registro SSF" (no se puede distinguir entre sin deuda y no consultado). 0 = "Sin saldo".'
M += [
    ('Saldo total de deuda (Central de Riesgo)', 'SALDO_SSF', 0.01, prm('DEU_L1', '0.01'), 'USD', 'D1 - Deuda baja',
     'Hasta aprox. P33 de saldos > 0.', 'RANGO_DEUDA', 'D1 - Deuda baja', deu_falt, deu_lim, USD),
    ('Saldo total de deuda (Central de Riesgo)', 'SALDO_SSF', prm('DEU_L1'), prm('DEU_L2', '0.01'), 'USD', 'D2 - Deuda intermedia',
     'P33 a P63 aprox.', 'RANGO_DEUDA', 'D2 - Deuda intermedia', deu_falt, deu_lim, USD),
    ('Saldo total de deuda (Central de Riesgo)', 'SALDO_SSF', prm('DEU_L2'), prm('DEU_L3', '0.01'), 'USD', 'D3 - Deuda elevada',
     'P63 a P87 aprox.', 'RANGO_DEUDA', 'D3 - Deuda elevada', deu_falt, deu_lim, USD),
    ('Saldo total de deuda (Central de Riesgo)', 'SALDO_SSF', prm('DEU_L3'), 'Sin límite', 'USD', 'D4 - Deuda muy elevada',
     'Sobre P87 aprox.; nivel de saldos hipotecarios.', 'RANGO_DEUDA', 'D4 - Deuda muy elevada', deu_falt, deu_lim, USD),
    ('Saldo total de deuda (Central de Riesgo)', 'SALDO_SSF', '≤ 0', 0, 'USD', 'Sin saldo (cero o a favor)',
     'Saldo 0 o saldos negativos mínimos (−$0.02 a −$0.89).', 'RANGO_DEUDA', 'Sin saldo (cero o a favor)', deu_falt,
     'Puede ser tarjeta sin uso o crédito cancelado.', None),
    ('Saldo total de deuda (Central de Riesgo)', 'SALDO_SSF', '—', '—', '—', 'Sin registro SSF', 'Sin saldo ni institución.',
     'RANGO_DEUDA', 'Sin registro SSF', deu_falt, 'No equivale a "sin deudas".', None),
]
dti_lim = 'Indicador de SALDO, no de cuota. Mezcla saldos de largo plazo (hipotecas) con consumo. Umbrales estadísticos, NO regulatorios.'
dti_falt = 'Ingreso 0/vacío o sin dato SSF = "No calculable". No se calcula una razón engañosa.'
M += [
    ('Relación deuda-ingreso', 'SALDO_SSF / INGRESO', '>0', prm('DTI_L1'), 'Veces ingreso mensual', 'R1 - Bajo',
     'Hasta P60 aprox. (< 3 veces).', 'RANGO_DEUDA_INGRESO', 'R1 - Bajo', dti_falt, dti_lim, MULT),
    ('Relación deuda-ingreso', 'SALDO_SSF / INGRESO', prm('DTI_L1'), prm('DTI_L2'), 'Veces ingreso mensual', 'R2 - Medio',
     'P60 a P82 aprox. (límite superior excluido).', 'RANGO_DEUDA_INGRESO', 'R2 - Medio', dti_falt, dti_lim, MULT),
    ('Relación deuda-ingreso', 'SALDO_SSF / INGRESO', prm('DTI_L2'), prm('DTI_L3'), 'Veces ingreso mensual', 'R3 - Alto',
     '1 a 2 años de ingreso (límite superior excluido).', 'RANGO_DEUDA_INGRESO', 'R3 - Alto', dti_falt, dti_lim, MULT),
    ('Relación deuda-ingreso', 'SALDO_SSF / INGRESO', prm('DTI_L3'), 'Sin límite', 'Veces ingreso mensual', 'R4 - Muy alto',
     '2 años o más de ingreso (≈ P92).', 'RANGO_DEUDA_INGRESO', 'R4 - Muy alto', dti_falt, dti_lim, MULT),
    ('Relación deuda-ingreso', 'SALDO_SSF / INGRESO', 0, 0, 'Veces', 'Sin deuda', 'Saldo SSF ≤ 0 con ingreso válido.',
     'RANGO_DEUDA_INGRESO', 'Sin deuda', dti_falt, dti_lim, None),
    ('Relación deuda-ingreso', 'SALDO_SSF / INGRESO', '—', '—', '—', 'No calculable - ingreso no válido', '',
     'RANGO_DEUDA_INGRESO', 'No calculable - ingreso no válido', dti_falt, '', None),
    ('Relación deuda-ingreso', 'SALDO_SSF / INGRESO', '—', '—', '—', 'No calculable - sin dato SSF', '',
     'RANGO_DEUDA_INGRESO', 'No calculable - sin dato SSF', dti_falt, '', None),
    ('Cuotas / carga financiera mensual', 'No disponible en la base', '—', '—', 'USD/mes', 'No disponible en la base',
     'La base no contiene cuotas mensuales de préstamos. CUOTAS_VENC_APORT es número de cuotas de APORTACIÓN vencidas, no cuotas de crédito.',
     None, None, 'No aplica', 'No se puede medir carga financiera ni relación cuota-ingreso.', None),
    ('Relación cuota-ingreso', 'No disponible en la base', '—', '—', '%', 'No disponible en la base',
     'Requiere cuotas mensuales de todas las obligaciones.', None, None, 'No aplica', 'No se sustituye por saldos.', None),
]
fon_lim = ('La base no indica si el fondo está disponible, retenido o pignorado. Se reporta como respaldo DESCRIPTIVO; '
           'no se le atribuye reducción de riesgo hasta validar su naturaleza.')
fon_falt = 'Fondo 0 = "Sin fondo (cero)". No hay vacíos en la columna.'
M += [
    ('Fondo compensatorio', 'FONDO_COMPENSATORIO', 0.01, prm('FON_L1', '0.01'), 'USD', 'F1 - Fondo bajo',
     'Hasta P28 aprox. de fondos > 0.', 'RANGO_FONDO', 'F1 - Fondo bajo', fon_falt, fon_lim, USD),
    ('Fondo compensatorio', 'FONDO_COMPENSATORIO', prm('FON_L1'), prm('FON_L2', '0.01'), 'USD', 'F2 - Fondo medio',
     'P28 a P53 aprox.', 'RANGO_FONDO', 'F2 - Fondo medio', fon_falt, fon_lim, USD),
    ('Fondo compensatorio', 'FONDO_COMPENSATORIO', prm('FON_L2'), prm('FON_L3', '0.01'), 'USD', 'F3 - Fondo alto',
     'P53 a P85 aprox.', 'RANGO_FONDO', 'F3 - Fondo alto', fon_falt, fon_lim, USD),
    ('Fondo compensatorio', 'FONDO_COMPENSATORIO', prm('FON_L3'), 'Sin límite', 'USD', 'F4 - Fondo muy alto',
     'Sobre P85 aprox.', 'RANGO_FONDO', 'F4 - Fondo muy alto', fon_falt, fon_lim, USD),
    ('Fondo compensatorio', 'FONDO_COMPENSATORIO', 0, 0, 'USD', 'Sin fondo (cero)', 'Fondo igual a 0.',
     'RANGO_FONDO', 'Sin fondo (cero)', fon_falt, 'No cumple criterio "existencia de fondos".', USD),
    ('Fondo / ingreso', 'FONDO_COMPENSATORIO / INGRESO', '>0', prm('FOI_L1'), 'Meses de ingreso', 'M1 - Cobertura baja',
     'Menos de 1 mes de ingreso.', 'RANGO_FONDO_INGRESO', 'M1 - Cobertura baja', 'Ingreso no válido = no calculable.', fon_lim, MULT),
    ('Fondo / ingreso', 'FONDO_COMPENSATORIO / INGRESO', prm('FOI_L1'), prm('FOI_L2'), 'Meses de ingreso', 'M2 - Cobertura media',
     '1 a < 6 meses.', 'RANGO_FONDO_INGRESO', 'M2 - Cobertura media', 'Ingreso no válido = no calculable.', fon_lim, MULT),
    ('Fondo / ingreso', 'FONDO_COMPENSATORIO / INGRESO', prm('FOI_L2'), prm('FOI_L3'), 'Meses de ingreso', 'M3 - Cobertura alta',
     '6 a < 12 meses.', 'RANGO_FONDO_INGRESO', 'M3 - Cobertura alta', 'Ingreso no válido = no calculable.', fon_lim, MULT),
    ('Fondo / ingreso', 'FONDO_COMPENSATORIO / INGRESO', prm('FOI_L3'), 'Sin límite', 'Meses de ingreso', 'M4 - Cobertura muy alta',
     '12 meses o más.', 'RANGO_FONDO_INGRESO', 'M4 - Cobertura muy alta', 'Ingreso no válido = no calculable.', fon_lim, MULT),
]
M += [
    ('Categoría interna', 'MAX_CAT_RIESGO', '—', '—', 'Categoría', 'A1', 'Valor del archivo.', 'CAT_INTERNA', 'A1',
     'Vacío sin crédito vigente = "Sin historial interno".', 'Es la máxima categoría registrada; no indica días de mora.', None),
    ('Categoría interna', 'MAX_CAT_RIESGO', '—', '—', 'Categoría', 'A2', 'Valor del archivo.', 'CAT_INTERNA', 'A2',
     '', 'A2 cumple el criterio, pero se marca como alerta.', None),
    ('Categoría interna', 'MAX_CAT_RIESGO', '—', '—', '—', 'Sin historial interno', 'Sin categoría y sin crédito vigente.',
     'CAT_INTERNA', 'Sin historial interno', '', 'Requiere definición: ¿se acepta sin historial?', None),
    ('Categoría interna', 'MAX_CAT_RIESGO', '—', '—', '—', 'Dato faltante (con crédito vigente)', 'Tiene crédito vigente pero sin categoría.',
     'CAT_INTERNA', 'Dato faltante (con crédito vigente)', '', 'Verificar en el sistema.', None),
    ('Calificación externa (SSF)', 'CALIFICACION_SSF', '—', '—', 'Categoría', 'A1', 'Valor del archivo.', 'CAT_EXTERNA_SSF', 'A1',
     '', 'Una sola calificación/institución por asociado.', None),
    ('Calificación externa (SSF)', 'CALIFICACION_SSF', '—', '—', 'Categoría', 'A2', 'Valor del archivo.', 'CAT_EXTERNA_SSF', 'A2',
     '', 'Alerta.', None),
    ('Calificación externa (SSF)', 'CALIFICACION_SSF', '—', '—', '—', 'Dato faltante (con saldo SSF)', 'Tiene saldo/institución, calificación en blanco.',
     'CAT_EXTERNA_SSF', 'Dato faltante (con saldo SSF)', 'Texto vacío en el archivo.', 'Verificar con la consulta SSF.', None),
    ('Calificación externa (SSF)', 'CALIFICACION_SSF', '—', '—', '—', 'Sin registro SSF', 'Sin saldo ni institución.',
     'CAT_EXTERNA_SSF', 'Sin registro SSF', '', 'No se puede validar.', None),
]
for var in ['Días de mora / obligaciones vencidas', 'Créditos castigados', 'Listas de exclusión (listas negras)',
            'Número de obligaciones', 'Historial de pago detallado', 'Solvencia tributaria', 'Monto solicitado / plazo / tasa',
            'Tipo de ingreso (bruto / neto / disponible)']:
    M.append((var, 'No disponible en la base', '—', '—', '—', 'No disponible en la base', 'La columna no existe en el archivo.',
              None, None, 'Requiere consulta a otra fuente', 'No se puede evaluar ni inferir.', None))

r = 4
for (var, org, lo, hi, uni, niv, jus, cv, cc, falt, lim, fmt) in M:
    put(wm, r, 1, var, fb)
    put(wm, r, 2, org)
    put(wm, r, 3, lo, fmt=fmt if isinstance(lo, (int, float)) or str(lo).startswith('=') else None)
    put(wm, r, 4, hi, fmt=fmt if isinstance(hi, (int, float)) or str(hi).startswith('=') else None)
    put(wm, r, 5, uni)
    put(wm, r, 6, niv, fill=fill_warn if 'No disponible' in niv else None)
    put(wm, r, 7, jus)
    if cv:
        put(wm, r, 8, f'=COUNTIF({RNG(cv)},"{cc}")', fmt='#,##0')
        put(wm, r, 9, f'=H{r}/{TOTAL}', fmt=PCT)
    else:
        put(wm, r, 8, 'No disponible', fill=fill_warn)
        put(wm, r, 9, '—')
    put(wm, r, 10, falt)
    put(wm, r, 11, lim)
    r += 1
wm.cell(r + 1, 1, 'Los límites inferiores son incluyentes; el límite superior mostrado es el último valor del rango '
        '(montos: hasta $0.01 menos que el siguiente rango; relaciones: hasta antes del siguiente límite). '
        'Columnas C y D están enlazadas a la hoja Parametros.').font = Font(name=F, size=9, italic=True)
wm.freeze_panes = 'C4'

# ---------------------------------------------------------------- Cruces (evaluacion conjunta)
wx = wb.create_sheet('Evaluacion_Conjunta')
wx['A1'] = 'EVALUACIÓN CONJUNTA: COMBINACIONES DE VARIABLES Y ALERTAS (COUNTIFS sobre BD_CLASIFICADA)'
wx['A1'].font = ftitle
r = 3


def cross(title, rv, rcats, cv, ccats, note=''):
    global r
    wx.cell(r, 1, title).font = fsub
    if note:
        wx.cell(r + 1, 1, note).font = Font(name=F, size=9, italic=True)
        r += 1
    r += 1
    hdr(wx, r, [rv + ' \\ ' + cv] + ccats + ['Total'], fill=fill_calc)
    r += 1
    s = r
    for rc in rcats:
        put(wx, r, 1, rc, fb)
        for j, cc in enumerate(ccats):
            put(wx, r, 2 + j, f'=COUNTIFS({RNG(rv)},$A{r},{RNG(cv)},{L(2+j)}${s-1})', fmt='#,##0')
        put(wx, r, 2 + len(ccats), f'=SUM(B{r}:{L(1+len(ccats))}{r})', fb, '#,##0')
        r += 1
    put(wx, r, 1, 'Total', fb, fill=fill_note)
    for j in range(len(ccats) + 1):
        put(wx, r, 2 + j, f'=SUM({L(2+j)}{s}:{L(2+j)}{r-1})', fb, '#,##0', fill_note)
    r += 3


wx.column_dimensions['A'].width = 44
for j in range(2, 12):
    wx.column_dimensions[L(j)].width = 17
cross('1. Ingreso × saldo de deuda SSF', 'RANGO_INGRESO', CATS['RANGO_INGRESO'], 'RANGO_DEUDA', CATS['RANGO_DEUDA'],
      'Combinación crítica: ingreso bajo con deuda elevada o muy elevada.')
cross('2. Relación deuda-ingreso × antigüedad', 'RANGO_DEUDA_INGRESO', CATS['RANGO_DEUDA_INGRESO'], 'NIVEL_ANTIGUEDAD',
      CATS['NIVEL_ANTIGUEDAD'])
cross('3. Fondo compensatorio × saldo de deuda SSF', 'RANGO_FONDO', CATS['RANGO_FONDO'], 'RANGO_DEUDA', CATS['RANGO_DEUDA'],
      'Fondos reducidos frente a deuda elevada. El fondo no se considera mitigante hasta validar su disponibilidad.')
cross('4. Antigüedad × categoría interna', 'NIVEL_ANTIGUEDAD', CATS['NIVEL_ANTIGUEDAD'], 'CAT_INTERNA', CATS['CAT_INTERNA'],
      'Antigüedad reducida y poca información crediticia.')
cross('5. Categoría interna × calificación SSF', 'CAT_INTERNA', CATS['CAT_INTERNA'], 'CAT_EXTERNA_SSF', CATS['CAT_EXTERNA_SSF'])
cross('6. Estado descriptivo × elegibilidad general', 'ESTADO_DESCRIPTIVO', CATS['ESTADO_DESCRIPTIVO'], 'ELEGIBILIDAD_GENERAL',
      CATS['ELEGIBILIDAD_GENERAL'], 'Indicadores favorables (cumple criterios) junto con alertas relevantes.')
wx.cell(r, 1, '7. Resumen de alertas por asociado').font = fsub
r += 1
hdr(wx, r, ['Alerta', 'Asociados con la alerta', '% del total', 'Tipo', 'Definición'], fill=fill_calc)
r += 1
AL_DEF = {
    'AL_ING_BAJO_DEUDA_ELEVADA': ('Riesgo', 'Rango ingreso I y deuda D3 o D4.'),
    'AL_DEUDA_ALTA_VS_INGRESO': ('Riesgo', 'Saldo SSF ≥ DTI_L2 veces el ingreso mensual (R3 o R4).'),
    'AL_ANTIG_BAJA_SIN_HISTORIAL': ('Riesgo', 'Antigüedad nivel Alto (0-1 año) y sin historial interno.'),
    'AL_FONDO_BAJO_DEUDA_ELEVADA': ('Riesgo', 'Fondo cero o F1 y deuda D3 o D4.'),
    'AL_CATEGORIA_A2': ('Riesgo', 'Categoría A2 interna o en SSF (cumple criterio pero es señal de atención).'),
    'AL_APORT_VENCIDAS': ('Riesgo', 'APORTACIONES_AL_DIA = NO o cuotas de aportación vencidas > 0.'),
    'DQ_PRECA_SIN_INGRESO': ('Calidad de datos', 'Tiene monto pre evaluado (TC, vehículo o hipoteca) pero ingreso vacío o 0.'),
    'DQ_PRECA_FUERA_TABLA': ('Calidad de datos', 'Monto pre evaluado distinto al que resulta de la tabla con el INGRESO de la base.'),
    'DQ_DATO_INCONSISTENTE': ('Calidad de datos', 'Antigüedad imposible o > edad−18, ingreso atípico, saldo negativo, edad que no cuadra o DUI ≠ 9 dígitos.'),
}
for a, (tp, df_) in AL_DEF.items():
    put(wx, r, 1, a, fb)
    put(wx, r, 2, f'=SUM({RNG(a)})', fmt='#,##0')
    put(wx, r, 3, f'=B{r}/{TOTAL}', fmt=PCT)
    put(wx, r, 4, tp)
    put(wx, r, 5, df_)
    r += 1
wx.column_dimensions['E'].width = 60
r += 1
wx.cell(r, 1, 'No se asigna un nivel global ponderado: la base no contiene resultados históricos de pago (mora, castigos) '
        'que permitan estimar o validar ponderaciones. El conteo de alertas es descriptivo.').font = Font(name=F, size=10, bold=True, color='C00000')

# ---------------------------------------------------------------- Productos
wpr = wb.create_sheet('Analisis_Producto')
wpr['A1'] = 'ANÁLISIS POR PRODUCTO: EXPOSICIÓN PRE EVALUADA Y CONSISTENCIA CON LAS TABLAS'
wpr['A1'].font = ftitle
hdr(wpr, 3, ['Indicador', 'Tarjeta de crédito', 'Vehículo', 'Hipotecario', 'Nota'], widths=[52, 18, 18, 18, 70])
prods = [('TC', 'TC', 'LIMITE_TC_SOBRE_INGRESO'), ('VEH', 'VEHICULO', 'VEHICULO_SOBRE_INGRESO'), ('HIPO', 'HIPO', 'HIPO_SOBRE_INGRESO')]
rows = [
    ('Asociados con monto pre evaluado', lambda k, nm, m: f'=COUNT({ORNG(k)})', '#,##0', 'Celdas "N/A" = no pre evaluado.'),
    ('Exposición total pre evaluada (USD)', lambda k, nm, m: f'=SUM({ORNG(k)})', USD,
     'Monto potencial si todos aceptaran; no es saldo desembolsado.'),
    ('Monto promedio (USD)', lambda k, nm, m: f'=AVERAGE({ORNG(k)})', USD, ''),
    ('Monto máximo (USD)', lambda k, nm, m: f'=MAX({ORNG(k)})', USD, 'Vehículo: un caso de 150,000 fuera de la tabla (máx. 40,000).'),
    ('Coincide con tabla', lambda k, nm, m: f'=COUNTIF({RNG("CONSIST_"+nm)},"Coincide con tabla")', '#,##0', ''),
    ('No coincide con tabla', lambda k, nm, m: f'=COUNTIF({RNG("CONSIST_"+nm)},"No coincide con tabla")', '#,##0',
     'Ej.: monto asignado con ingreso menor al mínimo de la tabla.'),
    ('Pre evaluado sin ingreso válido (vacío o 0)', lambda k, nm, m: f'=COUNTIF({RNG("CONSIST_"+nm)},"Pre evaluado sin ingreso válido")',
     '#,##0', 'El monto no puede reproducirse con la base: el ingreso usado no está en el archivo.'),
    ('% pre evaluados sin ingreso válido', lambda k, nm, m: f'=IFERROR({L(2+["TC","VEH","HIPO"].index(k))}{{r6}}/{L(2+["TC","VEH","HIPO"].index(k))}4,0)', PCT, ''),
    ('Mediana monto / ingreso mensual (veces)', lambda k, nm, m: f'=MEDIAN({RNG(m)})', MULT, 'Solo con ingreso válido.'),
    ('P90 monto / ingreso mensual (veces)', lambda k, nm, m: f'=PERCENTILE({RNG(m)},0.9)', MULT, ''),
    ('Pre evaluados que cumplen criterios verificables',
     lambda k, nm, m: f'=COUNTIFS({ORNG(k)},">0",{RNG("ELEGIBILIDAD_TC" if k=="TC" else "ELEGIBILIDAD_GENERAL")},"Cumple criterios verificables")',
     '#,##0', 'TC usa ELEGIBILIDAD_TC (excluye a quienes ya tienen tarjeta).'),
    ('Pre evaluados con información incompleta',
     lambda k, nm, m: f'=COUNTIFS({ORNG(k)},">0",{RNG("ELEGIBILIDAD_TC" if k=="TC" else "ELEGIBILIDAD_GENERAL")},"Información incompleta")', '#,##0', ''),
    ('Pre evaluados que NO cumplen algún criterio verificable',
     lambda k, nm, m: f'=COUNTIFS({ORNG(k)},">0",{RNG("ELEGIBILIDAD_TC" if k=="TC" else "ELEGIBILIDAD_GENERAL")},"No cumple criterio verificable")', '#,##0', ''),
    ('Pre evaluados con 1 o más alertas de riesgo',
     lambda k, nm, m: f'=COUNTIFS({ORNG(k)},">0",{RNG("TOTAL_ALERTAS_RIESGO")},">=1")', '#,##0', ''),
    ('Pre evaluados con deuda/ingreso R3 o R4',
     lambda k, nm, m: f'=COUNTIFS({ORNG(k)},">0",{RNG("RANGO_DEUDA_INGRESO")},"R3 - Alto")+COUNTIFS({ORNG(k)},">0",{RNG("RANGO_DEUDA_INGRESO")},"R4 - Muy alto")',
     '#,##0', ''),
]
r = 4
for lab, fx, fmt, note in rows:
    put(wpr, r, 1, lab, fb)
    for j, (k, nm, m) in enumerate(prods):
        f_ = fx(k, nm, m).replace('{r6}', str(r - 1))
        put(wpr, r, 2 + j, f_, fmt=fmt)
    put(wpr, r, 5, note)
    r += 1
r += 1
wpr.cell(r, 1, 'Distribución de montos pre evaluados').font = fsub
r += 1
for k, nm, title, amounts in [('TC', 'TC', 'Límite TC', [500, 1500, 2000, 2500, 3000, 3500, 4500, 5000, 7000, 10000]),
                              ('VEH', 'VEHICULO', 'Monto vehículo', [20000, 25000, 30000, 40000, 150000]),
                              ('HIPO', 'HIPO', 'Monto hipotecario', [110000, 150000, 205000])]:
    hdr(wpr, r, [title, 'Asociados', '% de pre evaluados', 'Exposición (USD)'], fill=fill_calc)
    r += 1
    s = r
    for a in amounts:
        put(wpr, r, 1, a, fb, '#,##0')
        put(wpr, r, 2, f'=COUNTIF({ORNG(k)},A{r})', fmt='#,##0')
        put(wpr, r, 3, f'=B{r}/COUNT({ORNG(k)})', fmt=PCT)
        put(wpr, r, 4, f'=A{r}*B{r}', fmt=USD)
        r += 1
    put(wpr, r, 1, 'Total', fb, fill=fill_note)
    put(wpr, r, 2, f'=SUM(B{s}:B{r-1})', fb, '#,##0', fill_note)
    put(wpr, r, 3, f'=SUM(C{s}:C{r-1})', fb, PCT, fill_note)
    put(wpr, r, 4, f'=SUM(D{s}:D{r-1})', fb, USD, fill_note)
    r += 2
lims = [
    'Limitaciones por producto:',
    '• No hay monto solicitado, plazo, tasa, garantía, valor del inmueble/vehículo ni prima: no se puede calcular la cuota '
    'de los montos pre evaluados ni la relación cuota-ingreso resultante.',
    '• Hipotecario: el archivo no trae tabla de montos; la regla de la hoja Parametros es inferida y debe confirmarse.',
    '• Vehículo: la hoja RANGOS VEHICULO excluye ingresos < $3,000 ("NO APLICA"), pero hay pre evaluados con ingreso menor.',
    '• Tarjeta: la tabla llama "DORADA" al segmento que la base llama "GOLD", y asigna $500 a "CLÁSICA LOCAL" mientras '
    'la base lo rotula "CLASICA INTERNACIONAL".',
    '• Otros créditos de consumo: no hay montos pre evaluados en la base; no es posible evaluarlos.',
    '• No se aplican los mismos umbrales de riesgo a todos los productos: deuda/ingreso es un indicador de saldo y '
    'debe complementarse con la cuota de cada producto cuando esté disponible.',
]
for t in lims:
    wpr.cell(r, 1, t).font = fb if t.endswith(':') else fn
    r += 1

# ---------------------------------------------------------------- Calidad y conciliacion
wq = wb.create_sheet('Calidad_Conciliacion')
wq['A1'] = 'VALIDACIÓN DE CALIDAD Y CONCILIACIÓN BASE ORIGINAL ↔ BASE CLASIFICADA'
wq['A1'].font = ftitle
hdr(wq, 3, ['Control', 'Base original (BD FINAL)', 'Base clasificada', 'Diferencia', 'Resultado'], widths=[58, 22, 22, 16, 40])
SRCR = lambda col: f"'BD FINAL'!${col}${R1}:${col}${RN}"
ctrl = [
    ('Registros (ASOCIADO no vacío)', f'=COUNTA({SRCR("A")})', f'={TOTAL}', '#,##0'),
    ('Suma INGRESO', f'=SUM({SRCR("P")})', f'=SUM({ORNG("ING")})', USD),
    ('Registros con INGRESO numérico', f'=COUNT({SRCR("P")})', f'=COUNT({ORNG("ING")})', '#,##0'),
    ('Suma FONDO_COMPENSATORIO', f'=SUM({SRCR("Q")})', f'=SUM({ORNG("FON")})', USD),
    ('Suma SALDO_SSF', f'=SUM({SRCR("AD")})', f'=SUM({ORNG("SSF")})', USD),
    ('Suma MONTO_TOTAL_APORTACIONES', f'=SUM({SRCR("AA")})', f'=SUM({ORNG("APM")})', USD),
    ('Suma PRECA HIPO', f'=SUM({SRCR("R")})', f'=SUM({ORNG("HIPO")})', USD),
    ('Suma PRECA VEHICULO', f'=SUM({SRCR("S")})', f'=SUM({ORNG("VEH")})', USD),
    ('Suma PRECA TC', f'=SUM({SRCR("T")})', f'=SUM({ORNG("TC")})', USD),
    ('Suma ANTIGUEDAD_ANIOS (incluye valores inconsistentes)', f'=SUM({SRCR("O")})', f'=SUM({ORNG("ANT")})', '#,##0'),
]
r = 4
for lab, a, b, fmt in ctrl:
    put(wq, r, 1, lab, fb)
    put(wq, r, 2, a, fmt=fmt)
    put(wq, r, 3, b, fmt=fmt)
    put(wq, r, 4, f'=ROUND(C{r}-B{r},2)', fmt=fmt)
    put(wq, r, 5, f'=IF(D{r}=0,"OK: concilia","REVISAR")')
    r += 1
r += 1
hdr(wq, r, ['Control de identificadores y datos', 'Cantidad', 'Resultado / nota'], fill=fill_calc)
r += 1
ctrl2 = [
    ('ASOCIADO duplicados', f'=SUMPRODUCT((COUNTIF({ORNG("ASOC")},{ORNG("ASOC")})>1)*1)', 'Debe ser 0: un registro por asociado.'),
    ('DUI duplicados', f'=SUMPRODUCT((COUNTIF({ORNG("DUI")},{ORNG("DUI")})>1)*1)', 'Debe ser 0.'),
    ('DUI con longitud distinta de 9 caracteres', f'=SUMPRODUCT((LEN({ORNG("DUI")})<>9)*1)', 'Verificar formato (posible pérdida de ceros o error de captura).'),
    ('Edad fuera de 30-55 a la fecha de corte del memorando', f'=COUNTIF({RNG("C_EDAD")},"No cumple")', 'Según la columna EDAD todos están en rango; recalculada al 30/09/2026 no.'),
    ('Edad que no coincide con fecha de nacimiento a la fecha de corte', f'=COUNTIF({RNG("CHK_EDAD")},"Difiere - verificar")', 'La columna EDAD parece calculada al 02/09/2026, no al corte del memorando.'),
    ('Antigüedad vacía', f'=COUNTIF({RNG("NIVEL_ANTIGUEDAD")},"Sin dato")', 'No se clasifica como 0.'),
    ('Antigüedad inconsistente (negativa o > edad)', f'=COUNTIF({RNG("NIVEL_ANTIGUEDAD")},"Inconsistente - revisar")', 'Asociados 23785 (−3992), 15868 (−1), 27463 (1776).'),
    ('Antigüedad > edad − EDAD_ADULTA (verificar)', f'=SUM({RNG("ALERTA_ANT_EDAD")})', 'Se clasifica, pero se marca alerta.'),
    ('INGRESO vacío', f'=COUNTIF({RNG("RANGO_INGRESO")},"Sin dato")', ''),
    ('INGRESO = 0', f'=COUNTIF({RNG("RANGO_INGRESO")},"Ingreso cero - revisar")', ''),
    ('INGRESO atípico (> umbral P75+3×RIC)', f'=SUM({RNG("ALERTA_ING_ATIPICO")})', 'No se excluye; verificar (máx. $132,222.90).'),
    ('SALDO_SSF negativo', f'=COUNTIF({ORNG("SSF")},"<0")', 'Asociados 23151, 17848, 18916 (DAVIVIENDA, −$0.02 a −$0.89).'),
    ('CALIFICACION_SSF en blanco con saldo SSF', f'=COUNTIF({RNG("CAT_EXTERNA_SSF")},"Dato faltante (con saldo SSF)")', 'Texto vacío ("") en la base.'),
    ('Con crédito vigente y sin categoría interna', f'=COUNTIF({RNG("CAT_INTERNA")},"Dato faltante (con crédito vigente)")', ''),
    ('PRECA VEHICULO fuera de la tabla (150,000)', f'=COUNTIF({ORNG("VEH")},150000)', 'Asociado 20403: probable error de captura.'),
    ('Asociados con alguna alerta de calidad de datos', f'=COUNTIF({RNG("TOTAL_ALERTAS_DATOS")},">0")', ''),
    ('Columnas sin encabezado en BD FINAL (AF:AI) con valores de prueba', 3,
     'Filas 2-4 tienen 1/2/3 y "Bajo"/"Medi"/"Alto" en AH:AI: parecen pruebas manuales; no se usan. '
     'La hoja "Variables" escribe "Meido" (Medio).'),
    ('Columna CONCATENADO', f"=COUNTA({SRCR('J')})", 'Vacía en todos los registros; no se usa.'),
]
for lab, f_, note in ctrl2:
    put(wq, r, 1, lab, fb)
    put(wq, r, 2, f_, fmt='#,##0')
    put(wq, r, 3, note)
    r += 1
wq.column_dimensions['C'].width = 70
r += 1
wq.cell(r, 1, 'Ningún registro fue eliminado. Las 31 columnas originales se copian sin modificación en BD_CLASIFICADA (A:AE); '
        'la hoja BD FINAL se conserva intacta.').font = fb

# ---------------------------------------------------------------- Metodologia
wmet = wb.create_sheet('Metodologia')
wmet.column_dimensions['A'].width = 34
wmet.column_dimensions['B'].width = 120
wmet['A1'] = 'METODOLOGÍA, FÓRMULAS, SUPUESTOS Y RESTRICCIONES'
wmet['A1'].font = ftitle
met = [
    ('SECCIÓN', 'A. Origen de la información'),
    ('Memorando', 'Memorando del 27/10/2026 "Validación de base pre evaluada de Crédito Hipotecario, vehículo y Tarjeta de Crédito" '
     '(Créditos, TC y Mercadeo → Jefe de Riesgos). Universo 10,314 asociados al 30/09/2026; base depurada declarada de 2,263 '
     'registros. Criterios, otros criterios, consolidado por producto y solicitud de opinión de Riesgos: ver hoja Cumplimiento_Memorando.'),
    ('Base', 'Archivo "Base_de_Asociados_pre_evaluada_creditos_y_tc_para_analisis_de_Riesgos.xlsx", hoja BD FINAL: '
     f'{N} registros, 31 columnas con encabezado (+4 columnas sin uso). Hojas auxiliares: Variables, RANGOS VEHICULO, RANGOS TC.'),
    ('Unidad de análisis', 'Un registro = un asociado (ASOCIADO y DUI sin duplicados). Saldo SSF ya viene en una sola fila por '
     'asociado, por lo que no se consolida ni se suma nada: no hay riesgo de duplicar saldos dentro de esta base.'),
    ('SECCIÓN', 'B. Fórmulas'),
    ('Edad recalculada', 'AÑO(corte) − AÑO(nacimiento) − 1 si aún no cumplía años a la fecha de corte.'),
    ('Nivel antigüedad', 'Vacío → Sin dato; <0 o > EDAD → Inconsistente; ≤ ANT_ALTO_MAX → Alto; ≤ ANT_MEDIO_MAX → Medio; resto → Bajo.'),
    ('Rango ingreso', 'Vacío → Sin dato; ≤0 → Ingreso cero; < ING_L1 → I; < ING_L2 → II; < ING_L3 → III; resto → IV.'),
    ('Ingreso atípico', 'INGRESO > P75 + 3 × (P75 − P25), calculado sobre ingresos > 0 (regla de Tukey, valla "extrema").'),
    ('Rango deuda', 'Sin saldo y sin institución → Sin registro SSF; ≤ 0 → Sin saldo; < DEU_L1 → D1; < DEU_L2 → D2; < DEU_L3 → D3; resto → D4.'),
    ('Relación deuda-ingreso', 'MAX(SALDO_SSF,0) / INGRESO, expresada en VECES el ingreso mensual. Solo si INGRESO > 0 y hay dato SSF.'),
    ('Relación cuota-ingreso', 'NO CALCULADA: la base no contiene cuotas mensuales. Fórmula prevista: Σ cuotas mensuales / INGRESO × 100.'),
    ('Carga financiera mensual', 'NO CALCULADA por la misma razón. Fórmula prevista: Σ cuotas mensuales de todas las obligaciones.'),
    ('Fondo / ingreso', 'FONDO_COMPENSATORIO / INGRESO, en meses de ingreso. Solo si INGRESO > 0.'),
    ('Criterios de elegibilidad', 'Cada criterio devuelve "Cumple", "No cumple" o "Sin información". Elegibilidad general: si algún '
     'criterio = No cumple → "No cumple criterio verificable"; si no, si alguno = Sin información → "Información incompleta"; si no → '
     '"Cumple criterios verificables". Listas, castigos y solvencia tributaria siempre quedan como "Requiere consulta externa".'),
    ('Elegibilidad TC', 'Igual a la general, pero quien ya tiene tarjeta (TIENE_TARJETA = SI) queda "Excluido: ya tiene tarjeta".'),
    ('Consistencia de producto', 'Monto esperado = LOOKUP(INGRESO, tabla de Parametros). Se compara con el monto PRECA de la base.'),
    ('Alertas', 'Seis alertas de riesgo (1/0) y tres de calidad de datos (1/0); ver hoja Evaluacion_Conjunta. Estado descriptivo = '
     'conteo de alertas de riesgo.'),
    ('SECCIÓN', 'C. Supuestos'),
    ('S1', 'FECHA_CORTE = 2026-09-30 (memorando). La columna EDAD parece calculada al 02/09/2026; la edad se recalcula al corte.'),
    ('S2', 'INGRESO es mensual y en USD (coherente con las tablas RANGOS TC / VEHICULO). No se sabe si es bruto o neto.'),
    ('S3', 'SALDO_SSF corresponde a la Central de Riesgo SSF y se toma como saldo reportado por asociado. La base muestra una sola '
     'INSTITUCION por asociado, por lo que podría no ser el total consolidado de todas sus deudas.'),
    ('S4', 'Un asociado sin saldo ni institución SSF se trata como "Sin registro SSF" (no como "sin deuda"): no se sabe si no tiene '
     'deudas o si no se consultó.'),
    ('S5', 'Fondo compensatorio se trata como respaldo descriptivo; no se supone disponible para pagar obligaciones.'),
    ('S6', 'Las categorías A1/A2 se consideran aceptables (memorando, criterios 5 y 6); A2 se marca como alerta.'),
    ('S7', 'Aportaciones no al día NO excluyen (memorando, otros 4): se reportan como condición para la TC y como alerta.'),
    ('S8', 'Criterio 4 "Antigüedad como asociado 0 meses" se interpreta como "sin antigüedad mínima": se exige solo que exista el dato.'),
    ('SECCIÓN', 'D. Restricciones'),
    ('R1', 'No existen en la base: días de mora, créditos castigados, listas de exclusión, número de obligaciones, historial de pago, '
     'solvencia tributaria, cuotas, plazos, tasas, montos solicitados.'),
    ('R2', 'Los rangos son estadísticos (percentiles de esta base) y en algunos casos alineados con tablas internas del archivo. NO son '
     'límites regulatorios. Cualquier referencia normativa (p. ej. normas de la SSF / BCR sobre gestión de riesgo crediticio) '
     'requiere verificación documental por Riesgos/Cumplimiento.'),
    ('R3', 'No hay variable de resultado (incumplimiento histórico), por lo que no se proponen ponderaciones ni un puntaje validado.'),
    ('R4', 'La segmentación es descriptiva y de alerta. La elegibilidad es preliminar. Ninguna columna constituye una decisión crediticia.'),
    ('SECCIÓN', 'E. Actualización'),
    ('Procedimiento', f'Para una base nueva: pegar los datos en BD FINAL y en BD_CLASIFICADA columnas A:AE respetando el orden; '
     f'extender las fórmulas de las columnas AF en adelante a todas las filas y ampliar los rangos de las fórmulas (actualmente filas '
     f'{R1}-{RN}). Alternativamente, ejecutar el script build_segmentacion.py con el archivo nuevo, que regenera todo el libro.'),
]
r = 3
for a, b in met:
    if a == 'SECCIÓN':
        wmet.cell(r, 1, b).font = fsub
    else:
        put(wmet, r, 1, a, fb)
        put(wmet, r, 2, b)
    r += 1

# ---------------------------------------------------------------- Pendientes
wpe = wb.create_sheet('Pendientes_Aprobacion')
wpe['A1'] = 'PARÁMETROS Y DEFINICIONES PENDIENTES DE APROBACIÓN INSTITUCIONAL'
wpe['A1'].font = ftitle
hdr(wpe, 3, ['#', 'Tema', 'Propuesta / situación actual', 'Decisión requerida', 'Responsable sugerido', 'Estado'],
    widths=[5, 34, 70, 56, 22, 16])
pend = [
    ('Número de registros', 'Memorando: 2,263; archivo: 2,262 asociados (2,263 filas contando el encabezado).', 'Confirmar cifra oficial.', 'Créditos'),
    ('Criterio 3 "No poseen TC"', '966 asociados SÍ tienen tarjeta (solo pre evaluados para hipotecario/vehículo).', 'Aclarar si el criterio aplica solo al universo TC.', 'Créditos / Comercial TC'),
    ('Criterio 4 "Antigüedad 0 meses"', 'Redacción ambigua; 843 sin antigüedad registrada.', 'Definir antigüedad mínima real.', 'Riesgos / Créditos'),
    ('Criterio 7 edad al corte', 'A la fecha de corte 30/09/2026, 8 asociados ya tienen 56 años.', 'Confirmar fecha de cálculo de edad; excluir o justificar.', 'Créditos'),
    ('Criterio 8 relacionados/empleados', 'La base no tiene campo para verificarlo.', 'Adjuntar validación contra listado de relacionados y empleados.', 'Créditos / Cumplimiento'),
    ('61 registros sin deuda SSF', 'No se reproduce: 342 sin registro SSF y 166 con saldo 0.', 'Explicar el criterio usado para 61.', 'Créditos'),
    ('54 asociados con aportaciones no al día', 'En la base hay 186 no al día; 54 son los pre evaluados en TC.', 'Confirmar tratamiento de los 132 de hipotecario/vehículo.', 'Créditos / Riesgos'),
    ('Fecha de edad / corte', 'Memorando: 30/09/2026; EDAD coincide con 02/09/2026.', 'Confirmar fecha de extracción.', 'TI / Riesgos'),
    ('Rangos de ingreso', 'I < $1,500; II $1,500-$2,999.99; III $3,000-$4,999.99; IV ≥ $5,000.', 'Aprobar o ajustar límites.', 'Riesgos'),
    ('Naturaleza del ingreso', 'Se desconoce si es bruto, neto o disponible, y su fecha de actualización.', 'Definir y documentar la fuente.', 'Negocios / TI'),
    ('Ingreso cero y vacío', '1,418 asociados (62.7%) sin ingreso válido y todos con algún monto pre evaluado.',
     'Definir si se excluyen, se completan o se evalúan con otra fuente.', 'Riesgos / Negocios'),
    ('Rangos de deuda SSF', 'D1 < $1,000; D2 < $10,000; D3 < $50,000; D4 ≥ $50,000.', 'Aprobar o ajustar.', 'Riesgos'),
    ('Saldo SSF consolidado', 'Una institución por asociado.', 'Confirmar si SALDO_SSF es total consolidado o sólo la principal.', 'Riesgos / TI'),
    ('Relación deuda-ingreso', 'R1 < 3x; R2 < 12x; R3 < 24x; R4 ≥ 24x (estadístico).', 'Aprobar; definir si se diferencia por producto.', 'Riesgos'),
    ('Relación cuota-ingreso', 'No calculable.', 'Proveer cuotas de todas las obligaciones (internas y SSF) y aprobar límite institucional por producto.', 'Riesgos / TI'),
    ('Rangos de fondo compensatorio', 'F1 < $3,000; F2 < $15,000; F3 < $50,000; F4 ≥ $50,000.', 'Aprobar o ajustar.', 'Riesgos'),
    ('Disponibilidad del fondo', 'Naturaleza no documentada.', 'Definir si es retirable, pignorable o garantía, y si mitiga riesgo.', 'Riesgos / Legal'),
    ('Asociados sin historial interno', 'Se marcan "Sin información" en categoría interna.', '¿Cumplen el criterio A1/A2 o requieren otro tratamiento?', 'Riesgos'),
    ('Categoría A2', 'Cumple criterio, pero se marca alerta.', 'Confirmar tratamiento.', 'Riesgos'),
    ('Aportaciones vencidas', '186 asociados no al día.', 'Confirmar si excluye o permite pre evaluación con condición.', 'Riesgos'),
    ('Tabla hipotecaria', 'Regla inferida (3,000→110,000; 5,000→150,000; 9,000→205,000).', 'Entregar tabla oficial; confirmar corte entre $8,900 y $9,000.', 'Negocios / Riesgos'),
    ('Montos fuera de tabla', 'Pre evaluados con ingreso por debajo del mínimo de la tabla y un vehículo de 150,000.', 'Explicar regla aplicada o corregir.', 'Negocios'),
    ('Segmentos TC', 'DORADA vs GOLD; $500 rotulado como CLÁSICA INTERNACIONAL.', 'Unificar nomenclatura.', 'Negocios'),
    ('Consultas externas', 'Listas, castigos y solvencia tributaria no están en la base.', 'Realizar consultas y adjuntar resultado por asociado.', 'Cumplimiento / Riesgos'),
    ('Ponderaciones / puntaje global', 'No se proponen por falta de datos de desempeño.', 'Proveer historial de pago para calibrar y validar (backtesting).', 'Riesgos'),
    ('Umbral de ingreso atípico', 'P75 + 3×RIC.', 'Aprobar método de verificación.', 'Riesgos'),
    ('Edad mínima de afiliación (alerta)', '18 años (supuesto).', 'Confirmar.', 'Riesgos'),
]
for i, (t, p, d, rsp) in enumerate(pend, 1):
    rr = 3 + i
    put(wpe, rr, 1, i)
    put(wpe, rr, 2, t, fb)
    put(wpe, rr, 3, p)
    put(wpe, rr, 4, d)
    put(wpe, rr, 5, rsp)
    put(wpe, rr, 6, 'Pendiente', Font(name=F, size=10, bold=True, color='C00000'))


# ---------------------------------------------------------------- Variables_Rangos
wv = wb.create_sheet('Variables_Rangos')
wv['A1'] = 'RANGOS Y NIVELES POR VARIABLE (mismo formato de la hoja "Variables" del archivo)'
wv['A1'].font = ftitle
wv['A2'] = ('Propuesta pendiente de aprobación por Riesgos. Límites enlazados a la hoja Parametros; el nivel de cada asociado '
            'está en BD_CLASIFICADA (columnas NIVEL_*). Antigüedad: niveles solicitados. Ingreso, deuda y fondo: rangos calculados '
            'con la distribución de esta base.')
wv['A2'].font = Font(name=F, size=9, italic=True)
TXT = lambda pid: f'TEXT({P[pid]},"$#,##0")'
TXTm = lambda pid: f'TEXT({P[pid]}-0.01,"$#,##0.00")'
TXTx = lambda pid: f'TEXT({P[pid]},"0")'
blocks = [
    ('ANTIGUEDAD_ANIOS', 'NIVEL_ANTIGUEDAD', 'Nivel de riesgo', [
        (f'="0 - "&{TXTx("ANT_ALTO_MAX")}', 'Alto'),
        (f'={TXTx("ANT_ALTO_MAX")}&"+1"', None),
    ]),
]
specs = [
    ('ANTIGUEDAD_ANIOS (años)', 'NIVEL_ANTIGUEDAD', 'Nivel de riesgo', [
        (f'="0 - "&{P["ANT_ALTO_MAX"]}', 'Alto'),
        (f'={P["ANT_ALTO_MAX"]}+1&" - "&{P["ANT_MEDIO_MAX"]}', 'Medio'),
        (f'={P["ANT_MEDIO_MAX"]}+1&" - Más"', 'Bajo'),
        ('Vacío', 'Sin dato'), ('Negativo o mayor que la edad', 'Inconsistente - revisar')]),
    ('INGRESO (USD mensual)', 'NIVEL_INGRESO', 'Nivel de riesgo', [
        (f'="$0.01 - "&{TXTm("ING_L1")}', 'Alto'),
        (f'={TXT("ING_L1")}&" - "&{TXTm("ING_L2")}', 'Medio'),
        (f'={TXT("ING_L2")}&" - "&{TXTm("ING_L3")}', 'Bajo'),
        (f'={TXT("ING_L3")}&" - Más"', 'Muy bajo'),
        ('Vacío o $0', 'Sin información')]),
    ('SALDO_SSF (USD, Central de Riesgo)', 'NIVEL_DEUDA', 'Nivel de riesgo', [
        ('$0 o a favor', 'Sin deuda'),
        (f'="$0.01 - "&{TXTm("DEU_L1")}', 'Bajo'),
        (f'={TXT("DEU_L1")}&" - "&{TXTm("DEU_L2")}', 'Medio'),
        (f'={TXT("DEU_L2")}&" - "&{TXTm("DEU_L3")}', 'Alto'),
        (f'={TXT("DEU_L3")}&" - Más"', 'Muy alto'),
        ('Sin registro SSF', 'Sin información')]),
    ('DEUDA / INGRESO (veces el ingreso mensual)', 'NIVEL_DEUDA_INGRESO', 'Nivel de riesgo', [
        ('0 veces', 'Sin deuda'),
        (f'="Más de 0 - menos de "&{P["DTI_L1"]}', 'Bajo'),
        (f'={P["DTI_L1"]}&" - menos de "&{P["DTI_L2"]}', 'Medio'),
        (f'={P["DTI_L2"]}&" - menos de "&{P["DTI_L3"]}', 'Alto'),
        (f'={P["DTI_L3"]}&" - Más"', 'Muy alto'),
        ('Ingreso vacío/$0 o sin dato SSF', 'Sin información')]),
    ('FONDO_COMPENSATORIO (USD)', 'NIVEL_RESPALDO_FONDO', 'Nivel de respaldo (no es riesgo)', [
        ('$0', 'Sin fondo'),
        (f'="$0.01 - "&{TXTm("FON_L1")}', 'Bajo'),
        (f'={TXT("FON_L1")}&" - "&{TXTm("FON_L2")}', 'Medio'),
        (f'={TXT("FON_L2")}&" - "&{TXTm("FON_L3")}', 'Alto'),
        (f'={TXT("FON_L3")}&" - Más"', 'Muy alto')]),
    ('OTROS CRITERIOS', None, None, None),
    ('FONDO / INGRESO (meses de ingreso)', 'NIVEL_COBERTURA_FONDO', 'Nivel de respaldo (no es riesgo)', [
        ('$0 de fondo', 'Sin fondo'),
        (f'="Más de 0 - menos de "&{P["FOI_L1"]}&" mes"', 'Bajo'),
        (f'={P["FOI_L1"]}&" - menos de "&{P["FOI_L2"]}&" meses"', 'Medio'),
        (f'={P["FOI_L2"]}&" - menos de "&{P["FOI_L3"]}&" meses"', 'Alto'),
        (f'={P["FOI_L3"]}&" meses - Más"', 'Muy alto'),
        ('Ingreso vacío o $0', 'Sin información')]),
    ('MONTO_TOTAL_APORTACIONES (USD)', 'NIVEL_APORTACIONES', 'Nivel de aportación (no es riesgo)', [
        (f'="$0.01 - "&{TXTm("APO_L1")}', 'Bajo'),
        (f'={TXT("APO_L1")}&" - "&{TXTm("APO_L2")}', 'Medio'),
        (f'={TXT("APO_L2")}&" - Más"', 'Alto')]),
    ('CUOTAS_VENC_APORT (cuotas de aportación vencidas)', 'NIVEL_MORA_APORTACIONES', 'Nivel de riesgo', [
        (f'="0 (al día)"', 'Bajo'),
        (f'={P["MORA_L1"]}&" - "&({P["MORA_L2"]}-1)', 'Medio'),
        (f'={P["MORA_L2"]}&" - Más"', 'Alto')]),
    ('MAX_CAT_RIESGO (categoría interna)', 'NIVEL_CAT_INTERNA', 'Nivel de riesgo', [
        ('A1', 'Bajo'), ('A2', 'Medio'), ('B, C, D o E', 'Alto'), ('Sin categoría', 'Sin información')]),
    ('CALIFICACION_SSF (categoría externa)', 'NIVEL_CAT_EXTERNA', 'Nivel de riesgo', [
        ('A1', 'Bajo'), ('A2', 'Medio'), ('B, C, D o E', 'Alto'), ('Sin calificación', 'Sin información')]),
    ('EDAD al 30/09/2026 (años)', 'TRAMO_EDAD', 'Tramo (descriptivo, sin nivel de riesgo)', [
        (f'="30 - "&({P["EDAD_T1"]}-1)', '30 a 39'), (f'={P["EDAD_T1"]}&" - "&({P["EDAD_T2"]}-1)', '40 a 49'),
        (f'={P["EDAD_T2"]}&" - "&{P["EDAD_MAX"]}', '50 a 55'), ('56 o más', 'Mayor de 55')]),
    ('TIENE_CREDITO_VIGENTE', 'ORIG:CRED', 'Valor', [('SI', 'SI'), ('NO', 'NO')]),
    ('TIENE_TARJETA', 'ORIG:TTC', 'Valor', [('SI', 'SI'), ('NO', 'NO')]),
    ('CUOTA / INGRESO', None, 'Nivel', [('No disponible en la base', 'No disponible en la base')]),
]
r = 4
for title, var, nivh, rows_ in specs:
    if rows_ is None:
        wv.cell(r, 1, title).font = fsub
        r += 2
        continue
    hdr(wv, r, [title, nivh, 'Asociados', '% del total'], fill=fill_h)
    r += 1
    s0 = r
    for rng_, lvl in rows_:
        put(wv, r, 1, rng_, fb)
        put(wv, r, 2, lvl, fill=fill_warn if 'No disponible' in lvl else None)
        if var:
            rg_ = ORNG(var[5:]) if var.startswith('ORIG:') else RNG(var)
            put(wv, r, 3, f'=COUNTIF({rg_},B{r})', fmt='#,##0')
            put(wv, r, 4, f'=C{r}/{TOTAL}', fmt=PCT)
        else:
            put(wv, r, 3, '—'); put(wv, r, 4, '—')
        r += 1
    if var:
        put(wv, r, 1, 'Total', fb, fill=fill_note)
        put(wv, r, 2, '', fill=fill_note)
        put(wv, r, 3, f'=SUM(C{s0}:C{r-1})', fb, '#,##0', fill_note)
        put(wv, r, 4, f'=SUM(D{s0}:D{r-1})', fb, PCT, fill_note)
        r += 1
    r += 1
for j, w in enumerate([40, 30, 12, 12], 1):
    wv.column_dimensions[L(j)].width = w

# ---------------------------------------------------------------- Cumplimiento memorando
wcm = wb.create_sheet('Cumplimiento_Memorando')
wcm['A1'] = '¿LA BASE APLICA CON LAS CONDICIONES DEL MEMORANDO? (Memorando 27/10/2026, corte 30/09/2026)'
wcm['A1'].font = ftitle
wcm['A2'] = ('Cada cifra de la base es una fórmula sobre BD_CLASIFICADA. Las cifras del memorando (azul) se transcribieron del Word. '
             '"NO verificable" = la base no contiene el dato; no se declara validado.')
wcm['A2'].font = Font(name=F, size=10, italic=True)
for j, w in enumerate([4, 44, 28, 11, 11, 13, 11, 44, 70], 1):
    wcm.column_dimensions[L(j)].width = w
r = 4
wcm.cell(r, 1, 'A. Criterios de la base depurada (memorando, criterios 1 a 8)').font = fsub
r += 1
hdr(wcm, r, ['#', 'Criterio del memorando', 'Columna(s) de la base', 'Cumplen', 'No cumplen', 'Sin información',
             '% que cumple', '¿La base aplica?', 'Observación'], fill=fill_calc)
r += 1
cm_rows = [
    (1, 'Poseen ingresos en sistema', 'INGRESO', 'C_INGRESO',
     'Criterio más incumplido: ingreso vacío o en cero, aunque todos tienen algún monto pre evaluado. '
     'El ingreso usado para asignar los límites no está en el archivo.'),
    (2, 'Poseen fondos compensatorios', 'FONDO_COMPENSATORIO', 'C_FONDO', 'Asociados con fondo igual a $0.00.'),
    (3, 'No poseen tarjeta de crédito', 'TIENE_TARJETA', 'C_SIN_TARJETA',
     'Los que tienen tarjeta NO tienen PRECA TC (solo hipotecario/vehículo). El criterio se cumple para el producto TC, '
     'pero no para la base completa como lo redacta el memorando.'),
    (4, 'Antigüedad como asociado 0 meses', 'ANTIGUEDAD_ANIOS', 'C_ANTIGUEDAD',
     'Redacción ambigua (se interpreta "sin antigüedad mínima"). Se verifica solo que exista el dato: hay vacíos y 3 valores imposibles.'),
    (5, 'Categoría A1 y A2 a nivel interno', 'MAX_CAT_RIESGO', 'C_CAT_INTERNA',
     'Nadie fuera de A1/A2. "Sin información" = sin categoría interna (en su mayoría sin crédito vigente).'),
    (6, 'Categoría A1 y A2 a nivel externo', 'CALIFICACION_SSF', 'C_CAT_EXTERNA',
     'Nadie fuera de A1/A2. "Sin información" = sin registro SSF o calificación en blanco con saldo.'),
    (7, 'Edad desde 30 hasta 55 años cumplidos', 'CLI_FECHA_NACIMIENTO (edad al 30/09/2026)', 'C_EDAD',
     'Según la columna EDAD todos cumplen, pero EDAD está calculada al 02/09/2026 aprox.; al corte del memorando '
     'algunos ya cumplieron 56. Con la vigencia de 3 meses habrá más casos.'),
]
for n_, cond, cols_, var, obs in cm_rows:
    put(wcm, r, 1, n_)
    put(wcm, r, 2, cond, fb)
    put(wcm, r, 3, cols_)
    put(wcm, r, 4, f'=COUNTIF({RNG(var)},"Cumple")', fmt='#,##0')
    put(wcm, r, 5, f'=COUNTIF({RNG(var)},"No cumple")', fmt='#,##0')
    put(wcm, r, 6, f'=COUNTIF({RNG(var)},"Sin información")', fmt='#,##0')
    put(wcm, r, 7, f'=D{r}/{TOTAL}', fmt=PCT)
    if var == 'C_SIN_TARJETA':
        put(wcm, r, 8, f'=IF(E{r}=0,"SÍ aplica","NO aplica en "&TEXT(E{r},"#,##0")&" registros; sí aplica al universo TC ("'
                       f'&TEXT(COUNTIFS({ORNG("TTC")},"SI",{ORNG("TC")},">0"),"0")&" tarjetahabientes con PRECA TC)")')
    else:
        put(wcm, r, 8, f'=IF(AND(E{r}=0,F{r}=0),"SÍ aplica (todos cumplen)",IF(E{r}>0,"NO aplica en "&TEXT(E{r},"#,##0")&" registros"'
                       f'&IF(F{r}>0," y sin dato en "&TEXT(F{r},"#,##0"),""),"NO verificable en "&TEXT(F{r},"#,##0")&" registros"))')
    put(wcm, r, 9, obs)
    r += 1
put(wcm, r, 1, 8)
put(wcm, r, 2, 'No incluye personas relacionadas ni empleados', fb)
put(wcm, r, 3, 'No disponible en la base', fill=fill_warn)
for k in (4, 5, 6, 7):
    put(wcm, r, k, '—')
put(wcm, r, 8, 'NO verificable con la base', fill=fill_warn)
put(wcm, r, 9, 'No hay campo de relación ni de empleado. Requiere cruce con el listado institucional.')
r += 2

wcm.cell(r, 1, 'B. Otros criterios del memorando').font = fsub
r += 1
hdr(wcm, r, ['#', 'Otro criterio', 'Columna(s)', 'Memorando', 'Base', 'Diferencia', '', '¿Se confirma con la base?', 'Observación'],
    fill=fill_calc)
r += 1
otros = [
    (1, 'Límites asignados según ingresos y fondos compensatorios', 'INGRESO, FONDO, PRECA *', None,
     f'=COUNTIF({RNG("DQ_PRECA_SIN_INGRESO")},1)', None,
     '="Parcial: las tablas solo se reproducen con INGRESO; "&TEXT(E{r},"#,##0")&" pre evaluados no tienen ingreso válido"',
     'Ninguna tabla del archivo usa el fondo compensatorio. Para quienes no tienen ingreso, el fondo no explica el monto asignado. '
     'Falta documentar la regla que combina ingreso y fondo.'),
    (2, 'Se excluyeron lista negra y créditos castigados con la Cooperativa', 'No disponible', None, None, None,
     '="NO verificable con la base"', 'Adjuntar evidencia del cruce.'),
    (3, 'Existen 61 registros sin deuda en la SSF', 'SALDO_SSF / INSTITUCION', 61,
     f'=COUNTIF({RNG("RANGO_DEUDA")},"Sin registro SSF")', '=E{r}-D{r}',
     '="NO coincide: "&TEXT(E{r},"#,##0")&" sin registro SSF y "&TEXT(COUNTIF({rd},"Sin saldo (cero o a favor)"),"#,##0")&" con saldo ≤ 0"',
     'Ninguna definición razonable da 61. Aclarar cómo se obtuvo la cifra.'),
    (4, '54 asociados no al día en aportaciones (condición: ponerse al día para TC)', 'APORTACIONES_AL_DIA, PRECA TC', 54,
     f'=COUNTIF({ORNG("APD")},"NO")', '=E{r}-D{r}',
     '="Coincide solo para TC: "&TEXT(COUNTIFS({apd},"NO",{tc},">0"),"0")&" con PRECA TC; total no al día = "&TEXT(E{r},"#,##0")',
     'Los no al día que solo tienen hipotecario/vehículo no están mencionados en el memorando.'),
    (5, 'TC > $4,000 sin solvencia en Hacienda → límite $4,000', 'PRECA TC', None,
     f'=COUNTIF({ORNG("TC")},">"&{P["TC_TOPE_SIN_SOLV"]})', None,
     '="Aplica a "&TEXT(E{r},"#,##0")&" asociados; exposición TC con tope = $"&TEXT(SUM({tcl}),"#,##0")',
     'La base no tiene solvencia tributaria: no se sabe a cuántos se aplicará el tope.'),
    (6, 'Hipotecario > $30,000: verificar solvencia en Hacienda', 'PRECA HIPO', None,
     f'=COUNTIF({ORNG("HIPO")},">"&{P["HIPO_SOLV_MIN"]})', None,
     '="Aplica a "&TEXT(E{r},"#,##0")&" de "&TEXT(COUNT({hipo}),"#,##0")&" pre evaluados (monto mínimo $110,000)"',
     'Todos los hipotecarios requieren la verificación; no está en la base.'),
    (7, 'Vigencia de 3 meses desde la aprobación del Comité', '—', None, None, None,
     '="No verificable con la base (fecha de aprobación pendiente)"',
     'Durante la vigencia cambian edad, aportaciones y Central de Riesgo: se recomienda refrescar antes de desembolsar.'),
]
for n_, cond, cols_, memo_v, base_f, diff_f, verdict, obs in otros:
    put(wcm, r, 1, n_)
    put(wcm, r, 2, cond, fb)
    put(wcm, r, 3, cols_)
    put(wcm, r, 4, memo_v if memo_v is not None else '—', finput if memo_v is not None else fn, '#,##0')
    put(wcm, r, 5, base_f if base_f else '—', fmt='#,##0')
    put(wcm, r, 6, diff_f.replace('{r}', str(r)) if diff_f else '—', fmt='#,##0')
    put(wcm, r, 7, '')
    v = verdict.replace('{r}', str(r)).replace('{rd}', RNG('RANGO_DEUDA')).replace('{apd}', ORNG('APD')) \
        .replace('{tc}', ORNG('TC')).replace('{tcl}', RNG('TC_LIMITE_SIN_SOLVENCIA')).replace('{hipo}', ORNG('HIPO'))
    put(wcm, r, 8, v)
    put(wcm, r, 9, obs)
    r += 1
r += 1

wcm.cell(r, 1, 'C. Consolidado por producto: memorando vs base').font = fsub
r += 1
hdr(wcm, r, ['', 'Producto / monto', 'Columna', 'Cant. memo', 'Cant. base', 'Dif. cant.', 'Monto memo', 'Monto base / diferencia', 'Resultado'],
    fill=fill_calc)
r += 1
cons = [('TC 500', 'TC', 500, 16, 8000), ('TC 1,500', 'TC', 1500, 242, 363000), ('TC 2,000', 'TC', 2000, 290, 580000),
        ('TC 2,500', 'TC', 2500, 183, 457500), ('TC 3,000', 'TC', 3000, 131, 393000), ('TC 3,500', 'TC', 3500, 89, 311500),
        ('TC 4,500', 'TC', 4500, 65, 292500), ('TC 5,000', 'TC', 5000, 91, 455000), ('TC 7,000', 'TC', 7000, 98, 686000),
        ('TC 10,000', 'TC', 10000, 44, 440000), ('TC TOTAL', 'TC', None, 1249, 3986500),
        ('Vehículo 20,000', 'VEH', 20000, 517, 10340000), ('Vehículo 25,000', 'VEH', 25000, 443, 11075000),
        ('Vehículo 30,000', 'VEH', 30000, 243, 7290000), ('Vehículo 40,000', 'VEH', 40000, 137, 5480000),
        ('Vehículo 150,000 (no está en el memorando)', 'VEH', 150000, 0, 0), ('Vehículo TOTAL', 'VEH', None, 1340, 34185000),
        ('Hipotecario 110,000', 'HIPO', 110000, 424, 46640000), ('Hipotecario 150,000', 'HIPO', 150000, 163, 24450000),
        ('Hipotecario 205,000', 'HIPO', 205000, 92, 18860000), ('Hipotecario TOTAL', 'HIPO', None, 679, 89950000)]
for lab, k, amt, qm, mm in cons:
    tot = amt is None
    put(wcm, r, 2, lab, fb)
    put(wcm, r, 3, orig_hdr[openpyxl.utils.column_index_from_string(C[k]) - 1])
    put(wcm, r, 4, qm, finput, '#,##0')
    put(wcm, r, 5, f'=COUNT({ORNG(k)})' if tot else f'=COUNTIF({ORNG(k)},{amt})', fb if tot else fn, '#,##0')
    put(wcm, r, 6, f'=E{r}-D{r}', fmt='#,##0')
    put(wcm, r, 7, mm, finput, USD)
    put(wcm, r, 8, f'=SUM({ORNG(k)})' if tot else f'=SUMIF({ORNG(k)},{amt})', fb if tot else fn, USD)
    put(wcm, r, 9, f'=IF(AND(F{r}=0,H{r}=G{r}),"Concilia","No concilia: diferencia de "&TEXT(F{r},"#,##0")&" registro(s) y $"&TEXT(H{r}-G{r},"#,##0"))')
    r += 1
put(wcm, r, 2, 'Registros de la base depurada', fb)
put(wcm, r, 3, 'ASOCIADO')
put(wcm, r, 4, 2263, finput, '#,##0')
put(wcm, r, 5, f'={TOTAL}', fb, '#,##0')
put(wcm, r, 6, f'=E{r}-D{r}', fmt='#,##0')
put(wcm, r, 9, f'=IF(F{r}=0,"Concilia","No concilia: el archivo tiene "&TEXT(E{r},"#,##0")&" asociados (2,263 filas incluyendo el encabezado)")')
r += 2

wcm.cell(r, 1, 'D. Opinión de Riesgo solicitada en el memorando → dónde se responde').font = fsub
r += 1
hdr(wcm, r, ['#', 'Solicitud', 'Respuesta en este libro', '', '', '', '', '', ''], fill=fill_calc)
r += 1
for n_, a, b in [(1, 'Validar la base actual', 'Secciones A-C de esta hoja; Calidad_Conciliacion.'),
                 (2, 'Incorporar nuevos parámetros (deudas de la Central de Riesgo)', 'Rangos D1-D4 de SALDO_SSF y relación deuda/ingreso R1-R4 (Matriz_Segmentacion).'),
                 (3, 'Validar el campo de fondo compensatorio', 'Estadisticas, rangos F1-F4, fondo/ingreso M1-M4 y limitaciones sobre su disponibilidad.'),
                 (4, 'Rango de ingreso', 'Rangos I-IV (Matriz_Segmentacion, Parametros).'),
                 (5, 'Rango de fondos compensatorios', 'Rangos F1-F4 (Matriz_Segmentacion, Parametros).')]:
    put(wcm, r, 1, n_)
    put(wcm, r, 2, a, fb)
    put(wcm, r, 3, b)
    wcm.merge_cells(start_row=r, start_column=3, end_row=r, end_column=9)
    r += 1
r += 1

wcm.cell(r, 1, 'E. Conclusión por asociado (criterios 1-7 verificables a la vez)').font = fsub
r += 1
hdr(wcm, r, ['', 'Resultado', 'Universo', 'Asociados', '% del total'], fill=fill_calc)
r += 1
for lab, var, cat in [('Cumple todos los criterios verificables', 'ELEGIBILIDAD_GENERAL', 'Cumple criterios verificables'),
                      ('Información incompleta (sin incumplimientos)', 'ELEGIBILIDAD_GENERAL', 'Información incompleta'),
                      ('Incumple al menos un criterio', 'ELEGIBILIDAD_GENERAL', 'No cumple criterio verificable'),
                      ('TC: cumple todos (sin tarjeta)', 'ELEGIBILIDAD_TC', 'Cumple criterios verificables'),
                      ('TC: excluido por tener tarjeta', 'ELEGIBILIDAD_TC', 'Excluido: ya tiene tarjeta')]:
    put(wcm, r, 2, lab, fb)
    put(wcm, r, 3, 'TC' if var == 'ELEGIBILIDAD_TC' else 'Hipotecario / vehículo / general')
    put(wcm, r, 4, f'=COUNTIF({RNG(var)},"{cat}")', fmt='#,##0')
    put(wcm, r, 5, f'=D{r}/{TOTAL}', fmt=PCT)
    r += 1
put(wcm, r, 2, 'Veredicto global', fb, fill=fill_note)
put(wcm, r, 3, '', fill=fill_note)
put(wcm, r, 4, f'=IF(COUNTIF({RNG("ELEGIBILIDAD_GENERAL")},"Cumple criterios verificables")={TOTAL},'
               f'"La base aplica en su totalidad (sujeto a consultas externas)","La base NO aplica en su totalidad con el memorando: '
               f'solo "&TEXT(COUNTIF({RNG("ELEGIBILIDAD_GENERAL")},"Cumple criterios verificables"),"#,##0")&" de "&TEXT({TOTAL},"#,##0")&'
               f'" asociados cumplen todos los criterios verificables. Lista negra, castigos, relacionados y solvencia no son verificables con la base.")',
    fb, fill=fill_note)
wcm.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9)
wcm.row_dimensions[r].height = 30

# ---------------------------------------------------------------- LEEME
wl = wb.create_sheet('LEEME')
wl.column_dimensions['A'].width = 28
wl.column_dimensions['B'].width = 110
wl['A1'] = 'SEGMENTACIÓN DE RIESGOS — BASE PRE EVALUADA (HIPOTECARIO, VEHÍCULO, TARJETA)'
wl['A1'].font = ftitle
wl['A2'] = 'Documento de trabajo. Parámetros PROPUESTOS, pendientes de validación por el área de Riesgos. No constituye decisión crediticia.'
wl['A2'].font = Font(name=F, size=10, bold=True, color='C00000')
guide = [
    ('Cumplimiento_Memorando', '¿La base aplica con las condiciones del memorando? Veredicto por condición y global.'),
    ('Variables_Rangos', 'Rangos y niveles (Alto/Medio/Bajo) de cada variable, en el formato de la hoja Variables del archivo.'),
    ('Parametros', 'Límites editables (celdas amarillas). Al cambiarlos, todo el libro se recalcula.'),
    ('Matriz_Segmentacion', 'Matriz de rangos y niveles: variable, columna, límites, unidad, nivel, justificación, cantidad, %, faltantes, limitaciones.'),
    ('BD_CLASIFICADA', 'Base con las 31 columnas originales (A:AE) + columnas calculadas con fórmulas (encabezado azul claro; ver comentarios).'),
    ('Consolidado', 'Conteo y % de asociados por variable y categoría, con control de que cada asociado aparezca una vez.'),
    ('Estadisticas', 'Mínimo, máximo, percentiles, promedio, vacíos, ceros y negativos (fórmulas vivas).'),
    ('Evaluacion_Conjunta', 'Cruces entre variables y resumen de alertas.'),
    ('Analisis_Producto', 'Exposición pre evaluada, consistencia con tablas y limitaciones por producto.'),
    ('Calidad_Conciliacion', 'Conciliación de registros y sumas contra BD FINAL; duplicados; inconsistencias.'),
    ('Metodologia', 'Fórmulas, supuestos, restricciones y procedimiento de actualización.'),
    ('Pendientes_Aprobacion', 'Decisiones que Riesgos debe tomar antes de usar los parámetros.'),
    ('BD FINAL, Variables, RANGOS ...', 'Hojas originales del archivo, sin modificaciones.'),
    ('Convención de colores', 'Azul en fondo amarillo = dato editable. Negro = fórmula. Encabezado azul oscuro = columna original; azul claro = calculada.'),
]
r = 4
for a, b in guide:
    put(wl, r, 1, a, fb)
    put(wl, r, 2, b)
    r += 1

order = ['LEEME', 'Variables_Rangos', 'Cumplimiento_Memorando', 'Parametros', 'Matriz_Segmentacion', 'Consolidado', 'Estadisticas', 'Evaluacion_Conjunta',
         'Analisis_Producto', 'Calidad_Conciliacion', 'Metodologia', 'Pendientes_Aprobacion', 'BD_CLASIFICADA',
         'BD FINAL', 'Variables', 'RANGOS VEHICULO', 'RANGOS TC']
wb._sheets = [wb[n] for n in order]
wb.active = 0
wb.save(OUT)

# exportar el mapa de columnas para la verificacion
import json
json.dump({'COL': COL, 'params': {k: v for k, v in P.items()}}, open(OUT + '.map.json', 'w'))
print('ok', OUT, 'calc cols', len(calc), 'last col', L(NCOL_ORIG + len(calc)))
