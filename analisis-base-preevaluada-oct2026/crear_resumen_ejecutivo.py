# -*- coding: utf-8 -*-
"""Genera el libro corto para decisión: Resumen_Ejecutivo_Base_PreEvaluada.xlsx.

Uso: python3 -I crear_resumen_ejecutivo.py <libro_tecnico_recalculado.xlsx> <salida.xlsx>
Toma los resultados ya calculados (y verificados) de BD_CLASIFICADA del libro técnico y asigna a cada
asociado un nivel de seguridad acumulativo (1 a 5). El libro técnico se conserva como respaldo.
"""
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]
CUMPLE = 'Cumple criterios verificables'
x = pd.read_excel(SRC, sheet_name='BD_CLASIFICADA', dtype={'ASOCIADO': str, 'DUI': str})
num = lambda c: pd.to_numeric(x[c], errors='coerce')
for c in ('PRECA HIPO', 'PRECA VEHICULO', 'PRECA TC'):
    x[c] = num(c)
TOTAL = len(x)

# ---------------------------------------------------------------- niveles de seguridad (acumulativos)
RISK = ['NIVEL_ANTIGUEDAD', 'NIVEL_INGRESO', 'NIVEL_DEUDA', 'NIVEL_DEUDA_INGRESO', 'NIVEL_MORA_APORTACIONES',
        'NIVEL_CAT_INTERNA', 'NIVEL_CAT_EXTERNA']
ALL_LOW = ['NIVEL_ANTIGUEDAD', 'NIVEL_INGRESO', 'NIVEL_DEUDA', 'NIVEL_DEUDA_INGRESO', 'NIVEL_MORA_APORTACIONES']
n1 = x.ELEGIBILIDAD_GENERAL == CUMPLE
n2 = n1 & (x.TOTAL_ALERTAS_RIESGO == 0)
n3 = n2 & ~(x[RISK] == 'Alto').any(axis=1)
n4 = n3 & (x.TOTAL_ALERTAS_DATOS == 0) & (x.NIVEL_CAT_INTERNA == 'Bajo') & (x.NIVEL_CAT_EXTERNA == 'Bajo')
n5 = n4 & x[ALL_LOW].isin(['Bajo', 'Sin deuda']).all(axis=1)
LEVELS = [
    (1, 'Cumple el memorando', 'Cumple los criterios 1 a 7 del memorando que se pueden verificar con la base.', n1),
    (2, 'Sin alertas de riesgo', 'Además: sin deuda alta frente al ingreso, sin fondo bajo con deuda alta, sin ingreso bajo con '
                                 'deuda alta, sin aportaciones vencidas y sin categoría A2.', n2),
    (3, 'Ningún nivel Alto', 'Además: ninguna variable de riesgo en nivel Alto (antigüedad, ingreso, deuda, deuda/ingreso, '
                             'aportaciones vencidas, categorías).', n3),
    (4, 'Máxima seguridad (RECOMENDADO)', 'Además: datos sin errores, monto igual al de la tabla oficial y categoría A1 '
                                          'interna y externa.', n4),
    (5, 'Máxima seguridad, todo en Bajo', 'Grupo dentro del nivel 4: además, todas las variables de riesgo en Bajo. '
                                           'Muy estricto: deja casi sin tarjetas.', n5),
]
x['NIVEL_SEGURIDAD'] = 0
for k, *_rest, m in LEVELS:
    x.loc[m, 'NIVEL_SEGURIDAD'] = k

AL_TXT = {'AL_ING_BAJO_DEUDA_ELEVADA': 'ingreso bajo con deuda alta', 'AL_DEUDA_ALTA_VS_INGRESO': 'deuda alta frente al ingreso',
          'AL_ANTIG_BAJA_SIN_HISTORIAL': 'poca antigüedad sin historial', 'AL_FONDO_BAJO_DEUDA_ELEVADA': 'fondo bajo con deuda alta',
          'AL_CATEGORIA_A2': 'categoría A2', 'AL_APORT_VENCIDAS': 'aportaciones vencidas'}
DQ_TXT = {'DQ_PRECA_FUERA_TABLA': 'monto distinto al de la tabla oficial', 'DQ_DATO_INCONSISTENTE': 'dato inconsistente en la base',
          'DQ_PRECA_SIN_INGRESO': 'monto asignado sin ingreso'}
RISK_TXT = {'NIVEL_ANTIGUEDAD': 'antigüedad', 'NIVEL_INGRESO': 'ingreso', 'NIVEL_DEUDA': 'deuda',
            'NIVEL_DEUDA_INGRESO': 'deuda/ingreso', 'NIVEL_MORA_APORTACIONES': 'aportaciones vencidas'}


def motivo(r):
    if r.NIVEL_SEGURIDAD >= 4:
        return ''
    out = []
    al = [t for k, t in AL_TXT.items() if r[k] == 1]
    if al:
        out.append('Alertas: ' + ', '.join(al))
    alto = [t for k, t in RISK_TXT.items() if r[k] == 'Alto']
    if alto:
        out.append('Nivel Alto en: ' + ', '.join(alto))
    dq = [t for k, t in DQ_TXT.items() if r[k] == 1]
    if dq:
        out.append('Revisar: ' + ', '.join(dq))
    if r.NIVEL_CAT_INTERNA != 'Bajo' or r.NIVEL_CAT_EXTERNA != 'Bajo':
        out.append('Categoría distinta de A1')
    return '. '.join(out)


x['MOTIVO_NO_MAXIMA'] = x.apply(motivo, axis=1)


def pendientes(r):
    p = []
    if pd.notna(r['PRECA HIPO']):
        p.append('Hipotecario: solvencia en Hacienda')
    if pd.notna(r['PRECA TC']) and r['PRECA TC'] > 4000:
        p.append('TC: solvencia en Hacienda o bajar a $4,000')
    if pd.notna(r['PRECA TC']) and r['APORTACIONES_AL_DIA'] == 'NO':
        p.append('TC: ponerse al día en aportaciones')
    return '; '.join(p)


x['PENDIENTE_ANTES_DE_OTORGAR'] = x.apply(pendientes, axis=1)

# ---------------------------------------------------------------- estilos
F = 'Arial'
thin = Side(style='thin', color='BFBFBF')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
NAVY, BLUE, GREEN, ORANGE, GRAY, RED = '1F3864', '2E75B6', '548235', 'C55A11', '7F7F7F', 'C00000'
USD, INT = '"$"#,##0.00', '"$"#,##0'
LEVEL_FILLS = {'Alto': 'F8CBAD', 'Medio': 'FFE699', 'Bajo': 'C6E0B4'}
wb = openpyxl.Workbook()
wb.remove(wb.active)


def font(size=10, bold=False, color='000000', italic=False):
    return Font(name=F, size=size, bold=bold, color=color, italic=italic)


def band(ws, r, text, color, ncols=6):
    for j in range(2, 2 + ncols):
        ws.cell(r, j).fill = PatternFill('solid', fgColor=color)
    c = ws.cell(r, 2, text)
    c.font = font(11, True, 'FFFFFF')
    return r + 1


def cellw(ws, r, c, v, fmt=None, bold=False, fill=None, align=None, color='000000', size=10, wrap=True):
    cc = ws.cell(r, c, v)
    cc.font = font(size, bold, color)
    cc.border = box
    cc.alignment = Alignment(wrap_text=wrap, vertical='center', horizontal=align)
    if fmt:
        cc.number_format = fmt
    if fill:
        cc.fill = PatternFill('solid', fgColor=fill)
    return cc


def link(ws, cell, sheet, text=None):
    c = ws[cell]
    c.value = text or sheet
    c.hyperlink = f"#'{sheet}'!A1"
    c.font = font(10, True, '0563C1')
    c.font = Font(name=F, size=10, bold=True, color='0563C1', underline='single')


cnt = lambda m: int(m.sum())
pct = lambda m: m.sum() / TOTAL
prod = {'Hipotecario': 'PRECA HIPO', 'Vehículo': 'PRECA VEHICULO', 'Tarjeta de crédito': 'PRECA TC'}
tc_ok = x.ELEGIBILIDAD_TC == CUMPLE

# ================================================================ HOJA 1: RESUMEN
ws = wb.create_sheet('1_RESUMEN')
ws.sheet_properties.tabColor = NAVY
ws.sheet_view.showGridLines = False
for col, w in zip('ABCDEFG', [2, 46, 16, 16, 16, 16, 40]):
    ws.column_dimensions[col].width = w
ws['B1'] = 'RESUMEN: ¿QUÉ ASOCIADOS DE LA BASE PRE EVALUADA SE PUEDEN OFRECER CON SEGURIDAD?'
ws['B1'].font = font(15, True, NAVY)
ws['B2'] = ('Memorando del 27/10/2026 (Créditos, TC y Mercadeo → Riesgos). Base con corte al 30/09/2026. '
            'Este archivo es un apoyo para decidir: no es una aprobación de crédito.')
ws['B2'].font = font(9, italic=True, color=RED)
r = 4

r = band(ws, r, '1. RESPUESTA CORTA', NAVY)
short = [
    f'• La base NO cumple completa con el memorando: de {TOTAL:,} asociados, solo {cnt(n1):,} ({pct(n1):.0%}) cumplen todos los '
    f'criterios que se pueden verificar.',
    f'• De esos {cnt(n1):,}, los más seguros son {cnt(n4):,} asociados (nivel 4 – máxima seguridad): sin alertas, sin niveles '
    f'altos de riesgo, sin errores de datos y con categoría A1.',
    f'• Recomendación: ofrecer primero a los {cnt(n4):,} del nivel 4 (hoja 3), después de hacer las consultas externas pendientes '
    f'(punto 6).',
]
for t in short:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = ws.cell(r, 2, t)
    c.font = font(11, True)
    c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[r].height = 32
    r += 1
r += 1

r = band(ws, r, '2. QUÉ SE REVISÓ', BLUE)
for t in [f'• Base recibida: {TOTAL:,} asociados pre evaluados para hipotecario, vehículo y tarjeta (el memorando dice 2,263; '
          f'el archivo tiene 2,262 personas).',
          '• Se revisaron los 7 criterios del memorando que se pueden comprobar: ingreso, fondo compensatorio, sin tarjeta (solo TC), '
          'antigüedad, categoría interna A1/A2, categoría SSF A1/A2 y edad de 30 a 55 años.',
          '• Se clasificó a cada asociado en niveles Bajo / Medio / Alto de riesgo (ingreso, deuda, deuda frente al ingreso, '
          'antigüedad, aportaciones vencidas).',
          '• No se pueden comprobar con la base: lista negra, créditos castigados, personas relacionadas o empleados y solvencia '
          'en Hacienda.']:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = ws.cell(r, 2, t)
    c.font = font(10)
    c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[r].height = 28
    r += 1
r += 1

r = band(ws, r, '3. RESULTADO CONTRA EL MEMORANDO', ORANGE)
for j, h in enumerate(['Resultado', 'Asociados', '% de la base', '', '', 'Qué significa'], 2):
    if h:
        cellw(ws, r, j, h, bold=True, fill='FBE5D6')
r += 1
inc = x.ELEGIBILIDAD_GENERAL == 'Información incompleta'
noc = x.ELEGIBILIDAD_GENERAL == 'No cumple criterio verificable'
for lab, m, fill, mean in [('Cumplen todos los criterios', n1, 'E2EFDA', 'Pasan a la evaluación de seguridad (punto 4).'),
                           ('Información incompleta', inc, 'FFF2CC', 'Les falta un dato (sobre todo el ingreso): no se pueden evaluar.'),
                           ('No cumplen algún criterio', noc, 'FCE4D6', 'Ingreso $0, fondo $0 o 56 años al corte.')]:
    cellw(ws, r, 2, lab, bold=True, fill=fill)
    cellw(ws, r, 3, cnt(m), '#,##0', align='right')
    cellw(ws, r, 4, pct(m), '0.0%', align='right')
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    cellw(ws, r, 5, mean)
    r += 1
cellw(ws, r, 2, 'Total', bold=True, fill='F2F2F2')
cellw(ws, r, 3, TOTAL, '#,##0', bold=True, align='right', fill='F2F2F2')
cellw(ws, r, 4, 1, '0.0%', bold=True, align='right', fill='F2F2F2')
r += 1
ing_vacio = int(x.INGRESO.isna().sum())
ing_cero = int((num('INGRESO') == 0).sum())
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(r, 2, f'Causa principal: {ing_vacio:,} asociados no tienen ingreso registrado y {ing_cero:,} tienen ingreso $0, '
                  f'aunque todos recibieron un monto pre evaluado.')
c.font = font(9, italic=True)
r += 2

r = band(ws, r, '4. NIVELES DE SEGURIDAD (cada nivel es más estricto que el anterior)', GREEN)
for j, h in enumerate(['Nivel', 'Asociados', 'Hipotecario', 'Vehículo', 'Tarjeta', 'Condición que se agrega'], 2):
    cellw(ws, r, j, h, bold=True, fill='E2EFDA', align='center')
r += 1
for k, name, cond, m in LEVELS:
    rec = k == 4
    fill = 'C6E0B4' if rec else None
    cellw(ws, r, 2, f'{k}. {name}', bold=True, fill=fill)
    cellw(ws, r, 3, cnt(m), '#,##0', bold=rec, align='right', fill=fill)
    for j, (pl, col) in enumerate(prod.items(), 4):
        mm = m & x[col].notna() & (tc_ok if col == 'PRECA TC' else True)
        cellw(ws, r, j, cnt(mm), '#,##0', bold=rec, align='right', fill=fill)
    cellw(ws, r, 7, cond, fill=fill, size=9)
    ws.row_dimensions[r].height = 40
    r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(r, 2, 'Un asociado puede tener más de un producto, por eso los productos suman más que el total.')
c.font = font(9, italic=True)
r += 2

r = band(ws, r, f'5. RECOMENDADO: NIVEL 4 – MÁXIMA SEGURIDAD ({cnt(n4)} asociados)', GREEN)
for j, h in enumerate(['Producto', 'Asociados', 'Monto total', 'Monto promedio', '', 'Dónde ver la lista'], 2):
    if h:
        cellw(ws, r, j, h, bold=True, fill='E2EFDA', align='center')
r += 1
sheet_of = {'Hipotecario': '4_Hipotecario', 'Vehículo': '5_Vehiculo', 'Tarjeta de crédito': '6_Tarjeta'}
for pl, col in prod.items():
    mm = n4 & x[col].notna() & (tc_ok if col == 'PRECA TC' else True)
    cellw(ws, r, 2, pl, bold=True)
    cellw(ws, r, 3, cnt(mm), '#,##0', align='right')
    cellw(ws, r, 4, float(x.loc[mm, col].sum()), INT, align='right')
    cellw(ws, r, 5, float(x.loc[mm, col].mean()) if mm.any() else 0, INT, align='right')
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    cellw(ws, r, 6, '')
    link(ws, f'F{r}', sheet_of[pl], f'Hoja {sheet_of[pl]}')
    r += 1
cellw(ws, r, 2, 'Todos (sin repetir)', bold=True, fill='F2F2F2')
cellw(ws, r, 3, cnt(n4), '#,##0', bold=True, align='right', fill='F2F2F2')
ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
cellw(ws, r, 6, '')
link(ws, f'F{r}', '3_Maxima_Seguridad', 'Hoja 3_Maxima_Seguridad')
r += 2

r = band(ws, r, '6. ANTES DE OTORGAR (obligatorio, no está en la base)', RED)
hip4 = cnt(n4 & x['PRECA HIPO'].notna())
tc4_big = cnt(n4 & tc_ok & (x['PRECA TC'] > 4000))
for t in ['• Consultar lista negra y créditos castigados con la Cooperativa.',
          '• Verificar que no sea persona relacionada ni empleado (memorando, criterio 8).',
          f'• Solvencia en Hacienda: obligatoria para los {hip4} hipotecarios del nivel 4 (monto > $30,000).',
          f'• Tarjeta con límite mayor de $4,000 ({tc4_big} en el nivel 4): si no está solvente en Hacienda, se otorga con $4,000.',
          '• La base vence 3 meses después de la aprobación del Comité: actualizar edad, aportaciones y Central de Riesgo.']:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = ws.cell(r, 2, t)
    c.font = font(10)
    c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[r].height = 20
    r += 1
r += 1

r = band(ws, r, '7. PROBLEMAS ENCONTRADOS EN LA BASE (para corregir con Créditos)', GRAY)
veh = num('PRECA VEHICULO')
age56 = cnt(x.EDAD_RECALCULADA > 55)
for t in [f'• {ing_vacio + ing_cero:,} asociados tienen monto pre evaluado sin ingreso válido: no se sabe con qué ingreso se calculó.',
          f'• {age56} asociados ya tienen 56 años al 30/09/2026 (el límite es 55).',
          f'• {cnt(num("FONDO_COMPENSATORIO") == 0)} asociados tienen fondo compensatorio $0 (el memorando exige fondo).',
          f'• {cnt(x.TIENE_TARJETA == "SI")} asociados ya tienen tarjeta: el memorando dice que nadie la tiene (solo aplica a TC).',
          f'• Vehículo: {cnt(veh == 150000)} asociado con $150,000, monto que no existe en la tabla (máximo $40,000).',
          '• El memorando dice 61 sin deuda en la SSF y 54 no al día en aportaciones; la base muestra 342 sin registro SSF y 186 no al día.',
          '• Crédito de consumo: no viene en la base ni en el memorando; se necesita su tabla de montos para evaluarlo.']:
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
    c = ws.cell(r, 2, t)
    c.font = font(10)
    c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[r].height = 20
    r += 1
r += 1

r = band(ws, r, '8. QUÉ HAY EN ESTE ARCHIVO', NAVY)
for sh, desc in [('2_Niveles_y_Rangos', 'Rangos Bajo / Medio / Alto de cada variable y cuántos asociados hay en cada nivel.'),
                 ('3_Maxima_Seguridad', f'Los {cnt(n4)} asociados recomendados (nivel 4), todos los productos.'),
                 ('4_Hipotecario', 'Nivel 4 con monto hipotecario.'),
                 ('5_Vehiculo', 'Nivel 4 con monto de vehículo.'),
                 ('6_Tarjeta', 'Nivel 4 sin tarjeta actual y con límite de tarjeta.'),
                 ('7_Los_528', f'Los {cnt(n1)} que cumplen el memorando: columna ¿RECOMENDADO? (SÍ {cnt(n4)} / NO {cnt(n1) - cnt(n4)}) '
                               f'y el motivo de cada NO.')]:
    cellw(ws, r, 2, '')
    link(ws, f'B{r}', sh)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    cellw(ws, r, 3, desc)
    r += 1
ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = ws.cell(r, 2, 'Respaldo técnico (fórmulas, metodología, conciliación): archivo Segmentacion_Riesgos_Base_PreEvaluada.xlsx.')
c.font = font(9, italic=True)

# ================================================================ HOJA 2: NIVELES Y RANGOS
wr = wb.create_sheet('2_Niveles_y_Rangos')
wr.sheet_properties.tabColor = BLUE
wr.sheet_view.showGridLines = False
for col, w in zip('ABCDEFG', [2, 34, 22, 22, 22, 16, 44]):
    wr.column_dimensions[col].width = w
wr['B1'] = 'RANGOS BAJO / MEDIO / ALTO POR VARIABLE'
wr['B1'].font = font(15, True, NAVY)
link(wr, 'B2', '1_RESUMEN', '◄ Volver al RESUMEN')
wb_src = openpyxl.load_workbook(SRC, data_only=True)
p = {row[0].value: row[2].value for row in wb_src['Parametros'].iter_rows(min_row=5, max_row=60) if row[0].value}
money = lambda v: f'${v:,.0f}'
money99 = lambda v: f'${v - 0.01:,.2f}'
cnt_lvl = lambda col, lv: cnt(x[col] == lv)
VARS = [
    ('Antigüedad como asociado', 'NIVEL_ANTIGUEDAD', f'{p["ANT_MEDIO_MAX"] + 1} años o más', f'{p["ANT_ALTO_MAX"] + 1} a {p["ANT_MEDIO_MAX"]} años',
     f'0 a {p["ANT_ALTO_MAX"]} año', ['Sin dato', 'Inconsistente - revisar'], 'Más antigüedad = menos riesgo.'),
    ('Ingreso mensual', 'NIVEL_INGRESO', f'{money(p["ING_L2"])} o más', f'{money(p["ING_L1"])} a {money99(p["ING_L2"])}',
     f'Menos de {money(p["ING_L1"])}', ['Sin información'], 'Más ingreso = menos riesgo.'),
    ('Deuda en la Central de Riesgo (SSF)', 'NIVEL_DEUDA', f'Menos de {money(p["DEU_L1"])}', f'{money(p["DEU_L1"])} a {money99(p["DEU_L2"])}',
     f'{money(p["DEU_L2"])} o más', ['Sin deuda', 'Sin información'], 'Deuda alta no significa que no pague.'),
    ('Deuda ÷ ingreso mensual', 'NIVEL_DEUDA_INGRESO', f'Menos de {p["DTI_L1"]:g} veces', f'{p["DTI_L1"]:g} a menos de {p["DTI_L2"]:g} veces',
     f'{p["DTI_L2"]:g} veces o más', ['Sin deuda', 'Sin información'], 'Cuántos meses de ingreso equivale la deuda.'),
    ('Cuotas de aportación vencidas', 'NIVEL_MORA_APORTACIONES', '0 (al día)', f'{p["MORA_L1"]} a {p["MORA_L2"] - 1}',
     f'{p["MORA_L2"]} o más', [], 'Atraso en aportaciones con la Cooperativa.'),
    ('Categoría interna', 'NIVEL_CAT_INTERNA', 'A1', 'A2', 'B, C, D o E', ['Sin información'], 'Sin información = sin historial interno.'),
    ('Categoría SSF (externa)', 'NIVEL_CAT_EXTERNA', 'A1', 'A2', 'B, C, D o E', ['Sin información'], ''),
]
r = 4
for j, h in enumerate(['Variable (nivel de RIESGO)', 'BAJO', 'MEDIO', 'ALTO', 'Sin dato / sin deuda', 'Nota'], 2):
    cellw(wr, r, j, h, bold=True, fill=NAVY, color='FFFFFF', align='center')
r += 1
for name, col, lo, me, hi, other, note in VARS:
    cellw(wr, r, 2, name, bold=True)
    for j, (rng, lv) in enumerate([(lo, 'Bajo'), (me, 'Medio'), (hi, 'Alto')], 3):
        cellw(wr, r, j, f'{rng}\n{cnt_lvl(col, lv):,} asociados', fill=LEVEL_FILLS[lv], align='center')
    cellw(wr, r, 6, f'{sum(cnt_lvl(col, o) for o in other):,}' if other else '0', align='center')
    cellw(wr, r, 7, note, size=9)
    wr.row_dimensions[r].height = 34
    r += 1
r += 1
for j, h in enumerate(['Variable (nivel de RESPALDO, no es riesgo)', 'BAJO', 'MEDIO', 'ALTO', 'Sin fondo', 'Nota'], 2):
    cellw(wr, r, j, h, bold=True, fill=GRAY, color='FFFFFF', align='center')
r += 1
for name, col, lo, me, hi, other, note in [
        ('Fondo compensatorio', 'NIVEL_RESPALDO_FONDO', f'Menos de {money(p["FON_L1"])}', f'{money(p["FON_L1"])} a {money99(p["FON_L2"])}',
         f'{money(p["FON_L2"])} o más', ['Sin fondo'], 'No se sabe si el fondo está disponible para pagar.'),
        ('Aportaciones acumuladas', 'NIVEL_APORTACIONES', f'Menos de {money(p["APO_L1"])}', f'{money(p["APO_L1"])} a {money99(p["APO_L2"])}',
         f'{money(p["APO_L2"])} o más', [], 'Muestra el vínculo con la Cooperativa.')]:
    cellw(wr, r, 2, name, bold=True)
    for j, (rng, lv) in enumerate([(lo, 'Bajo'), (me, 'Medio'), (hi, 'Alto')], 3):
        cellw(wr, r, j, f'{rng}\n{cnt_lvl(col, lv):,} asociados', fill='DDEBF7', align='center')
    cellw(wr, r, 6, f'{sum(cnt_lvl(col, o) for o in other):,}' if other else '0', align='center')
    cellw(wr, r, 7, note, size=9)
    wr.row_dimensions[r].height = 34
    r += 1
r += 1
wr.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
c = wr.cell(r, 2, 'Rangos propuestos, pendientes de aprobación de Riesgos. Se calcularon con la distribución de esta base '
                  '(cortes cercanos a tercios). No son límites de la normativa.')
c.font = font(9, italic=True, color=RED)
c.alignment = Alignment(wrap_text=True)
wr.row_dimensions[r].height = 28

# ================================================================ HOJAS DE LISTAS
BASE_COLS = [('N° ASOCIADO', 'ASOCIADO', '@', 11), ('NOMBRE', 'NOMBRE', None, 36), ('DUI', 'DUI', '@', 11),
             ('AGENCIA', 'AGENCIA', None, 16), ('INGRESO MENSUAL', 'INGRESO', USD, 13),
             ('FONDO COMPENSATORIO', 'FONDO_COMPENSATORIO', USD, 14), ('DEUDA SSF', 'SALDO_SSF', USD, 14)]
LVL_COLS = [('NIVEL ANTIGÜEDAD', 'NIVEL_ANTIGUEDAD', None, 11), ('NIVEL INGRESO', 'NIVEL_INGRESO', None, 11),
            ('NIVEL DEUDA', 'NIVEL_DEUDA', None, 11), ('NIVEL DEUDA / INGRESO', 'NIVEL_DEUDA_INGRESO', None, 11)]
PEND = [('PENDIENTE ANTES DE OTORGAR', 'PENDIENTE_ANTES_DE_OTORGAR', None, 44)]


def list_sheet(name, color, title, note, df, cols):
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = color
    ws.sheet_view.showGridLines = False
    ws['A1'] = f'{title} — {len(df)} asociados'
    ws['A1'].font = font(14, True, NAVY)
    ws['A2'] = note
    ws['A2'].font = font(9, italic=True, color=RED)
    link(ws, 'A3', '1_RESUMEN', '◄ Volver al RESUMEN')
    for j, (h, src, fmt, w) in enumerate(cols, 1):
        c = ws.cell(5, j, h)
        c.font = font(9, True, 'FFFFFF')
        c.fill = PatternFill('solid', fgColor=color)
        c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        c.border = box
        ws.column_dimensions[L(j)].width = w
    ws.row_dimensions[5].height = 32
    df = df.sort_values('ASOCIADO', key=lambda s: s.astype(int))
    for i, (_, row) in enumerate(df.iterrows()):
        rr = 6 + i
        for j, (h, src, fmt, w) in enumerate(cols, 1):
            v = row[src]
            if isinstance(v, float) and pd.isna(v):
                v = None
            c = ws.cell(rr, j, v)
            c.font = font(9, bold=(j == 1))
            c.border = box
            if fmt:
                c.number_format = fmt
            if src.startswith('NIVEL_') and src != 'NIVEL_SEGURIDAD' or src == 'RECOMENDADO':
                c.alignment = Alignment(horizontal='center', vertical='center')
            elif src in ('MOTIVO_NO_MAXIMA', 'PENDIENTE_ANTES_DE_OTORGAR'):
                c.alignment = Alignment(wrap_text=True, vertical='center')
            if i % 2:
                c.fill = PatternFill('solid', fgColor='F7F9FC')
    end = 5 + len(df)
    for j, (h, src, fmt, w) in enumerate(cols, 1):
        if src.startswith('NIVEL_') and src != 'NIVEL_SEGURIDAD':
            for lv, fc in LEVEL_FILLS.items():
                ws.conditional_formatting.add(f'{L(j)}6:{L(j)}{end}', CellIsRule(
                    operator='equal', formula=[f'"{lv}"'], fill=PatternFill('solid', fgColor=fc, bgColor=fc)))
    for j, (h, src, fmt, w) in enumerate(cols, 1):
        if src == 'RECOMENDADO':
            for txt, fc in (('SÍ', 'C6E0B4'), ('NO', 'F8CBAD')):
                ws.conditional_formatting.add(f'{L(j)}6:{L(j)}{end}', CellIsRule(
                    operator='equal', formula=[f'"{txt}"'], fill=PatternFill('solid', fgColor=fc, bgColor=fc),
                    font=Font(name=F, bold=True)))
    ws.freeze_panes = 'C6'
    ws.auto_filter.ref = f'A5:{L(len(cols))}{end}'
    # totales de montos
    rr = end + 2
    monto = [j for j, (h, src, fmt, w) in enumerate(cols, 1) if src in ('PRECA HIPO', 'PRECA VEHICULO', 'PRECA TC')]
    if monto:
        ws.cell(rr, monto[0] - 1, 'Total:').font = font(10, True)
    for j in monto:
        c = ws.cell(rr, j, f'=SUM({L(j)}6:{L(j)}{end})')
        c.number_format = INT
        c.font = font(10, True)
    return ws


M_H = ('HIPOTECARIO', 'PRECA HIPO', INT, 13)
M_V = ('VEHÍCULO', 'PRECA VEHICULO', INT, 12)
M_T = ('TARJETA', 'PRECA TC', INT, 11)
EXT = 'Falta: lista negra, créditos castigados, relacionados/empleados y solvencia en Hacienda. No es aprobación de crédito.'
list_sheet('3_Maxima_Seguridad', GREEN, 'NIVEL 4 – MÁXIMA SEGURIDAD (RECOMENDADO)',
           'Cumplen el memorando, sin alertas, sin niveles Alto, sin errores de datos y con A1 interna y externa. ' + EXT,
           x[n4], BASE_COLS + [M_H, M_V, M_T] + LVL_COLS + PEND)
list_sheet('4_Hipotecario', GREEN, 'HIPOTECARIO – NIVEL 4', 'Todos requieren solvencia en Hacienda. ' + EXT,
           x[n4 & x['PRECA HIPO'].notna()], BASE_COLS + [M_H] + LVL_COLS + PEND)
list_sheet('5_Vehiculo', GREEN, 'VEHÍCULO – NIVEL 4', EXT, x[n4 & x['PRECA VEHICULO'].notna()], BASE_COLS + [M_V] + LVL_COLS + PEND)
list_sheet('6_Tarjeta', GREEN, 'TARJETA DE CRÉDITO – NIVEL 4',
           'No tienen tarjeta actualmente. Límite > $4,000 sin solvencia en Hacienda: se otorga con $4,000. ' + EXT,
           x[n4 & tc_ok & x['PRECA TC'].notna()],
           BASE_COLS + [M_T, ('SEGMENTO', 'SEGMENTO', None, 20)] + LVL_COLS + PEND)
x['NIVEL_SEGURIDAD_TXT'] = x.NIVEL_SEGURIDAD.map({k: f'{k}. {nm}' for k, nm, *_ in LEVELS})
x['RECOMENDADO'] = x.NIVEL_SEGURIDAD.ge(4).map({True: 'SÍ', False: 'NO'})
ws7 = list_sheet('7_Los_528', ORANGE, 'LOS QUE CUMPLEN EL MEMORANDO, CON SU NIVEL DE SEGURIDAD',
                 '¿RECOMENDADO? SÍ = nivel 4 o 5 (máxima seguridad). NO = cumple el memorando pero tiene alguna señal de riesgo; '
                 'la última columna dice cuál. ' + EXT,
                 x[n1], BASE_COLS[:4] + [('¿RECOMENDADO?', 'RECOMENDADO', None, 13), ('NIVEL DE SEGURIDAD', 'NIVEL_SEGURIDAD_TXT', None, 30)] + [M_H, M_V, M_T]
                 + LVL_COLS + [('POR QUÉ NO ES RECOMENDADO', 'MOTIVO_NO_MAXIMA', None, 70)])

wb.active = 0
wb.save(OUT)
print({f'nivel {k}': cnt(m) for k, _, _, m in LEVELS})
print('hojas:', wb.sheetnames)
