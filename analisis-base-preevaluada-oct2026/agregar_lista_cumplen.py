# -*- coding: utf-8 -*-
"""Agrega las hojas de asociados que cumplen los criterios verificables del memorando.

Uso: python3 -I agregar_lista_cumplen.py <libro_recalculado.xlsx>
Crea Cumplen_Memorando (todos), Cumplen_Hipotecario, Cumplen_Vehiculo, Cumplen_TC y Consumo.
La lista de ASOCIADO se toma del libro ya recalculado; el resto de columnas son formulas INDEX/MATCH
contra BD_CLASIFICADA, y la columna de control verifica que cada asociado siga cumpliendo.
Despues de ejecutarlo hay que recalcular el libro otra vez.
"""
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L

OUT = sys.argv[1]
CUMPLE = 'Cumple criterios verificables'
x = pd.read_excel(OUT, sheet_name='BD_CLASIFICADA', dtype={'ASOCIADO': str})
num = lambda c: pd.to_numeric(x[c], errors='coerce')
ok = x.ELEGIBILIDAD_GENERAL == CUMPLE
ok_tc = x.ELEGIBILIDAD_TC == CUMPLE
sel = lambda m: sorted(x.loc[m, 'ASOCIADO'].astype(str), key=int)

wb = openpyxl.load_workbook(OUT)
bd = wb['BD_CLASIFICADA']
hdr = {bd.cell(1, j).value: L(j) for j in range(1, bd.max_column + 1)}
last = bd.max_row
KEY = f"BD_CLASIFICADA!$A$2:$A${last}"
RNG = lambda c: f"BD_CLASIFICADA!${hdr[c]}$2:${hdr[c]}${last}"

F = 'Arial'
thin = Side(style='thin', color='BFBFBF')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
COLS = [  # (encabezado, columna de BD_CLASIFICADA, formato, ancho)
    ('NOMBRE', 'NOMBRE', None, 38), ('DUI', 'DUI', '@', 12), ('EDAD AL 30/09/2026', 'EDAD_RECALCULADA', '0', 9),
    ('AGENCIA', 'AGENCIA', None, 16), ('INGRESO', 'INGRESO', '"$"#,##0.00', 13),
    ('FONDO COMPENSATORIO', 'FONDO_COMPENSATORIO', '"$"#,##0.00', 14), ('SALDO SSF', 'SALDO_SSF', '"$"#,##0.00', 14),
    ('CAT. INTERNA', 'CAT_INTERNA', None, 9), ('CAT. SSF', 'CAT_EXTERNA_SSF', None, 9),
    ('PRECA HIPO', 'PRECA HIPO', '#,##0', 11), ('PRECA VEHÍCULO', 'PRECA VEHICULO', '#,##0', 11),
    ('PRECA TC', 'PRECA TC', '#,##0', 10), ('SEGMENTO TC', 'SEGMENTO', None, 20),
    ('TIENE TARJETA', 'TIENE_TARJETA', None, 8), ('APORTACIONES AL DÍA', 'APORTACIONES_AL_DIA', None, 10),
    ('NIVEL ANTIGÜEDAD', 'NIVEL_ANTIGUEDAD', None, 10), ('NIVEL INGRESO', 'NIVEL_INGRESO', None, 10),
    ('NIVEL DEUDA', 'NIVEL_DEUDA', None, 10), ('NIVEL DEUDA/INGRESO', 'NIVEL_DEUDA_INGRESO', None, 11),
    ('NIVEL RESPALDO FONDO', 'NIVEL_RESPALDO_FONDO', None, 10), ('ALERTAS DE RIESGO', 'TOTAL_ALERTAS_RIESGO', '0', 9),
    ('CONDICIONES DEL MEMORANDO', 'CONDICIONES_MEMO', None, 50), ('ELEGIBILIDAD TC', 'ELEGIBILIDAD_TC', None, 26),
]
# columnas que pueden venir vacias o "N/A" en la base: se muestran en blanco
BLANKABLE = {'PRECA HIPO', 'PRECA VEHICULO', 'PRECA TC', 'SEGMENTO', 'SALDO_SSF', 'CONDICIONES_MEMO'}
PROD_COL = {'HIPO': 'K', 'VEH': 'L', 'TC': 'M'}  # columnas de monto en la hoja generada


def make_sheet(name, pos, ids, title, note, elig_col, count_formula, product=None):
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name, pos)
    n = len(ids)
    end = 5 + n
    ws['A1'] = f'{title} ({n})'
    ws['A1'].font = Font(name=F, size=14, bold=True, color='1F3864')
    ws['A2'] = note
    ws['A2'].font = Font(name=F, size=9, italic=True, color='C00000')
    ws['A3'] = 'Control: total en lista vs total en BD_CLASIFICADA →'
    ws['A3'].font = Font(name=F, size=9, bold=True)
    flag = L(len(COLS) + 2)
    c = ws.cell(3, 6, f'=IF(AND(COUNTA(A6:A{end})={count_formula},COUNTIF({flag}6:{flag}{end},"Sí")=COUNTA(A6:A{end})),'
                         f'"OK: "&COUNTA(A6:A{end})&" asociados, todos siguen cumpliendo","REVISAR: la base cambió; regenerar la lista")')
    c.font = Font(name=F, size=9, bold=True, color='006100')
    heads = ['ASOCIADO'] + [h for h, *_ in COLS] + ['¿SIGUE CUMPLIENDO?']
    for j, h in enumerate(heads, 1):
        cc = ws.cell(5, j, h)
        cc.font = Font(name=F, size=9, bold=True, color='FFFFFF')
        cc.fill = PatternFill('solid', fgColor='1F3864')
        cc.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        cc.border = box
    ws.row_dimensions[5].height = 42
    elig = RNG(elig_col)
    for i, a in enumerate(ids):
        r = 6 + i
        cc = ws.cell(r, 1, a)
        cc.font = Font(name=F, size=9, bold=True)
        cc.number_format = '@'
        cc.border = box
        for j, (h, src, fmt, w) in enumerate(COLS, 2):
            look = f'INDEX({RNG(src)},MATCH($A{r},{KEY},0))'
            f = f'={look}'
            if src in BLANKABLE:
                f = f'=IF(ISNUMBER({look})+(LEN({look}&"")>0)*({look}<>"N/A"),{look},"")'
            cc = ws.cell(r, j, f)
            cc.font = Font(name=F, size=9)
            cc.border = box
            if fmt:
                cc.number_format = fmt
        cond = f'INDEX({elig},MATCH($A{r},{KEY},0))="{CUMPLE}"'
        if product:
            cond = f'AND({cond},ISNUMBER(INDEX({RNG(product)},MATCH($A{r},{KEY},0))))'
        cc = ws.cell(r, len(COLS) + 2, f'=IF({cond},"Sí","No")')
        cc.font = Font(name=F, size=9)
        cc.border = box
    ws.column_dimensions['A'].width = 10
    for j, (h, src, fmt, w) in enumerate(COLS, 2):
        ws.column_dimensions[L(j)].width = w
    ws.column_dimensions[flag].width = 11
    ws.freeze_panes = 'C6'
    ws.auto_filter.ref = f'A5:{flag}{end}'
    r = end + 2
    ws.cell(r, 1, 'Resumen por producto (asociados de esta hoja con monto pre evaluado)').font = Font(name=F, size=11, bold=True, color='1F3864')
    for k, (lab, key) in enumerate([('Hipotecario', 'HIPO'), ('Vehículo', 'VEH'), ('Tarjeta de crédito', 'TC')]):
        rr = r + 1 + k
        col = PROD_COL[key]
        ws.cell(rr, 1, lab).font = Font(name=F, size=9, bold=True)
        ws.cell(rr, 2, f'=COUNT({col}6:{col}{end})').number_format = '#,##0'
        ws.cell(rr, 3, f'=SUM({col}6:{col}{end})').number_format = '"$"#,##0'
        ws.cell(rr, 4, 'asociados / monto total pre evaluado').font = Font(name=F, size=9, italic=True)
    return ws


EXT = ('Aún requieren la consulta externa de lista negra, créditos castigados, personas relacionadas o empleados '
       'y solvencia en Hacienda. No es una aprobación de crédito.')
make_sheet('Cumplen_Memorando', 2, sel(ok), 'ASOCIADOS QUE CUMPLEN TODOS LOS CRITERIOS VERIFICABLES DEL MEMORANDO',
           'Cumplen los criterios 1 a 7 verificables con la base. ' + EXT, 'ELEGIBILIDAD_GENERAL',
           f'COUNTIF({RNG("ELEGIBILIDAD_GENERAL")},"{CUMPLE}")')
make_sheet('Cumplen_Hipotecario', 3, sel(ok & num('PRECA HIPO').notna()), 'CRÉDITO HIPOTECARIO: CUMPLEN Y TIENEN MONTO PRE EVALUADO',
           'Todos requieren verificar solvencia en Hacienda (memorando: hipotecario > $30,000; el monto mínimo es $110,000). ' + EXT,
           'ELEGIBILIDAD_GENERAL', f'COUNTIFS({RNG("ELEGIBILIDAD_GENERAL")},"{CUMPLE}",{RNG("PRECA HIPO")},">0")', 'PRECA HIPO')
make_sheet('Cumplen_Vehiculo', 4, sel(ok & num('PRECA VEHICULO').notna()), 'CRÉDITO DE VEHÍCULO: CUMPLEN Y TIENEN MONTO PRE EVALUADO',
           'Monto según tabla RANGOS VEHICULO (ingreso ≥ $3,000). Revisar los casos cuyo monto no coincide con la tabla '
           '(hoja Analisis_Producto). ' + EXT,
           'ELEGIBILIDAD_GENERAL', f'COUNTIFS({RNG("ELEGIBILIDAD_GENERAL")},"{CUMPLE}",{RNG("PRECA VEHICULO")},">0")', 'PRECA VEHICULO')
make_sheet('Cumplen_TC', 5, sel(ok_tc & num('PRECA TC').notna()), 'TARJETA DE CRÉDITO: CUMPLEN, NO TIENEN TARJETA Y TIENEN LÍMITE PRE EVALUADO',
           'Límite > $4,000 sin solvencia en Hacienda → se otorga con $4,000. Si no está al día en aportaciones debe ponerse al día. ' + EXT,
           'ELEGIBILIDAD_TC', f'COUNTIFS({RNG("ELEGIBILIDAD_TC")},"{CUMPLE}",{RNG("PRECA TC")},">0")', 'PRECA TC')

# ---- Consumo: no existe en la base ni en el memorando
if 'Consumo' in wb.sheetnames:
    del wb['Consumo']
wc = wb.create_sheet('Consumo', 6)
wc.column_dimensions['A'].width = 40
wc.column_dimensions['B'].width = 100
wc['A1'] = 'CRÉDITO DE CONSUMO: NO SE PUEDE EVALUAR CON ESTA BASE'
wc['A1'].font = Font(name=F, size=14, bold=True, color='C00000')
rows = [
    ('Situación', 'El memorando solo pide validar Crédito Hipotecario, Vehículo y Tarjeta de Crédito. La base no tiene columna '
                  'de monto pre evaluado de consumo (solo PRECA HIPO, PRECA VEHICULO y PRECA TC) ni tabla de montos de consumo.'),
    ('Qué NO se hizo', 'No se asignaron montos de consumo: inventar una tabla o un monto sería presentar como dato algo que la '
                       'institución no ha definido.'),
    ('Universo que sí cumple los criterios generales',
     f'=COUNTIF({RNG("ELEGIBILIDAD_GENERAL")},"{CUMPLE}")&" asociados (hoja Cumplen_Memorando) cumplen los criterios verificables y '
     f'podrían ser el universo de partida para consumo, si Riesgos lo aprueba."'),
    ('Para evaluar consumo se necesita', '1) Tabla institucional de montos de consumo por ingreso (y/o fondo). '
                                          '2) Plazo y tasa para calcular la cuota. '
                                          '3) Cuotas actuales de todas las obligaciones para medir la relación cuota-ingreso. '
                                          '4) Límite institucional de cuota/ingreso para consumo.'),
    ('Siguiente paso', 'Con esa tabla se agrega una columna PRECA CONSUMO a la base y se genera la hoja Cumplen_Consumo con el mismo '
                       'formato que las de hipotecario, vehículo y tarjeta.'),
]
for i, (a, b) in enumerate(rows, 3):
    wc.cell(i, 1, a).font = Font(name=F, size=10, bold=True)
    c = wc.cell(i, 2, b)
    c.font = Font(name=F, size=10)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    wc.row_dimensions[i].height = 45

leeme = wb['LEEME']
existing = {leeme.cell(r, 1).value for r in range(1, leeme.max_row + 1)}
for nm, desc in [('Cumplen_Memorando', 'Asociados que cumplen todos los criterios verificables (todos los productos).'),
                 ('Cumplen_Hipotecario', 'Los que cumplen y tienen monto hipotecario pre evaluado.'),
                 ('Cumplen_Vehiculo', 'Los que cumplen y tienen monto de vehículo pre evaluado.'),
                 ('Cumplen_TC', 'Los que cumplen, no tienen tarjeta y tienen límite de TC pre evaluado.'),
                 ('Consumo', 'Por qué no se puede evaluar consumo con esta base y qué se necesita.')]:
    if nm not in existing:
        nr = leeme.max_row + 1
        leeme.cell(nr, 1, nm).font = Font(name=F, size=10, bold=True)
        leeme.cell(nr, 2, desc).font = Font(name=F, size=10)
wb.save(OUT)
print({k: len(v) for k, v in [('todos', sel(ok)), ('hipo', sel(ok & num('PRECA HIPO').notna())),
                              ('veh', sel(ok & num('PRECA VEHICULO').notna())), ('tc', sel(ok_tc & num('PRECA TC').notna()))]})
