# -*- coding: utf-8 -*-
"""Agrega la hoja Cumplen_Memorando (asociados con ELEGIBILIDAD_GENERAL = "Cumple criterios verificables").

Uso: python3 -I agregar_lista_cumplen.py <libro_recalculado.xlsx>
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
ids = x.loc[x.ELEGIBILIDAD_GENERAL == CUMPLE, 'ASOCIADO'].astype(str)
ids = sorted(ids, key=int)

wb = openpyxl.load_workbook(OUT)
bd = wb['BD_CLASIFICADA']
hdr = {bd.cell(1, j).value: L(j) for j in range(1, bd.max_column + 1)}
last = bd.max_row
if 'Cumplen_Memorando' in wb.sheetnames:
    del wb['Cumplen_Memorando']
ws = wb.create_sheet('Cumplen_Memorando', 2)

F = 'Arial'
thin = Side(style='thin', color='BFBFBF')
box = Border(left=thin, right=thin, top=thin, bottom=thin)
cols = [  # (encabezado, columna de BD_CLASIFICADA, formato, ancho)
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
ws['A1'] = f'ASOCIADOS QUE CUMPLEN TODOS LOS CRITERIOS VERIFICABLES DEL MEMORANDO ({len(ids)})'
ws['A1'].font = Font(name=F, size=14, bold=True, color='1F3864')
ws['A2'] = ('Cumplen los criterios 1 a 7 verificables con la base. Aún requieren la consulta externa de lista negra, '
            'créditos castigados, personas relacionadas o empleados y solvencia en Hacienda. No es una aprobación de crédito.')
ws['A2'].font = Font(name=F, size=9, italic=True, color='C00000')
ws['A3'] = 'Control: total en lista vs total que cumple en BD_CLASIFICADA →'
ws['A3'].font = Font(name=F, size=9, bold=True)
elig = f"BD_CLASIFICADA!${hdr['ELEGIBILIDAD_GENERAL']}$2:${hdr['ELEGIBILIDAD_GENERAL']}${last}"
c = ws.cell(3, 6, f'=IF(AND(COUNTA(A6:A{5+len(ids)})=COUNTIF({elig},"{CUMPLE}"),COUNTIF({L(len(cols)+2)}6:{L(len(cols)+2)}{5+len(ids)},"Sí")=COUNTA(A6:A{5+len(ids)})),'
                     f'"OK: "&COUNTA(A6:A{5+len(ids)})&" asociados, todos siguen cumpliendo","REVISAR: la base cambió; regenerar la lista")')
c.font = Font(name=F, size=9, bold=True, color='006100')

heads = ['ASOCIADO'] + [h for h, *_ in cols] + ['¿SIGUE CUMPLIENDO?']
for j, h in enumerate(heads, 1):
    cc = ws.cell(5, j, h)
    cc.font = Font(name=F, size=9, bold=True, color='FFFFFF')
    cc.fill = PatternFill('solid', fgColor='1F3864')
    cc.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    cc.border = box
ws.row_dimensions[5].height = 42
key = f"BD_CLASIFICADA!$A$2:$A${last}"
for i, a in enumerate(ids):
    r = 6 + i
    cc = ws.cell(r, 1, a)
    cc.font = Font(name=F, size=9, bold=True)
    cc.number_format = '@'
    cc.border = box
    for j, (h, src, fmt, w) in enumerate(cols, 2):
        rng = f"BD_CLASIFICADA!${hdr[src]}$2:${hdr[src]}${last}"
        f = f'=INDEX({rng},MATCH($A{r},{key},0))'
        if src in ('PRECA HIPO', 'PRECA VEHICULO', 'PRECA TC', 'SEGMENTO', 'SALDO_SSF', 'CONDICIONES_MEMO'):
            f = f'=IF(ISNUMBER({f[1:]})+(LEN({f[1:]}&"")>0)*({f[1:]}<>"N/A"),{f[1:]},"")'
        cc = ws.cell(r, j, f)
        cc.font = Font(name=F, size=9)
        cc.border = box
        if fmt:
            cc.number_format = fmt
    cc = ws.cell(r, len(cols) + 2, f'=IF(INDEX({elig},MATCH($A{r},{key},0))="{CUMPLE}","Sí","No")')
    cc.font = Font(name=F, size=9)
    cc.border = box
ws.column_dimensions['A'].width = 10
for j, (h, src, fmt, w) in enumerate(cols, 2):
    ws.column_dimensions[L(j)].width = w
ws.column_dimensions[L(len(cols) + 2)].width = 11
ws.freeze_panes = 'C6'
ws.auto_filter.ref = f'A5:{L(len(cols)+2)}{5+len(ids)}'

r = 6 + len(ids) + 1
ws.cell(r, 1, 'Resumen por producto (asociados de esta lista con monto pre evaluado)').font = Font(name=F, size=11, bold=True, color='1F3864')
for k, (lab, col) in enumerate([('Hipotecario', 'K'), ('Vehículo', 'L'), ('Tarjeta de crédito', 'M')]):
    rr = r + 1 + k
    ws.cell(rr, 1, lab).font = Font(name=F, size=9, bold=True)
    ws.cell(rr, 2, f'=COUNT({col}6:{col}{5+len(ids)})').number_format = '#,##0'
    ws.cell(rr, 3, f'=SUM({col}6:{col}{5+len(ids)})').number_format = '"$"#,##0'
    ws.cell(rr, 4, 'asociados / monto total pre evaluado').font = Font(name=F, size=9, italic=True)

leeme = wb['LEEME']
for row in leeme.iter_rows(min_row=4):
    if row[0].value == 'Cumplimiento_Memorando':
        break
nr = leeme.max_row + 1
leeme.cell(nr, 1, 'Cumplen_Memorando').font = Font(name=F, size=10, bold=True)
leeme.cell(nr, 2, f'Lista de los {len(ids)} asociados que cumplen todos los criterios verificables (datos por fórmula desde BD_CLASIFICADA).').font = Font(name=F, size=10)
wb.save(OUT)
print('lista', len(ids))
