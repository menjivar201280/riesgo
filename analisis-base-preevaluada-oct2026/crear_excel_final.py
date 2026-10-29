# -*- coding: utf-8 -*-
"""Genera Recomendacion_Base_PreEvaluada_COMEDICA_FINAL_CORREGIDA.xlsx (presentación a Gerencia).

Uso: python3 crear_excel_final.py <archivo_original.xlsx> <salida.xlsx>

- Conserva la CLASIFICACIÓN ORIGINAL del archivo (492 recomendados + 36 con observaciones = 528) y la muestra tal cual.
- Encima agrega la REVISIÓN de cada asociado: alertas (riesgo alto, ingreso bajo, deuda alta, montos fuera de tabla, datos por
  corregir) sin reclasificar automáticamente; cada alerta dice por qué se revisa.
- Separa por producto: monto registrado en la base · monto que podría evaluarse · capacidad de pago (siempre PENDIENTE: la base
  no trae cuota, tasa ni plazo) · resultado.
- Hojas visibles: 1 Resumen · 2 Listado de los 528 · 3 Criterios · 4 Hipotecarios a validar (todos los de la base).
  Las hojas técnicas de respaldo quedan ocultas (base original, clasificación original, parámetros y cálculo).
Todo con fórmulas: si cambia la base, se recalcula.
"""
import os
import sys

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_FILE, OUT_FILE = sys.argv[1:3]
motor = open(os.path.join(HERE, 'crear_excel_comedica.py'), encoding='utf-8').read()
motor = motor[:motor.index('# ============================================================== listas dinámicas')]
for viejo, nuevo in [("SRC, OUT = sys.argv[1:3]", "SRC, OUT = SRC_FILE, OUT_FILE"),
                     ("BASE, CALC = '6_Base_Original', '7_Calculo'", "BASE, CALC = 'Respaldo_Base_Original', 'Respaldo_Calculo'"),
                     ("'2_Criterios_y_Rangos'", "'Respaldo_Parametros'"),
                     ("'CRITERIOS, RANGOS Y PARÁMETROS'", "'RESPALDO: PARÁMETROS DEL CÁLCULO'"),
                     ('", deuda SSF"', '", deuda bancaria"'), ('", categoría SSF"', '", categoría en la central de riesgo"')]:
    assert viejo in motor, viejo
    motor = motor.replace(viejo, nuevo)

# ---------------- clasificación original (valores del archivo auditado, sin cambios)
src_wb = openpyxl.load_workbook(SRC_FILE, data_only=True)
ORIG = []  # (id, clasificación, nivel anterior, puntaje anterior, motivo anterior)
for sh, clas in (('RECOMENDADOS', 'Recomendado'), ('CON OBSERVACIONES', 'Con observaciones')):
    w = src_wb[sh]
    hdr = [c.value for c in w[5]]
    for row in w.iter_rows(min_row=6, values_only=True):
        if row[0] in (None, '') or row[1] in (None, 'TOTAL'):
            continue
        d = dict(zip(hdr, row))
        ORIG.append((str(d['N° ASOCIADO']), clas, d.get('NIVEL DE RIESGO') or 'Sin nivel (no se calculó)', d.get('PUNTAJE (1 a 3)'),
                     d.get('¿POR QUÉ SÍ?') or d.get('OBSERVACIÓN: QUÉ VERIFICAR')))
N_REC = sum(1 for o in ORIG if o[1] == 'Recomendado')
N_OBS = sum(1 for o in ORIG if o[1] == 'Con observaciones')
assert (N_REC, N_OBS) == (492, 36), (N_REC, N_OBS)

EXTRA = r'''
OR_RNG = "'Respaldo_Clasif_Original'!$A$2:$A$600"
orig = lambda col: f"INDEX('Respaldo_Clasif_Original'!${col}$2:${col}$600,MATCH({{ID}},{OR_RNG},0))"
add('O_CLAS', 'CLASIFICACIÓN ORIGINAL', '=IF({ID}="","",IFERROR(' + orig('B') + ',"No está en los 528"))', None, 16)
add('O_NIV', 'NIVEL ANTERIOR (puntaje)', '=IF(OR({ID}="",{O_CLAS}="No está en los 528"),"",' + orig('C') + ')', None, 11)
add('O_MOT', 'MOTIVO ORIGINAL', '=IF(OR({ID}="",{O_CLAS}="No está en los 528"),"",' + orig('E') + '&"")', None, 30)
# alertas (no reclasifican: documentan por qué se revisa)
add('A_ALTO', 'ALERTA RIESGO ALTO', '=IF({NIVEL}="ALTO",1,0)', '0', 7)
add('A_ING', 'ALERTA INGRESO BAJO', '=IF(AND({ING}<>"",N({ING})>0,N({ING})<P_ING_L1),1,0)', '0', 7)
add('A_DEU', 'ALERTA DEUDA ALTA PARA SU INGRESO', '=IF({CAPAC}="Deuda alta para su ingreso",1,0)', '0', 7)
add('A_VEH', 'ALERTA VEHÍCULO FUERA DE TABLA', '=IF({VEH_OK}="No",1,0)', '0', 7)
add('A_INF', 'ALERTA INFORMACIÓN INCOMPLETA', '=IF(OR({ELEG}="Información insuficiente",{SININF}>0),1,0)', '0', 7)
add('A_DATO', 'ALERTA DATO POR CORREGIR', f'=IF({{ID}}="",0,IF(OR({b("EDAD")}<>{{EDADC}},LEN({b("DUI")}&"")<>9),1,0))', '0', 7)
add('N_AL', 'NÚMERO DE ALERTAS', '={A_ALTO}+{A_ING}+{A_DEU}+{A_VEH}+{A_INF}+{A_DATO}', '0', 7)
add('ALTXT', 'ALERTAS (por qué se revisa)', '=IF({ID}="","",IF({N_AL}=0,"Sin alertas",MID(IF({A_ALTO}=1,"; Riesgo ALTO ("&{ROJTXT}&")","")'
    '&IF({A_ING}=1,"; Ingreso bajo: $"&TEXT({ING},"#,##0")&" al mes (menos de $"&TEXT(P_ING_L1,"#,##0")&")","")'
    '&IF({A_DEU}=1,"; Deuda bancaria de $"&TEXT({SSF},"#,##0")&" alta para su ingreso","")'
    '&IF({A_VEH}=1,"; Vehículo de $"&TEXT({VEH},"#,##0")&" fuera de la tabla de vehículos (le corresponde: "&IF({VEH_T}="NO APLICA","ninguno",TEXT({VEH_T},"$#,##0"))&")","")'
    '&IF({A_INF}=1,"; Información incompleta","")&IF({A_DATO}=1,"; Dato por corregir (edad o DUI)",""),3,400)))', None, 50)
add('REV', 'RESULTADO DE LA REVISIÓN', '=IF(OR({ID}="",{O_CLAS}="No está en los 528"),"",IF({O_CLAS}="Recomendado",'
    'IF({N_AL}=0,"Recomendado – sin alertas","Recomendado – requiere revisión"),"Con observaciones – requiere revisión"))', None, 24)
# por producto: monto registrado · monto que podría evaluarse · capacidad de pago
PEND = '"Pendiente de evaluación"'
add('EV_HIPO', 'HIPOTECARIO QUE PODRÍA EVALUARSE', '=IF({HIPO}="","",IF(AND({O_CLAS}="Recomendado",{A_ALTO}=0,{A_ING}=0,{A_DEU}=0,{A_INF}=0),{HIPO},' + PEND + '))', '$#,##0', 12)
add('EV_VEH', 'VEHÍCULO QUE PODRÍA EVALUARSE', '=IF({VEH}="","",IF(AND({O_CLAS}="Recomendado",{A_ALTO}=0,{A_DEU}=0,{A_VEH}=0,{A_INF}=0),{VEH},' + PEND + '))', '$#,##0', 12)
add('EV_TC', 'TARJETA QUE PODRÍA EVALUARSE', '=IF({TC}="","",IF(AND({O_CLAS}<>"No está en los 528",{A_ALTO}=0,{A_INF}=0,{TTC}="NO",{TC_OK}="Sí"),{TC},' + PEND + '))', '$#,##0', 11)
add('CAP', 'CAPACIDAD DE PAGO', '=IF(AND({HIPO}="",{VEH}="",{TC}=""),"","Pendiente de validar (la base no trae cuota, tasa ni plazo)")', None, 30)
add('NOTA_H', 'NOTA HIPOTECARIO', '=IF({HIPO}="","","Monto registrado en la base; capacidad de pago pendiente de validar")', None, 34)
add('H_SIT', 'SITUACIÓN DEL HIPOTECARIO', '=IF({HIPO}="","",IF({O_CLAS}="No está en los 528","No cumple criterios iniciales o le falta información",'
    'IF({EV_HIPO}=' + PEND + ',"No sustentado con los datos disponibles","Sin alertas; capacidad de pago pendiente")))', None, 30)
add('K_528', 'ORDEN LISTADO', '=IF(OR({ID}="",{O_CLAS}="No está en los 528"),"",IF({REV}="Recomendado – sin alertas",1,IF({REV}="Recomendado – requiere revisión",2,3))*100000+ROW())', '0', 8)
add('K_HIP', 'ORDEN HIPOTECARIOS', '=IF({HIPO}="","",IF({H_SIT}="No sustentado con los datos disponibles",1,IF({H_SIT}="Sin alertas; capacidad de pago pendiente",2,3))*100000+ROW())', '0', 8)
'''
i = motor.index('COL = {k: L(i + 1)')
motor = motor[:i] + EXTRA + motor[i:]
exec(compile(motor, 'motor_comedica', 'exec'))  # noqa: S102 (código propio del proyecto)

from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.chart.series import DataPoint  # noqa: E402
from openpyxl.chart.label import DataLabelList  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402,F811
from openpyxl.styles import Alignment  # noqa: E402,F811
from openpyxl.worksheet.hyperlink import Hyperlink  # noqa: E402
from openpyxl.worksheet.properties import PageSetupProperties  # noqa: E402

LIGHT = 'F2F6FC'
RISK = {'BAJO': GREEN_F, 'MEDIO': YEL_F, 'ALTO': RED_F}  # noqa: F821
REVC = {'Recomendado – sin alertas': 'C6E0B4', 'Recomendado – requiere revisión': 'FFE699', 'Con observaciones – requiere revisión': 'F8CBAD'}


def cf_text(w, rng_, colors):
    for t_, c_ in colors.items():
        w.conditional_formatting.add(rng_, CellIsRule(operator='equal', formula=[f'"{t_}"'], fill=fill(c_)))  # noqa: F821


# ---------------- hoja oculta con la clasificación original
wo = wb.create_sheet('Respaldo_Clasif_Original')  # noqa: F821
head(wo, 1, ['N° ASOCIADO', 'CLASIFICACIÓN ORIGINAL', 'NIVEL DE RIESGO ORIGINAL', 'PUNTAJE ORIGINAL', 'MOTIVO / OBSERVACIÓN ORIGINAL'],  # noqa: F821
     [12, 20, 16, 12, 60], NAVY)  # noqa: F821
for k, row in enumerate(ORIG, 2):
    for j, v in enumerate(row, 1):
        wo.cell(k, j, v).font = font(9)  # noqa: F821
put(wo, 'G1', 'Copia de los valores de las hojas RECOMENDADOS y CON OBSERVACIONES del archivo original (sin cambios).', italic=True, color=GRAY)  # noqa: F821

# ---------------- verificaciones (en Respaldo_Parametros)
wp = wb['Respaldo_Parametros']  # noqa: F821
rv = wp.max_row + 3
put(wp, f'A{rv}', 'VERIFICACIONES AUTOMÁTICAS', size=12, bold=True, color=NAVY)  # noqa: F821
VER = [
    ('Todos los registros de la base están calculados', "COUNTIF(C_ID,\"?*\")=COUNTIF('Respaldo_Base_Original'!$A$2:$A$2601,\"?*\")"),
    ('Clasificación original: 492 recomendados + 36 con observaciones = 528',
     "AND(COUNTIF('Respaldo_Clasif_Original'!$B$2:$B$600,\"Recomendado\")=492,COUNTIF('Respaldo_Clasif_Original'!$B$2:$B$600,\"Con observaciones\")=36)"),
    ('Los 528 de la clasificación original existen en la base', 'COUNTIF(C_O_CLAS,"Recomendado")+COUNTIF(C_O_CLAS,"Con observaciones")=528'),
    ('Los 528 cumplen los criterios iniciales del memorando', f'COUNTIFS(C_O_CLAS,"<>No está en los 528",C_ELEG,"{E1}",C_ID,"?*")=528'),  # noqa: F821
    ('Hoja 2 tiene los 528', 'COUNT(C_K_528)=528'),
    ('Hoja 4 tiene todos los hipotecarios de la base', 'COUNT(C_K_HIP)=COUNT(C_HIPO)'),
    ('Ningún riesgo ALTO como "Recomendado – sin alertas"', 'COUNTIFS(C_REV,"Recomendado – sin alertas",C_NIVEL,"ALTO")=0'),
    ('Ningún hipotecario se declara con capacidad validada', 'COUNTIFS(C_HIPO,">0",C_CAP,"<>Pendiente*")=0'),
    ('Ingreso vacío no se cuenta como cero', "COUNTIF(C_C_ING,\"Sin información\")=SUMPRODUCT(('Respaldo_Base_Original'!$A$2:$A$2601<>\"\")"
     "*NOT(ISNUMBER('Respaldo_Base_Original'!$P$2:$P$2601)))"),
    ('Sin errores en el cálculo', f"SUMPRODUCT(--ISERROR('Respaldo_Calculo'!$A$2:${L(len(COLS))}$2601))=0"),  # noqa: F821
]
v0 = rv + 1
for k, (lab, cond) in enumerate(VER):
    put(wp, f'A{v0 + k}', lab, border=True)  # noqa: F821
    put(wp, f'C{v0 + k}', f'=IF({cond},"OK","REVISAR")', bold=True, border=True, h='center')  # noqa: F821
VERIF = f"'Respaldo_Parametros'!$C${v0}:$C${v0 + len(VER) - 1}"


# ---------------- listas dinámicas
def lista(nombre, pos, titulo, sub, cuenta, key, cols, cf):
    w = wb.create_sheet(nombre, pos)  # noqa: F821
    w.sheet_view.showGridLines = False
    put(w, 'A1', titulo, size=15, bold=True, color=NAVY)  # noqa: F821
    put(w, 'A2', sub, size=9, italic=True, color=GRAY)  # noqa: F821
    w['A3'] = cuenta
    w['A3'].font = font(12, True, NAVY)  # noqa: F821
    head(w, 5, [h for h, *_ in cols] + ['fila'], [wd for *_, wd in cols] + [6], NAVY, 50)  # noqa: F821
    last = len(cols) + 1
    for i_ in range(LCAP):  # noqa: F821
        rr = 6 + i_
        c = w.cell(rr, last, f'=IFERROR(MATCH(SMALL(C_{key},{i_ + 1}),C_{key},0),"")')
        c.font = font(8, color='BFBFBF')  # noqa: F821
        for j, (h, k, fmt, wd) in enumerate(cols, 1):
            c = w.cell(rr, j, f'=IF(${L(last)}{rr}="","",IF(INDEX(C_{k},${L(last)}{rr})="","",INDEX(C_{k},${L(last)}{rr})))')  # noqa: F821
            c.font = font(9)  # noqa: F821
            if fmt:
                c.number_format = fmt
            c.alignment = Alignment(wrap_text=True, vertical='top')
    w.column_dimensions[L(last)].hidden = True  # noqa: F821
    end = 5 + LCAP  # noqa: F821
    w.freeze_panes = 'C6'
    w.auto_filter.ref = f'A5:{L(len(cols))}{end}'  # noqa: F821
    for colr, colors in cf:
        cf_text(w, f'{colr}6:{colr}{end}', colors)
    return w


H2, H4 = '2_Listado_528', '4_Hipotecarios_a_Validar'
C528 = [('Número de asociado', 'ID', '@', 11), ('Nombre', 'NOM', None, 28), ('Ingreso al mes', 'ING', '$#,##0', 10),
        ('Deuda bancaria', 'SSF', '$#,##0', 10), ('Antigüedad (años)', 'ANT', '0', 8),
        ('Clasificación original', 'O_CLAS', None, 14), ('Nivel de riesgo anterior', 'O_NIV', None, 10), ('Nivel de riesgo actual', 'NIVEL', None, 10),
        ('Resultado de la revisión', 'REV', None, 19), ('Alertas: por qué se revisa', 'ALTXT', None, 44),
        ('Hipotecario registrado en la base', 'HIPO', '$#,##0', 12), ('Hipotecario que podría evaluarse', 'EV_HIPO', '$#,##0', 13),
        ('Vehículo registrado en la base', 'VEH', '$#,##0', 11), ('Vehículo que podría evaluarse', 'EV_VEH', '$#,##0', 12),
        ('Tarjeta registrada en la base', 'TC', '$#,##0', 10), ('Tarjeta que podría evaluarse', 'EV_TC', '$#,##0', 11),
        ('Capacidad de pago', 'CAP', None, 22), ('Motivo / observación original', 'O_MOT', None, 30)]
lista(H2, 1, 'LISTADO DE LOS 528 ASOCIADOS QUE CUMPLEN LOS CRITERIOS INICIALES',
      'Clasificación original (492 recomendados + 36 con observaciones) y su revisión. Orden: recomendados sin alertas, recomendados que requieren '
      'revisión, con observaciones. Un monto registrado en la base NO es un monto recomendado ni aprobado.',
      '=TEXT(COUNT(C_K_528),"#,##0")&" asociados  ·  "&TEXT(COUNTIF(C_REV,"Recomendado – sin alertas"),"#,##0")&" recomendados sin alertas  ·  "'
      '&TEXT(COUNTIF(C_REV,"Recomendado – requiere revisión"),"#,##0")&" recomendados que requieren revisión  ·  "'
      '&TEXT(COUNTIF(C_REV,"Con observaciones – requiere revisión"),"#,##0")&" con observaciones"',
      'K_528', C528, [('G', RISK), ('H', RISK), ('I', REVC)])
CHIP = [('Número de asociado', 'ID', '@', 11), ('Nombre', 'NOM', None, 28), ('Ingreso al mes', 'ING', '$#,##0', 10),
        ('Deuda bancaria', 'SSF', '$#,##0', 11), ('Clasificación original', 'O_CLAS', None, 16), ('Nivel de riesgo actual', 'NIVEL', None, 10),
        ('Hipotecario registrado en la base', 'HIPO', '$#,##0', 12), ('Hipotecario que podría evaluarse', 'EV_HIPO', '$#,##0', 13),
        ('Situación', 'H_SIT', None, 24), ('Alertas: por qué no se sustenta', 'ALTXT', None, 44), ('Nota', 'NOTA_H', None, 30)]
lista(H4, 3, 'HIPOTECARIOS REGISTRADOS EN LA BASE: TODOS REQUIEREN VALIDAR LA CAPACIDAD DE PAGO',
      'La base no trae cuota mensual, tasa ni plazo: ningún hipotecario tiene capacidad de pago validada. Primero los que, además, tienen '
      'ingreso bajo, deuda alta, riesgo alto o información incompleta.',
      '=TEXT(COUNT(C_K_HIP),"#,##0")&" hipotecarios en la base  ·  "&TEXT(COUNTIF(C_H_SIT,"No sustentado*"),"#,##0")&" no sustentados con los datos '
      'disponibles  ·  "&TEXT(COUNTIF(C_H_SIT,"Sin alertas*"),"#,##0")&" sin alertas (capacidad pendiente)  ·  "&TEXT(COUNTIF(C_H_SIT,"No cumple*"),"#,##0")'
      '&" fuera de los 528"', 'K_HIP', CHIP, [('F', RISK)])

# ---------------- 1 Resumen
ws = wb.create_sheet('1_Resumen', 0)  # noqa: F821
ws.sheet_view.showGridLines = False
for col, w in zip('ABCDEFGHIJKLM', [2, 16, 16, 2, 16, 16, 2, 14, 12, 12, 16, 16, 2]):
    ws.column_dimensions[col].width = w
put(ws, 'B2', 'COMÉDICA · BASE PREEVALUADA (hipotecario, vehículo y tarjeta)', size=20, bold=True, color=NAVY)  # noqa: F821
ws['B3'] = '="Base al "&TEXT(P_FECHA_CORTE,"dd/mm/yyyy")&" · Resultados de PREEVALUACIÓN: no son aprobación ni montos aprobados."'
ws['B3'].font = font(11, True, 'C00000')  # noqa: F821
CARDS = [('B', 'C', '=COUNTIF(C_ID,"?*")', 'Asociados en la base', NAVY, None),
         ('E', 'F', '=COUNT(C_K_528)', 'Cumplen criterios iniciales', BLUE, H2),
         ('H', 'I', '=COUNTIF(C_O_CLAS,"Recomendado")', 'Recomendados (clasificación original)', '548235', H2),
         ('K', 'L', '=COUNTIF(C_O_CLAS,"Con observaciones")', 'Con observaciones', 'BF8F00', H2)]
for a, b2, f, lab, col, dest in CARDS:
    for rr, v, sz in ((5, f, 32), (6, lab, 11)):
        ws.merge_cells(f'{a}{rr}:{b2}{rr}')
        c = ws[f'{a}{rr}']
        c.value = v
        c.font = font(sz, True, WHITE)  # noqa: F821
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.number_format = '#,##0'
        for cc in (a, b2):
            ws[f'{cc}{rr}'].fill = fill(col)  # noqa: F821
    if dest:
        ws[f'{a}5'].hyperlink = Hyperlink(ref=f'{a}5', location=f"'{dest}'!A1")
ws.row_dimensions[5].height = 50
ws.row_dimensions[6].height = 30
ws.merge_cells('B7:L7')
ws['B7'] = ('="Cómo se relacionan: de "&TEXT(COUNTIF(C_ID,"?*"),"#,##0")&" asociados, "&TEXT(COUNT(C_K_528),"#,##0")&" cumplen los criterios iniciales del '
            'memorando. Esos "&TEXT(COUNT(C_K_528),"#,##0")&" = "&TEXT(COUNTIF(C_O_CLAS,"Recomendado"),"#,##0")&" recomendados + "'
            '&TEXT(COUNTIF(C_O_CLAS,"Con observaciones"),"#,##0")&" con observaciones. Los otros "&TEXT(COUNTIF(C_ID,"?*")-COUNT(C_K_528),"#,##0")'
            '&" no cumplen o les falta información (sobre todo el ingreso)."')
ws['B7'].font = font(10, color=GRAY)  # noqa: F821
ws['B7'].alignment = Alignment(wrap_text=True, vertical='center')
ws.row_dimensions[7].height = 32

put(ws, 'B9', 'REVISIÓN DE LOS 492 RECOMENDADOS', size=12, bold=True, color=NAVY)  # noqa: F821
for k, (lab, f, col) in enumerate([('Recomendados sin alertas', '=COUNTIF(C_REV,"Recomendado – sin alertas")', 'C6E0B4'),
                                   ('Recomendados que requieren revisión', '=COUNTIF(C_REV,"Recomendado – requiere revisión")', 'FFE699'),
                                   ('Con observaciones (requieren revisión)', '=COUNTIF(C_REV,"Con observaciones – requiere revisión")', 'F8CBAD')]):
    rr = 10 + k
    ws.merge_cells(f'B{rr}:E{rr}')
    put(ws, f'B{rr}', lab, bold=True, bg=col, border=True)  # noqa: F821
    put(ws, f'F{rr}', f, '#,##0', bold=True, border=True, h='center', size=12)  # noqa: F821
ws.merge_cells('B13:F13')
put(ws, 'B13', 'No se cambió la clasificación original: se marca quién necesita revisión y por qué (hoja 2).', italic=True, color=GRAY, size=9)  # noqa: F821

put(ws, 'H9', 'NIVEL DE RIESGO DE LOS 528', size=12, bold=True, color=NAVY)  # noqa: F821
for j, h in zip('HIJ', ['Nivel', 'Anterior', 'Actual']):
    put(ws, f'{j}10', h, bold=True, color=WHITE, bg=NAVY, h='center')  # noqa: F821
for k, lv in enumerate(['BAJO', 'MEDIO', 'ALTO']):
    rr = 11 + k
    put(ws, f'H{rr}', lv, bold=True, bg=RISK[lv], border=True, h='center')  # noqa: F821
    put(ws, f'I{rr}', f'=COUNTIF(C_O_NIV,"{lv}")', '#,##0', border=True, h='center')  # noqa: F821
    put(ws, f'J{rr}', f'=COUNTIFS(C_NIVEL,"{lv}",C_K_528,">0")', '#,##0', bold=True, border=True, h='center')  # noqa: F821
put(ws, 'H14', '=IF(COUNTIF(C_O_NIV,"Sin nivel*")>0,"Anterior: "&COUNTIF(C_O_NIV,"Sin nivel*")&" con observaciones no tenían nivel","")', italic=True, color=GRAY, size=8)  # noqa: F821
ws.merge_cells('H15:J16')
put(ws, 'H15', 'Anterior = por puntaje (cortes 2.2 y 2.6). Actual = datos en riesgo alto (0 bajo · 1 medio · 2+ alto).', italic=True, color=GRAY, size=8, wrap=True)  # noqa: F821
ch = BarChart()
ch.type = 'col'
ch.title = 'Nivel de riesgo actual (528)'
ch.add_data(Reference(ws, min_col=10, min_row=11, max_row=13), titles_from_data=False)
ch.set_categories(Reference(ws, min_col=8, min_row=11, max_row=13))
ch.legend = None
ch.y_axis.majorGridlines = None
s_ = ch.series[0]
for idx, colr in enumerate(['70AD47', 'FFC000', 'C00000']):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = colr
    s_.dPt.append(pt)
s_.dLbls = DataLabelList()
s_.dLbls.showVal, s_.dLbls.showSerName, s_.dLbls.showCatName, s_.dLbls.showLegendKey = True, False, False, False
ch.height, ch.width = 5.2, 8.5
ws.add_chart(ch, 'K9')

put(ws, 'B18', 'ALERTAS IMPORTANTES', size=12, bold=True, color=WHITE, bg='C00000')  # noqa: F821
for col in 'CDEFGHIJKL':
    ws[f'{col}18'].fill = fill('C00000')  # noqa: F821
ALERTAS = [
    ('Hipotecarios: ninguno tiene la capacidad de pago validada (la base no trae cuota, tasa ni plazo)',
     '=COUNTIFS(C_HIPO,">0",C_K_528,">0")', 'entre los 528'),
    ('Hipotecarios NO sustentados con los datos disponibles (ingreso bajo, deuda alta, riesgo alto o información incompleta)',
     '=COUNTIFS(C_H_SIT,"No sustentado*")', 'ver hoja 4'),
    ('Recomendados con riesgo ALTO', '=COUNTIFS(C_O_CLAS,"Recomendado",C_A_ALTO,1)', 'ver hoja 2'),
    ('Recomendados con ingreso menor de $1,500 al mes', '=COUNTIFS(C_O_CLAS,"Recomendado",C_A_ING,1)', 'ver hoja 2'),
    ('Recomendados con deuda bancaria alta para su ingreso', '=COUNTIFS(C_O_CLAS,"Recomendado",C_A_DEU,1)', 'ver hoja 2'),
    ('Recomendados con vehículo fuera de la tabla de vehículos', '=COUNTIFS(C_O_CLAS,"Recomendado",C_A_VEH,1)', 'ver hoja 2'),
    ('Recomendados con un dato por corregir (edad o DUI)', '=COUNTIFS(C_O_CLAS,"Recomendado",C_A_DATO,1)', 'ver hoja 2'),
]
for k, (lab, f, donde) in enumerate(ALERTAS):
    rr = 19 + k
    ws.merge_cells(f'B{rr}:I{rr}')
    put(ws, f'B{rr}', lab, border=True, wrap=True)  # noqa: F821
    put(ws, f'J{rr}', f, '#,##0', bold=True, border=True, h='center', size=12, color='C00000')  # noqa: F821
    put(ws, f'K{rr}', donde, italic=True, color=GRAY, size=9)  # noqa: F821
    ws.row_dimensions[rr].height = 28 if len(lab) > 80 else 20
    if k % 2:
        ws[f'B{rr}'].fill = fill(LIGHT)  # noqa: F821
rr = 19 + len(ALERTAS) + 1
ws.merge_cells(f'B{rr}:L{rr}')
ws[f'B{rr}'] = f'=IF(COUNTIF({VERIF},"OK")={len(VER)},"✔ Verificación automática: los 528 concilian con 492 + 36 y con la base; ningún monto se presenta como aprobado.",' \
               '"✖ Verificación automática: hay algo que revisar.")'
ws[f'B{rr}'].font = font(10, True, '548235')  # noqa: F821
ws.merge_cells(f'B{rr + 1}:L{rr + 1}')
put(ws, f'B{rr + 1}', 'Hoja 2: los 528, uno por uno, con su clasificación original y su revisión · Hoja 3: criterios · Hoja 4: todos los hipotecarios a validar.',
    italic=True, color=GRAY, size=9)  # noqa: F821

# ---------------- 3 Criterios
wcr3 = wb.create_sheet('3_Criterios', 2)  # noqa: F821
wcr3.sheet_view.showGridLines = False
for col, w in zip('ABC', [30, 70, 2]):
    wcr3.column_dimensions[col].width = w
put(wcr3, 'A1', 'CRITERIOS (en pocas palabras)', size=15, bold=True, color=NAVY)  # noqa: F821
BLOQUES = [
    ('CRITERIOS INICIALES (memorando 27/10/2026)', [
        ('Cumple criterios iniciales', 'Edad de 30 a 55 años, ingreso y fondo compensatorio mayores de 0, categoría A1 o A2 (interna y central de riesgo).'),
        ('Clasificación original', '492 "Recomendado" y 36 "Con observaciones", tal como venían en el archivo. No se cambió.')]),
    ('ALERTAS: POR QUÉ UN ASOCIADO SE REVISA', [
        ('Riesgo ALTO', '2 o más datos en riesgo alto (antigüedad 0-1 año, ingreso menor de $1,500, deuda de $10,000 o más, 3+ cuotas de aportación vencidas, categoría fuera de A1/A2).'),
        ('Ingreso bajo', 'Gana menos de $1,500 al mes.'),
        ('Deuda alta para su ingreso', 'Su deuda bancaria es muy alta para lo que gana (ejemplo: gana $1,000 y debe más de $12,000), o gana menos de $1,500 y debe $10,000 o más.'),
        ('Vehículo fuera de tabla', 'El monto de vehículo de la base no coincide con la tabla de vehículos de la Cooperativa (menos de $3,000 de ingreso: no aplica).'),
        ('Información incompleta / dato por corregir', 'Falta un dato de riesgo, o la edad o el DUI no coinciden.')]),
    ('MONTOS', [
        ('Monto registrado en la base', 'El que trae la base (PRECA HIPO, PRECA VEHICULO, PRECA TC). Es una referencia, no una recomendación ni una aprobación.'),
        ('Monto que podría evaluarse', 'El monto de la base solo si el asociado no tiene alertas que lo impidan; si no, "Pendiente de evaluación".'),
        ('Capacidad de pago', 'Pendiente de validar en todos los casos: la base no trae la cuota mensual de las deudas, ni la tasa ni el plazo del crédito.')]),
    ('CAMBIOS RESPECTO AL ARCHIVO ORIGINAL (documentados)', [
        ('Clasificación', 'Sin cambios: se conservan los 492 recomendados y los 36 con observaciones. Se agregó la columna "Resultado de la revisión".'),
        ('Nivel de riesgo', 'Se muestran el anterior (puntaje) y el actual (datos en riesgo alto). Cambia en algunos asociados; ambos quedan visibles.'),
        ('Montos', 'El archivo original no mostraba 27 montos que sí están en la base (23 vehículos de "paga su deuda con el hipotecario", 3 vehículos y '
         '1 hipotecario de "solo tarjeta"). Aquí se muestra el monto registrado y, por sus alertas, queda "Pendiente de evaluación".'),
        ('Hipotecarios', 'Ninguno se presenta como recomendado: todos dicen "Monto registrado en la base; capacidad de pago pendiente de validar" (hoja 4).')]),
    ('PENDIENTE DE APROBACIÓN INSTITUCIONAL', [
        ('Rangos', 'Ingreso ($1,500 / $3,000), deuda ($1,000 / $10,000) y cuotas vencidas (1 / 3) son propuestas de Riesgo; la antigüedad viene de la hoja Variables.'),
        ('Nivel de riesgo', 'Regla por datos en riesgo alto (actual) en lugar de cortes de puntaje (anterior).')]),
]
r = 3
for tit, filas in BLOQUES:
    put(wcr3, f'A{r}', tit, size=11, bold=True, color=WHITE, bg=NAVY)  # noqa: F821
    wcr3[f'B{r}'].fill = fill(NAVY)  # noqa: F821
    r += 1
    for a_, b_ in filas:
        put(wcr3, f'A{r}', a_, bold=True, border=True, wrap=True)  # noqa: F821
        put(wcr3, f'B{r}', b_, border=True, wrap=True)  # noqa: F821
        wcr3.row_dimensions[r].height = 32
        r += 1
    r += 1

# ---------------- orden, visibilidad, impresión
order = ['1_Resumen', H2, '3_Criterios', H4, 'Respaldo_Parametros', 'Respaldo_Calculo', 'Respaldo_Clasif_Original', 'Respaldo_Base_Original']
wb._sheets = [wb[n] for n in order]  # noqa: F821
for n, colr in zip(order, [NAVY, BLUE, BLUE, BLUE] + ['A6A6A6'] * 4):  # noqa: F821
    w_ = wb[n]  # noqa: F821
    w_.sheet_properties.tabColor = colr
    w_.page_setup.orientation = 'landscape'
    w_.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    w_.page_setup.fitToWidth = 1
    w_.page_setup.fitToHeight = 0
    if n.startswith('Respaldo_'):
        w_.sheet_state = 'hidden'
wb.active = 0  # noqa: F821
wb.calculation.fullCalcOnLoad = True  # noqa: F821
wb.save(OUT_FILE)  # noqa: F821
print('ok', OUT_FILE, 'originales', len(ORIG))
