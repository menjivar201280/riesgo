# -*- coding: utf-8 -*-
"""Genera Recomendacion_Base_PreEvaluada_COMEDICA_SIMPLE.xlsx: versión sencilla para Gerencia de Negocios.

Uso: python3 crear_excel_simple.py <archivo_auditado.xlsx> <salida.xlsx>

Reutiliza el motor de cálculo ya verificado de crear_excel_comedica.py (base original, parámetros con nombre y
cálculo por registro con fórmulas), que queda al final como hojas de RESPALDO. Encima construye 4 hojas sencillas:
  1_Resumen_General · 2_Listado_Asociados · 3_Criterios · 4_Casos_por_Revisar
Todo se calcula con fórmulas: si se pega una base nueva en Respaldo_Base_Original, el libro se actualiza solo.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_FILE, OUT_FILE = sys.argv[1:3]
motor = open(os.path.join(HERE, 'crear_excel_comedica.py'), encoding='utf-8').read()
motor = motor[:motor.index('# ============================================================== listas dinámicas')]
SUSTITUCIONES = [
    ("SRC, OUT = sys.argv[1:3]", "SRC, OUT = SRC_FILE, OUT_FILE"),
    ("BASE, CALC = '6_Base_Original', '7_Calculo'", "BASE, CALC = 'Respaldo_Base_Original', 'Respaldo_Calculo'"),
    ("'2_Criterios_y_Rangos'", "'Respaldo_Parametros'"),
    ("E2, E3 = '2. Recomendado preliminarmente', '3. Requiere revisión adicional'", "E2, E3 = 'Puede continuar', 'Requiere revisión'"),
    ("E4, E5 = '4. Pendiente por información insuficiente', '5. No elegible (restricción del memorando)'",
     "E4, E5 = 'Información incompleta', 'No cumple criterios iniciales'"),
    ("'CRITERIOS, RANGOS Y PARÁMETROS'", "'RESPALDO: PARÁMETROS DEL CÁLCULO (no se necesita para leer los resultados)'"),
    ('"6. Sujeto a evaluación crediticia formal"', '"Pasa a evaluación crediticia formal"'),
    # textos sin siglas técnicas para la Gerencia
    ('", deuda SSF"', '", deuda bancaria"'), ('", categoría SSF"', '", categoría en la central de riesgo"'),
    ('"Deuda SSF de $"', '"Deuda bancaria de $"'), ('"Deuda SSF alta para su ingreso, pero', '"Deuda bancaria alta para su ingreso, pero'),
    ('"; el hipotecario debe pagar primero su deuda SSF"', '"; el hipotecario debe pagar primero su deuda bancaria"'),
    ('"Sin categoría SSF; "', '"Sin categoría en la central de riesgo; "'),
    ('"Categoría SSF no es A1/A2; "', '"Categoría en la central de riesgo no es A1/A2; "'),
]
for viejo, nuevo in SUSTITUCIONES:
    assert viejo in motor, viejo
    motor = motor.replace(viejo, nuevo)
# columnas extra del motor para la versión sencilla (antes de numerar las columnas)
EXTRA = '''
add('ALERTA', 'INFORMACIÓN QUE FALTA O ALERTA', '=IF({ESTADO}="' + E3 + '",TRIM(IF({NIVEL}="ALTO","Riesgo ALTO por: "&{ROJTXT}&". ","")'
    '&IF({CAPAC}<>"Adecuada","Deuda bancaria alta para su ingreso"&IF({HIPO}="",", sin hipotecario para pagarla",IF({CABE}=0,", mayor que el '
    'hipotecario ofrecido",""))&".","")),IF({ESTADO}="' + E4 + '",IF({M_ELEG}<>"",{M_ELEG},"Algún dato de riesgo sin información"),""))', None, 40)
add('MOTC', 'MOTIVO CORTO', '=IF({ESTADO}="' + E3 + '",IF({NIVEL}="ALTO","Riesgo alto","")&IF(AND({NIVEL}="ALTO",{CAPAC}<>"Adecuada")," y ","")'
    '&IF({CAPAC}<>"Adecuada",IF({NIVEL}="ALTO","deuda alta para su ingreso","Deuda alta para su ingreso"),""),IF({ESTADO}="' + E4 + '","Falta información para evaluar",""))', None, 24)
add('FALTA', 'QUÉ FALTA REVISAR', '=IF({ESTADO}="' + E2 + '",{PEND},IF({ESTADO}="' + E3 + '",{VALID},IF({ESTADO}="' + E4 + '",{VALID},"")))', None, 46)
add('ACC2', 'ACCIÓN RECOMENDADA', '=IF(OR({ESTADO}="' + E3 + '",{ESTADO}="' + E4 + '"),IF({ACCION}="","Pendiente de validación",{ACCION}),"")', None, 34)
'''
i = motor.index('COL = {k: L(i + 1)')
motor = motor[:i] + EXTRA + motor[i:]
exec(compile(motor, 'motor_comedica', 'exec'))  # define wb, COL, P, head, put, title, etc.

from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.chart.series import DataPoint  # noqa: E402
from openpyxl.chart.label import DataLabelList  # noqa: E402
from openpyxl.formatting.rule import CellIsRule  # noqa: E402,F811
from openpyxl.styles import Alignment, PatternFill  # noqa: E402,F811
from openpyxl.worksheet.properties import PageSetupProperties  # noqa: E402

LIGHT = 'F2F6FC'
RISK = {'BAJO': GREEN_F, 'MEDIO': YEL_F, 'ALTO': RED_F}  # noqa: F821 (definidos en el motor)
RES = {E2: 'C6E0B4', E3: 'FFE699', E4: 'DDEBF7', E5: 'EDEDED'}  # noqa: F821


def cf_text(ws, rng, colors):
    for txt, col in colors.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=[f'"{txt}"'], fill=fill(col)))  # noqa: F821


# ============================================================== verificaciones (en Respaldo_Parametros)
wp = wb['Respaldo_Parametros']  # noqa: F821
rv = wp.max_row + 3
put(wp, f'A{rv}', 'VERIFICACIONES AUTOMÁTICAS DEL ARCHIVO', size=12, bold=True, color=NAVY)  # noqa: F821
VER = [
    ('Todos los registros de la base están calculados', 'COUNTIF(C_ID,"?*")=COUNTIF(\'Respaldo_Base_Original\'!$A$2:$A$2601,"?*")'),
    ('Cada asociado tiene un solo resultado', 'COUNTIF(C_ESTADO,"?*")=COUNTIF(C_ID,"?*")'),
    ('Ningún "Puede continuar" tiene riesgo ALTO', f'COUNTIFS(C_ESTADO,"{E2}",C_NIVEL,"ALTO")=0'),  # noqa: F821
    ('Ingreso vacío no se cuenta como cero', 'COUNTIF(C_C_ING,"Sin información")=SUMPRODUCT((\'Respaldo_Base_Original\'!$A$2:$A$2601<>"")'
     '*NOT(ISNUMBER(\'Respaldo_Base_Original\'!$P$2:$P$2601)))'),
    ('Pesos del puntaje suman 100%', 'ROUND(P_W_ANT+P_W_ING+P_W_DEU+P_W_MORA+P_W_CATE,6)=1'),
    ('Rangos sin solapamiento', 'AND(P_ING_L1<P_ING_L2,P_DEU_L1<P_DEU_L2,P_DTI_L1<P_DTI_L2,P_MORA_L1<P_MORA_L2,P_ANT_ALTO_MAX<P_ANT_BAJO_MIN)'),
    ('Sin errores en el cálculo', f"SUMPRODUCT(--ISERROR('Respaldo_Calculo'!$A$2:${L(len(COLS))}$2601))=0"),  # noqa: F821
    ('Casos por revisar = Requiere revisión + Información incompleta', f'COUNT(C_K_CAS)=COUNTIF(C_ESTADO,"{E3}")+COUNTIF(C_ESTADO,"{E4}")'),
]
v0 = rv + 1
for k, (lab, cond) in enumerate(VER):
    put(wp, f'A{v0 + k}', lab, border=True)  # noqa: F821
    put(wp, f'C{v0 + k}', f'=IF({cond},"OK","REVISAR")', bold=True, border=True, h='center')  # noqa: F821
VERIF = f"'Respaldo_Parametros'!$C${v0}:$C${v0 + len(VER) - 1}"
cf_text(wp, f'C{v0}:C{v0 + len(VER) - 1}', {'OK': GREEN_F, 'REVISAR': RED_F})  # noqa: F821

# ============================================================== 1_Resumen_General
ws = wb.create_sheet('1_Resumen_General', 0)  # noqa: F821
ws.sheet_view.showGridLines = False
for col, w in zip('ABCDEFGHIJKL', [2, 15, 15, 2, 15, 15, 2, 15, 15, 2, 15, 15]):
    ws.column_dimensions[col].width = w
ws.column_dimensions['M'].width = 2
ws.column_dimensions['N'].width = 15
ws.column_dimensions['O'].width = 15
put(ws, 'B2', 'COMÉDICA · PREEVALUACIÓN CREDITICIA', size=20, bold=True, color=NAVY)  # noqa: F821
ws['B3'] = '="Base al "&TEXT(P_FECHA_CORTE,"dd/mm/yyyy")&" · Hipotecario, vehículo y tarjeta de crédito · Resultado PRELIMINAR: no es aprobación de crédito"'
ws['B3'].font = font(11, True, 'C00000')  # noqa: F821
CARDS = [('B', 'C', '=COUNTIF(C_ID,"?*")', 'Asociados evaluados', 'Todos los de la base', NAVY),
         ('E', 'F', f'=COUNTIF(C_ELEG,"{E1}")', 'Cumplen criterios iniciales', 'Requisitos del memorando', BLUE),  # noqa: F821
         ('H', 'I', f'=COUNTIF(C_ESTADO,"{E2}")', 'Pueden continuar', 'Pasan a la siguiente etapa', '548235'),
         ('K', 'L', f'=COUNTIF(C_ESTADO,"{E3}")', 'Requieren revisión', 'Riesgo alto o deuda alta', 'BF8F00'),
         ('N', 'O', f'=COUNTIF(C_ESTADO,"{E4}")', 'Información incompleta', 'Faltan datos para evaluar', '7F7F7F')]
for a, b2, f, lab, desc, col in CARDS:
    for rr, (v, sz, bold) in zip((5, 6, 7), ((f, 30, True), (lab, 12, True), (desc, 9, False))):
        ws.merge_cells(f'{a}{rr}:{b2}{rr}')
        c = ws[f'{a}{rr}']
        c.value = v
        c.font = font(sz, bold, WHITE)  # noqa: F821
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.number_format = '#,##0'
        for cc in (a, b2):
            ws[f'{cc}{rr}'].fill = fill(col)  # noqa: F821
ws.row_dimensions[5].height = 46
ws.row_dimensions[6].height = 20
ws.row_dimensions[7].height = 18
ws['B8'] = f'="Además, "&TEXT(COUNTIF(C_ESTADO,"{E5}"),"#,##0")&" asociados no cumplen los criterios iniciales del memorando (por ejemplo, ingreso o fondo igual a 0)."'
ws['B8'].font = font(9, italic=True, color=GRAY)  # noqa: F821

# nivel de riesgo (tabla + gráfico)
put(ws, 'B10', 'NIVEL DE RIESGO DE LOS QUE CUMPLEN CRITERIOS INICIALES', size=12, bold=True, color=NAVY)  # noqa: F821
for j, h in zip('BCE', ['Nivel', 'Asociados', 'Qué significa']):
    put(ws, f'{j}11', h, bold=True, color=WHITE, bg=NAVY, h='center')  # noqa: F821
ws.merge_cells('E11:I11')
for k, (lv, desc) in enumerate([('BAJO', 'Ningún dato en riesgo alto'), ('MEDIO', '1 dato en riesgo alto'),
                                ('ALTO', '2 o más datos en riesgo alto: va a revisión')]):
    rr = 12 + k
    put(ws, f'B{rr}', f'Riesgo {lv}', bold=True, bg=RISK[lv], border=True)  # noqa: F821
    put(ws, f'C{rr}', f'=COUNTIFS(C_ELEG,"{E1}",C_NIVEL,"{lv}")', '#,##0', bold=True, border=True, h='center')  # noqa: F821
    ws.merge_cells(f'E{rr}:I{rr}')
    put(ws, f'E{rr}', desc, border=True)  # noqa: F821
ch = BarChart()
ch.type = 'col'
ch.title = 'Asociados por nivel de riesgo'
ch.add_data(Reference(ws, min_col=3, min_row=12, max_row=14), titles_from_data=False)
ch.set_categories(Reference(ws, min_col=2, min_row=12, max_row=14))
ch.legend = None
ch.y_axis.majorGridlines = None
s = ch.series[0]
for idx, colr in enumerate(['70AD47', 'FFC000', 'C00000']):
    pt = DataPoint(idx=idx)
    pt.graphicalProperties.solidFill = colr
    s.dPt.append(pt)
s.dLbls = DataLabelList()
s.dLbls.showVal = True
s.dLbls.showSerName = False
s.dLbls.showCatName = False
s.dLbls.showLegendKey = False
ch.height, ch.width = 6.2, 11
ws.add_chart(ch, 'K10')

# productos
put(ws, 'B17', 'PRODUCTOS DE LOS QUE PUEDEN CONTINUAR (montos de la base, por evaluar)', size=12, bold=True, color=NAVY)  # noqa: F821
for j, h in zip('BEF', ['Producto', 'Asociados', 'Monto']):
    put(ws, f'{j}18', h, bold=True, color=WHITE, bg=NAVY, h='center')  # noqa: F821
ws.merge_cells('B18:C18')
for k, (lab, key) in enumerate((('Hipotecario', 'P_HIPO'), ('Vehículo', 'P_VEH'), ('Tarjeta de crédito', 'P_TC'))):
    rr = 19 + k
    ws.merge_cells(f'B{rr}:C{rr}')
    put(ws, f'B{rr}', lab, border=True)  # noqa: F821
    put(ws, f'E{rr}', f'=COUNT(C_{key})', '#,##0', border=True, h='center')  # noqa: F821
    put(ws, f'F{rr}', f'=SUM(C_{key})', '$#,##0', border=True, h='right')  # noqa: F821

# conclusiones
put(ws, 'B23', '¿QUÉ NOS DICEN ESTOS RESULTADOS?', size=13, bold=True, color=WHITE, bg=NAVY)  # noqa: F821
for col in 'CDEFGHIJKLMNO':
    ws[f'{col}23'].fill = fill(NAVY)  # noqa: F821
T, C1 = 'COUNTIF(C_ID,"?*")', f'COUNTIF(C_ELEG,"{E1}")'  # noqa: F821
CONC = [
    f'="1. De "&TEXT({T},"#,##0")&" asociados de la base, "&TEXT({C1},"#,##0")&" ("&TEXT({C1}/{T},"0%")&") cumplen los criterios iniciales del '
    f'memorando. A "&TEXT(COUNTIF(C_ESTADO,"{E4}"),"#,##0")&" les falta información, sobre todo el ingreso ("&TEXT(COUNTIF(C_C_ING,"Sin información"),"#,##0")&" sin ingreso registrado)."',
    f'="2. De los que cumplen, "&TEXT(COUNTIF(C_ESTADO,"{E2}"),"#,##0")&" pueden continuar a la siguiente etapa de evaluación y "'
    f'&TEXT(COUNTIF(C_ESTADO,"{E3}"),"#,##0")&" requieren revisión antes de ofrecerles un producto (hoja 4)."',
    f'="3. "&TEXT(COUNTIFS(C_ELEG,"{E1}",C_NIVEL,"BAJO")/{C1},"0%")&" de los que cumplen tienen riesgo BAJO. Los "&TEXT(COUNTIFS(C_ELEG,"{E1}",C_NIVEL,"ALTO"),"#,##0")'
    '&" con riesgo ALTO NO están entre los que pueden continuar: están en revisión."',
    f'="4. Antes de otorgar falta verificar lista negra, créditos castigados y solvencia en Hacienda. Además, "&TEXT(COUNTIFS(C_ESTADO,"{E2}",C_VEH_OK,"No",C_P_VEH,">0"),"#,##0")'
    '&" montos de vehículo no coinciden con la tabla de vehículos de la Cooperativa: Créditos debe validarlos."',
]
for k, f in enumerate(CONC):
    rr = 24 + k
    ws.merge_cells(f'B{rr}:O{rr}')
    ws[f'B{rr}'] = f
    ws[f'B{rr}'].font = font(11)  # noqa: F821
    ws[f'B{rr}'].alignment = Alignment(wrap_text=True, vertical='center')
    ws.row_dimensions[rr].height = 34
    ws[f'B{rr}'].fill = fill(LIGHT)  # noqa: F821
ws['B29'] = f'=IF(COUNTIF({VERIF},"OK")={len(VER)},"✔ Verificación automática del archivo: todo cuadra (totales, listas y niveles).",' \
            '"✖ Verificación automática: hay algo que revisar (ver hoja Respaldo_Parametros).")'
ws['B29'].font = font(10, True, '548235')  # noqa: F821
ws['B30'] = 'Cómo leer: hoja 2 = todos los asociados con su resultado (use los filtros) · hoja 3 = criterios usados · hoja 4 = quiénes requieren revisión y qué hacer.'
ws['B30'].font = font(9, italic=True, color=GRAY)  # noqa: F821

# ============================================================== 2_Listado_Asociados (una fila por asociado, misma fila que el cálculo)
wl = wb.create_sheet('2_Listado_Asociados', 1)  # noqa: F821
wl.sheet_view.showGridLines = False
put(wl, 'A1', 'LISTADO DE ASOCIADOS Y RESULTADO DE LA PREEVALUACIÓN', size=15, bold=True, color=NAVY)  # noqa: F821
put(wl, 'A2', 'Use las flechas de los encabezados para filtrar por nivel de riesgo, producto o resultado. Resultado preliminar: no es aprobación.',
    size=9, italic=True, color=GRAY)  # noqa: F821
c_ = lambda k, r: f"'Respaldo_Calculo'!{COL[k]}{r}"  # noqa: F821
LCOLS = [
    ('Número de asociado', lambda r: f'={c_("ID", r)}', '@', 12),
    ('Nombre', lambda r: f'={c_("NOM", r)}', None, 32),
    ('Agencia', lambda r: f'={c_("AGE", r)}', None, 15),
    ('Ingreso registrado (al mes)', lambda r: f'=IF({c_("ID", r)}="","",IF({c_("ING", r)}="","No registrado",{c_("ING", r)}))', '$#,##0', 13),
    ('Deuda bancaria', lambda r: f'=IF({c_("ID", r)}="","",IF({c_("SSF", r)}="","No registrada",{c_("SSF", r)}))', '$#,##0', 13),
    ('Antigüedad como asociado (años)', lambda r: f'=IF({c_("ID", r)}="","",IF({c_("C_ANT", r)}="Sin información","No válida",{c_("ANT", r)}))', '0', 12),
    ('Nivel de riesgo', lambda r: f'=IF({c_("ID", r)}="","",IF({c_("NIVEL", r)}="No aplica","No se calcula",{c_("NIVEL", r)}))', None, 12),
    ('Producto que se propone evaluar', lambda r: f'={c_("PROD", r)}', None, 24),
    ('Monto hipotecario', lambda r: f'={c_("P_HIPO", r)}', '$#,##0', 12),
    ('Monto vehículo', lambda r: f'={c_("P_VEH", r)}', '$#,##0', 11),
    ('Monto tarjeta', lambda r: f'=IF({c_("P_TC", r)}<>"",{c_("P_TC", r)},{c_("E_TC", r)})', '$#,##0', 11),
    ('Resultado de la preevaluación', lambda r: f'={c_("ESTADO", r)}', None, 20),
    ('Motivo del resultado', lambda r: f'={c_("MOTIVO", r)}', None, 46),
    ('Qué falta revisar', lambda r: f'={c_("FALTA", r)}', None, 52),
]
HR = 4
head(wl, HR, [h for h, *_ in LCOLS], [w for *_, w in LCOLS], NAVY, 54)  # noqa: F821
for rr in range(HR + 1, HR + 1 + CAP):  # noqa: F821
    for j, (h, f, fmt, w) in enumerate(LCOLS, 1):
        c = wl.cell(rr, j, f(rr - HR + 1))
        c.font = font(9)  # noqa: F821
        if fmt:
            c.number_format = fmt
        c.alignment = Alignment(wrap_text=w >= 40, vertical='top')
ENDL = HR + CAP  # noqa: F821
wl.freeze_panes = 'C5'
wl.auto_filter.ref = f'A{HR}:{L(len(LCOLS))}{ENDL}'  # noqa: F821
cf_text(wl, f'G{HR + 1}:G{ENDL}', RISK)
cf_text(wl, f'L{HR + 1}:L{ENDL}', RES)

# ============================================================== 3_Criterios
wcri = wb.create_sheet('3_Criterios', 2)  # noqa: F821
wcri.sheet_view.showGridLines = False
for col, w in zip('ABCD', [26, 44, 62, 2]):
    wcri.column_dimensions[col].width = w
put(wcri, 'A1', 'CRITERIOS DE EVALUACIÓN', size=15, bold=True, color=NAVY)  # noqa: F821
put(wcri, 'A2', 'Qué revisamos y cómo se interpreta. Solo se usan datos que existen en la base.', size=9, italic=True, color=GRAY)  # noqa: F821
head(wcri, 4, ['Qué revisamos', 'Qué significa', 'Cómo se interpreta'], None, NAVY, 24)  # noqa: F821
CRI = [
    ('Criterios iniciales (memorando)', 'Requisitos del memorando del 27/10/2026', 'Edad de 30 a 55 años, ingreso y fondo compensatorio mayores de 0, '
     'categoría A1 o A2 interna y en la central de riesgo. Si no cumple uno: "No cumple criterios iniciales".'),
    ('Información completa', 'Que estén los datos necesarios', 'Si falta el ingreso, el fondo, la antigüedad o la categoría: "Información incompleta". '
     'Un dato vacío nunca se cuenta como cero.'),
    ('Antigüedad', 'Años de pertenecer a COMÉDICA', '="0 a "&P_ANT_ALTO_MAX&" año = riesgo alto · "&(P_ANT_ALTO_MAX+1)&" a "&(P_ANT_BAJO_MIN-1)&" años = medio · "&P_ANT_BAJO_MIN&" años o más = bajo (hoja Variables de la base)."'),
    ('Ingresos', 'Ingreso mensual registrado', '="Menos de "&TEXT(P_ING_L1,"$#,##0")&" = riesgo alto · de "&TEXT(P_ING_L1,"$#,##0")&" a menos de "&TEXT(P_ING_L2,"$#,##0")&" = medio · "&TEXT(P_ING_L2,"$#,##0")&" o más = bajo."'),
    ('Deuda bancaria', 'Saldo que debe en bancos y otras entidades (central de riesgo)', '="Menos de "&TEXT(P_DEU_L1,"$#,##0")&" o sin deuda = bajo · de "&TEXT(P_DEU_L1,"$#,##0")&" a menos de "&TEXT(P_DEU_L2,"$#,##0")&" = medio · "&TEXT(P_DEU_L2,"$#,##0")&" o más = alto."'),
    ('Deuda frente al ingreso', 'Si lo que debe es mucho para lo que gana', 'Deuda muy alta para lo que gana (ejemplo: gana $1,000 al mes y debe más de $12,000), o gana menos de $1,500 y debe $10,000 o más: '
     '"Requiere revisión" (salvo que su deuda sea menor que el hipotecario ofrecido, que podría pagarla).'),
    ('Aportaciones', 'Cuotas de aportación vencidas', '="0 = bajo · "&P_MORA_L1&" a "&(P_MORA_L2-1)&" = medio · "&P_MORA_L2&" o más = alto. No impide continuar: debe ponerse al día."'),
    ('Historial crediticio', 'Categoría en la central de riesgo (SSF)', 'A1 = bajo · A2 = medio. Es el único dato de historial que trae la base.'),
]
for k, row in enumerate(CRI):
    rr = 5 + k
    for j, v in enumerate(row, 1):
        c = wcri.cell(rr, j, v)
        c.font = font(10, j == 1)  # noqa: F821
        c.border = box  # noqa: F821
        c.alignment = Alignment(wrap_text=True, vertical='center')
        if k % 2:
            c.fill = fill(LIGHT)  # noqa: F821
    wcri.row_dimensions[rr].height = 44
r = 5 + len(CRI) + 1
put(wcri, f'A{r}', 'NIVEL DE RIESGO', size=12, bold=True, color=NAVY)  # noqa: F821
r += 1
for lv, txt in (('BAJO', 'Ningún dato en riesgo alto (antigüedad, ingreso, deuda, aportaciones, categoría en la central de riesgo).'),
                ('MEDIO', '1 dato en riesgo alto.'), ('ALTO', '2 o más datos en riesgo alto. Siempre va a "Requiere revisión".')):
    put(wcri, f'A{r}', f'Riesgo {lv}', bold=True, bg=RISK[lv], border=True)  # noqa: F821
    put(wcri, f'B{r}', txt, border=True, wrap=True)  # noqa: F821
    wcri.merge_cells(f'B{r}:C{r}')
    r += 1
put(wcri, f'A{r}', 'Ejemplo: gana $1,101 al mes (alto) y debe $11,640 (alto); lo demás bien = 2 datos en riesgo alto = Riesgo ALTO.',
    italic=True, color=GRAY, size=9)  # noqa: F821
r += 2
put(wcri, f'A{r}', 'RESULTADO DE LA PREEVALUACIÓN (distinto del nivel de riesgo)', size=12, bold=True, color=NAVY)  # noqa: F821
r += 1
for lab, txt in ((E2, 'Cumple los criterios iniciales, su deuda es adecuada para su ingreso y su riesgo no es alto. Pasa a la siguiente etapa.'),  # noqa: F821
                 (E3, 'Cumple los criterios iniciales, pero tiene riesgo alto o deuda alta para su ingreso. Hay que verificar antes de ofrecer.'),  # noqa: F821
                 (E4, 'Faltan datos para dar una conclusión.'), (E5, 'Tiene una condición del memorando que impide continuar.')):  # noqa: F821
    put(wcri, f'A{r}', lab, bold=True, bg=RES[lab], border=True)  # noqa: F821
    put(wcri, f'B{r}', txt, border=True, wrap=True)  # noqa: F821
    wcri.merge_cells(f'B{r}:C{r}')
    wcri.row_dimensions[r].height = 30
    r += 1
r += 1
put(wcri, f'A{r}', 'MONTOS Y PENDIENTES', size=12, bold=True, color=NAVY)  # noqa: F821
for t in ['Los montos son los que trae la base (PRECA HIPO, PRECA VEHICULO, PRECA TC). Riesgo no los cambia.',
          'Tarjeta: su límite se compara con la tabla de tarjeta de la base. Vehículo: si no coincide con su tabla, se avisa en "Qué falta revisar".',
          'Pendiente de aprobación institucional: los rangos de ingreso, deuda y aportaciones, y la regla del nivel de riesgo (propuestas de Riesgo).',
          'La base no trae la cuota mensual de las deudas: la capacidad de pago real se revisa en la evaluación formal.']:
    r += 1
    put(wcri, f'A{r}', '• ' + t, size=10)  # noqa: F821
for rr in range(1, r + 1):
    pass

# ============================================================== 4_Casos_por_Revisar (lista dinámica)
wcs = wb.create_sheet('4_Casos_por_Revisar', 3)  # noqa: F821
wcs.sheet_view.showGridLines = False
put(wcs, 'A1', 'CASOS POR REVISAR', size=15, bold=True, color=NAVY)  # noqa: F821
put(wcs, 'A2', 'Primero los que "Requieren revisión" (riesgo alto o deuda alta); después los de "Información incompleta". Cada fila dice qué revisar y qué hacer.',
    size=9, italic=True, color=GRAY)  # noqa: F821
wcs['A3'] = '=COUNT(C_K_CAS)&" asociados en esta lista"'
wcs['A3'].font = font(11, True, NAVY)  # noqa: F821
CCOLS = [('Número de asociado', 'ID', '@', 12), ('Nombre', 'NOM', None, 30), ('Resultado', 'ESTADO', None, 18),
         ('Nivel de riesgo', 'NIVEL', None, 11), ('Motivo por el que requiere revisión', 'MOTC', None, 24),
         ('Información que falta o alerta', 'ALERTA', None, 42), ('Acción que se recomienda', 'ACC2', None, 38),
         ('Tarjeta a evaluar', 'E_TC', '$#,##0', 11)]
head(wcs, 5, [h for h, *_ in CCOLS] + ['fila'], [w for *_, w in CCOLS] + [6], NAVY, 40)  # noqa: F821
last = len(CCOLS) + 1
for i in range(LCAP):  # noqa: F821
    rr = 6 + i
    c = wcs.cell(rr, last, f'=IFERROR(MATCH(SMALL(C_K_CAS,{i + 1}),C_K_CAS,0),"")')
    c.font = font(8, color='BFBFBF')  # noqa: F821
    for j, (h, k, fmt, w) in enumerate(CCOLS, 1):
        c = wcs.cell(rr, j, f'=IF(${L(last)}{rr}="","",IF(INDEX(C_{k},${L(last)}{rr})="","",INDEX(C_{k},${L(last)}{rr})))')  # noqa: F821
        c.font = font(9)  # noqa: F821
        if fmt:
            c.number_format = fmt
        c.alignment = Alignment(wrap_text=w >= 24, vertical='top')
wcs.column_dimensions[L(last)].hidden = True  # noqa: F821
ENDC = 5 + LCAP  # noqa: F821
wcs.freeze_panes = 'C6'
wcs.auto_filter.ref = f'A5:{L(len(CCOLS))}{ENDC}'  # noqa: F821
cf_text(wcs, f'D6:D{ENDC}', RISK)
cf_text(wcs, f'C6:C{ENDC}', RES)

# ============================================================== orden, colores de pestaña, impresión
order = ['1_Resumen_General', '2_Listado_Asociados', '3_Criterios', '4_Casos_por_Revisar', 'Respaldo_Parametros', 'Respaldo_Calculo',
         'Respaldo_Base_Original']
wb._sheets = [wb[n] for n in order]  # noqa: F821
for n, colr in zip(order, [NAVY, BLUE, BLUE, BLUE, 'A6A6A6', 'A6A6A6', 'A6A6A6']):  # noqa: F821
    wb[n].sheet_properties.tabColor = colr  # noqa: F821
    w_ = wb[n]  # noqa: F821
    w_.page_setup.orientation = 'landscape'
    w_.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    w_.page_setup.fitToWidth = 1
    w_.page_setup.fitToHeight = 0
wb.active = 0  # noqa: F821
wb.calculation.fullCalcOnLoad = True  # noqa: F821
wb.save(OUT_FILE)  # noqa: F821
print('ok', OUT_FILE)
