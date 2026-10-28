# -*- coding: utf-8 -*-
"""Genera el libro corto para decisión: Resumen_Ejecutivo_Base_PreEvaluada.xlsx.

Uso: python3 -I crear_resumen_ejecutivo.py <libro_tecnico_recalculado.xlsx> <salida.xlsx> [<base_original.xlsx>]
El libro resultante no lleva fórmulas (solo valores y enlaces), así que no necesita recalcularse.
Toma los resultados ya calculados (y verificados) de BD_CLASIFICADA del libro técnico y asigna a cada
asociado un nivel de seguridad acumulativo (1 a 5). El libro técnico se conserva como respaldo.
"""
import re
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

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
# Decisión POR PRODUCTO. El hipotecario se toma tal como viene en el Excel (no tiene tabla de rangos).
# El vehículo y la tarjeta se comparan con su tabla oficial: si el monto no coincide, se bloquea SOLO ese producto.
VBAD = x.CONSIST_VEHICULO == 'No coincide con tabla'
TBAD = x.CONSIST_TC == 'No coincide con tabla'
x['DQ_FUERA_TABLA_OFICIAL'] = (VBAD | TBAD).astype(int)
tc_ok = x.ELEGIBILIDAD_TC == CUMPLE
# Ganar mucho NO es un dato malo: el ingreso muy alto solo se confirma antes de otorgar (no decide SÍ / NO).
DATO_MAL = ((x.NIVEL_ANTIGUEDAD == 'Inconsistente - revisar') | (x.ALERTA_ANT_EDAD == 1) | (x.CHK_EDAD != 'Coincide')
            | x.ESTADO_DEUDA_SSF.astype(str).str.startswith('Saldo negativo') | (x.DUI.astype(str).str.len() != 9))
assert ((DATO_MAL | (x.ALERTA_ING_ATIPICO == 1)).astype(int) == x.DQ_DATO_INCONSISTENTE).all()
x['DQ_DATO_MAL'] = DATO_MAL.astype(int)
BASE_OK = (n3 & (x.DQ_DATO_MAL == 0) & (x.DQ_PRECA_SIN_INGRESO == 0)
           & (x.NIVEL_CAT_INTERNA == 'Bajo') & (x.NIVEL_CAT_EXTERNA == 'Bajo'))
# Hipotecario "trayendo la deuda": la base les dio hipotecario porque tienen capacidad o se les puede traer la deuda.
# Si su ÚNICO problema es la deuda (debe mucho / deuda vs. sueldo / poco fondo para lo que debe) y esa deuda cabe dentro
# del monto del hipotecario, se les recomienda el hipotecario (decisión del usuario). Vehículo y tarjeta siguen en NO.
for c in ('SALDO_SSF', 'INGRESO', 'FONDO_COMPENSATORIO'):
    x[c] = num(c)
RISK_NO_DEUDA = [c for c in RISK if c not in ('NIVEL_DEUDA', 'NIVEL_DEUDA_INGRESO')]
AL_NO_DEUDA = ['AL_ING_BAJO_DEUDA_ELEVADA', 'AL_ANTIG_BAJA_SIN_HISTORIAL', 'AL_CATEGORIA_A2', 'AL_APORT_VENCIDAS']
SOLO_DEUDA = (n1 & (x[AL_NO_DEUDA].sum(axis=1) == 0) & ~(x[RISK_NO_DEUDA] == 'Alto').any(axis=1)
              & (x.DQ_DATO_MAL == 0) & (x.DQ_PRECA_SIN_INGRESO == 0)
              & (x.NIVEL_CAT_INTERNA == 'Bajo') & (x.NIVEL_CAT_EXTERNA == 'Bajo'))
HIPO_DEUDA = SOLO_DEUDA & ~BASE_OK & x['PRECA HIPO'].notna() & (x.SALDO_SSF <= x['PRECA HIPO'])
REC_H = (BASE_OK & x['PRECA HIPO'].notna()) | HIPO_DEUDA
REC_V = BASE_OK & x['PRECA VEHICULO'].notna() & ~VBAD
REC_T = BASE_OK & tc_ok & x['PRECA TC'].notna() & ~TBAD
n4 = REC_H | REC_V | REC_T
PARCIAL = n4 & ((x['PRECA VEHICULO'].notna() & VBAD) | (x['PRECA TC'].notna() & TBAD))
x['RECOMENDADO'] = 'NO'
x.loc[n4, 'RECOMENDADO'] = 'SÍ'
x.loc[PARCIAL, 'RECOMENDADO'] = 'SÍ, sin vehículo'
L_HD = 'SÍ, hipotecario trayendo la deuda'
x.loc[HIPO_DEUDA, 'RECOMENDADO'] = L_HD
x['DEUDA_A_TRAER'] = x.SALDO_SSF.where(HIPO_DEUDA)
x['LE_SOBRA'] = (x['PRECA HIPO'] - x.SALDO_SSF).where(HIPO_DEUDA)
x['PRODUCTOS_REC'] = [', '.join(p for p, f in (('Hipotecario', h), ('Vehículo', v), ('Tarjeta', t)) if f)
                      for h, v, t in zip(REC_H, REC_V, REC_T)]
x['HIPO_REC'] = x['PRECA HIPO'].where(REC_H)
x['VEH_REC'] = x['PRECA VEHICULO'].where(REC_V)
x['TC_REC'] = x['PRECA TC'].where(REC_T)
n5 = n4 & x[ALL_LOW].isin(['Bajo', 'Sin deuda']).all(axis=1)
LEVELS = [
    (1, 'Cumple el memorando', 'Pasa los requisitos del memorando que se pueden revisar con la base.', n1),
    (2, 'Sin señales de riesgo', 'Además: no debe demasiado para lo que gana, no está atrasado en aportaciones y no es categoría A2.', n2),
    (3, 'Nada en rojo', 'Además: ningún dato en rojo (ingreso, deuda, antigüedad, aportaciones).', n3),
    (4, 'Máxima seguridad (RECOMENDADO)', 'Además: sus datos están bien, el monto que le dieron sí le corresponde según la tabla y es categoría A1.', n4),
    (5, 'Máxima seguridad, todo en verde', 'Parte del nivel 4 con todos sus datos en verde. Muy estricto: casi no deja tarjetas.', n5),
]
# ---------------------------------------------------------------- puntaje con los pesos indicados por el usuario
# Pesos indicados por el usuario en la columna "Peso" de la hoja 2: verde (Bajo) = 3, amarillo (Medio) = 2, rojo (Alto) = 1.
PESOS = [('Antigüedad', 'NIVEL_ANTIGUEDAD', 0.13, False), ('Ingreso', 'NIVEL_INGRESO', 0.30, False),
         ('Deuda (Central de Riesgo)', 'NIVEL_DEUDA', 0.20, False),
         ('Cuotas de aportación vencidas', 'NIVEL_MORA_APORTACIONES', 0.12, False),
         ('Categoría SSF (externa)', 'NIVEL_CAT_EXTERNA', 0.25, False)]
PESO_DE = {col: w for _lab, col, w, _inv in PESOS}
assert abs(sum(w for _, _, w, _ in PESOS) - 1) < 1e-9


def puntos(nivel, al_reves):
    if nivel == 'Sin deuda':
        return 3
    if nivel == 'Sin fondo':
        return 1
    return ({'Bajo': 1, 'Medio': 2, 'Alto': 3} if al_reves else {'Bajo': 3, 'Medio': 2, 'Alto': 1}).get(nivel)


for _lab, col, _w, inv in PESOS:
    x['PTS_' + col] = x[col].map(lambda v, inv=inv: puntos(v, inv))
x['PUNTAJE'] = sum(x['PTS_' + col] * w for _lab, col, w, _inv in PESOS).round(2)
BANDAS = [(2.6, 3.01, 'Muy bueno', '548235'), (2.2, 2.6, 'Bueno', '2E75B6'), (1.8, 2.2, 'Regular', 'BF8F00'),
          (1.0, 1.8, 'Bajo', 'C00000')]
x['CALIFICACION'] = x.PUNTAJE.map(lambda v: next((b for lo, hi, b, _ in BANDAS if lo <= v < hi), '') if pd.notna(v) else '')
x['NIVEL_SEGURIDAD'] = 0
for k, *_rest, m in LEVELS:
    x.loc[m, 'NIVEL_SEGURIDAD'] = k

def motivo(r):
    """Motivo de "no recomendado" en palabras simples, con el dato que lo comprueba."""
    if r.RECOMENDADO == 'SÍ':
        return ''
    if r.RECOMENDADO == L_HD:
        return (f'Se le recomienda SOLO el hipotecario, para traer su deuda: debe ${r.SALDO_SSF:,.0f} en bancos y otras '
                f'entidades y el hipotecario es de ${r["PRECA HIPO"]:,.0f} (le sobran ${r.LE_SOBRA:,.0f}). '
                f'Vehículo y tarjeta NO, por su deuda.')
    if r.RECOMENDADO == 'SÍ, sin vehículo':
        esp = r.ESPERADO_VEHICULO if isinstance(r.ESPERADO_VEHICULO, str) else f'${r.ESPERADO_VEHICULO:,.0f}'
        return (f'Se le recomienda: {r.PRODUCTOS_REC}. El vehículo NO: le dieron ${r["PRECA VEHICULO"]:,.0f}, '
                f'pero con su ingreso la tabla de vehículo dice {esp}.')
    out = []
    if r.NIVEL_INGRESO == 'Alto':
        out.append(f'Gana poco: ${r.INGRESO:,.0f} al mes (menos de $1,500)')
    if r.NIVEL_DEUDA == 'Alto':
        out.append(f'Debe mucho en bancos y otras entidades: ${r.SALDO_SSF:,.0f} ($10,000 o más)')
    if r.NIVEL_DEUDA_INGRESO == 'Alto':
        out.append(f'Lo que debe equivale a {r.MULT_DEUDA_INGRESO:,.0f} meses de su sueldo (12 o más)')
    if r.AL_FONDO_BAJO_DEUDA_ELEVADA == 1:
        out.append(f'Tiene poco fondo compensatorio (${r.FONDO_COMPENSATORIO:,.0f}) para lo que debe')
    if r.AL_APORT_VENCIDAS == 1:
        out.append(f'Está atrasado en sus aportaciones ({int(r.CUOTAS_VENC_APORT)} cuotas)')
    if r.NIVEL_ANTIGUEDAD == 'Alto':
        out.append(f'Es asociado nuevo ({int(r.ANTIGUEDAD_ANIOS)} año(s))')
    if r.AL_CATEGORIA_A2 == 1:
        out.append('Su categoría es A2, no A1')
    if r.CONSIST_VEHICULO == 'No coincide con tabla':
        esp = r.ESPERADO_VEHICULO if isinstance(r.ESPERADO_VEHICULO, str) else f'${r.ESPERADO_VEHICULO:,.0f}'
        out.append(f'Le dieron ${r["PRECA VEHICULO"]:,.0f} de vehículo, pero con su ingreso la tabla dice {esp}')
        if pd.isna(r['PRECA HIPO']) and pd.isna(r['PRECA TC']):
            out.append('En la base no tiene hipotecario ni tarjeta (N/A): no queda otro producto para ofrecerle')
    if r.CONSIST_TC == 'No coincide con tabla':
        out.append('El límite de tarjeta no coincide con la tabla')
    if r.CHK_EDAD != 'Coincide':
        out.append('La edad anotada no coincide con su fecha de nacimiento')
    if len(str(r.DUI)) != 9:
        out.append('El DUI está mal escrito en la base')
    if r.ALERTA_ANT_EDAD == 1:
        out.append('Los años de asociado no cuadran con su edad')
    if str(r.ESTADO_DEUDA_SSF).startswith('Saldo negativo'):
        out.append('Su deuda aparece en negativo')
    return '; '.join(out) + '.' if out else 'Revisar caso'


x['MOTIVO_NO_MAXIMA'] = x.apply(motivo, axis=1)

T_VEH, T_TC, T_ANT = 'Tabla 1 (vehículo)', 'Tabla 2 (tarjeta)', 'Tabla 3 (antigüedad)'
T_LIM, T_DATOS = 'Tabla 5 (límites de riesgo)', 'Tabla 6 (revisión de datos)'


def tablas(r):
    """Tablas de la hoja 8 con las que se comparó al asociado para decidir el NO."""
    if r.RECOMENDADO == 'SÍ':
        return ''
    if r.RECOMENDADO == 'SÍ, sin vehículo':
        return T_VEH
    if r.RECOMENDADO == L_HD:
        return T_LIM
    t = []
    if r.CONSIST_VEHICULO == 'No coincide con tabla':
        t.append(T_VEH)
    if r.CONSIST_TC == 'No coincide con tabla':
        t.append(T_TC)
    if r.NIVEL_ANTIGUEDAD == 'Alto':
        t.append(T_ANT)
    if (r.NIVEL_INGRESO == 'Alto' or r.NIVEL_DEUDA == 'Alto' or r.NIVEL_DEUDA_INGRESO == 'Alto'
            or r.AL_FONDO_BAJO_DEUDA_ELEVADA == 1 or r.AL_APORT_VENCIDAS == 1 or r.AL_CATEGORIA_A2 == 1):
        t.append(T_LIM)
    if (r.CHK_EDAD != 'Coincide' or len(str(r.DUI)) != 9 or r.ALERTA_ANT_EDAD == 1
            or str(r.ESTADO_DEUDA_SSF).startswith('Saldo negativo')):
        t.append(T_DATOS)
    return ', '.join(t)


x['TABLA_USADA'] = x.apply(tablas, axis=1)


VEH_SOLV_MIN = 25000  # regla propuesta por el usuario (no está en el memorando): vehículo mayor de $25,000 pide solvencia


def pendientes(r):
    p = []
    if pd.notna(r.HIPO_REC):
        p.append('Hipotecario: solvencia en Hacienda')
    if r.ALERTA_ING_ATIPICO == 1:
        p.append(f'Confirmar ingreso (es muy alto: ${r.INGRESO:,.0f})')
    if r.RECOMENDADO == L_HD:
        p.append('Hipotecario: confirmar garantía y que el crédito pague primero su deuda en la SSF')
    if pd.notna(r.VEH_REC) and r.VEH_REC > VEH_SOLV_MIN:
        p.append('Vehículo: solvencia en Hacienda')
    if pd.notna(r.TC_REC) and r.TC_REC > 4000:
        p.append('TC: solvencia en Hacienda o bajar a $4,000')
    if pd.notna(r.TC_REC) and r['APORTACIONES_AL_DIA'] == 'NO':
        p.append('TC: ponerse al día en aportaciones')
    return '; '.join(p)


x['PENDIENTE_ANTES_DE_OTORGAR'] = x.apply(pendientes, axis=1)
GEN = 'Solo las revisiones generales (lista negra, castigados, relacionados)'
x['PEND_HIPO'] = 'Solvencia en Hacienda (todo hipotecario es mayor de $30,000)'
x.loc[HIPO_DEUDA, 'PEND_HIPO'] = ('Solvencia en Hacienda. Además: confirmar garantía y que el hipotecario pague primero '
                                  'su deuda en bancos y otras entidades (por eso se le recomienda)')
x['PEND_VEH'] = x['PRECA VEHICULO'].map(
    lambda v: 'Solvencia en Hacienda (vehículo mayor de $25,000: regla propuesta, no está en el memorando)'
    if pd.notna(v) and v > VEH_SOLV_MIN else GEN)


def pend_tc(r):
    p = []
    if pd.notna(r['PRECA TC']) and r['PRECA TC'] > 4000:
        p.append('Solvencia en Hacienda; si no está solvente, darle $4,000')
    if r['APORTACIONES_AL_DIA'] == 'NO':
        p.append('Ponerse al día en aportaciones')
    return '; '.join(p) if p else GEN


x['PEND_TC'] = x.apply(pend_tc, axis=1)

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

# ================================================================ HOJA 1: RESUMEN
ws = wb.create_sheet('1_RESUMEN')
ws.sheet_properties.tabColor = NAVY
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 100
for col, w in zip('ABCDEFGHIJ', [2, 17, 17, 4, 17, 17, 4, 17, 17, 2]):
    ws.column_dimensions[col].width = w
ws.column_dimensions['K'].width = 4
for col in 'LM':
    ws.column_dimensions[col].width = 17
RES_LINKS = []  # (celda, clave de la hoja 10): se enlazan cuando exista la hoja 10
nb = n1 & ~n4
LAST = 'M'


def title_bar(r, text, color):
    ws.merge_cells(f'B{r}:{LAST}{r}')
    c = ws[f'B{r}']
    c.value = text
    c.font = font(12, True, 'FFFFFF')
    c.fill = PatternFill('solid', fgColor=color)
    c.alignment = Alignment(vertical='center', indent=1)
    ws.row_dimensions[r].height = 24
    return r + 2


def card(r, c1, c2, number, label, note, color, target=None):
    """Cuadro grande: número, etiqueta y nota, ocupando columnas c1:c2 y filas r a r+2."""
    for rr in range(r, r + 3):
        for cc in range(CI(c1), CI(c2) + 1):
            ws.cell(rr, cc).fill = PatternFill('solid', fgColor=color)
    ws.merge_cells(f'{c1}{r}:{c2}{r}')
    ws.merge_cells(f'{c1}{r + 1}:{c2}{r + 1}')
    ws.merge_cells(f'{c1}{r + 2}:{c2}{r + 2}')
    c = ws[f'{c1}{r}']
    c.value = number
    c.number_format = '#,##0'
    c.font = Font(name=F, size=26, bold=True, color='FFFFFF', underline='single' if target else None)
    c.alignment = Alignment(horizontal='center', vertical='center')
    if target:
        c.hyperlink = target
    c = ws[f'{c1}{r + 1}']
    c.value = label
    c.font = font(11, True, 'FFFFFF')
    c.alignment = Alignment(horizontal='center', vertical='center')
    c = ws[f'{c1}{r + 2}']
    c.value = note
    c.font = font(9, color='FFFFFF')
    c.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)


ws['B1'] = 'RESULTADO DE LA REVISIÓN DE LA BASE PRE EVALUADA'
ws['B1'].font = font(18, True, NAVY)
ws['B2'] = ('Memorando del 27/10/2026 · Base al 30/09/2026 · Hipotecario, vehículo y tarjeta de crédito. '
            'Es un apoyo para decidir: no es una aprobación de crédito.')
ws['B2'].font = font(9, italic=True, color=GRAY)

# ---- cuadros con las cifras principales
r = 4
ws.row_dimensions[r].height = 42
ws.row_dimensions[r + 1].height = 20
ws.row_dimensions[r + 2].height = 30
card(r, 'B', 'C', TOTAL, 'ASOCIADOS EN LA BASE', 'Todos los que vienen en su archivo', NAVY, "#'BD FINAL'!A1")
card(r, 'E', 'F', cnt(n1), 'CUMPLEN EL MEMORANDO', 'Pasan los requisitos que se pueden revisar', BLUE, "#'7_Los_528'!A1")
card(r, 'H', 'I', cnt(n4), 'RECOMENDADOS', 'Se les puede ofrecer al menos un producto', GREEN, "#'3_Recomendados'!A1")
card(r, 'L', 'M', cnt(nb), 'NO RECOMENDADOS', 'Cumplen el memorando pero tienen algún problema', ORANGE,
     "#'10_Asociados_por_tabla'!A1")
ws.merge_cells(f'B{r + 3}:{LAST}{r + 3}')  # nota bajo los cuadros
c = ws[f'B{r + 3}']
c.value = 'Haga clic en cualquier número para ver la lista de esos asociados.'
c.font = font(9, italic=True, color=GRAY)
c.alignment = Alignment(horizontal='center')

# ---- cómo se llegó a los recomendados
r = title_bar(r + 5, f'¿CÓMO SE LLEGÓ A LOS {cnt(n4)} RECOMENDADOS?', NAVY)
pasos = [
    ('PASO 1', 'Base recibida', TOTAL, 'Todos los asociados de su archivo.', 'D9E1F2'),
    ('PASO 2', 'Cumplen el memorando', cnt(n1),
     f'Se quitaron {TOTAL - cnt(n1):,}: a la mayoría les falta el ingreso; otros tienen ingreso $0, fondo $0 o ya cumplieron 56 años.', 'DDEBF7'),
    ('PASO 3', 'RECOMENDADOS', cnt(n4),
     f'Se quitaron {cnt(nb):,} que tienen algún problema: deben mucho, ganan poco, están atrasados, o les dieron un monto '
     f'que no corresponde a su tabla. Incluye {cnt(HIPO_DEUDA)} a los que solo se les da el hipotecario para traer su deuda.', 'C6E0B4'),
]
for paso, lab, val, expl, fill in pasos:
    cellw(ws, r, 2, paso, bold=True, fill=fill, align='center')
    ws.merge_cells(f'C{r}:D{r}')
    cellw(ws, r, 3, lab, bold=True, fill=fill)
    cellw(ws, r, 5, val, '#,##0', bold=True, fill=fill, align='center', size=12)
    ws.merge_cells(f'F{r}:{LAST}{r}')
    cellw(ws, r, 6, expl, size=9)
    ws.row_dimensions[r].height = 32
    r += 1

# ---- recomendados por producto
r = title_bar(r + 1, 'RECOMENDADOS POR PRODUCTO', GREEN)
heads = [('B', 'D', 'Producto'), ('E', 'E', 'Asociados'), ('F', 'G', 'Monto total'), ('H', 'I', 'Ver la lista')]
for a, b, h in heads:
    if a != b:
        ws.merge_cells(f'{a}{r}:{b}{r}')
    cellw(ws, r, CI(a), h, bold=True, fill='E2EFDA', align='center')
r += 1
for pl, col, sh in [('Hipotecario', 'PRECA HIPO', '4_Hipotecario'), ('Vehículo', 'PRECA VEHICULO', '5_Vehiculo'),
                    ('Tarjeta de crédito', 'PRECA TC', '6_Tarjeta')]:
    mm = {'PRECA HIPO': REC_H, 'PRECA VEHICULO': REC_V, 'PRECA TC': REC_T}[col]
    ws.merge_cells(f'B{r}:D{r}')
    cellw(ws, r, 2, pl, bold=True)
    cellw(ws, r, 5, cnt(mm), '#,##0', align='center')
    ws.merge_cells(f'F{r}:G{r}')
    cellw(ws, r, 6, float(x.loc[mm, col].sum()), INT, align='right')
    ws.merge_cells(f'H{r}:I{r}')
    cellw(ws, r, 8, '')
    link(ws, f'H{r}', sh, f'▶ {sh}')
    r += 1
ws.merge_cells(f'B{r}:{LAST}{r}')
c = ws[f'B{r}']
c.value = (f'Un asociado puede tener más de un producto. En total son {cnt(n4)} personas distintas '
           f'(lista completa en la hoja 3_Recomendados). Del hipotecario, {cnt(HIPO_DEUDA)} son "trayendo la deuda": su único '
           f'problema es que deben mucho, pero esa deuda cabe en el hipotecario que les dio la base.')
c.font = font(9, italic=True, color=GRAY)

# ---- por qué no se recomiendan
r = title_bar(r + 2, f'ORDEN DE LOS {cnt(n4)} RECOMENDADOS SEGÚN SU PUNTAJE (los NO no llevan puntaje)', NAVY)
ws.merge_cells(f'B{r}:{LAST}{r}')
c = ws[f'B{r}']
c.value = ('Primero se decide SÍ o NO con la lista de revisión. Después, SOLO a los SÍ se les da un puntaje con sus pesos, '
           'para ofrecer primero a los mejores.')
c.font = font(10, True)
c.alignment = Alignment(wrap_text=True, vertical='center')
ws.row_dimensions[r].height = 30
r += 1
for a_, b_, h in [('B', 'D', 'Calificación'), ('E', 'F', 'Puntaje'), ('G', 'H', 'Recomendados'), ('I', 'I', 'Ver')]:
    if a_ != b_:
        ws.merge_cells(f'{a_}{r}:{b_}{r}')
    cellw(ws, r, CI(a_), h, bold=True, fill='D9E1F2', align='center')
r += 1
for lo, hi, lab, col in BANDAS:
    m = (x.PUNTAJE >= lo) & (x.PUNTAJE < hi)
    ws.merge_cells(f'B{r}:D{r}')
    cellw(ws, r, 2, lab, bold=True, color=col)
    ws.merge_cells(f'E{r}:F{r}')
    cellw(ws, r, 5, f'{lo:.1f} a {hi - 0.01:.2f}' if hi < 3 else f'{lo:.1f} a 3.0', align='center')
    ws.merge_cells(f'G{r}:H{r}')
    cellw(ws, r, 7, cnt(m & n4), '#,##0', align='center')
    cellw(ws, r, 9, '')
    link(ws, f'I{r}', '12_Puntaje', '▶ Ver')
    r += 1
ws.merge_cells(f'B{r}:{LAST}{r}')
c = ws[f'B{r}']
c.value = ('Pesos (hoja 2): antigüedad 13%, ingreso 30%, deuda 20%, cuotas de aportación vencidas 12%, categoría SSF 25%. '
           'Verde = 3 puntos, amarillo = 2, rojo = 1. Las listas de recomendados ya van ordenadas de mayor a menor puntaje.')
c.font = font(9, italic=True, color=GRAY)
c.alignment = Alignment(wrap_text=True)
ws.row_dimensions[r].height = 26
r += 1

r = title_bar(r + 1, f'¿POR QUÉ {cnt(nb)} NO SE RECOMIENDAN?  (un asociado puede tener varios motivos)', ORANGE)
for a, b, h in [('B', 'G', 'Motivo'), ('H', 'H', 'Asociados'), ('I', 'I', 'Ver quiénes')]:
    if a != b:
        ws.merge_cells(f'{a}{r}:{b}{r}')
    cellw(ws, r, CI(a), h, bold=True, fill='FBE5D6', align='center')
r += 1
motivos = [
    ('Deben $10,000 o más en bancos y otras entidades', nb & (x.NIVEL_DEUDA == 'Alto'), 'T5b'),
    (f'No se les da el VEHÍCULO: su tabla dice NO APLICA (ganan menos de $3,000). '
     f'Ojo: {cnt(PARCIAL)} de ellos SÍ reciben hipotecario o tarjeta', n1 & VBAD, 'T1'),
    ('Ganan menos de $1,500 al mes', nb & (x.NIVEL_INGRESO == 'Alto'), 'T5a'),
    ('Tienen poco fondo compensatorio para lo que deben', nb & (x.AL_FONDO_BAJO_DEUDA_ELEVADA == 1), 'T5d'),
    ('Deben 12 meses o más de su sueldo', nb & (x.NIVEL_DEUDA_INGRESO == 'Alto'), 'T5c'),
    ('Están atrasados en sus aportaciones', nb & (x.AL_APORT_VENCIDAS == 1), 'T5e'),
    ('Tienen un dato mal anotado en la base (edad, DUI, años de asociado)', None, 'T6'),
    ('Son asociados nuevos (0 a 1 año)', nb & (x.NIVEL_ANTIGUEDAD == 'Alto'), 'T3'),
    ('Su categoría es A2 (no A1)', nb & (x.AL_CATEGORIA_A2 == 1), 'T5f'),
]
for lab, m, key in motivos:
    ws.merge_cells(f'B{r}:G{r}')
    cellw(ws, r, 2, lab)
    cellw(ws, r, 8, cnt(m) if m is not None else '', '#,##0', align='center')
    cellw(ws, r, 9, '▶ Ver', align='center')
    RES_LINKS.append((f'I{r}', key, f'H{r}'))
    r += 1

# ---- antes de otorgar
r = title_bar(r + 1, 'ANTES DE OTORGAR A LOS RECOMENDADOS (no viene en la base)', RED)
hip4 = cnt(REC_H)
tc4_big = cnt(REC_T & (x['PRECA TC'] > 4000))
veh4_big = cnt(REC_V & (x['PRECA VEHICULO'] > VEH_SOLV_MIN))
for t in ['✔ Revisar lista negra y créditos castigados con la Cooperativa.',
          '✔ Revisar que no sea persona relacionada ni empleado.',
          f'✔ Solvencia en Hacienda: obligatoria para los {hip4} hipotecarios recomendados.',
          f'✔ Solvencia en Hacienda también para vehículo mayor de $25,000 ({veh4_big} recomendados). '
          f'Regla propuesta por usted: el memorando no la pide.',
          f'✔ Tarjeta mayor de $4,000 ({tc4_big} recomendados): si no está solvente en Hacienda, se da con $4,000.',
          f'✔ Confirmar el ingreso de los {cnt(n4 & (x.ALERTA_ING_ATIPICO == 1))} recomendados que ganan muy alto (más de $12,000 al mes). '
          'Ganar mucho no es motivo de NO.',
          '✔ La base vence 3 meses después de que la apruebe el Comité.']:
    ws.merge_cells(f'B{r}:{LAST}{r}')
    c = ws[f'B{r}']
    c.value = t
    c.font = font(10)
    c.alignment = Alignment(vertical='center', indent=1)
    ws.row_dimensions[r].height = 18
    r += 1

# ---- problemas de la base
r = title_bar(r + 1, 'PROBLEMAS DE LA BASE (para que Créditos los corrija)', GRAY)
ing_vacio = int(x.INGRESO.isna().sum())
ing_cero = int((num('INGRESO') == 0).sum())
veh = num('PRECA VEHICULO')
for a_, b_, h in [('B', 'G', 'Problema'), ('H', 'H', 'Asociados'), ('I', 'I', 'Ver quiénes')]:
    if a_ != b_:
        ws.merge_cells(f'{a_}{r}:{b_}{r}')
    cellw(ws, r, CI(a_), h, bold=True, fill='E7E6E6', align='center')
r += 1
PROB_ROWS = []  # (fila, clave de la hoja 11)
for lab, val, key in [('Tienen monto asignado pero no tienen ingreso (vacío o $0)', ing_vacio + ing_cero, 'P1'),
                      ('Ya tienen 56 años al 30/09/2026 (el memorando pide hasta 55)', cnt(x.EDAD_RECALCULADA > 55), 'P2'),
                      ('Tienen fondo compensatorio $0 (el memorando pide fondo)', cnt(num('FONDO_COMPENSATORIO') == 0), 'P3'),
                      ('Están atrasados en aportaciones (el memorando dice 54; son los que tienen tarjeta)',
                       cnt(x.APORTACIONES_AL_DIA == 'NO'), 'P4'),
                      ('Datos mal anotados: antigüedad imposible, deuda en negativo, DUI mal escrito, vehículo de $150,000',
                       cnt((x.NIVEL_ANTIGUEDAD == 'Inconsistente - revisar') | (num('SALDO_SSF') < 0)
                           | (x.DUI.astype(str).str.len() != 9) | (veh == 150000)), 'P5')]:
    ws.merge_cells(f'B{r}:G{r}')
    cellw(ws, r, 2, lab)
    cellw(ws, r, 8, val, '#,##0', align='center')
    cellw(ws, r, 9, '▶ Ver', align='center')
    PROB_ROWS.append((r, key))
    r += 1
for t in ['El memorando dice 2,263 registros y 61 sin deuda; la base tiene 2,262 y 342 sin deuda registrada en la Central de Riesgo.',
          'Crédito de consumo: no viene en la base ni en el memorando.']:
    ws.merge_cells(f'B{r}:{LAST}{r}')
    c = ws[f'B{r}']
    c.value = '• ' + t
    c.font = font(9, italic=True, color=GRAY)
    r += 1

# ---- índice de hojas
r = title_bar(r + 1, 'HOJAS DE ESTE ARCHIVO (clic para abrir)', NAVY)
for sh, desc in [('2_Niveles_y_Rangos', 'Los rangos Bajo / Medio / Alto de cada dato.'),
                 ('3_Recomendados', f'Los {cnt(n4)} recomendados, con todos sus datos.'),
                 ('4_Hipotecario', 'Recomendados con hipotecario.'),
                 ('5_Vehiculo', 'Recomendados con vehículo.'),
                 ('6_Tarjeta', 'Recomendados con tarjeta.'),
                 ('7_Los_528', 'Los que cumplen el memorando: SÍ o NO recomendado y por qué.'),
                 ('8_Tablas_usadas', 'Las tablas con las que se comparó a cada asociado.'),
                 ('9_Como_explicarlo', 'Ejemplos paso a paso para explicar un NO.'),
                 ('10_Asociados_por_tabla', 'Los NO recomendados agrupados por motivo.'),
                 ('11_Problemas_de_la_base', 'Los asociados con problemas en la base (sin ingreso, 56 años, fondo $0...).'),
                 ('12_Puntaje', 'Cómo se calcula el puntaje y los recomendados ordenados del mejor al peor.'),
                 ('BD FINAL', 'Su base original, sin cambios.')]:
    ws.merge_cells(f'B{r}:D{r}')
    cellw(ws, r, 2, '')
    link(ws, f'B{r}', sh)
    ws.merge_cells(f'E{r}:{LAST}{r}')
    cellw(ws, r, 5, desc)
    r += 1
ws.merge_cells(f'B{r + 1}:{LAST}{r + 1}')
c = ws[f'B{r + 1}']
c.value = ('Los límites marcados como propuesta (deuda, ingreso, fondo, tabla de hipotecario) deben ser aprobados por Riesgos. '
           'Respaldo técnico: archivo Segmentacion_Riesgos_Base_PreEvaluada.xlsx.')
c.font = font(9, italic=True, color=GRAY)
c.alignment = Alignment(wrap_text=True)
ws.row_dimensions[r + 1].height = 26
ws.print_area = f'A1:{LAST}{r + 1}'


# ================================================================ HOJA 2: NIVELES Y RANGOS
wr = wb.create_sheet('2_Niveles_y_Rangos')
wr.sheet_properties.tabColor = BLUE
wr.sheet_view.showGridLines = False
for col, w in zip('ABCDEFGH', [2, 34, 22, 22, 22, 16, 44, 10]):
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
for j, h in enumerate(['Variable (nivel de RIESGO)', 'BAJO', 'MEDIO', 'ALTO', 'Sin dato / sin deuda', 'Nota', 'Peso'], 2):
    cellw(wr, r, j, h, bold=True, fill=NAVY, color='FFFFFF', align='center')
r += 1
for name, col, lo, me, hi, other, note in VARS:
    cellw(wr, r, 8, PESO_DE.get(col), '0%', bold=True, align='center', color=NAVY)
    cellw(wr, r, 2, name, bold=True)
    for j, (rng, lv) in enumerate([(lo, 'Bajo'), (me, 'Medio'), (hi, 'Alto')], 3):
        cellw(wr, r, j, f'{rng}\n{cnt_lvl(col, lv):,} asociados', fill=LEVEL_FILLS[lv], align='center')
    cellw(wr, r, 6, f'{sum(cnt_lvl(col, o) for o in other):,}' if other else '0', align='center')
    cellw(wr, r, 7, note, size=9)
    wr.row_dimensions[r].height = 34
    r += 1
cellw(wr, r, 7, 'Total de pesos', bold=True, align='right')
cellw(wr, r, 8, sum(PESO_DE.values()), '0%', bold=True, align='center', fill='F2F2F2')
r += 2
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
SCORE_COLS = [('PUNTAJE (1 a 3)', 'PUNTAJE', '0.00', 10), ('CALIFICACIÓN', 'CALIFICACION', None, 12)]
LVL_COLS = [('RIESGO POR ANTIGÜEDAD', 'NIVEL_ANTIGUEDAD', None, 11), ('RIESGO POR INGRESO', 'NIVEL_INGRESO', None, 11),
            ('RIESGO POR DEUDA', 'NIVEL_DEUDA', None, 11), ('RIESGO POR DEUDA vs SUELDO', 'NIVEL_DEUDA_INGRESO', None, 12)]
PEND = [('PENDIENTE ANTES DE OTORGAR (por producto)', 'PENDIENTE_ANTES_DE_OTORGAR', None, 44)]
PEND_H = [('PENDIENTE ANTES DE OTORGAR EL HIPOTECARIO', 'PEND_HIPO', None, 44)]
PEND_V = [('PENDIENTE ANTES DE OTORGAR EL VEHÍCULO', 'PEND_VEH', None, 44)]
PEND_T = [('PENDIENTE ANTES DE OTORGAR LA TARJETA', 'PEND_TC', None, 44)]


def list_sheet(name, color, title, note, df, cols, by_score=False):
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
    df = (df.sort_values(['PUNTAJE', 'NOMBRE'], ascending=[False, True]) if by_score
          else df.sort_values('ASOCIADO', key=lambda s: s.astype(int)))
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
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=src == 'RECOMENDADO')
            elif src in ('MOTIVO_NO_MAXIMA', 'PENDIENTE_ANTES_DE_OTORGAR', 'TABLA_USADA', 'PEND_HIPO', 'PEND_VEH', 'PEND_TC'):
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
            for txt, fc in (('SÍ', 'C6E0B4'), ('NO', 'F8CBAD'), ('SÍ, sin vehículo', 'FFF2CC'), (L_HD, 'DDEBF7')):
                ws.conditional_formatting.add(f'{L(j)}6:{L(j)}{end}', CellIsRule(
                    operator='equal', formula=[f'"{txt}"'], fill=PatternFill('solid', fgColor=fc, bgColor=fc),
                    font=Font(name=F, bold=True)))
    ws.freeze_panes = 'C6'
    ws.auto_filter.ref = f'A5:{L(len(cols))}{end}'
    # totales de montos
    rr = end + 2
    monto = [j for j, (h, src, fmt, w) in enumerate(cols, 1) if src in ('HIPO_REC', 'VEH_REC', 'TC_REC')]
    if monto:
        ws.cell(rr, monto[0] - 1, 'Total:').font = font(10, True)
    for j in monto:
        c = ws.cell(rr, j, float(df[cols[j - 1][1]].sum()))
        c.number_format = INT
        c.font = font(10, True)
    return ws


M_H = ('HIPOTECARIO', 'HIPO_REC', INT, 13)
M_V = ('VEHÍCULO', 'VEH_REC', INT, 12)
M_T = ('TARJETA', 'TC_REC', INT, 11)
M_P = ('PRODUCTOS RECOMENDADOS', 'PRODUCTOS_REC', None, 26)
EXT = 'Falta: lista negra, créditos castigados, relacionados/empleados y solvencia en Hacienda. No es aprobación de crédito.'
list_sheet('3_Recomendados', GREEN, 'RECOMENDADOS (se pueden ofrecer)',
           'Cumplen el memorando, no deben mucho, no ganan poco, están al día y son categoría A1. Se decide por producto: '
           'el hipotecario se toma como viene en el Excel (si dice N/A, no aplica); vehículo y tarjeta deben coincidir con su tabla (si el vehículo no '
           'coincide, se le ofrece solo lo demás). Los montos vacíos = ese producto no se recomienda. ' + EXT,
           x[n4], BASE_COLS[:2] + SCORE_COLS + BASE_COLS[2:] + [M_P, M_H, M_V, M_T] + LVL_COLS + PEND, by_score=True)
list_sheet('4_Hipotecario', GREEN, 'HIPOTECARIO – RECOMENDADOS',
           'Solo los que traen monto de hipotecario en la base (si dice N/A, no aplica). '
           f'{cnt(HIPO_DEUDA)} son "trayendo la deuda" (columna ¿CÓMO?): deben mucho, pero esa deuda cabe en el hipotecario. ' + EXT,
           x[REC_H], BASE_COLS[:2] + SCORE_COLS + [('¿CÓMO?', 'RECOMENDADO', None, 18)] + BASE_COLS[2:] + [M_H,
           ('DEUDA QUE SE TRAE', 'DEUDA_A_TRAER', INT, 13), ('LE SOBRA', 'LE_SOBRA', INT, 12)] + LVL_COLS + PEND_H, by_score=True)
list_sheet('5_Vehiculo', GREEN, 'VEHÍCULO – RECOMENDADOS',
           'Vehículo mayor de $25,000 ($30,000 y $40,000): pedir solvencia en Hacienda (regla propuesta, el memorando no la pide). '
           'Falta además: lista negra, créditos castigados y relacionados/empleados. No es aprobación de crédito.', x[REC_V], BASE_COLS[:2] + SCORE_COLS + BASE_COLS[2:] + [M_V] + LVL_COLS + PEND_V, by_score=True)
list_sheet('6_Tarjeta', GREEN, 'TARJETA DE CRÉDITO – RECOMENDADOS',
           'No tienen tarjeta actualmente. Límite > $4,000 sin solvencia en Hacienda: se otorga con $4,000. ' + EXT,
           x[REC_T],
           BASE_COLS[:2] + SCORE_COLS + BASE_COLS[2:] + [M_T, ('SEGMENTO', 'SEGMENTO', None, 20)] + LVL_COLS + PEND_T, by_score=True)
x['NIVEL_SEGURIDAD_TXT'] = x.NIVEL_SEGURIDAD.map({k: f'{k}. {nm}' for k, nm, *_ in LEVELS})
# ================================================================ HOJA 7: LOS 528, CON LISTA DE REVISIÓN ✔ / ✗
CHECKS = [  # (encabezado, ¿está bien?)
    ('GANA $1,500 O MÁS AL MES', lambda e: e.NIVEL_INGRESO != 'Alto'),
    ('DEBE MENOS DE $10,000', lambda e: e.NIVEL_DEUDA != 'Alto'),
    ('SU DEUDA ES MENOR A 12 MESES DE SUELDO', lambda e: e.NIVEL_DEUDA_INGRESO != 'Alto'),
    ('TIENE FONDO SUFICIENTE PARA LO QUE DEBE', lambda e: e.AL_FONDO_BAJO_DEUDA_ELEVADA != 1),
    ('AL DÍA EN APORTACIONES', lambda e: e.AL_APORT_VENCIDAS != 1),
    ('MÁS DE 1 AÑO DE ASOCIADO', lambda e: e.NIVEL_ANTIGUEDAD != 'Alto'),
    ('CATEGORÍA A1', lambda e: e.AL_CATEGORIA_A2 != 1),
    ('VEHÍCULO / TARJETA SEGÚN SU TABLA', lambda e: e.DQ_FUERA_TABLA_OFICIAL != 1),
    ('SUS DATOS ESTÁN BIEN ANOTADOS', lambda e: e.DQ_DATO_MAL != 1),
]
d7 = x[n1].sort_values('ASOCIADO', key=lambda s_: s_.astype(int))
for idx, e in d7.iterrows():  # todo ✔ debe ser exactamente lo mismo que RECOMENDADO
    assert all(f(e) for _h, f in CHECKS) == (e.RECOMENDADO == 'SÍ'), e.ASOCIADO
H7 = 9          # fila de encabezados de la hoja 7
FIRST7 = H7 + 1  # primera fila de datos
w7 = wb.create_sheet('7_Los_528')
w7.sheet_properties.tabColor = ORANGE
w7.sheet_view.showGridLines = False
w7['A1'] = f'LOS {len(d7)} QUE CUMPLEN EL MEMORANDO: ¿SE RECOMIENDAN O NO?'
w7['A1'].font = font(15, True, NAVY)
link(w7, 'A2', '1_RESUMEN', '◄ Volver al RESUMEN')
guia = [
    ('CÓMO LEER ESTA HOJA', None),
    ('1. Cada fila es un asociado. La columna ¿RECOMENDADO? dice SÍ (verde), NO (rojo), SÍ, sin vehículo (amarillo) o SÍ, hipotecario trayendo la deuda (celeste).', None),
    ('2. Las columnas del medio son una lista de revisión: ✔ verde = está bien;  ✗ rojo = ahí está el problema.', None),
    ('3. Si tiene aunque sea un ✗, es NO. Excepciones: si el ✗ es solo el vehículo = "SÍ, sin vehículo"; si el ✗ es solo la deuda y cabe en su hipotecario = "SÍ, hipotecario trayendo la deuda".', None),
    ('4. Para ver solo los NO (o solo los SÍ), use la flechita del encabezado ¿RECOMENDADO?.', None),
]
for i, (t, _) in enumerate(guia):
    rr = 3 + i
    w7.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=12)
    c = w7.cell(rr, 1, t)
    c.font = font(11 if i == 0 else 10, i == 0, 'FFFFFF' if i == 0 else '000000')
    c.fill = PatternFill('solid', fgColor=NAVY if i == 0 else 'EAF1FB')
    c.alignment = Alignment(vertical='center', indent=1)
    w7.row_dimensions[rr].height = 18
# fila de grupos
groups = [(1, 3, 'ASOCIADO', NAVY), (4, 4, 'RESULTADO', ORANGE), (5, 4 + len(CHECKS), 'LISTA DE REVISIÓN:  ✔ = bien   ✗ = problema', BLUE),
          (5 + len(CHECKS), 6 + len(CHECKS), 'EXPLICACIÓN', GRAY), (7 + len(CHECKS), 9 + len(CHECKS), 'MONTOS PRE EVALUADOS', GREEN),
          (10 + len(CHECKS), 11 + len(CHECKS), 'PUNTAJE', NAVY)]
for c1, c2, t, col in groups:
    w7.merge_cells(start_row=H7 - 1, start_column=c1, end_row=H7 - 1, end_column=c2)
    c = w7.cell(H7 - 1, c1, t)
    c.font = font(10, True, 'FFFFFF')
    c.alignment = Alignment(horizontal='center')
    for k in range(c1, c2 + 1):
        w7.cell(H7 - 1, k).fill = PatternFill('solid', fgColor=col)
        w7.cell(H7, k).fill = PatternFill('solid', fgColor=col)
heads = (['N° ASOCIADO', 'NOMBRE', 'AGENCIA', '¿RECOMENDADO?'] + [h for h, _ in CHECKS]
         + ['EXPLICACIÓN: POR QUÉ SÍ O POR QUÉ NO', 'TABLA USADA (clic)', 'HIPOTECARIO', 'VEHÍCULO', 'TARJETA', 'PUNTAJE (1 a 3)', 'CALIFICACIÓN'])
widths = [11, 34, 15, 17] + [12] * len(CHECKS) + [62, 18, 12, 11, 10, 10, 12]
for j, (h, w) in enumerate(zip(heads, widths), 1):
    c = w7.cell(H7, j, h)
    c.font = font(8 if 5 <= j <= 4 + len(CHECKS) else 9, True, 'FFFFFF')
    c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
    c.border = box
    w7.column_dimensions[L(j)].width = w
w7.row_dimensions[H7].height = 52
ROW7 = {}
for i, (idx, e) in enumerate(d7.iterrows()):
    rr = FIRST7 + i
    a = str(e.ASOCIADO)
    ROW7[a] = rr
    vals = ([a, e.NOMBRE, e.AGENCIA, e.RECOMENDADO] + ['✔' if f(e) else '✗' for _h, f in CHECKS]
            + [e.MOTIVO_NO_MAXIMA or '', e.TABLA_USADA or '']
            + [None if pd.isna(e[k]) else float(e[k]) for k in ('PRECA HIPO', 'PRECA VEHICULO', 'PRECA TC')]
            + ([float(e.PUNTAJE), e.CALIFICACION] if n4[idx] else [None, 'No se califica (es NO)']))
    for j, v in enumerate(vals, 1):
        c = w7.cell(rr, j, v)
        c.border = box
        c.font = font(9, bold=j in (1, 4))
        if j == 4 or 5 <= j <= 4 + len(CHECKS):
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=j == 4)
        elif j == 5 + len(CHECKS):
            c.alignment = Alignment(wrap_text=True, vertical='center')
        else:
            c.alignment = Alignment(vertical='center')
        if 7 + len(CHECKS) <= j <= 9 + len(CHECKS):
            c.number_format = INT
        elif j == 10 + len(CHECKS):
            c.number_format = '0.00'
            c.alignment = Alignment(horizontal='center', vertical='center')
        elif j == 11 + len(CHECKS):
            c.alignment = Alignment(horizontal='center', vertical='center')
LAST7 = FIRST7 + len(d7) - 1
rngR = f'D{FIRST7}:D{LAST7}'
rngC = f'E{FIRST7}:{L(4 + len(CHECKS))}{LAST7}'
for rng_, txt, fc, fcol in ((rngR, 'SÍ', 'C6E0B4', '006100'), (rngR, 'NO', 'F8CBAD', '9C0006'),
                            (rngR, 'SÍ, sin vehículo', 'FFF2CC', '7F6000'), (rngR, L_HD, 'DDEBF7', '1F3864'),
                            (rngC, '✔', 'E2EFDA', '548235'), (rngC, '✗', 'F8CBAD', 'C00000')):
    w7.conditional_formatting.add(rng_, CellIsRule(operator='equal', formula=[f'"{txt}"'],
                                                   fill=PatternFill('solid', fgColor=fc, bgColor=fc),
                                                   font=Font(name=F, bold=True, color=fcol)))
w7.freeze_panes = w7.cell(FIRST7, 5)
w7.auto_filter.ref = f'A{H7}:{L(len(heads))}{LAST7}'
TCOL7 = 6 + len(CHECKS)

# ================================================================ HOJA 8: TABLAS USADAS
w8 = wb.create_sheet('8_Tablas_usadas')
w8.sheet_properties.tabColor = GRAY
w8.sheet_view.showGridLines = False
for col, w in zip('ABCDEF', [2, 40, 30, 30, 34, 2]):
    w8.column_dimensions[col].width = w
w8['B1'] = 'TABLAS CON LAS QUE SE COMPARÓ A CADA ASOCIADO'
w8['B1'].font = font(15, True, NAVY)
link(w8, 'B2', '1_RESUMEN', '◄ Volver al RESUMEN')
w8['B3'] = ('La hoja 7 dice, para cada asociado, con qué tabla se comparó. Aquí están las tablas completas. '
            'Las tablas 1, 2 y 3 son de SU archivo; la 4 es solo informativa (no decide nada); la 5 es una propuesta hecha con sus datos.')
w8['B3'].font = font(10, True)
w8['B3'].alignment = Alignment(wrap_text=True)
w8.merge_cells('B3:E3')
w8.row_dimensions[3].height = 30
r = 5
noc = lambda col: cnt(n1 & ~n4 & (x[col] == 'No coincide con tabla'))


T_ROW = {}  # fila del título de cada tabla en la hoja 8 (para los enlaces)


def table_title(r, num, title, source, color, nota):
    T_ROW[num] = r
    w8.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    c = w8.cell(r, 2, f'TABLA {num}: {title}')
    c.font = font(12, True, 'FFFFFF')
    c.fill = PatternFill('solid', fgColor=color)
    w8.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=5)
    c = w8.cell(r + 1, 2, f'De dónde sale: {source}.  {nota}')
    c.font = font(9, italic=True)
    c.alignment = Alignment(wrap_text=True)
    w8.row_dimensions[r + 1].height = 26
    return r + 2


def table_rows(r, head, rows, fills=None):
    for j, h in enumerate(head, 2):
        cellw(w8, r, j, h, bold=True, fill='D9E1F2', align='center')
    r += 1
    for i, row in enumerate(rows):
        for j, v in enumerate(row, 2):
            fill = fills[i] if fills and j == 1 + len(row) else None
            cellw(w8, r, j, v, align='center' if j > 2 else None, fill=fill)
        r += 1
    return r + 1


OWN, PROP = '2E75B6', 'BF8F00'
NA = 'F8CBAD'
r = table_title(r, 1, 'VEHÍCULO (monto según ingreso)', 'su hoja "RANGOS VEHICULO" (copiada tal cual)', OWN,
                f'{cnt(n1 & VBAD & (x.INGRESO < 3000))} de los 528 ganan menos de $3,000 '
                f'y aun así la base les dio vehículo (esta tabla dice NO APLICA). '
                f'{cnt(n1 & VBAD & (x.INGRESO >= 3000))} tiene $150,000, monto que no existe en la tabla. '
                f'A ellos no se les recomienda el vehículo (lo demás sí, si no tienen otro problema).')
r = table_rows(r, ['Ingreso mensual', 'Monto de vehículo'],
               [('$15,000 o más', '$40,000'), ('$10,000 a $14,999.99', '$30,000'), ('$5,000 a $9,999.99', '$25,000'),
                ('$3,000 a $4,999.99', '$20,000'), ('Menos de $3,000', 'NO APLICA')], [None] * 4 + [NA])
r = table_title(r, 2, 'TARJETA DE CRÉDITO (límite según ingreso)', 'su hoja "RANGOS TC" (copiada tal cual)', OWN,
                'Ningún asociado de los 528 tiene un límite de tarjeta distinto al de esta tabla.')
r = table_rows(r, ['Ingreso mensual', 'Límite', 'Segmento'],
               [('$500 a $749.99', '$500', 'Clásica local'), ('$750 a $999.99', '$1,500', 'Clásica internacional'),
                ('$1,000 a $1,499.99', '$2,000', 'Clásica internacional'), ('$1,500 a $1,699.99', '$2,500', 'Dorada'),
                ('$1,700 a $1,999.99', '$3,000', 'Dorada'), ('$2,000 a $2,499.99', '$3,500', 'Platinum'),
                ('$2,500 a $2,999.99', '$4,500', 'Platinum'), ('$3,000 a $3,499.99', '$5,000', 'Platinum'),
                ('$3,500 a $7,499.99', '$7,000', 'Black'), ('$7,500 a $7,999.99', '$8,000', 'Black'),
                ('$8,000 o más', '$10,000', 'Black')])
r = table_title(r, 3, 'ANTIGÜEDAD COMO ASOCIADO', 'su hoja "Variables" (la que usted pidió)', OWN,
                f'{cnt(n1 & ~n4 & (x.NIVEL_ANTIGUEDAD == "Alto"))} de los NO son asociados nuevos (0 a 1 año).')
r = table_rows(r, ['Años como asociado', 'Nivel de riesgo'],
               [('0 a 1', 'Alto'), ('2 a 4', 'Medio'), ('5 o más', 'Bajo')], [NA, 'FFE699', 'C6E0B4'])
r = table_title(r, 4, 'HIPOTECARIO – SOLO INFORMATIVA (NO se usa para decidir SÍ / NO)',
                'su archivo NO trae tabla de hipotecario. Esta solo muestra cómo se asignaron los montos en su base',
                GRAY, 'Como no es una tabla oficial, ningún asociado queda en NO por el monto hipotecario. '
                'Créditos debe entregar la tabla oficial para poder revisarlo.')
r = table_rows(r, ['Ingreso mensual', 'Monto hipotecario'],
               [('$9,000 o más', '$205,000'), ('$5,000 a $8,999.99', '$150,000'), ('$3,000 a $4,999.99', '$110,000'),
                ('Menos de $3,000', 'NO APLICA')], [None] * 3 + [NA])
# evidencia: como se asignaron los montos hipotecarios en la base
hh = x[x['PRECA HIPO'].notna() & (num('INGRESO') > 0)]
tramos = [('Menos de $3,000', 0, 3000), ('$3,000 a $4,999.99', 3000, 5000), ('$5,000 a $8,999.99', 5000, 9000), ('$9,000 o más', 9000, 1e12)]
w8.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=5)
c = w8.cell(r - 1, 2, f'¿Cómo se armó? Se contó en SU base a los {len(hh)} asociados con hipotecario e ingreso registrado:')
c.font = font(10, True, NAVY)
ev = []
for lab, lo, hi in tramos:
    m = (hh.INGRESO >= lo) & (hh.INGRESO < hi)
    ev.append((lab,) + tuple(f'{int(((hh["PRECA HIPO"] == amt) & m).sum()):,}' for amt in (110000, 150000, 205000)))
r = table_rows(r, ['Ingreso mensual', 'Les dieron $110,000', 'Les dieron $150,000', 'Les dieron $205,000'], ev)
sin_ing = cnt(x['PRECA HIPO'].notna() & ~(num('INGRESO') > 0))
w8.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=5)
c = w8.cell(r - 1, 2, f'Para información: casi todos siguen el mismo patrón, pero {ev[0][1]} + {ev[0][2]} + {ev[0][3]} asociados recibieron hipotecario '
                      f'ganando menos de $3,000, y {sin_ing} no tienen ingreso registrado. Conviene preguntarle a Créditos qué regla usó.')
c.font = font(9, italic=True)
c.alignment = Alignment(wrap_text=True)
w8.row_dimensions[r - 1].height = 30
r += 1
r = table_title(r, 5, 'LÍMITES DE RIESGO', 'PROPUESTA hecha con los datos de su base (Riesgos debe aprobarla o cambiarla)', PROP,
                'Si el asociado cae en cualquier fila de esta tabla, se marca NO recomendado.')
r = table_rows(r, ['Se marca NO si…', 'Dato de su base', 'Por qué ese límite'],
               [('Gana menos de $1,500 al mes', 'INGRESO', 'Es el corte de $1,500 de su tabla de tarjeta'),
                ('Debe $10,000 o más', 'SALDO_SSF', 'Es el tercio de asociados que más debe'),
                ('Debe 12 meses o más de su sueldo', 'SALDO_SSF ÷ INGRESO', 'Debe un año de sueldo o más'),
                ('Tiene fondo menor de $4,000 y debe $10,000 o más', 'FONDO_COMPENSATORIO y SALDO_SSF', 'Poco respaldo para lo que debe'),
                ('Está atrasado en aportaciones', 'APORTACIONES_AL_DIA', 'El memorando pide ponerse al día'),
                ('Su categoría es A2 (no A1)', 'MAX_CAT_RIESGO y CALIFICACION_SSF', 'Para recomendar se pide A1')])
r = table_title(r, 6, 'REVISIÓN DE DATOS', 'errores encontrados en su propia base', GRAY,
                'No es riesgo del asociado: es un dato mal anotado. Si Créditos lo corrige, el asociado puede pasar a SÍ.')
r = table_rows(r, ['Se marca NO si…', 'Dato de su base'],
               [('La edad no coincide con la fecha de nacimiento', 'EDAD y CLI_FECHA_NACIMIENTO'),
                ('El DUI no tiene 9 dígitos', 'DUI'),
                ('Los años de asociado no cuadran con la edad', 'ANTIGUEDAD_ANIOS y EDAD'),
                ('La deuda aparece en negativo', 'SALDO_SSF')])

# ================================================================ HOJA 9: CÓMO EXPLICARLO (ejemplos paso a paso)
w9 = wb.create_sheet('9_Como_explicarlo')
w9.sheet_properties.tabColor = NAVY
w9.sheet_view.showGridLines = False
w9.column_dimensions['A'].width = 2
w9.column_dimensions['B'].width = 16
w9.column_dimensions['C'].width = 110
w9['B1'] = 'CÓMO EXPLICAR POR QUÉ UN ASOCIADO ES "NO RECOMENDADO" (ejemplos reales, paso a paso)'
w9['B1'].font = font(15, True, NAVY)
link(w9, 'B2', '1_RESUMEN', '◄ Volver al RESUMEN')
r = 4


EX = []  # (fila del paso 1, fila de "Cuántos", asociado) de cada ejemplo


def ejemplo(r, titulo, pasos, resultado, frase, total, ok=False):
    EX.append([r + 1, r + len(pasos) + 3, re.search(r'(\d+) –', pasos[0]).group(1)])
    w9.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = w9.cell(r, 2, titulo)
    c.font = font(12, True, 'FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    r += 1
    for i, p in enumerate(pasos, 1):
        cellw(w9, r, 2, f'Paso {i}', bold=True, fill='D9E1F2', align='center')
        cellw(w9, r, 3, p)
        r += 1
    rf = 'C6E0B4' if ok else 'F8CBAD'
    cellw(w9, r, 2, 'Resultado', bold=True, fill=rf, align='center')
    cellw(w9, r, 3, resultado, bold=True, fill=rf)
    r += 1
    cellw(w9, r, 2, 'Cómo decirlo', bold=True, fill='E2EFDA', align='center')
    cellw(w9, r, 3, frase, fill='E2EFDA')
    r += 1
    cellw(w9, r, 2, 'Cuántos', bold=True, fill='F2F2F2', align='center')
    cellw(w9, r, 3, total, fill='F2F2F2')
    return r + 2


nb = n1 & ~n4
e = x[nb & (x.CONSIST_VEHICULO == 'No coincide con tabla') & (x.INGRESO < 3000)].iloc[0]
r = ejemplo(r, 'EJEMPLO 1 – VEHÍCULO: "Le dieron vehículo, pero la tabla dice NO APLICA"', [
    f'Busque al asociado en su base: {e.ASOCIADO} – {e.NOMBRE}.',
    f'Mire cuánto gana (columna INGRESO): ${e.INGRESO:,.2f} al mes.',
    'Busque ese ingreso en la Tabla 1 (su hoja RANGOS VEHICULO): la última fila dice "Menor a $3,000 → NO APLICA".',
    f'Mire cuánto le asignaron en la base (columna PRECA VEHICULO): ${e["PRECA VEHICULO"]:,.0f}.'],
    f'Su propia tabla dice que no le corresponde vehículo, pero la base le dio ${e["PRECA VEHICULO"]:,.0f}. No cuadra.',
    '"Según nuestra tabla de vehículo, con lo que gana no le corresponde ese crédito. Hay que confirmar con Créditos '
    'por qué se le asignó."',
    f'{cnt(nb & (x.CONSIST_VEHICULO == "No coincide con tabla") & (x.INGRESO < 3000))} asociados de los 528 están en este caso.')
e = x[nb & (x.NIVEL_DEUDA == 'Alto') & (x.NIVEL_INGRESO != 'Alto') & (x.CONSIST_VEHICULO != 'No coincide con tabla')].iloc[0]
r = ejemplo(r, 'EJEMPLO 2 – DEUDA: "Debe mucho en bancos y otras entidades"', [
    f'Busque al asociado: {e.ASOCIADO} – {e.NOMBRE}.',
    f'Mire cuánto debe en la Central de Riesgo (columna SALDO_SSF): ${e.SALDO_SSF:,.2f} con {e.INSTITUCION}.',
    'Compare con la Tabla 5 (límites de riesgo): deber $10,000 o más se marca como riesgo Alto.'],
    'Debe $10,000 o más: no se recomienda.',
    '"Ya tiene una deuda alta; prestarle más aumenta el riesgo."',
    f'{cnt(nb & (x.NIVEL_DEUDA == "Alto"))} asociados de los 528 deben $10,000 o más.')
e = x[nb & (x.AL_APORT_VENCIDAS == 1)].iloc[0]
r = ejemplo(r, 'EJEMPLO 3 – APORTACIONES: "Está atrasado en sus aportaciones"', [
    f'Busque al asociado: {e.ASOCIADO} – {e.NOMBRE}.',
    f'Mire la columna APORTACIONES_AL_DIA: dice "{e.APORTACIONES_AL_DIA}", y CUOTAS_VENC_APORT: {int(e.CUOTAS_VENC_APORT)} cuota(s) vencidas.',
    'El memorando dice que quien no está al día debe ponerse al día antes de darle el producto.'],
    'Está atrasado con la Cooperativa: no se recomienda hasta que se ponga al día.',
    '"Primero debe ponerse al día con sus aportaciones."',
    f'{cnt(nb & (x.AL_APORT_VENCIDAS == 1))} asociados de los 528 están atrasados.')
e = x[(x.RECOMENDADO == 'SÍ') & x['PRECA VEHICULO'].notna()].iloc[0]
prods_e = [f'{t}: ${e[col]:,.0f} (correcto según {tb})' for t, col, tb in
           [('vehículo', 'PRECA VEHICULO', 'la Tabla 1'), ('tarjeta', 'PRECA TC', 'la Tabla 2')]
           if pd.notna(e[col])]
if pd.notna(e['PRECA HIPO']):
    prods_e.append(f'hipotecario: ${e["PRECA HIPO"]:,.0f} (sin tabla oficial: no se revisa)')
r = ejemplo(r, 'EJEMPLO 4 – UN "SÍ": por qué sí se recomienda', [
    f'Asociado: {e.ASOCIADO} – {e.NOMBRE}.',
    f'Gana ${e.INGRESO:,.2f} al mes; debe ${e.SALDO_SSF:,.2f}; tiene ${e.FONDO_COMPENSATORIO:,.2f} de fondo; está al día en aportaciones; '
    f'categoría A1.',
    'Lo que le dieron: ' + '; '.join(prods_e) + '. Ningún dato cae en la Tabla 5 ni en la Tabla 6.'],
    'Pasa todo: es RECOMENDADO.',
    '"Cumple el memorando, no tiene ninguna señal de riesgo y sus datos están correctos."',
    f'{cnt(x.RECOMENDADO == "SÍ")} asociados están en este caso (y {cnt(PARCIAL)} más con "SÍ, sin vehículo").', ok=True)
for row in w9.iter_rows(min_row=4, max_row=r):
    for c in row:
        if c.column == 3:
            c.alignment = Alignment(wrap_text=True, vertical='center')

# ================================================================ HOJA 10: ASOCIADOS POR TABLA (destino de los enlaces)
w10 = wb.create_sheet('10_Asociados_por_tabla')
w10.sheet_properties.tabColor = GRAY
w10.sheet_view.showGridLines = False
for col, w in zip('ABCDEFGHI', [2, 12, 38, 16, 18, 18, 18, 18, 30]):
    w10.column_dimensions[col].width = w
w10['B1'] = 'ASOCIADOS "NO RECOMENDADOS", AGRUPADOS POR LA TABLA QUE LOS DEJÓ FUERA'
w10['B1'].font = font(15, True, NAVY)
link(w10, 'B2', '1_RESUMEN', '◄ Volver al RESUMEN')
w10['B3'] = ('Cada bloque explica qué se revisó y muestra el dato de cada asociado. Haga clic en el N° de asociado para ver su fila '
             'completa en la hoja 7. Un asociado puede aparecer en más de un bloque.')
w10['B3'].font = font(10, italic=True)
SEC = {}
r = 5
nb = n1 & ~n4
money0 = lambda v: '' if pd.isna(v) else f'${v:,.0f}'
money2 = lambda v: '' if pd.isna(v) else f'${v:,.2f}'
to_txt = lambda v: v if isinstance(v, str) else money0(v)


def section(r, key, title, explica, table_num, mask, cols):
    """cols: [(encabezado, funcion(fila)->valor)]"""
    d = x[mask].sort_values('ASOCIADO', key=lambda s_: s_.astype(int))
    SEC[key] = (r, len(d))
    w10.merge_cells(start_row=r, start_column=2, end_row=r, end_column=9)
    c = w10.cell(r, 2, f'{title} — {len(d)} asociados')
    c.font = font(12, True, 'FFFFFF')
    c.fill = PatternFill('solid', fgColor=NAVY)
    w10.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=8)
    c = w10.cell(r + 1, 2, explica)
    c.font = font(9)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    w10.row_dimensions[r + 1].height = 30
    c = w10.cell(r + 1, 9, f'◄ Ver la Tabla {table_num}')
    c.hyperlink = f"#'8_Tablas_usadas'!B{T_ROW[table_num]}"
    c.font = Font(name=F, size=10, bold=True, color='0563C1', underline='single')
    r += 2
    heads = ['N° ASOCIADO', 'NOMBRE', 'AGENCIA'] + [h for h, _ in cols]
    for j, h in enumerate(heads, 2):
        cellw(w10, r, j, h, bold=True, fill='D9E1F2', align='center', size=9)
    r += 1
    for _, e in d.iterrows():
        a = str(e.ASOCIADO)
        c = cellw(w10, r, 2, a, size=9, bold=True)
        c.hyperlink = f"#'7_Los_528'!A{ROW7[a]}"
        c.font = Font(name=F, size=9, bold=True, color='0563C1', underline='single')
        cellw(w10, r, 3, e.NOMBRE, size=9)
        cellw(w10, r, 4, e.AGENCIA, size=9)
        for j, (h, fn) in enumerate(cols, 5):
            cellw(w10, r, j, fn(e), size=9, align='center')
        r += 1
    return r + 2


r = section(r, 'T1', 'TABLA 1 – VEHÍCULO: le dieron vehículo pero no cuadra con su tabla',
            'Qué se revisó: se buscó lo que gana cada asociado en su tabla RANGOS VEHICULO y se comparó con el vehículo que le asignó '
            'la base. Si gana menos de $3,000, la tabla dice NO APLICA. Solo se bloquea el VEHÍCULO: si no tiene otro problema, '
            'se le recomienda lo demás (columna ¿SE LE OFRECE OTRO PRODUCTO?).', 1, n1 & VBAD,
            [('GANA AL MES', lambda e: money2(e.INGRESO)), ('LE DIERON DE VEHÍCULO', lambda e: money0(e['PRECA VEHICULO'])),
             ('SEGÚN LA TABLA 1 LE TOCA', lambda e: to_txt(e.ESPERADO_VEHICULO)), ('¿CUADRA?', lambda e: 'NO'),
             ('¿SE LE OFRECE OTRO PRODUCTO?', lambda e: f'SÍ: {e.PRODUCTOS_REC}' if e.RECOMENDADO != 'NO' else 'NO')])
r = section(r, 'T3', 'TABLA 3 – ANTIGÜEDAD: asociados nuevos (0 a 1 año)',
            'Qué se revisó: los años como asociado (columna ANTIGUEDAD_ANIOS). Según su hoja Variables, 0 a 1 año es riesgo Alto.', 3,
            nb & (x.NIVEL_ANTIGUEDAD == 'Alto'),
            [('AÑOS DE ASOCIADO', lambda e: int(e.ANTIGUEDAD_ANIOS)), ('NIVEL', lambda e: 'Alto')])
r = section(r, 'T5a', 'TABLA 5 – GANA POCO (menos de $1,500 al mes)', 'Qué se revisó: la columna INGRESO.', 5,
            nb & (x.NIVEL_INGRESO == 'Alto'), [('GANA AL MES', lambda e: money2(e.INGRESO))])
r = section(r, 'T5b', 'TABLA 5 – DEBE MUCHO ($10,000 o más)',
            'Qué se revisó: lo que debe en bancos y otras entidades según la Central de Riesgo (columna SALDO_SSF).', 5,
            nb & (x.NIVEL_DEUDA == 'Alto'), [('DEBE', lambda e: money2(e.SALDO_SSF)), ('INSTITUCIÓN', lambda e: e.INSTITUCION)])
r = section(r, 'T5c', 'TABLA 5 – DEBE 12 MESES O MÁS DE SU SUELDO',
            'Qué se revisó: se dividió lo que debe entre lo que gana al mes. 12 o más quiere decir que debe un año de sueldo o más.', 5,
            nb & (x.NIVEL_DEUDA_INGRESO == 'Alto'),
            [('DEBE', lambda e: money0(e.SALDO_SSF)), ('GANA AL MES', lambda e: money0(e.INGRESO)),
             ('MESES DE SUELDO', lambda e: f'{e.MULT_DEUDA_INGRESO:,.0f}')])
r = section(r, 'T5d', 'TABLA 5 – POCO FONDO PARA LO QUE DEBE (fondo menor de $4,000 y deuda de $10,000 o más)',
            'Qué se revisó: el fondo compensatorio contra lo que debe.', 5, nb & (x.AL_FONDO_BAJO_DEUDA_ELEVADA == 1),
            [('FONDO COMPENSATORIO', lambda e: money2(e.FONDO_COMPENSATORIO)), ('DEBE', lambda e: money0(e.SALDO_SSF))])
r = section(r, 'T5e', 'TABLA 5 – ATRASADO EN APORTACIONES',
            'Qué se revisó: las columnas APORTACIONES_AL_DIA y CUOTAS_VENC_APORT. El memorando pide que se ponga al día.', 5,
            nb & (x.AL_APORT_VENCIDAS == 1),
            [('¿AL DÍA?', lambda e: e.APORTACIONES_AL_DIA), ('CUOTAS VENCIDAS', lambda e: int(e.CUOTAS_VENC_APORT))])
r = section(r, 'T5f', 'TABLA 5 – CATEGORÍA A2 (no A1)', 'Qué se revisó: la categoría interna y la de la Central de Riesgo.', 5,
            nb & (x.AL_CATEGORIA_A2 == 1),
            [('CATEGORÍA INTERNA', lambda e: '' if pd.isna(e.MAX_CAT_RIESGO) else e.MAX_CAT_RIESGO),
             ('CATEGORÍA SSF', lambda e: '' if pd.isna(e.CALIFICACION_SSF) else e.CALIFICACION_SSF)])


def dato_malo(e):
    out = []
    if e.CHK_EDAD != 'Coincide':
        out.append(f'Edad anotada {int(e.EDAD)}, por su fecha de nacimiento ({e.CLI_FECHA_NACIMIENTO}) tiene {int(e.EDAD_RECALCULADA)}')
    if len(str(e.DUI)) != 9:
        out.append(f'DUI "{e.DUI}" no tiene 9 dígitos')
    if e.ALERTA_ANT_EDAD == 1:
        out.append(f'{int(e.ANTIGUEDAD_ANIOS)} años de asociado con {int(e.EDAD)} años de edad')
    if str(e.ESTADO_DEUDA_SSF).startswith('Saldo negativo'):
        out.append('Deuda en negativo')
    return '; '.join(out)


w10.column_dimensions['E'].width = 24
r = section(r, 'T6', 'TABLA 6 – DATOS MAL ANOTADOS EN LA BASE',
            'No es riesgo del asociado: es un dato que hay que corregir o confirmar con Créditos. Si se corrige, puede pasar a SÍ.', 6,
            nb & (x.TOTAL_ALERTAS_DATOS > 0) & (x.MOTIVO_NO_MAXIMA.str.contains('edad anotada|DUI|no cuadran|negativo', case=False)),
            [('QUÉ DATO ESTÁ MAL', dato_malo)])
# la hoja 7 se ordena por N° de asociado: sus filas sirven de destino para los enlaces
for rr in range(FIRST7, LAST7 + 1):
    c = w7.cell(rr, TCOL7)
    mm = re.search(r'Tabla (\d)', str(c.value or ''))
    if mm:
        c.hyperlink = f"#'8_Tablas_usadas'!B{T_ROW[int(mm.group(1))]}"
        c.font = Font(name=F, size=9, color='0563C1', underline='single')

# ---- enlaces desde la hoja 8 hacia las listas
w8.column_dimensions['F'].width = 30


def link_cell(ws_, cell, target, text):
    c = ws_[cell]
    c.value = text
    c.hyperlink = target
    c.font = Font(name=F, size=10, bold=True, color='0563C1', underline='single')


SEC_LINK = lambda key: f"#'10_Asociados_por_tabla'!B{SEC[key][0]}"
for num, key in ((1, 'T1'), (3, 'T3'), (6, 'T6')):
    link_cell(w8, f'F{T_ROW[num]}', SEC_LINK(key), f'▶ Ver los {SEC[key][1]} asociados')
link_cell(w8, f'F{T_ROW[5]}', SEC_LINK('T5a'), '▶ Ver los asociados (por fila)')
for i, key in enumerate(['T5a', 'T5b', 'T5c', 'T5d', 'T5e', 'T5f']):
    link_cell(w8, f'F{T_ROW[5] + 3 + i}', SEC_LINK(key), f'▶ Ver los {SEC[key][1]}')
w8.cell(T_ROW[2], 6, 'Nadie: 0 asociados').font = font(10, True, GRAY)
w8.cell(T_ROW[4], 6, 'No decide nada').font = font(10, True, GRAY)

# ---- enlaces desde la hoja 9 (ejemplos)
for (paso1, cuantos, aso), key in zip(EX, ['T1', 'T5b', 'T5e', None]):
    c = w9.cell(paso1, 3)
    c.hyperlink = f"#'7_Los_528'!A{ROW7[aso]}"
    c.font = Font(name=F, size=10, color='0563C1', underline='single')
    c = w9.cell(cuantos, 3)
    if key:
        c.value = f'{c.value}  ▶ Clic aquí para ver la lista.'
        c.hyperlink = SEC_LINK(key)
    else:
        c.value = f'{c.value}  ▶ Clic aquí para ver la lista.'
        c.hyperlink = "#'3_Recomendados'!A1"
    c.font = Font(name=F, size=10, bold=True, color='0563C1', underline='single')

# ---- enlaces desde el RESUMEN (motivos -> hoja 10)
for cell, key, cnt_cell in RES_LINKS:
    ws[cell].hyperlink = SEC_LINK(key)
    ws[cell].font = Font(name=F, size=10, bold=True, color='0563C1', underline='single')
    if key == 'T6':
        ws[cnt_cell].value = SEC[key][1]
        ws[cnt_cell].number_format = '#,##0'

# ================================================================ HOJA 11: PROBLEMAS DE LA BASE (toda la base)
w11 = wb.create_sheet('11_Problemas_de_la_base')
w11.sheet_properties.tabColor = GRAY
w11.sheet_view.showGridLines = False
for col, w in zip('ABCDEFGHIJ', [2, 12, 38, 12, 16, 16, 16, 16, 16, 16]):
    w11.column_dimensions[col].width = w
w11['B1'] = 'PROBLEMAS ENCONTRADOS EN LA BASE (para que Créditos los revise y corrija)'
w11['B1'].font = font(15, True, NAVY)
link(w11, 'B2', '1_RESUMEN', '◄ Volver al RESUMEN')
w11['B3'] = ('Aquí aparecen asociados de TODA la base (2,262), no solo de los 528. Son errores o faltantes del archivo: '
             'no se pudo decidir con ellos hasta que se corrijan.')
w11['B3'].font = font(10, italic=True)
PSEC = {}
numc = lambda c_: pd.to_numeric(x[c_], errors='coerce')  # 'num' se reutiliza como variable de ciclo más arriba
r = 5
age_txt = lambda e: f'{int(e.EDAD_RECALCULADA)}'


def psection(r, key, title, explica, mask, cols):
    d = x[mask].sort_values('ASOCIADO', key=lambda s_: s_.astype(int))
    PSEC[key] = r
    w11.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
    c = w11.cell(r, 2, f'{title} — {len(d):,} asociados')
    c.font = font(12, True, 'FFFFFF')
    c.fill = PatternFill('solid', fgColor=GRAY)
    w11.merge_cells(start_row=r + 1, start_column=2, end_row=r + 1, end_column=9)
    c = w11.cell(r + 1, 2, explica)
    c.font = font(9)
    c.alignment = Alignment(wrap_text=True, vertical='top')
    w11.row_dimensions[r + 1].height = 28
    c = w11.cell(r + 1, 10, '◄ RESUMEN')
    c.hyperlink = "#'1_RESUMEN'!A1"
    c.font = Font(name=F, size=10, bold=True, color='0563C1', underline='single')
    r += 2
    for j, h in enumerate(['N° ASOCIADO', 'NOMBRE', 'DUI', 'AGENCIA'] + [h for h, _ in cols], 2):
        cellw(w11, r, j, h, bold=True, fill='E7E6E6', align='center', size=9)
    r += 1
    for _, e in d.iterrows():
        a = str(e.ASOCIADO)
        c = cellw(w11, r, 2, a, size=9, bold=True)
        if a in ROW7:
            c.hyperlink = f"#'7_Los_528'!A{ROW7[a]}"
            c.font = Font(name=F, size=9, bold=True, color='0563C1', underline='single')
        cellw(w11, r, 3, e.NOMBRE, size=9)
        cellw(w11, r, 4, str(e.DUI), size=9)
        cellw(w11, r, 5, e.AGENCIA, size=9)
        for j, (h, fn) in enumerate(cols, 6):
            cellw(w11, r, j, fn(e), size=9, align='center')
        r += 1
    return r + 2


ing_txt = lambda e: 'VACÍO' if pd.isna(e.INGRESO) else money2(e.INGRESO)
montos = [('HIPOTECARIO', lambda e: money0(e['PRECA HIPO'])), ('VEHÍCULO', lambda e: money0(e['PRECA VEHICULO'])),
          ('TARJETA', lambda e: money0(e['PRECA TC']))]
r = psection(r, 'P1', 'MONTO ASIGNADO SIN INGRESO',
             'La base les asignó un monto, pero la columna INGRESO está vacía o en $0. Con qué ingreso se calculó el monto, '
             'no se sabe. Créditos debe completar el ingreso.', ~(numc('INGRESO') > 0),
             [('INGRESO', ing_txt)] + montos)
r = psection(r, 'P2', 'YA TIENEN 56 AÑOS AL 30/09/2026',
             'El memorando pide de 30 a 55 años cumplidos. La columna EDAD dice 55 porque se calculó antes; por la fecha de '
             'nacimiento, al 30/09/2026 ya tienen 56.', x.EDAD_RECALCULADA > 55,
             [('FECHA DE NACIMIENTO', lambda e: e.CLI_FECHA_NACIMIENTO), ('EDAD EN LA BASE', lambda e: int(e.EDAD)),
              ('EDAD AL 30/09/2026', age_txt)])
r = psection(r, 'P3', 'FONDO COMPENSATORIO EN $0',
             'El memorando dice que todos poseen fondo compensatorio; estos tienen $0.', numc('FONDO_COMPENSATORIO') == 0,
             [('FONDO', lambda e: money2(e.FONDO_COMPENSATORIO))] + montos)
r = psection(r, 'P4', 'ATRASADOS EN APORTACIONES',
             'El memorando habla de 54 (son los que tienen tarjeta pre evaluada). En la base hay más: también los que solo '
             'tienen hipotecario o vehículo.', x.APORTACIONES_AL_DIA == 'NO',
             [('CUOTAS VENCIDAS', lambda e: int(e.CUOTAS_VENC_APORT)),
              ('¿TIENE TARJETA PRE EVALUADA?', lambda e: 'SÍ' if pd.notna(e['PRECA TC']) else 'NO')] + montos[:2])


def que_esta_mal(e):
    out = []
    if e.NIVEL_ANTIGUEDAD == 'Inconsistente - revisar':
        out.append(f'Antigüedad imposible: {int(e.ANTIGUEDAD_ANIOS)} años')
    if pd.notna(e.SALDO_SSF) and e.SALDO_SSF < 0:
        out.append(f'Deuda en negativo: ${e.SALDO_SSF:,.2f}')
    if len(str(e.DUI)) != 9:
        out.append(f'DUI mal escrito: "{e.DUI}"')
    if e['PRECA VEHICULO'] == 150000:
        out.append('Vehículo de $150,000 (no existe en la tabla)')
    return '; '.join(out)


w11.column_dimensions['F'].width = 16
r = psection(r, 'P5', 'DATOS MAL ANOTADOS',
             'Datos que no pueden ser correctos y deben corregirse en el sistema.',
             (x.NIVEL_ANTIGUEDAD == 'Inconsistente - revisar') | (numc('SALDO_SSF') < 0)
             | (x.DUI.astype(str).str.len() != 9) | (numc('PRECA VEHICULO') == 150000),
             [('QUÉ ESTÁ MAL', que_esta_mal)])
for rr in range(1, w11.max_row + 1):
    c = w11.cell(rr, 6)
    if isinstance(c.value, str) and len(c.value) > 20:
        c.alignment = Alignment(wrap_text=True, vertical='center')
for rr, key in PROB_ROWS:
    c = ws.cell(rr, 9)
    c.hyperlink = f"#'11_Problemas_de_la_base'!B{PSEC[key]}"
    c.font = Font(name=F, size=10, bold=True, color='0563C1', underline='single')

# ================================================================ HOJA 12: PUNTAJE CON LOS PESOS
w12 = wb.create_sheet('12_Puntaje')
w12.sheet_properties.tabColor = NAVY
w12.sheet_view.showGridLines = False
for col, w in zip('ABCDEFGHIJKL', [2, 12, 34, 13, 14, 14, 14, 14, 14, 12, 13, 2]):
    w12.column_dimensions[col].width = w
w12['B1'] = 'PUNTAJE DE LOS RECOMENDADOS (para ofrecer primero a los mejores)'
w12['B1'].font = font(15, True, NAVY)
link(w12, 'B2', '1_RESUMEN', '◄ Volver al RESUMEN')
w12['B3'] = ('Primero se decide SÍ o NO (hoja 7). El puntaje se da SOLO a los SÍ, para ordenarlos del mejor al peor. '
             'Los NO no llevan puntaje. Pesos: columna "Peso" de la hoja 2.')
w12['B3'].font = font(9, italic=True, color=RED)
r = 5
w12.merge_cells(f'B{r}:K{r}')
c = w12.cell(r, 2, 'PASO 1 – Cada dato da puntos según su color')
c.font = font(12, True, 'FFFFFF')
c.fill = PatternFill('solid', fgColor=NAVY)
r += 1
for j, h in enumerate(['Dato', '', 'Peso', 'Verde (Bajo)', 'Amarillo (Medio)', 'Rojo (Alto)'], 2):
    if j != 3:
        cellw(w12, r, j, h, bold=True, fill='D9E1F2', align='center')
w12.merge_cells(f'B{r}:C{r}')
r += 1
for lab, col, w, inv in PESOS:
    cellw(w12, r, 2, lab, bold=True)
    w12.merge_cells(f'B{r}:C{r}')
    cellw(w12, r, 4, w, '0%', bold=True, align='center', color=NAVY)
    trio = ('3 puntos', '2 puntos', '1 punto')
    for j, (t, fc) in enumerate(zip(trio, ('C6E0B4', 'FFE699', 'F8CBAD')), 5):
        cellw(w12, r, j, t, align='center', fill=None if inv else fc, size=9)
    r += 1
cellw(w12, r, 2, 'Total', bold=True, fill='F2F2F2')
w12.merge_cells(f'B{r}:C{r}')
cellw(w12, r, 4, 1, '0%', bold=True, align='center', fill='F2F2F2')
r += 2
w12.merge_cells(f'B{r}:K{r}')
c = w12.cell(r, 2, 'PASO 2 – Puntaje = suma de (puntos × peso). Va de 1 (peor) a 3 (mejor).   PASO 3 – Se califica:')
c.font = font(12, True, 'FFFFFF')
c.fill = PatternFill('solid', fgColor=NAVY)
r += 1
for j, h in enumerate(['Calificación', '', 'Puntaje', 'Recomendados'], 2):
    if j != 3:
        cellw(w12, r, j, h, bold=True, fill='D9E1F2', align='center')
w12.merge_cells(f'B{r}:C{r}')
r += 1
for lo, hi, lab, colr in BANDAS:
    m = (x.PUNTAJE >= lo) & (x.PUNTAJE < hi)
    cellw(w12, r, 2, lab, bold=True, color=colr)
    w12.merge_cells(f'B{r}:C{r}')
    cellw(w12, r, 4, f'{lo:.1f} a {hi - 0.01:.2f}' if hi < 3 else f'{lo:.1f} a 3.0', align='center')
    cellw(w12, r, 5, cnt(m & n4), '#,##0', align='center')
    r += 1
r += 1
w12.merge_cells(f'B{r}:K{r}')
c = w12.cell(r, 2, f'LOS {cnt(n4)} RECOMENDADOS, DEL MEJOR AL PEOR PUNTAJE (clic en el N° para ver su fila en la hoja 7)')
c.font = font(12, True, 'FFFFFF')
c.fill = PatternFill('solid', fgColor=NAVY)
r += 1
heads12 = ['N° ASOCIADO', 'NOMBRE', 'AGENCIA'] + [f'{lab}\n({w:.0%})' for lab, _, w, _ in PESOS] + ['PUNTAJE', 'CALIFICACIÓN']
for j, h in enumerate(heads12, 2):
    cellw(w12, r, j, h, bold=True, fill='D9E1F2', align='center', size=9)
w12.row_dimensions[r].height = 46
H12 = r
r += 1
d12 = x[n4].sort_values(['PUNTAJE', 'NOMBRE'], ascending=[False, True])
for _, e in d12.iterrows():
    a = str(e.ASOCIADO)
    c = cellw(w12, r, 2, a, size=9)
    c.hyperlink = f"#'7_Los_528'!A{ROW7[a]}"
    c.font = Font(name=F, size=9, bold=True, color='0563C1', underline='single')
    cellw(w12, r, 3, e.NOMBRE, size=9)
    cellw(w12, r, 4, e.AGENCIA, size=9)
    for j, (lab, col, w, inv) in enumerate(PESOS, 5):
        cellw(w12, r, j, f'{int(e["PTS_" + col])}  ({e[col]})', size=9, align='center')
    cellw(w12, r, 10, float(e.PUNTAJE), '0.00', size=10, bold=True, align='center')
    cellw(w12, r, 11, e.CALIFICACION, size=9, bold=True, align='center')
    r += 1
last12 = r - 1
for lo, hi, lab, colr in BANDAS:
    w12.conditional_formatting.add(f'K{H12 + 1}:K{last12}', CellIsRule(
        operator='equal', formula=[f'"{lab}"'], font=Font(name=F, bold=True, color=colr)))
w12.freeze_panes = w12.cell(H12 + 1, 4)
w12.auto_filter.ref = f'B{H12}:K{last12}'

# ================================================================ BD FINAL original (copia sin cambios)
BD_SRC = sys.argv[3] if len(sys.argv) > 3 else None
if BD_SRC:
    src_ws = openpyxl.load_workbook(BD_SRC)['BD FINAL']
    wbd = wb.create_sheet('BD FINAL')
    wbd.sheet_properties.tabColor = '404040'
    for row in src_ws.iter_rows():
        for c in row:
            if c.value is None:
                continue
            d = wbd.cell(c.row, c.column, c.value)
            d.number_format = c.number_format
            d.font = font(9, bold=(c.row == 1), color='FFFFFF' if c.row == 1 else '000000')
            if c.row == 1:
                d.fill = PatternFill('solid', fgColor='404040')
    for j in range(1, src_ws.max_column + 1):
        wbd.column_dimensions[L(j)].width = 14
    wbd.freeze_panes = 'C2'
    wbd.auto_filter.ref = f'A1:{L(src_ws.max_column)}{src_ws.max_row}'

wb.active = 0
wb.save(OUT)
print({f'nivel {k}': cnt(m) for k, _, _, m in LEVELS})
print('hojas:', wb.sheetnames)
