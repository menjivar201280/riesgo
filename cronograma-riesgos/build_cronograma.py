# -*- coding: utf-8 -*-
"""
Genera el libro Excel «Cronograma_Dashboard_Riesgos.xlsx».

Uso:
    python build_cronograma.py [ruta_salida.xlsx]

El script también expone build(...) con parámetros opcionales de prueba
(actividades temporales, semana actual, fecha de referencia) que usa
test_cronograma.py. El archivo entregable se genera SIN datos de prueba.

Las fórmulas se escriben con nombres de función en inglés porque así las
almacena el formato .xlsx; Excel en español las muestra traducidas
(LOOKUP = BUSCAR, IFERROR = SI.ERROR, SUMPRODUCT = SUMAPRODUCTO, etc.).
"""
import re
import sys
import datetime as dt

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.formatting.rule import FormulaRule, DataBarRule, CellIsRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo, TableFormula
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

# --------------------------------------------------------------------------
# Paleta y estilos
# --------------------------------------------------------------------------
NAVY = "1F3864"       # azul oscuro: encabezados principales
MID = "2E75B6"        # azul medio: indicadores
LIGHT = "DDEBF7"      # azul muy claro: apoyo
INPUT = "EDEDED"      # gris suave: celdas que se llenan (desbloqueadas)
CALC = "EAF3FB"       # azul muy claro: celdas automáticas (bloqueadas)
GREEN_F, GREEN_T = "C6EFCE", "006100"
DKGREEN = "1E7B34"
AMBER_F, AMBER_T = "FFE699", "7F6000"
ORANGE_F = "F8CBAD"
RED_F, RED_T = "FFC7CE", "9C0006"
BLUE_F, BLUE_T = "BDD7EE", "1F3864"
TXT = "1A1A1A"

FONT = "Arial"


def font(size=10, bold=False, color=TXT, italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
LEFT_TOP = Alignment(horizontal="left", vertical="top", wrap_text=True)

PUESTOS = [
    "Riesgo Operacional",
    "Riesgo Normativo",
    "Riesgo Financiero",
    "Asistente Técnico de Riesgo",
]
SHORT = ["R. Operacional", "R. Normativo", "R. Financiero", "Asistente Técnico"]
ESTADOS = ["Pendiente", "En curso", "Completado"]
SEMANAS = ["Semana 1", "Semana 2", "Semana 3", "Semana 4"]

S_DASH = "Dashboard - Consolidado"
# (hoja de cronograma, hoja de actividades (vista automática), código, color propio del puesto) – nombres ≤ 31 caracteres
AREAS = [
    ("Cronograma R. Operacional", "Actividades R. Operacional", "RO", "C55A11"),       # naranja
    ("Cronograma R. Normativo", "Actividades R. Normativo", "RN", "7030A0"),           # morado
    ("Cronograma R. Financiero", "Actividades R. Financiero", "RF", "548235"),         # verde
    ("Cronograma Asistente Técnico", "Actividades Asistente Técnico", "ATR", "2E75B6"),  # azul
]
# Hoja de la jefatura: mismo formato que un Cronograma (título, hoja, código, color)
ADMIN = ("Administración (jefatura)", "Administración", "ADM", "595959")
MES_FMT = "[$-440A]mmmm yyyy"   # nombre del mes en español
S_HIST = "Historial mensual"
S_CONF = "Configuración"
S_INST = "Instrucciones"
S_CALC = "Calculos"

FIRST = 5            # primera fila de datos de las tablas
N_ACT_ROWS = 300     # filas preformateadas por hoja de cronograma (se acumulan los meses)
N_VISTA_ROWS = 100   # filas de la hoja Actividades (actividades del mes mostrado)
LAST_RNG = 1000      # límite de los rangos con nombre que leen el Dashboard

# Columnas de la hoja Cronograma (clave, encabezado). Las claves se usan en las fórmulas
# para no depender de letras fijas.
COL_SPEC = [
    ("id", "N°"), ("tipo", "Tipo"), ("act", "Actividad"), ("deleg", "Delegar a"),
    ("ini", "Fecha de inicio"), ("lim", "Fecha límite"), ("cump", "Fecha de cumplimiento"),
    ("peso", "Peso (%)"),
    ("s1", "Avance Semana 1 (%)"), ("s2", "Avance Semana 2 (%)"), ("s3", "Avance Semana 3 (%)"),
    ("s4", "Avance Semana 4 (%)"), ("av", "% de Avance Total Actual"), ("est", "Estado"),
    ("alerta", "Alerta de plazo"), ("seg", "Seguimiento de la delegación"), ("obs", "Observaciones"), ("rev", "Revisión automática (alertas)"),
    ("mes", "Mes"),
    # auxiliares ocultas
    ("orden", "Orden en el mes (aux)"), ("grupo", "Grupo (aux)"), ("t", "Tipo (aux)"),
    ("nsub", "N subactividades (aux)"), ("ap", "Avance propio (aux)"),
    ("o1", "Propio S1 (aux)"), ("o2", "Propio S2 (aux)"), ("o3", "Propio S3 (aux)"), ("o4", "Propio S4 (aux)"),
    ("e1", "Efectivo S1 (aux)"), ("e2", "Efectivo S2 (aux)"), ("e3", "Efectivo S3 (aux)"), ("e4", "Efectivo S4 (aux)"),
    ("pef", "Peso efectivo (aux)"),
    ("dfila", "Fila en el puesto (aux)"), ("dn", "N° delegada en el mes (aux)"), ("dkey", "Clave delegada (aux)"),
]
ADMIN_ONLY = ("deleg", "seg")       # columnas visibles solo en la hoja Administración
PREFIJO_JEF = "[Jefatura] "         # marca de las actividades delegadas por la jefatura
COLS_ACT = [h for _, h in COL_SPEC]
C = {k: get_column_letter(i + 1) for i, (k, _) in enumerate(COL_SPEC)}       # clave -> letra
CN = {k: i + 1 for i, (k, _) in enumerate(COL_SPEC)}                          # clave -> n° de columna
COLS_VISTA = [
    "N°", "Actividad", "Fecha de inicio", "Fecha límite", "Semana de entrega",
    "Semana 1", "Semana 2", "Semana 3", "Semana 4", "% Avance actual", "Estado", "Alerta de plazo",
    "Observaciones",
]

# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def q(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def style_range(ws, ref, **kw):
    for row in ws[ref]:
        for c in row:
            for k, v in kw.items():
                setattr(c, k, v)


def merge_set(ws, ref, value=None, **kw):
    ws.merge_cells(ref)
    first = ws[ref.split(":")[0]]
    if value is not None:
        first.value = value
    style_range(ws, ref, **kw)
    return first


def add_name(wb, name, ref):
    dn = DefinedName(name, attr_text=ref)
    wb.defined_names[name] = dn


# --------------------------------------------------------------------------
# Fórmulas de la hoja Cronograma (fila r). Ver COL_SPEC para las columnas.
# - Tipo: Macroactividad (o vacío) / Subactividad (pertenece a la macroactividad más cercana hacia arriba).
# - Fecha de cumplimiento: si se escribe, la fila queda al 100% (Completado / Cumplida).
# - Avance de una macroactividad con subactividades = promedio de sus subactividades (valen igual).
# - Peso: solo macroactividades; se ajusta solo a 100% en el mes (vacío = partes iguales).
# --------------------------------------------------------------------------
LR = LAST_RNG
TIPOS = ["Macroactividad", "Subactividad"]
WK = [C["s1"], C["s2"], C["s3"], C["s4"]]
OWN = [C["o1"], C["o2"], C["o3"], C["o4"]]
EFF = [C["e1"], C["e2"], C["e3"], C["e4"]]


def col_abs(k):
    c = C.get(k, k)
    return f"${c}${FIRST}:${c}${LR}"


def col_upto(k, r):
    c = C.get(k, k)
    return f"${c}${FIRST}:${c}{r}"


def X(k, r):
    """Referencia a la celda de la columna k en la fila r (columna absoluta)."""
    return f"${C[k]}{r}"


def reg(r):
    """Fila registrada: hay algún dato de entrada."""
    return f"(COUNTA({X('tipo', r)}:{X('s4', r)})+COUNTA({X('obs', r)}))>0"


def f_tipo(r):
    return f'=IF(NOT({reg(r)}),"",IF({X("tipo", r)}="Subactividad","Sub","Macro"))'


def f_grupo(r):
    t = col_upto("t", r)
    return (f'=IF({X("t", r)}="","",IF({X("t", r)}="Macro",ROW(),'
            f'IFERROR(LOOKUP(2,1/({t}="Macro"),ROW({t})),0)))')


def f_nsub(r):
    return f'=IF({X("t", r)}="Macro",COUNTIFS({col_abs("grupo")},ROW(),{col_abs("t")},"Sub"),"")'


def ciclo_ini(r):
    """Lunes de inicio del ciclo del mes de la fila (el de Configuración si es el mes mostrado)."""
    m = X("mes", r)
    return f'IF({m}=MesCiclo,InicioCiclo,{m}+MOD(2-WEEKDAY({m}),7))'


def cumplida_en(r, w):
    """La fila tiene fecha de cumplimiento dentro o antes de la semana w (1–4) de su ciclo."""
    return f'AND(ISNUMBER({X("cump", r)}),{X("cump", r)}<={ciclo_ini(r)}+{7 * w - 1})'


def f_ap(r):
    """Avance propio: 100% si tiene fecha de cumplimiento; si no, último avance semanal escrito."""
    s1, s4 = X("s1", r), X("s4", r)
    return (f'=IF({X("t", r)}="","",IF(ISNUMBER({X("cump", r)}),1,'
            f'IFERROR(LOOKUP(2,1/({s1}:{s4}<>""),{s1}:{s4}),0)))')


def f_own(r, w):
    c = X(["s1", "s2", "s3", "s4"][w], r)
    prev = f"${OWN[w - 1]}{r}" if w else "0"
    return (f'=IF({X("t", r)}="","",IF({cumplida_en(r, w + 1)},1,'
            f'IF(ISNUMBER({c}),{c},{prev})))')


def f_eff(r, w):
    grp = f'{col_abs("grupo")},ROW(),{col_abs("t")},"Sub"'
    return (f'=IF({X("t", r)}="","",IF({cumplida_en(r, w + 1)},1,IF(N({X("nsub", r)})>0,'
            f'AVERAGEIFS({col_abs(OWN[w])},{grp}),${OWN[w]}{r})))')


def _avance_core(r):
    grp = f'{col_abs("grupo")},ROW(),{col_abs("t")},"Sub"'
    return (f'IF(ISNUMBER({X("cump", r)}),1,IF(N({X("nsub", r)})>0,'
            f'AVERAGEIFS({col_abs("ap")},{grp}),{X("ap", r)}))')


def f_avance(r):
    return f'=IF(NOT({reg(r)}),"",{_avance_core(r)})'


def f_estado(r):
    av = X("av", r)
    return (f'=IF({av}="","",IF({av}>=1,"Completado",'
            f'IF({av}>0,"En curso","Pendiente")))')


def _alerta_core(r):
    """Alerta de plazo: cumplida (a tiempo o con retraso) / vencida / vence hoy / por vencer / en plazo."""
    av, lim, cump = X("av", r), X("lim", r), X("cump", r)
    dias = lambda a, b: f'({a}-{b})&IF({a}-{b}=1," día"," días")'
    return (f'IF(ISNUMBER({cump}),IF(AND(ISNUMBER({lim}),{cump}>{lim}),"Cumplida con "&{dias(cump, lim)}&" de retraso","Cumplida"),'
            f'IF(N({av})>=1,"Cumplida",IF(NOT(ISNUMBER({lim})),"Sin fecha límite",'
            f'IF({lim}<FechaRef,"Vencida hace "&{dias("FechaRef", lim)},'
            f'IF({lim}=FechaRef,"Vence hoy",IF({lim}-FechaRef<=DiasAlerta,"Vence en "&{dias(lim, "FechaRef")},"En plazo"))))))')


def f_alerta(r):
    return f'=IF(NOT({reg(r)}),"",{_alerta_core(r)})'


def _mes_de(ini, lim):
    return (f'IF(ISNUMBER({ini}),DATE(YEAR({ini}),MONTH({ini}),1),'
            f'IF(ISNUMBER({lim}),DATE(YEAR({lim}),MONTH({lim}),1),MesCiclo))')


def f_mes(r):
    """Mes de la actividad (fecha de inicio o límite). Las subactividades usan el mes de su macroactividad."""
    g = f'{X("grupo", r)}-{FIRST - 1}'
    m_ini = f"INDEX({col_abs('ini')},{g})"
    m_lim = f"INDEX({col_abs('lim')},{g})"
    return (f'=IF(NOT({reg(r)}),"",IF(AND({X("t", r)}="Sub",N({X("grupo", r)})>0),{_mes_de(m_ini, m_lim)},'
            f'{_mes_de(X("ini", r), X("lim", r))}))')


def f_orden(r):
    """Auxiliar: posición de la fila dentro del mes mostrado (para la hoja Actividades)."""
    m = X("mes", r)
    return f'=IF(OR({m}="",{m}<>MesCiclo),"",COUNTIF({col_upto("mes", r)},MesCiclo))'


def f_pef(r):
    """Peso efectivo de la macroactividad en su mes: se ajusta solo a 100% (vacío = partes iguales)."""
    mm = f'{col_abs("mes")},{X("mes", r)},{col_abs("t")},"Macro"'
    n_mac = f'COUNTIFS({mm})'
    n_w = f'COUNTIFS({mm},{col_abs("peso")},">0")'
    s_w = f'SUMIFS({col_abs("peso")},{mm})'
    p = X("peso", r)
    return (f'=IF({X("t", r)}<>"Macro","",IF({n_w}=0,1/{n_mac},'
            f'IF(AND(ISNUMBER({p}),N({p})>0),{p},{s_w}/{n_w})/({n_mac}*{s_w}/{n_w})))')


def f_id(r):
    """N° automático: 1, 2, 3… para macroactividades (por mes) y 1.1, 1.2… para subactividades."""
    g = f'{X("grupo", r)}-{FIRST - 1}'
    T, M = col_abs("t"), col_abs("mes")
    macro_num = (f'COUNTIFS(INDEX({T},1):INDEX({T},{g}),"Macro",'
                 f'INDEX({M},1):INDEX({M},{g}),INDEX({M},{g}))')
    return (f'=IF(NOT({reg(r)}),"",IF({X("t", r)}="Macro",COUNTIFS({col_upto("t", r)},"Macro",{col_upto("mes", r)},{X("mes", r)}),'
            f'IF(N({X("grupo", r)})=0,"?",{macro_num}&"."&COUNTIFS({col_upto("grupo", r)},{X("grupo", r)},{col_upto("t", r)},"Sub"))))')


def checks(r):
    act, ini, lim, cump, peso = (X(k, r) for k in ("act", "ini", "lim", "cump", "peso"))
    s1, s2, s3, s4, av, obs = (X(k, r) for k in ("s1", "s2", "s3", "s4", "av", "obs"))
    t, grupo, nsub = X("t", r), X("grupo", r), X("nsub", r)
    macro, sub = f'{t}="Macro"', f'{t}="Sub"'
    return [
        (f'{act}=""', "Falta el nombre de la actividad"),
        (f'AND({sub},N({grupo})=0)', "Subactividad sin macroactividad arriba"),
        (f'{ini}=""', "Sin fecha de inicio"),
        (f'{lim}=""', "Sin fecha límite"),
        (f'OR(AND({ini}<>"",NOT(ISNUMBER({ini}))),AND({lim}<>"",NOT(ISNUMBER({lim}))),AND({cump}<>"",NOT(ISNUMBER({cump}))))',
         "Fecha no válida"),
        (f'AND(ISNUMBER({ini}),ISNUMBER({lim}),{ini}>{lim})', "Inicio posterior a fecha límite"),
        (f'AND({macro},{peso}<>"",OR(NOT(ISNUMBER({peso})),{peso}<0,{peso}>1))', "Peso fuera de rango"),
        (f'AND({sub},{peso}<>"")', "Las subactividades no llevan peso (valen igual dentro de su macroactividad)"),
        (f'AND(N({nsub})>0,COUNT({s1}:{s4})>0)', "Con subactividades: deje vacías las semanas (se calculan solas)"),
        (f'(COUNTIF({s1}:{s4},">1")+COUNTIF({s1}:{s4},"<0")+COUNTA({s1}:{s4})-COUNT({s1}:{s4}))>0',
         "Avance fuera de rango 0-100%"),
        (f'OR(AND({s2}<>"",{s2}<{s1}),AND({s3}<>"",{s3}<MAX({s1}:{s2})),AND({s4}<>"",{s4}<MAX({s1}:{s3})))',
         "Avance semanal decreciente"),
        (f'AND(ISNUMBER({lim}),{lim}<FechaRef,N({av})<1)', "Vencida"),
        (f'AND(ObsObligatoria="Sí",ISNUMBER({lim}),{lim}<FechaRef,N({av})<1,{obs}="")',
         "Vencida sin observación"),
    ]


def f_validacion(r):
    parts = []
    for cond, msg in checks(r):
        m = msg if msg.startswith('"') else f'"{msg}"'
        parts.append(f'IF({cond},{m}&" · ","")')
    s = "&".join(parts)
    return (f'=IF(NOT({reg(r)}),"",IF(({s})="","✔ OK",'
            f'"⚠ "&LEFT({s},LEN({s})-3)))')


# --- Delegación (solo hoja Administración) ---
def _por_puesto(r, fn):
    """CHOOSE según el puesto elegido en «Delegar a»: fn(hoja) devuelve la expresión para esa hoja."""
    opciones = ",".join(fn(q(a[0]) + "!") for a in AREAS)
    return f'CHOOSE(MATCH({X("deleg", r)},ListaPuestos,0),{opciones})'


def _delegada(r):
    return f'AND({X("deleg", r)}<>"",{X("t", r)}="Macro")'


def f_dfila(r):
    """Fila (1…) de la actividad delegada dentro del Cronograma del puesto; 0 = el puesto aún no la agregó."""
    clave = f'"{PREFIJO_JEF}"&{X("act", r)}'
    return (f'=IF(NOT({_delegada(r)}),"",IFERROR('
            f'{_por_puesto(r, lambda sh: f"MATCH({clave},{sh}{col_abs('act')},0)")},0))')


def f_av_adm(r):
    val = _por_puesto(r, lambda sh: f'INDEX({sh}{col_abs("av")},{X("dfila", r)})')
    return (f'=IF(NOT({reg(r)}),"",IF({_delegada(r)},IF(N({X("dfila", r)})>0,N({val}),0),{_avance_core(r)}))')


def f_alerta_adm(r):
    val = _por_puesto(r, lambda sh: f'INDEX({sh}{col_abs("alerta")},{X("dfila", r)})')
    return (f'=IF(NOT({reg(r)}),"",IF(AND({_delegada(r)},N({X("dfila", r)})>0),{val}&"",{_alerta_core(r)}))')


def f_seg_adm(r):
    d = X("deleg", r)
    return (f'=IF(NOT({reg(r)}),"",IF({d}="","Propia",IF({X("t", r)}<>"Macro","Solo se delegan macroactividades",'
            f'IF(N({X("dfila", r)})>0,"✔ Recibida por "&{d},"⚠ "&{d}&" aún no la agrega"))))')


def f_dn(r):
    m = X("mes", r)
    return (f'=IF(AND({_delegada(r)},{m}=MesCiclo),COUNTIFS({col_upto("deleg", r)},{X("deleg", r)},'
            f'{col_upto("mes", r)},MesCiclo,{col_upto("t", r)},"Macro"),"")')


def f_dkey(r):
    return f'=IF({X("dn", r)}="","",{X("deleg", r)}&"|"&{X("dn", r)})'


ADMIN_FUNCS = {"av": f_av_adm, "alerta": f_alerta_adm, "seg": f_seg_adm,
               "dfila": f_dfila, "dn": f_dn, "dkey": f_dkey}


# Columnas calculadas (clave → función) visibles y auxiliares
CALC_FUNCS = [("id", f_id), ("av", f_avance), ("est", f_estado), ("alerta", f_alerta),
              ("rev", f_validacion), ("mes", f_mes)]
AUX_FUNCS = [("orden", f_orden), ("grupo", f_grupo), ("t", f_tipo), ("nsub", f_nsub), ("ap", f_ap)] + \
    [(f"o{w + 1}", (lambda w: (lambda r: f_own(r, w)))(w)) for w in range(4)] + \
    [(f"e{w + 1}", (lambda w: (lambda r: f_eff(r, w)))(w)) for w in range(4)] + [("pef", f_pef)]


# Traducción a referencias estructuradas para la definición de columna calculada de la
# tabla (permite que Excel extienda las fórmulas al agregar filas nuevas).
def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def to_structured(formula, table, r=FIRST):
    f = formula.lstrip("=")
    name = lambda c: COLS_ACT[_col_index(c)]
    L = "[A-Z]{1,2}"
    f = re.sub(rf"\$({L})\${FIRST}:\$?({L}){r}(?!\d)",
               lambda m: f"INDEX({table}[{name(m.group(1))}],1):{table}[[#This Row],[{name(m.group(2))}]]", f)
    f = re.sub(rf"(?<![A-Za-z0-9_\]])\$?({L}){r}:\$?({L}){r}(?!\d)",
               lambda m: f"{table}[[#This Row],[{name(m.group(1))}]:[{name(m.group(2))}]]", f)
    f = re.sub(rf"(?<![\$A-Za-z0-9_:\]])\$?({L}){r}(?!\d)",
               lambda m: f"{table}[[#This Row],[{name(m.group(1))}]]", f)
    return f


# --------------------------------------------------------------------------
# Construcción
# --------------------------------------------------------------------------
def build(path, test_rows=None, semana=None, fecha_ref=None, extra_rows=0,
          inicio_ciclo=None, mes=None, obs="Sí"):
    """semana/fecha_ref/inicio_ciclo/mes = None → fórmulas automáticas (entregable)."""
    wb = Workbook()
    ws_d = wb.active
    ws_d.title = S_DASH
    ws_adm = wb.create_sheet(ADMIN[1])
    ws_adm.sheet_properties.tabColor = ADMIN[3]
    ws_hist = wb.create_sheet(S_HIST)
    area_ws, vista_ws = [], []
    for a in AREAS:
        area_ws.append(wb.create_sheet(a[0]))
        vista_ws.append(wb.create_sheet(a[1]))
        area_ws[-1].sheet_properties.tabColor = a[3]
        vista_ws[-1].sheet_properties.tabColor = a[3]
    ws_i = wb.create_sheet(S_INST)
    ws_k = wb.create_sheet(S_CONF)
    ws_x = wb.create_sheet(S_CALC)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 90

    build_config(wb, ws_k, semana, fecha_ref, inicio_ciclo, mes, obs)
    test_rows = test_rows or []
    for i, ws_a in enumerate(area_ws):
        rows_i = [t for t in test_rows if t.get("area") == PUESTOS[i]]
        build_cronograma(wb, ws_a, i, rows_i, extra_rows if any("_row" in t for t in rows_i) else 0)
    adm_rows = [t for t in test_rows if t.get("area") == ADMIN[0]]
    build_cronograma(wb, ws_adm, None, adm_rows, 0, area=ADMIN)
    for i, ws_h in enumerate(vista_ws):
        build_vista(wb, ws_h, i)
    build_calculos(wb, ws_x)
    build_dashboard(wb, ws_d)
    build_historial(wb, ws_hist)
    build_instrucciones(ws_i)

    ws_x.sheet_state = "hidden"
    wb.active = 0
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.save(path)
    return path


# --------------------------------------------------------------------------
def build_config(wb, ws, semana, fecha_ref, inicio_ciclo, mes, obs):
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 70

    merge_set(ws, "B1:D1", "Configuración del ciclo mensual",
              font=font(14, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[1].height = 28
    ws["B2"] = ("Todo se calcula solo a partir de la fecha de hoy. Celdas GRIS SUAVE = puede escribir encima si necesita "
                "fijar otro valor; celdas AZUL CLARO = automáticas y bloqueadas.")
    ws["B2"].font = font(9, italic=True)

    hdr = ["Parámetro", "Valor", "Uso / nota"]
    for i, h in enumerate(hdr):
        c = ws.cell(row=3, column=2 + i, value=h)
        c.font = font(10, True, "FFFFFF"); c.fill = fill(MID); c.alignment = CENTER; c.border = BORDER

    # (etiqueta, valor, nota, formato, nombre, editable)
    params = [
        ("Mes que se muestra", mes if mes else "=DATE(YEAR(TODAY()),MONTH(TODAY()),1)",
         "AUTOMÁTICO: mes actual. Para revisar un mes anterior (historial) escriba cualquier fecha de ese mes "
         "(ej. 01/09/2026); para volver al mes actual escriba =FECHA(AÑO(HOY());MES(HOY());1).",
         MES_FMT, "MesEntrada", True),
        ("Fecha de inicio del ciclo (lunes de la Semana 1)",
         inicio_ciclo if inicio_ciclo else "=MesCiclo+MOD(2-WEEKDAY(MesCiclo),7)",
         "AUTOMÁTICO: primer lunes del mes. Puede escribir otra fecha. Define en qué semana cae cada actividad en las hojas «Actividades».",
         "dd/mm/yyyy", "InicioCiclo", True),
        ("Fecha de corte", fecha_ref if fecha_ref else "=TODAY()",
         "AUTOMÁTICO: hoy. Puede escribir una fecha fija para emitir un reporte de corte. Determina actividades vencidas.",
         "dd/mm/yyyy", "FechaRef", True),
        ("Semana actual del ciclo (1 a 4)",
         "=IF(SemanaManual<>\"\",SemanaManual,MIN(4,MAX(1,INT((FechaRef-InicioCiclo)/7)+1)))",
         "AUTOMÁTICA: se calcula con la fecha de inicio y la fecha de corte. Controla los mensajes de finalización anticipada.",
         '"Semana "0', "SemanaActual", False),
        ("Semana manual (opcional)", semana,
         "Déjela VACÍA para usar la semana automática. Escriba 1–4 solo si necesita forzar otra semana.",
         "0", "SemanaManual", True),
        ("Días de anticipación para alerta «próxima a vencer»", 3,
         "Actividades no completadas cuya fecha límite cae dentro de este número de días se señalan en ámbar.",
         "0", "DiasAlerta", True),
        ("¿Observación obligatoria si una actividad vence sin cumplirse?", obs,
         "Si es «Sí», una actividad vencida sin texto en Observaciones genera alerta (explicar por qué no se cumplió).",
         "@", "ObsObligatoria", True),
        ("Meta mensual", 1, "Meta de cumplimiento de las actividades al cierre de la Semana 4 de cada mes.", "0%", "MetaFinal", False),
        ("Primer mes del historial", dt.date(2026, 10, 1),
         "El Historial mensual empieza en este mes (octubre 2026) y muestra 24 meses seguidos.",
         MES_FMT, "HistInicio", True),
    ]
    for i, (lbl, val, note, fmt, nm, editable) in enumerate(params):
        r = 4 + i
        ws.cell(row=r, column=2, value=lbl).font = font(10, True)
        v = ws.cell(row=r, column=3, value=val)
        v.number_format = fmt; v.font = font(10, True, NAVY); v.fill = fill(INPUT if editable else CALC)
        v.alignment = CENTER; v.protection = Protection(locked=not editable)
        ws.cell(row=r, column=4, value=note).font = font(9)
        for col in range(2, 5):
            ws.cell(row=r, column=col).border = BORDER
            if col != 3:
                ws.cell(row=r, column=col).alignment = LEFT
        ws.row_dimensions[r].height = 30
        add_name(wb, nm, f"{q(S_CONF)}!$C${r}")

    dvd = DataValidation(type="date", operator="greaterThan", formula1="36526",
                         showErrorMessage=True, error="Ingrese una fecha válida (dd/mm/aaaa).")
    for c in ("C4", "C5", "C6", "C12"):
        dvd.add(c)
    ws.add_data_validation(dvd)
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="4", allow_blank=True,
                        showErrorMessage=True, errorTitle="Semana no válida",
                        error="Ingrese un número entero entre 1 y 4, o deje vacío para la semana automática.")
    dv.add("C8"); ws.add_data_validation(dv)
    dvn = DataValidation(type="whole", operator="between", formula1="0", formula2="30",
                         showErrorMessage=True, error="Ingrese un número de días entre 0 y 30.")
    dvn.add("C9"); ws.add_data_validation(dvn)
    dvs = DataValidation(type="list", formula1="ListaSiNo", showErrorMessage=True)
    dvs.add("C10"); ws.add_data_validation(dvs)

    # Plan lineal de referencia
    r0 = 14
    merge_set(ws, f"B{r0}:D{r0}", "Plan de referencia (avance esperado acumulado por semana) – supuesto editable",
              font=font(10, True, "FFFFFF"), fill=fill(MID), alignment=LEFT)
    for i in range(4):
        r = r0 + 1 + i
        ws.cell(row=r, column=2, value=f"Semana {i + 1}").font = font(10, True)
        v = ws.cell(row=r, column=3, value=(i + 1) * 0.25)
        v.number_format = "0%"; v.fill = fill(INPUT); v.font = font(10, True, NAVY)
        v.alignment = CENTER; v.protection = Protection(locked=False)
        ws.cell(row=r, column=4, value="Referencia lineal solo para comparación visual; no afecta el cálculo del avance real.").font = font(9)
        for col in range(2, 5):
            ws.cell(row=r, column=col).border = BORDER
    add_name(wb, "PlanRef", f"{q(S_CONF)}!$C${r0 + 1}:$C${r0 + 4}")

    # Listas
    r1 = 20
    merge_set(ws, f"B{r1}:D{r1}", "Listas desplegables (no editar sin actualizar las fórmulas)",
              font=font(10, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    lists = [("Áreas / Puestos", PUESTOS, "ListaPuestos"),
             ("Estados", ESTADOS, "ListaEstados"),
             ("Semanas", SEMANAS, "ListaSemanas"),
             ("Sí / No", ["Sí", "No"], "ListaSiNo")]
    r = r1 + 1
    for title, items, nm in lists:
        ws.cell(row=r, column=2, value=title).font = font(10, True, MID)
        r += 1
        start = r
        for it in items:
            c = ws.cell(row=r, column=2, value=it)
            c.font = font(10); c.border = BORDER
            r += 1
        add_name(wb, nm, f"{q(S_CONF)}!$B${start}:$B${r - 1}")
        r += 1

    # Metodología
    r += 1
    merge_set(ws, f"B{r}:D{r}", "Metodología de cálculo (documentación)",
              font=font(10, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    notes = [
        "1. Avance de una actividad = último avance semanal acumulado registrado (Semana 4 → 1), ignorando semanas vacías. Los avances semanales NO se suman.",
        "2. Estado automático: 0% = Pendiente; >0% y <100% = En curso; 100% = Completado.",
        "3. Avance ponderado del puesto = Σ (peso efectivo × avance actual) de las macroactividades del mes. Es válido si el puesto tiene ≥1 macroactividad en el mes y ningún peso fuera de 0–100%.",
        "0. Cada actividad pertenece al mes de su fecha de inicio (columna «Mes» del Cronograma). Dashboard, hojas Actividades y cálculos usan solo el mes mostrado; los meses anteriores quedan en el Historial mensual.",
        "0b. Subactividades: filas con Tipo «Subactividad» pertenecen a la macroactividad más cercana hacia arriba y valen igual. Avance de la macroactividad = promedio de sus subactividades (semana por semana). Una fila con Fecha de cumplimiento cuenta 100% desde la semana en que se cumplió. Peso efectivo de cada macroactividad = su peso / suma de pesos del mes (vacío = partes iguales), por lo que siempre suma 100%.",
        "4. Avance general de las actividades = promedio simple de los 4 puestos (cada puesto pesa 25%, sin importar cuántas actividades tenga). Si algún puesto no es válido, el Dashboard muestra «No definitivo» y un valor provisional rotulado como tal.",
        "5. Avance semanal acumulado (Semana n) = Σ peso × último avance registrado hasta la Semana n. Incremento semanal = acumulado Semana n − acumulado Semana n−1.",
        "6. «¡Enhorabuena, completado anticipadamente!»: el puesto (o el total de actividades del mes) es válido, llega al 100% y la semana en que lo alcanzó (la menor entre la primera semana con 100% registrada y la Semana actual) es anterior a la Semana 4. Si lo alcanza en la Semana 4: «Completado». Si no: «En ejecución».",
        "7. Una semana se muestra en el seguimiento si es ≤ Semana actual o si ya tiene algún avance registrado.",
        "8. Cada puesto tiene su propia hoja «Cronograma» y su hoja «Actividades» semanales. El Dashboard consolida los cuatro cronogramas (filas 5 a 1000) y las cuatro hojas de actividades (filas 5 a 500).",
        "9. Semana actual = ENTERO((fecha de corte − fecha de inicio) / 7) + 1, limitada entre 1 y 4, salvo que se escriba una semana manual.",
        "10. Hojas «Actividades»: se llenan solas desde el Cronograma del puesto. Semana de entrega = semana en que cae la fecha límite. Estado: Completado (100%), Vencida (pasó la fecha límite sin llegar a 100%), En curso o Pendiente.",
    ]
    for n in notes:
        r += 1
        merge_set(ws, f"B{r}:D{r}", n, font=font(9), alignment=LEFT_TOP)
        ws.row_dimensions[r].height = 30

    ws.freeze_panes = "A4"
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# --------------------------------------------------------------------------
def protect_and_print(ws, hr, last_col, last):
    ws.protection.sheet = True
    ws.protection.autoFilter = False
    ws.protection.sort = False
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{hr}:{hr}"
    ws.print_area = f"A1:{last_col}{last}"
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.oddFooter.center.text = "&A – Página &P de &N"


def tint(hexcolor, k=0.8):
    """Color claro derivado del color del puesto (mezcla con blanco)."""
    rgb = [int(hexcolor[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join(f"{int(c + (255 - c) * k):02X}" for c in rgb)


AUX_FIRST = CN["orden"]   # primera columna auxiliar oculta
N_COLS = len(COLS_ACT)
LASTV = C["mes"]          # última columna visible
SUB_NUMFMT = '"      ↳  "@'   # sangría visual de las subactividades (formato condicional)


def alerta_cf(cf, rng, first_cell):
    """Colores de la columna «Alerta de plazo»."""
    cf.add(rng, FormulaRule(formula=[f'LEFT({first_cell},7)="Vencida"'], stopIfTrue=True,
                            fill=fill("C00000"), font=Font(color="FFFFFF", bold=True)))
    cf.add(rng, FormulaRule(formula=[f'LEFT({first_cell},5)="Vence"'], stopIfTrue=True,
                            fill=fill("FFC000"), font=Font(color="3F2F00", bold=True)))
    cf.add(rng, FormulaRule(formula=[f'LEFT({first_cell},8)="Cumplida"'], stopIfTrue=True,
                            fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(rng, FormulaRule(formula=[f'{first_cell}="En plazo"'], stopIfTrue=True,
                            fill=fill("E2EFDA"), font=Font(color="375623")))
    cf.add(rng, FormulaRule(formula=[f'{first_cell}="Sin fecha límite"'],
                            fill=fill(RED_F), font=Font(color=RED_T)))


def build_cronograma(wb, ws, idx, test_rows, extra_rows, area=None):
    if area is None:
        puesto = PUESTOS[idx]
        sheet, _, code, color = AREAS[idx]
    else:
        puesto, sheet, code, color = area
    es_admin = area is not None
    light = tint(color)
    funcs = CALC_FUNCS + AUX_FUNCS
    if es_admin:
        funcs = [(k, ADMIN_FUNCS.get(k, fn)) for k, fn in funcs] + \
                [(k, ADMIN_FUNCS[k]) for k in ("seg", "dfila", "dn", "dkey")]
    widths = {"id": 7, "tipo": 15, "act": 48, "deleg": 22, "seg": 30, "ini": 12, "lim": 12, "cump": 14, "peso": 12,
              "s1": 10, "s2": 10, "s3": 10, "s4": 10, "av": 12, "est": 12, "alerta": 24,
              "obs": 36, "rev": 44, "mes": 14}
    for k, w in widths.items():
        ws.column_dimensions[C[k]].width = w
    for ci in range(AUX_FIRST, N_COLS + 1):
        ws.column_dimensions[get_column_letter(ci)].hidden = True
    if not es_admin:
        for k in ADMIN_ONLY:
            ws.column_dimensions[C[k]].hidden = True

    merge_set(ws, f"A1:{LASTV}1", f"Cronograma – {puesto}" if area is None else
              f"{puesto} – Actividades de la jefatura",
              font=font(14, True, "FFFFFF"), fill=fill(color), alignment=LEFT)
    ws.row_dimensions[1].height = 28
    merge_set(ws, f"A2:{LASTV}2",
              "Llene solo las celdas GRIS SUAVE; las AZUL CLARO son automáticas. En «Tipo» elija Macroactividad o Subactividad: "
              "las subactividades van en las filas de ABAJO de su macroactividad (se ven con ↳) y su promedio da el % de la macroactividad. "
              "Cuando termine algo, escriba la FECHA DE CUMPLIMIENTO: queda al 100% y ya no aparece vencido. "
              "Si quiere registrar avances parciales, escriba en cada semana el % total logrado hasta esa semana. "
              "Cada mes agregue sus actividades debajo de las anteriores (no borre los meses pasados: quedan como historial).",
              font=font(9, False, NAVY), fill=fill(LIGHT), alignment=LEFT)
    ws.row_dimensions[2].height = 44
    ws["A3"] = "Llenar"; ws["A3"].fill = fill(INPUT); ws["A3"].font = font(8); ws["A3"].border = BORDER
    ws["B3"] = "Automático"; ws["B3"].fill = fill(CALC); ws["B3"].font = font(8); ws["B3"].border = BORDER
    ws["C3"] = "Macroactividad (fila resaltada)"; ws["C3"].fill = fill(light); ws["C3"].font = font(8, True); ws["C3"].border = BORDER
    merge_set(ws, f"{C['ini']}3:{C['s2']}3",
              "Peso: solo macroactividades. Si lo deja vacío valen igual; si lo escribe, se ajusta solo a 100%.",
              font=font(8, italic=True, color=NAVY), alignment=LEFT)
    # Resumen de alertas de plazo del mes mostrado
    ws[f"{C['s3']}3"] = "Plazos del mes:"; ws.merge_cells(f"{C['s3']}3:{C['s4']}3")
    ws[f"{C['s3']}3"].font = font(9, True, NAVY)
    ws[f"{C['s3']}3"].alignment = Alignment(horizontal="right", vertical="center")
    venc = f'COUNTIFS({col_abs("mes")},MesCiclo,{col_abs("alerta")},"Vencida*")'
    porv = f'COUNTIFS({col_abs("mes")},MesCiclo,{col_abs("alerta")},"Vence*")'
    rng3 = f"{C['av']}3:{C['alerta']}3"
    merge_set(ws, rng3, f'={venc}&" vencida(s)  ·  "&{porv}&" por vencer"',
              font=font(10, True), alignment=CENTER, border=BORDER)
    cf3 = ws.conditional_formatting
    cf3.add(rng3, FormulaRule(formula=[f'{venc}>0'], stopIfTrue=True, fill=fill("C00000"), font=Font(color="FFFFFF", bold=True)))
    cf3.add(rng3, FormulaRule(formula=[f'{porv}>0'], stopIfTrue=True, fill=fill("FFC000"), font=Font(color="3F2F00", bold=True)))
    cf3.add(rng3, FormulaRule(formula=['TRUE'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    rng_b = f"{C['obs']}3:{LASTV}3"
    if es_admin:
        merge_set(ws, rng_b, "En «Delegar a» elija el puesto: la actividad le llega a su Cronograma y aquí verá su avance.",
                  font=font(9, True, NAVY), fill=fill(LIGHT), alignment=LEFT, border=BORDER)
    else:
        adm = lambda k: f"{q(ADMIN[1])}!{col_abs(k)}"
        n_del = f'COUNTIFS({adm("deleg")},"{puesto}",{adm("mes")},MesCiclo,{adm("t")},"Macro")'
        n_rec = f'COUNTIFS({adm("deleg")},"{puesto}",{adm("mes")},MesCiclo,{adm("t")},"Macro",{adm("dfila")},">0")'
        merge_set(ws, rng_b,
                  f'=IF({n_del}=0,"Sin actividades delegadas por la jefatura este mes",'
                  f'"★ La jefatura le delegó "&{n_del}&" actividad(es) este mes · "&({n_del}-{n_rec})&'
                  f'" pendiente(s) de agregar (elíjalas en la lista de «Actividad»)")',
                  font=font(9, True), alignment=LEFT, border=BORDER)
        cf3.add(rng_b, FormulaRule(formula=[f'({n_del}-{n_rec})>0'], stopIfTrue=True,
                                   fill=fill("FFC000"), font=Font(color="3F2F00", bold=True)))
        cf3.add(rng_b, FormulaRule(formula=[f'{n_del}>0'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    ws.row_dimensions[3].height = 22

    hr = FIRST - 1
    calc_cols = {CN[k] for k, _ in CALC_FUNCS} | {CN["seg"]} | set(range(AUX_FIRST, N_COLS + 1))
    for i, h in enumerate(COLS_ACT):
        c = ws.cell(row=hr, column=i + 1, value=h)
        c.font = font(10, True, "FFFFFF")
        c.fill = fill(color if (i + 1) not in calc_cols else NAVY)
        c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[hr].height = 42

    last = FIRST + N_ACT_ROWS - 1 + extra_rows
    fmt = {CN["ini"]: "dd/mm/yyyy", CN["lim"]: "dd/mm/yyyy", CN["cump"]: "dd/mm/yyyy", CN["peso"]: "0%",
           CN["s1"]: "0%", CN["s2"]: "0%", CN["s3"]: "0%", CN["s4"]: "0%", CN["av"]: "0%", CN["mes"]: MES_FMT}
    centered = set(fmt) | {CN["id"], CN["tipo"], CN["est"], CN["alerta"]}
    for r in range(FIRST, last + 1):
        for col in range(1, N_COLS + 1):
            c = ws.cell(row=r, column=col)
            c.border = BORDER
            c.font = font(10)
            locked = col in calc_cols
            c.fill = fill(CALC if locked else INPUT)
            c.protection = Protection(locked=locked)
            if col in fmt:
                c.number_format = fmt[col]
            c.alignment = CENTER if col in centered else LEFT
        for k, fn in funcs:
            ws.cell(row=r, column=CN[k], value=fn(r))
        ws.cell(row=r, column=CN["rev"]).font = font(9)

    # Datos de prueba (solo en ejecuciones de test)
    for i, row in enumerate(test_rows):
        r = row.get("_row", FIRST + i)
        if "sub" in row:
            ws.cell(row=r, column=CN["tipo"], value="Subactividad")
            ws.cell(row=r, column=CN["act"], value=row["sub"])
        for key, k in [("macro", "act"), ("deleg", "deleg"), ("ini", "ini"), ("fin", "lim"), ("cump", "cump"), ("peso", "peso"),
                       ("s1", "s1"), ("s2", "s2"), ("s3", "s3"), ("s4", "s4"), ("obs", "obs")]:
            if key in row and row[key] is not None:
                ws.cell(row=r, column=CN[k], value=row[key])

    # Tabla estructurada (una por puesto) con columnas calculadas
    tname = f"tblAct_{code}"
    tab = Table(displayName=tname, ref=f"A{hr}:{get_column_letter(N_COLS)}{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    tab._initialise_columns()
    for i, col in enumerate(tab.tableColumns):
        col.name = COLS_ACT[i]
    for k, fn in funcs:
        tab.tableColumns[CN[k] - 1].calculatedColumnFormula = TableFormula(
            attr_text=to_structured(fn(FIRST), tname))
    ws.add_table(tab)

    # Nombres de rango por puesto para Dashboard / Cálculos
    names = [("rFLim", "lim"), ("rPeso", "pef"), ("rPesoEscrito", "peso"), ("rS1", "s1"), ("rS2", "s2"),
             ("rS3", "s3"), ("rS4", "s4"), ("rAvance", "av"), ("rEstado", "est"), ("rAlerta", "alerta"),
             ("rValid", "rev"), ("rMes", "mes"), ("rOrden", "orden"), ("rTipo", "t"),
             ("rEf1", "e1"), ("rEf2", "e2"), ("rEf3", "e3"), ("rEf4", "e4"),
             ("rAct", "act"), ("rDeleg", "deleg"), ("rDkey", "dkey")]
    for nm, k in names:
        add_name(wb, f"{nm}_{code}", f"{q(sheet)}!{col_abs(k)}")

    # Validaciones de datos
    R = lambda k: f"{C[k]}{FIRST}:{C[k]}{LAST_RNG}"
    dv_tipo = DataValidation(type="list", formula1='"Macroactividad,Subactividad"', allow_blank=True,
                             showErrorMessage=True, errorTitle="Tipo no válido",
                             error="Elija Macroactividad o Subactividad.",
                             showInputMessage=True, promptTitle="Tipo",
                             prompt="Macroactividad (lleva peso) o Subactividad (va debajo de su macroactividad). Vacío = Macroactividad.")
    dv_tipo.add(R("tipo"))
    dv_date = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                             showErrorMessage=True, errorTitle="Fecha no válida",
                             error="Ingrese una fecha válida (dd/mm/aaaa).")
    for k in ("ini", "lim"):
        dv_date.add(R(k))
    dv_cump = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True,
                             showErrorMessage=True, errorTitle="Fecha no válida",
                             error="Ingrese una fecha válida (dd/mm/aaaa).",
                             showInputMessage=True, promptTitle="Fecha de cumplimiento",
                             prompt="Escriba la fecha en que se cumplió. La actividad queda al 100% y ya no aparece vencida.")
    dv_cump.add(R("cump"))
    dv_peso = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                             allow_blank=True, showErrorMessage=True, errorTitle="Peso no válido",
                             error="Ingrese un porcentaje entre 0% y 100% (por ejemplo 40%).",
                             showInputMessage=True, promptTitle="Peso (solo macroactividades)",
                             prompt="Opcional. Vacío = todas las macroactividades del mes valen igual. Si escribe pesos, se ajustan solos a 100%. Las subactividades no llevan peso.")
    dv_peso.add(R("peso"))
    dv_av = DataValidation(type="decimal", operator="between", formula1="0", formula2="1",
                           allow_blank=True, showErrorMessage=True, errorTitle="Avance no válido",
                           error="Ingrese un porcentaje entre 0% y 100%. Deje la celda vacía si la semana no ha llegado.",
                           showInputMessage=True, promptTitle="Avance parcial (opcional)",
                           prompt="% TOTAL logrado hasta esta semana. Cuando termine, escriba la Fecha de cumplimiento. Si la macroactividad tiene subactividades, déjelo vacío.")
    dv_av.add(f"{C['s1']}{FIRST}:{C['s4']}{LAST_RNG}")
    for d in (dv_tipo, dv_date, dv_cump, dv_peso, dv_av):
        ws.add_data_validation(d)
    if es_admin:
        dv_del = DataValidation(type="list", formula1="ListaPuestos", allow_blank=True, showErrorMessage=True,
                                errorTitle="Puesto no válido", error="Elija un puesto de la lista o deje vacío (actividad de la jefatura).",
                                showInputMessage=True, promptTitle="Delegar a",
                                prompt="Elija el puesto al que delega esta macroactividad. Vacío = la hace la jefatura.")
        dv_del.add(R("deleg"))
        ws.add_data_validation(dv_del)
    else:
        # Lista de actividades delegadas por la jefatura (también se puede escribir cualquier otro nombre)
        dv_act = DataValidation(type="list", formula1=f"ListaDeleg_{code}", allow_blank=True, showErrorMessage=False,
                                showInputMessage=True, promptTitle="Actividad",
                                prompt="Escriba el nombre, o elija en la lista una actividad delegada por la jefatura.")
        dv_act.add(R("act"))
        ws.add_data_validation(dv_act)

    # Formato condicional
    F_ = FIRST
    x = lambda k: f"${C[k]}{F_}"
    regf = reg(F_)
    red = dict(fill=fill(RED_F), font=Font(color=RED_T))
    amb = dict(fill=fill(AMBER_F), font=Font(color=AMBER_T))
    org = dict(fill=fill(ORANGE_F), font=Font(color="833C0B", bold=True))
    cf = ws.conditional_formatting
    macro_line = Border(top=Side(style="medium", color=color))
    ABC = f"A{F_}:{C['act']}{LAST_RNG}"
    cf.add(R("act"), FormulaRule(stopIfTrue=True, formula=[f'LEFT({x("act")},{len(PREFIJO_JEF)})="{PREFIJO_JEF}"'],
                                 fill=fill("FFF2CC"), font=Font(bold=True, color="7F6000")))
    cf.add(ABC, FormulaRule(formula=[f'AND({x("t")}="Macro",{x("act")}="")'], stopIfTrue=True, **red))
    cf.add(ABC, FormulaRule(formula=[f'{x("t")}="Macro"'], fill=fill(light), font=Font(bold=True), border=macro_line))
    cf.add(R("act"), Rule(type="expression", formula=[f'{x("t")}="Sub"'],
                          dxf=DifferentialStyle(numFmt=NumberFormat(numFmtId=200, formatCode=SUB_NUMFMT),
                                                font=Font(color="404040"))))
    if es_admin:
        cf.add(R("seg"), FormulaRule(formula=[f'LEFT({x("seg")},1)="✔"'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
        cf.add(R("seg"), FormulaRule(formula=[f'LEFT({x("seg")},1)="⚠"'], fill=fill(RED_F), font=Font(color=RED_T, bold=True)))
    cf.add(R("est"), FormulaRule(formula=[f'{x("est")}="Completado"'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(R("est"), FormulaRule(formula=[f'{x("est")}="En curso"'], fill=fill(BLUE_F), font=Font(color=BLUE_T, bold=True)))
    cf.add(R("est"), FormulaRule(formula=[f'{x("est")}="Pendiente"'], fill=fill(RED_F), font=Font(color=RED_T)))
    alerta_cf(cf, R("alerta"), x("alerta"))
    cf.add(R("av"), FormulaRule(formula=[f'AND(ISNUMBER({x("av")}),{x("av")}>=1)'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(R("av"), DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5", showValue=True))
    cf.add(R("cump"), FormulaRule(formula=[f'ISNUMBER({x("cump")})'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(R("lim"), FormulaRule(formula=[f'AND(ISNUMBER({x("lim")}),{x("lim")}<FechaRef,N({x("av")})<1)'], stopIfTrue=True, **org))
    cf.add(R("lim"), FormulaRule(formula=[f'AND(ISNUMBER({x("lim")}),{x("lim")}>=FechaRef,{x("lim")}-FechaRef<=DiasAlerta,N({x("av")})<1)'], **amb))
    DE = f"{C['ini']}{F_}:{C['lim']}{LAST_RNG}"
    cf.add(DE, FormulaRule(formula=[f'AND({regf},{C["ini"]}{F_}="")'], stopIfTrue=True, **red))
    cf.add(DE, FormulaRule(formula=[f'AND(ISNUMBER({x("ini")}),ISNUMBER({x("lim")}),{x("ini")}>{x("lim")})'], **org))
    cf.add(R("peso"), FormulaRule(formula=[f'AND({x("t")}="Sub",{x("peso")}<>"")'], stopIfTrue=True, **red))
    cf.add(R("peso"), FormulaRule(formula=[f'{x("t")}="Sub"'], stopIfTrue=True, fill=fill(CALC)))
    WKR = f"{C['s1']}{F_}:{C['s4']}{LAST_RNG}"
    s1 = f"{C['s1']}{F_}"
    cf.add(WKR, FormulaRule(formula=[f'AND(N({x("nsub")})>0,{s1}<>"")'], stopIfTrue=True, **red))
    cf.add(WKR, FormulaRule(formula=[f'OR(N({x("nsub")})>0,ISNUMBER({x("cump")}))'], stopIfTrue=True, fill=fill(CALC)))
    cf.add(WKR, FormulaRule(formula=[f'AND({s1}<>"",OR(NOT(ISNUMBER({s1})),{s1}>1,{s1}<0))'], stopIfTrue=True, **red))
    cf.add(f"{C['s2']}{F_}:{C['s4']}{LAST_RNG}",
           FormulaRule(formula=[f'AND({C["s2"]}{F_}<>"",{C["s2"]}{F_}<MAX(${C["s1"]}{F_}:{C["s1"]}{F_}))'], **org))
    cf.add(R("obs"), FormulaRule(formula=[f'AND(ObsObligatoria="Sí",ISNUMBER({x("lim")}),{x("lim")}<FechaRef,N({x("av")})<1,{x("obs")}="")'], **red))
    cf.add(R("rev"), FormulaRule(formula=[f'LEFT({x("rev")},1)="⚠"'], fill=fill(RED_F), font=Font(color=RED_T)))
    cf.add(R("rev"), FormulaRule(formula=[f'LEFT({x("rev")},1)="✔"'], font=Font(color=GREEN_T, bold=True)))
    cf.add(R("mes"), FormulaRule(formula=[f'AND({x("mes")}<>"",{x("mes")}=MesCiclo)'], fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))

    ws.freeze_panes = f"D{FIRST}"
    protect_and_print(ws, hr, LASTV, last)
    return last


# --------------------------------------------------------------------------
# Hoja «Actividades» por puesto: vista semanal AUTOMÁTICA del Cronograma.
# Muestra, en orden y sin huecos, las actividades del Cronograma cuyo Mes = mes mostrado
# (columna auxiliar oculta N = fila de origen). No se escribe nada en esta hoja.
# (ver build_vista para las columnas de la hoja Actividades)
#           F Semana de entrega · G–J Semana 1–4 (avance + barra) · K % Avance · L Estado · M Observaciones
# --------------------------------------------------------------------------
def week_start(w):
    return f"(InicioCiclo+{7 * (w - 1)})"


def build_vista(wb, ws, idx):
    puesto = PUESTOS[idx]
    sheet, vsheet, code, color = AREAS[idx]
    src = q(sheet) + "!"
    # Columnas: A N° · B Actividad · C Inicio · D Límite · E Semana de entrega · F–I Semana 1–4
    #           J % Avance · K Estado · L Alerta de plazo · M Observaciones  (N–P auxiliares ocultas)
    widths = [7, 50, 12, 12, 16, 11, 11, 11, 11, 12, 13, 20, 46]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 1)].width = w
    merge_set(ws, "A1:M1", f"Actividades del mes por semana – {puesto}",
              font=font(14, True, "FFFFFF"), fill=fill(color), alignment=LEFT)
    ws.row_dimensions[1].height = 28
    merge_set(ws, "A2:M2",
              f"HOJA AUTOMÁTICA Y BLOQUEADA: muestra solo las actividades del mes indicado abajo, tomadas de «{sheet}». No hay que escribir nada aquí. "
              "Cada macroactividad (resaltada) va seguida de sus subactividades (↳). La barra de color marca las semanas en que cada actividad "
              "está programada y dentro aparece el avance de esa semana. La columna «Alerta de plazo» indica lo vencido y lo que está por vencer.",
              font=font(9, True, NAVY), fill=fill(LIGHT), alignment=LEFT)
    ws.row_dimensions[2].height = 34
    # Fila 3: mes mostrado y fechas de cada semana
    ws["A3"] = "Mes:"; ws["A3"].font = font(9, True, NAVY)
    ws["B3"] = "=MesCiclo"; ws["B3"].number_format = MES_FMT; ws["B3"].font = font(10, True, NAVY)
    ws["B3"].fill = fill(LIGHT); ws["B3"].alignment = LEFT
    ws["E3"] = "Semana del:"; ws["E3"].font = font(8, True, NAVY)
    ws["E3"].alignment = Alignment(horizontal="right", vertical="center")
    for w in range(1, 5):
        c = ws.cell(row=3, column=5 + w, value=f"={week_start(w)}")
        c.number_format = "dd/mm"; c.font = font(8, True, NAVY); c.alignment = CENTER
        c.fill = fill(LIGHT); c.border = BORDER

    hr = FIRST - 1
    for i, h in enumerate(COLS_VISTA):
        c = ws.cell(row=hr, column=i + 1, value=h)
        c.font = font(10, True, "FFFFFF"); c.fill = fill(color if 6 <= i + 1 <= 9 else NAVY)
        c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[hr].height = 34

    last = FIRST + N_VISTA_ROWS - 1
    for hc in "NOPQ":
        ws.column_dimensions[hc].hidden = True
    ws["N4"] = "Fila origen (aux)"; ws["O4"] = "Tipo (aux)"; ws["P4"] = "N subactividades (aux)"
    ws["Q4"] = "N° alerta de plazo (aux)"
    for r in range(FIRST, last + 1):
        ws[f"N{r}"] = f'=IFERROR(MATCH(ROW()-{FIRST - 1},rOrden_{code},0),"")'
        on = f'$N{r}=""'
        g = lambda col: f"INDEX({src}${col}${FIRST}:${col}${LAST_RNG},$N{r})"
        ws[f"O{r}"] = f'=IF({on},"",{g(C["t"])})'          # tipo (aux)
        ws[f"P{r}"] = f'=IF({on},"",N({g(C["nsub"])}))'       # n° subactividades (aux)
        alerta_k = f'OR(LEFT($L{r},7)="Vencida",LEFT($L{r},5)="Vence")'
        ws[f"Q{r}"] = (f'=IF({alerta_k},SUMPRODUCT(--(LEFT($L${FIRST}:$L{r},7)="Vencida"))'
                       f'+SUMPRODUCT(--(LEFT($L${FIRST}:$L{r},5)="Vence")),"")')   # n° de alerta (aux)
        vals = {
            1: f'=IF({on},"",{g(C["id"])})',
            2: f'=IF({on},"",IF($O{r}="Sub","      ↳  "&{g(C["act"])},{g(C["act"])}&""))',
            3: f'=IF({on},"",IF({g(C["ini"])}="","",{g(C["ini"])}))',
            4: f'=IF({on},"",IF({g(C["lim"])}="","",{g(C["lim"])}))',
            5: (f'=IF({on},"",IF(NOT(ISNUMBER($D{r})),"",IF($D{r}<InicioCiclo,"Antes del ciclo",'
                f'IF($D{r}>=InicioCiclo+28,"Después del ciclo","Semana "&(INT(($D{r}-InicioCiclo)/7)+1)))))'),
            10: f'=IF({on},"",{g(C["av"])})',
            11: (f'=IF({on},"",IF(N($J{r})>=1,"Completado",IF(AND(ISNUMBER($D{r}),$D{r}<FechaRef),"Vencida",'
                 f'IF(N($J{r})>0,"En curso","Pendiente"))))'),
            12: f'=IF({on},"",{g(C["alerta"])}&"")',
            13: f'=IF({on},"",{g(C["obs"])}&"")',
        }
        for w in range(1, 5):
            sc = WK[w - 1]
            ef = EFF[w - 1]
            vals[5 + w] = (f'=IF({on},"",IF($P{r}>0,IF({w}<=SemanaActual,{g(ef)},""),'
                           f'IF({g(sc)}="","",{g(sc)})))')
        for col in range(1, 14):
            c = ws.cell(row=r, column=col, value=vals[col])
            c.border = BORDER; c.font = font(10); c.fill = fill(CALC)
            c.alignment = LEFT if col in (2, 13) else CENTER
            if col in (3, 4):
                c.number_format = "dd/mm/yyyy"
            elif col in (6, 7, 8, 9, 10):
                c.number_format = "0%"

    tname = f"tblSem_{code}"
    tab = Table(displayName=tname, ref=f"A{hr}:M{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(tab)

    for nm, col in [("hSemana", "E"), ("hEstado", "K"), ("hTipo", "O"), ("hAlertaN", "Q"), ("hN", "A"),
                    ("hAct", "B"), ("hLim", "D"), ("hAlerta", "L"), ("hObs", "M")]:
        add_name(wb, f"{nm}_{code}", f"{q(vsheet)}!${col}${FIRST}:${col}${last}")

    cf = ws.conditional_formatting
    rngc = lambda c: f"{c}{FIRST}:{c}{last}"
    # Barra de semanas (Gantt): semana activa entre fecha de inicio y fecha límite
    gantt = (f'AND(ISNUMBER($C{FIRST}),ISNUMBER($D{FIRST}),$C{FIRST}<=InicioCiclo+7*(COLUMN(F{FIRST})-6)+6,'
             f'$D{FIRST}>=InicioCiclo+7*(COLUMN(F{FIRST})-6))')
    cf.add(f"F{FIRST}:I{last}", FormulaRule(formula=[f'AND({gantt},$K{FIRST}="Vencida")'], stopIfTrue=True,
                                           fill=fill(ORANGE_F), font=Font(color="833C0B", bold=True)))
    cf.add(f"F{FIRST}:I{last}", FormulaRule(formula=[gantt], fill=fill(color), font=Font(color="FFFFFF", bold=True)))
    cf.add(f"A{FIRST}:B{last}", FormulaRule(formula=[f'$O{FIRST}="Macro"'], fill=fill(tint(color)), font=Font(bold=True),
                                           border=Border(top=Side(style="medium", color=color))))
    cf.add(rngc("K"), FormulaRule(formula=[f'$K{FIRST}="Completado"'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(rngc("K"), FormulaRule(formula=[f'$K{FIRST}="Vencida"'], fill=fill(ORANGE_F), font=Font(color="833C0B", bold=True)))
    cf.add(rngc("K"), FormulaRule(formula=[f'$K{FIRST}="En curso"'], fill=fill(BLUE_F), font=Font(color=BLUE_T, bold=True)))
    cf.add(rngc("K"), FormulaRule(formula=[f'$K{FIRST}="Pendiente"'], fill=fill(RED_F), font=Font(color=RED_T)))
    alerta_cf(cf, rngc("L"), f"$L{FIRST}")
    cf.add(rngc("E"), FormulaRule(formula=[f'$E{FIRST}="Semana "&SemanaActual'], fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))
    cf.add(rngc("J"), DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="5B9BD5", showValue=True))
    # Semana actual resaltada en el encabezado
    cf.add("F4:I4", FormulaRule(formula=['COLUMN(F4)-5=SemanaActual'], fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))
    cf.add("F3:I3", FormulaRule(formula=['COLUMN(F3)-5=SemanaActual'], fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))

    ws.freeze_panes = f"C{FIRST}"
    protect_and_print(ws, hr, "M", last)


# --------------------------------------------------------------------------
# Hoja oculta de cálculos por puesto
# --------------------------------------------------------------------------
CALC_HDR = ["Puesto", "N° actividades", "Suma de pesos", "Pesos vacíos o inválidos",
            "Configuración válida", "Avance ponderado actual",
            "Acum. Semana 1", "Acum. Semana 2", "Acum. Semana 3", "Acum. Semana 4",
            "Primera semana con 100%", "Completado (100%)", "Completado anticipadamente",
            "Completadas", "En curso", "Pendientes", "Vencidas", "Con alertas",
            "Próximas a vencer", "Hoja de cronograma"]


def build_calculos(wb, ws):
    ws["A1"] = ("Hoja auxiliar OCULTA: cálculos intermedios por puesto (cada fila lee la hoja de cronograma "
                "de su puesto) que alimentan el Dashboard. No editar. Ver metodología en Configuración.")
    ws["A1"].font = font(9, True)
    for i, h in enumerate(CALC_HDR):
        c = ws.cell(row=3, column=i + 1, value=h)
        c.font = font(9, True, "FFFFFF"); c.fill = fill(NAVY); c.alignment = CENTER
        ws.column_dimensions[get_column_letter(i + 1)].width = 14
    ws.column_dimensions["A"].width = 28
    for i, p in enumerate(PUESTOS):
        r = 4 + i
        code = AREAS[i][2]
        ws[f"A{r}"] = f"=INDEX(ListaPuestos,{i + 1})"
        mes = f'rMes_{code},MesCiclo,rTipo_{code},"Macro"'
        sel = f'--((rMes_{code}=MesCiclo)*(rTipo_{code}="Macro"))'
        ws[f"B{r}"] = f'=COUNTIFS(rValid_{code},"?*",{mes})'
        ws[f"C{r}"] = f"=SUMIFS(rPeso_{code},{mes})"
        ws[f"D{r}"] = f'=COUNTIFS(rPesoEscrito_{code},">1",{mes})+COUNTIFS(rPesoEscrito_{code},"<0",{mes})'
        ws[f"E{r}"] = f"=AND(B{r}>0,D{r}=0,ABS(C{r}-1)<=0.0001)"
        ws[f"F{r}"] = f"=IFERROR(SUMPRODUCT({sel},rPeso_{code},rAvance_{code}),0)"
        for w in range(1, 5):
            col = get_column_letter(6 + w)
            ws[f"{col}{r}"] = f"=IFERROR(SUMPRODUCT({sel},rPeso_{code},rEf{w}_{code}),0)"
        ws[f"K{r}"] = (f'=IF(ROUND(G{r},4)>=1,1,IF(ROUND(H{r},4)>=1,2,'
                       f'IF(ROUND(I{r},4)>=1,3,IF(ROUND(J{r},4)>=1,4,""))))')
        ws[f"L{r}"] = f"=AND(E{r},ROUND(F{r},4)>=1)"
        ws[f"M{r}"] = f"=AND(L{r},MIN(IF(K{r}=\"\",4,K{r}),SemanaActual)<4)"
        ws[f"N{r}"] = f'=COUNTIFS(rEstado_{code},"Completado",{mes})'
        ws[f"O{r}"] = f'=COUNTIFS(rEstado_{code},"En curso",{mes})'
        ws[f"P{r}"] = f'=COUNTIFS(rEstado_{code},"Pendiente",{mes})'
        ws[f"Q{r}"] = f'=SUMPRODUCT((rMes_{code}=MesCiclo)*(rTipo_{code}="Macro")*ISNUMBER(SEARCH("Vencida",rValid_{code})))'
        ws[f"R{r}"] = f'=SUMPRODUCT((rMes_{code}=MesCiclo)*(LEFT(rValid_{code},1)="⚠"))'
        ws[f"S{r}"] = (f'=COUNTIFS(rFLim_{code},">="&FechaRef,rFLim_{code},"<="&(FechaRef+DiasAlerta),'
                       f'rAvance_{code},"<1",{mes})')
        ws[f"T{r}"] = AREAS[i][0]
    r = 8
    ws[f"A{r}"] = "Total actividades del mes (promedio de 4 puestos)"
    ws[f"B{r}"] = "=SUM(B4:B7)"
    ws[f"C{r}"] = '=COUNTIF(E4:E7,TRUE)'          # n° de puestos válidos
    ws[f"D{r}"] = '=IF(C8>0,SUMPRODUCT(F4:F7*(E4:E7=TRUE))/C8,"")'  # provisional
    ws[f"E{r}"] = "=AND(E4,E5,E6,E7)"
    ws[f"F{r}"] = '=IF(E8,AVERAGE(F4:F7),"")'
    for w in range(1, 5):
        col = get_column_letter(6 + w)
        ws[f"{col}{r}"] = f'=IF($E$8,AVERAGE({col}4:{col}7),"")'
    ws[f"K{r}"] = (f'=IF(NOT(E8),"",IF(ROUND(G8,4)>=1,1,IF(ROUND(H8,4)>=1,2,'
                   f'IF(ROUND(I8,4)>=1,3,IF(ROUND(J8,4)>=1,4,"")))))')
    ws[f"L{r}"] = "=AND(E8,ROUND(N(F8),4)>=1)"
    ws[f"M{r}"] = "=AND(L8,MIN(IF(K8=\"\",4,K8),SemanaActual)<4)"
    for col in "NOPQRS":
        ws[f"{col}{r}"] = f"=SUM({col}4:{col}7)"
    ws["A9"] = "Fila 8: C = n° puestos válidos; D = avance provisional (promedio solo de puestos válidos)."
    ws["A9"].font = font(8, italic=True)

    ws["A10"] = "Semana visible en seguimiento"
    for w in range(1, 5):
        col = get_column_letter(6 + w)
        cnt = "+".join(f'COUNTIFS(rMes_{a[2]},MesCiclo,rS{w}_{a[2]},"<>")' for a in AREAS)
        ws[f"{col}10"] = f"=OR({w}<=SemanaActual,({cnt})>0)"
    ws["A11"] = "Total registros"
    ws["B11"] = "=B8"
    ws["A12"] = "Puestos al 100%"
    ws["B12"] = "=COUNTIF(L4:L7,TRUE)"
    ws["A13"] = "Próximas a vencer"
    ws["B13"] = "=S8"
    for rr in range(4, 9):
        for cc in "CDFGHIJ":
            ws[f"{cc}{rr}"].number_format = "0.0%"
    ws["C8"].number_format = "0"
    # Mes mostrado normalizado al día 1 (cualquier fecha escrita en Configuración sirve)
    ws["A15"] = "Mes mostrado (día 1)"
    ws["B15"] = "=DATE(YEAR(MesEntrada),MONTH(MesEntrada),1)"
    # Listas de actividades delegadas por la jefatura a cada puesto (mes mostrado)
    ws["A19"] = "Actividades delegadas por la jefatura (listas desplegables de cada Cronograma)"
    ws["A19"].font = font(9, True)
    for i, a in enumerate(AREAS):
        col = get_column_letter(2 + i)
        ws[f"{col}20"] = f"=A{4 + i}"
        for k in range(1, 41):
            r = 20 + k
            ws[f"{col}{r}"] = (f'=IFERROR("{PREFIJO_JEF}"&INDEX(rAct_ADM,MATCH($A${4 + i}&"|"&{k},rDkey_ADM,0)),"")')
        add_name(wb, f"ListaDeleg_{a[2]}",
                 f"OFFSET({q(S_CALC)}!${col}$21,0,0,MAX(1,COUNTIF({q(S_CALC)}!${col}$21:${col}$60,\"?*\")),1)")
    ws["B15"].number_format = MES_FMT
    add_name(wb, "MesCiclo", f"{q(S_CALC)}!$B$15")
    ws.protection.sheet = True


# --------------------------------------------------------------------------
# Historial mensual: avance de cada puesto mes a mes desde el primer mes del historial (oct-2026)
# --------------------------------------------------------------------------
N_HIST = 24


def build_historial(wb, ws):
    widths = {"A": 2, "B": 16}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v
    for c in "CDEFGHIJ":
        ws.column_dimensions[c].width = 14
    ws.column_dimensions["K"].width = 18
    merge_set(ws, "B1:K1", "Historial mensual de actividades",
              font=font(14, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[1].height = 28
    merge_set(ws, "B2:K2",
              "HOJA AUTOMÁTICA: resume cada mes con las actividades guardadas en los Cronogramas (nada se borra al empezar un mes). "
              "Empieza en octubre 2026 y muestra 24 meses (el mes actual se resalta en amarillo). «Sin datos» = el puesto no registró actividades ese mes; "
              "«Pesos ≠ 100%» = los pesos de ese mes no suman 100%.",
              font=font(9, False, NAVY), fill=fill(LIGHT), alignment=LEFT)
    ws.row_dimensions[2].height = 32

    hr = 4
    ws[f"B{hr}"] = "Mes"
    ws.merge_cells(f"B{hr}:B{hr + 1}")
    for i, a in enumerate(AREAS):
        c1 = get_column_letter(3 + 2 * i)
        c2 = get_column_letter(4 + 2 * i)
        merge_set(ws, f"{c1}{hr}:{c2}{hr}", PUESTOS[i], font=font(10, True, "FFFFFF"), fill=fill(a[3]),
                  alignment=CENTER, border=BORDER)
        ws[f"{c1}{hr + 1}"] = "Avance"
        ws[f"{c2}{hr + 1}"] = "Completadas / total"
    ws[f"K{hr}"] = "Avance general"
    ws.merge_cells(f"K{hr}:K{hr + 1}")
    for col in "BCDEFGHIJK":
        for rr in (hr, hr + 1):
            c = ws[f"{col}{rr}"]
            c.border = BORDER; c.alignment = CENTER
            if rr == hr + 1 or col in "BK":
                c.font = font(9, True, "FFFFFF" if col in "BK" else NAVY)
                c.fill = fill(NAVY if col in "BK" else LIGHT)
    ws.row_dimensions[hr + 1].height = 28

    first = hr + 2
    for k in range(N_HIST):
        r = first + k
        ws[f"B{r}"] = f"=DATE(YEAR(HistInicio),MONTH(HistInicio)+{k},1)"
        ws[f"B{r}"].number_format = MES_FMT
        for i, a in enumerate(AREAS):
            code = a[2]
            c1 = get_column_letter(3 + 2 * i)
            c2 = get_column_letter(4 + 2 * i)
            mm = f'rMes_{code},$B{r},rTipo_{code},"Macro"'
            n = f'COUNTIFS({mm})'
            ws[f"{c1}{r}"] = (f'=IF({n}=0,"Sin datos",IF(ABS(SUMIFS(rPeso_{code},{mm})-1)>0.0001,'
                              f'"Pesos ≠ 100%",SUMPRODUCT(--((rMes_{code}=$B{r})*(rTipo_{code}="Macro")),rPeso_{code},rAvance_{code})))')
            ws[f"{c2}{r}"] = f'=IF({n}=0,"",COUNTIFS(rEstado_{code},"Completado",{mm})&" / "&{n})'
            ws[f"{c1}{r}"].number_format = "0%"
        ws[f"K{r}"] = (f'=IF(COUNT(C{r},E{r},G{r},I{r})=4,AVERAGE(C{r},E{r},G{r},I{r}),'
                       f'IF(COUNTIF(C{r}:J{r},"Sin datos")=4,"Sin datos","No definitivo"))')
        ws[f"K{r}"].number_format = "0%"
        for col in "BCDEFGHIJK":
            c = ws[f"{col}{r}"]
            c.border = BORDER; c.alignment = CENTER; c.font = font(10, col in "BK")
        ws.row_dimensions[r].height = 20
    last = first + N_HIST - 1
    cf = ws.conditional_formatting
    rng = f"C{first}:K{last}"
    cf.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(C{first}),C{first}>=1)'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(rng, FormulaRule(formula=[f'AND(ISNUMBER(C{first}),C{first}<1)'], fill=fill(BLUE_F), font=Font(color=BLUE_T)))
    cf.add(rng, FormulaRule(formula=[f'OR(C{first}="Pesos ≠ 100%",C{first}="No definitivo")'], fill=fill(AMBER_F), font=Font(color=AMBER_T)))
    cf.add(rng, FormulaRule(formula=[f'C{first}="Sin datos"'], font=Font(color="7F7F7F", italic=True)))
    cf.add(f"B{first}:B{last}", FormulaRule(formula=[f'$B{first}=MesCiclo'], fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))

    # Datos auxiliares del gráfico (columnas ocultas M–Q): 6 meses que terminan en el mes mostrado
    # (sin empezar antes del primer mes del historial), para que el gráfico no se amontone.
    N_CH = 6
    for hc in "MNOPQ":
        ws.column_dimensions[hc].hidden = True
    ws["M5"] = "Mes (gráfico)"
    for i in range(4):
        ws[f"{get_column_letter(14 + i)}5"] = SHORT[i]
    g0 = 6
    for k in range(N_CH):
        r = g0 + k
        ws[f"M{r}"] = (f'=CHOOSE(MONTH(DATE(YEAR(MAX(HistInicio,EDATE(MesCiclo,-{N_CH - 1}))),'
                       f'MONTH(MAX(HistInicio,EDATE(MesCiclo,-{N_CH - 1})))+{k},1)),'
                       f'"ene","feb","mar","abr","may","jun","jul","ago","sep","oct","nov","dic")&" "&'
                       f'YEAR(DATE(YEAR(MAX(HistInicio,EDATE(MesCiclo,-{N_CH - 1}))),'
                       f'MONTH(MAX(HistInicio,EDATE(MesCiclo,-{N_CH - 1})))+{k},1))')
        mes_k = (f'DATE(YEAR(MAX(HistInicio,EDATE(MesCiclo,-{N_CH - 1}))),'
                 f'MONTH(MAX(HistInicio,EDATE(MesCiclo,-{N_CH - 1})))+{k},1)')
        for i in range(4):
            src = get_column_letter(3 + 2 * i)
            ws[f"{get_column_letter(14 + i)}{r}"] = (
                f'=IFERROR(N(INDEX({src}{first}:{src}{last},MATCH({mes_k},B{first}:B{last},0))),0)')
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.visible_cells_only = False   # graficar aunque las columnas auxiliares estén ocultas
    ch.title = "Avance por puesto – últimos 6 meses"
    ch.add_data(Reference(ws, min_col=14, max_col=17, min_row=5, max_row=g0 + N_CH - 1), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=13, min_row=g0, max_row=g0 + N_CH - 1))
    ch.y_axis.scaling.min = 0; ch.y_axis.scaling.max = 1
    ch.y_axis.number_format = "0%"; ch.y_axis.majorUnit = 0.25
    ch.gapWidth = 120
    ch.overlap = -10
    ch.legend.position = "b"
    for i, a in enumerate(AREAS):
        ch.series[i].graphicalProperties.solidFill = a[3]
        ch.series[i].graphicalProperties.line.solidFill = a[3]
    ch.x_axis.delete = False; ch.y_axis.delete = False
    ch.height = 8.5; ch.width = 26
    ws.add_chart(ch, f"B{last + 2}")
    ws.row_breaks.append(__import__("openpyxl").worksheet.pagebreak.Break(id=last + 1))

    ws.freeze_panes = f"C{first}"
    ws.protection.sheet = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


# --------------------------------------------------------------------------
# Dashboard
# --------------------------------------------------------------------------
def count_msgs(*msgs):
    terms = []
    for a in AREAS:
        parts = "+".join(f'ISNUMBER(SEARCH("{m}",rValid_{a[2]}))' for m in msgs)
        terms.append(f"SUMPRODUCT((rMes_{a[2]}=MesCiclo)*(({parts})>0))")
    return "=" + "+".join(terms)


def build_dashboard(wb, ws):
    C = q(S_CALC) + "!"
    widths = {"A": 2, "B": 30, "C": 18, "D": 16, "E": 13, "F": 13, "G": 13, "H": 13,
              "I": 13, "J": 13, "K": 22, "L": 20, "M": 24, "N": 2}
    for k, v in widths.items():
        ws.column_dimensions[k].width = v

    merge_set(ws, "B1:M1", "Dashboard Consolidado – Gestión de Riesgos · Ciclo mensual de 4 semanas",
              font=font(16, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[1].height = 34
    right = Alignment(horizontal="right", vertical="center")
    for ref, lbl in (("B2", "Período:"), ("D2", "Semana actual:"), ("F2", "Fecha de corte:")):
        ws[ref] = lbl; ws[ref].font = font(10, True, NAVY); ws[ref].alignment = right
    for ref, val, fmt in (("C2", "=MesCiclo", MES_FMT), ("E2", "=SemanaActual", '0" de 4"'), ("G2", "=FechaRef", "dd/mm/yyyy")):
        ws[ref] = val; ws[ref].number_format = fmt; ws[ref].font = font(11, True, NAVY); ws[ref].alignment = CENTER
        ws[ref].fill = fill(LIGHT); ws[ref].border = BORDER
    merge_set(ws, "H2:M2", "Se actualiza solo con las hojas «Cronograma …» y «Actividades …» de cada puesto. El período y la semana se calculan con la fecha de hoy (ver Configuración).",
              font=font(8, italic=True, color=NAVY), alignment=LEFT)
    ws.row_dimensions[2].height = 22

    # --- Tarjetas KPI ---
    cards = [
        ("Avance general de las actividades", f'=IF({C}$E$8,{C}$F$8,"No definitivo")', "0%",
         f'=IF({C}$E$8,"Promedio simple de los 4 puestos",IF({C}$C$8=0,"Ningún puesto con datos válidos",'
         f'"Provisional "&TEXT({C}$D$8,"0%")&" ("&{C}$C$8&" de 4 puestos válidos)"))'),
        ("Meta de cumplimiento", "=MetaFinal", "0%", "Al cierre de la Semana 4"),
        ("Macroactividades registradas", f"={C}$B$11", "0", "En los 4 cronogramas"),
        ("Actividades completadas", f"={C}$N$8", "0",
         f'=IF({C}$B$11=0,"—",TEXT({C}$N$8/{C}$B$11,"0%")&" del total")'),
        ("Actividades en curso", f"={C}$O$8", "0", "Avance > 0% y < 100%"),
        ("Actividades pendientes", f"={C}$P$8", "0", f'="De todas las actividades, vencidas: "&{C}$Q$8'),
        ("Puestos que alcanzaron el 100%", f"={C}$B$12", '0" de 4"', "Con ponderación válida"),
        ("Actividades con alertas", f"={C}$R$8", "0", f'="Próximas a vencer: "&{C}$B$13'),
    ]
    groups = [("B", "D"), ("E", "G"), ("H", "J"), ("K", "M")]
    for idx, (lbl, val, fmt, sub) in enumerate(cards):
        g = groups[idx % 4]
        r = 4 if idx < 4 else 8
        a, b = g
        merge_set(ws, f"{a}{r}:{b}{r}", lbl, font=font(10, True, "FFFFFF"), fill=fill(MID),
                  alignment=CENTER, border=BORDER)
        v = merge_set(ws, f"{a}{r + 1}:{b}{r + 1}", val, font=font(24, True, NAVY), fill=fill("FFFFFF"),
                      alignment=CENTER, border=BORDER)
        v.number_format = fmt
        merge_set(ws, f"{a}{r + 2}:{b}{r + 2}", sub, font=font(9, False, NAVY), fill=fill(LIGHT),
                  alignment=CENTER, border=BORDER)
    for r in (4, 8):
        ws.row_dimensions[r].height = 20
        ws.row_dimensions[r + 1].height = 38
        ws.row_dimensions[r + 2].height = 18
    ws.row_dimensions[7].height = 8
    cf = ws.conditional_formatting
    cf.add("B5", FormulaRule(formula=['AND(ISNUMBER($B$5),$B$5>=1)'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add("B5", FormulaRule(formula=['NOT(ISNUMBER($B$5))'], fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))
    cf.add("K9", FormulaRule(formula=['$K$9>0'], fill=fill(RED_F), font=Font(color=RED_T, bold=True)))
    cf.add("K9", FormulaRule(formula=['$K$9=0'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add("H9", FormulaRule(formula=['$H$9=4'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add("E9", FormulaRule(formula=['$E$9>0'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))

    # Barra de avance general
    ws["B12"] = f'=IF({C}$E$8,"Avance general vs meta 100%","Avance PROVISIONAL vs meta")'
    ws["B12"].font = font(10, True, NAVY)
    merge_set(ws, "C12:F12", f'=IF({C}$E$8,{C}$F$8,N({C}$D$8))', font=font(10, True, NAVY), alignment=LEFT)
    ws["C12"].number_format = "0%"
    cf.add("C12:F12", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                  color="2E75B6", showValue=True))
    # Mensaje general de las actividades del mes
    proj_msg = (
        f'=IF({C}$B$11=0,"⚠ No hay actividades registradas en este mes. Cada puesto registra sus actividades y pesos en su hoja «Cronograma …».",'
        f'IF(NOT({C}$E$8),"⚠ Avance general NO definitivo: "&(4-{C}$C$8)&" puesto(s) sin actividades válidas o con pesos que no suman 100%.",'
        f'IF({C}$M$8,"Actividades del mes: ¡Enhorabuena, completado anticipadamente!",'
        f'IF({C}$L$8,"Actividades del mes completadas: meta del 100% alcanzada.",'
        f'"Actividades del mes en ejecución – Semana "&SemanaActual&" de 4."))))'
    )
    merge_set(ws, "G12:M12", proj_msg, font=font(11, True, NAVY), alignment=CENTER, border=BORDER)
    ws.row_dimensions[12].height = 30
    msg_cf(cf, "G12:M12", "$G$12")

    # --- Avance individual por puesto ---
    r = 14
    merge_set(ws, f"B{r}:M{r}", "Avance individual por puesto (ponderado por el peso de cada actividad)",
              font=font(12, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[r].height = 24
    hdrs = ["Puesto", "Avance ponderado", "Progreso", "Total actividades", "Completadas", "En curso",
            "Pendientes", "Vencidas", "Con alertas", "Estado general", "Mensaje de finalización"]
    cols = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
    for h, col in zip(hdrs, cols):
        c = ws[f"{col}{r + 1}"]
        c.value = h; c.font = font(10, True, "FFFFFF"); c.fill = fill(MID); c.alignment = CENTER; c.border = BORDER
    ws.merge_cells(f"L{r + 1}:M{r + 1}")
    ws[f"M{r + 1}"].border = BORDER
    ws.row_dimensions[r + 1].height = 30
    for i in range(4):
        rr = r + 2 + i
        k = 4 + i
        ws[f"B{rr}"] = f"={C}A{k}"
        ws[f"B{rr}"].hyperlink = f"#{q(AREAS[i][0])}!A1"
        ws[f"C{rr}"] = (f'=IF({C}B{k}=0,"Sin actividades",IF(NOT({C}E{k}),"Revisar pesos",{C}F{k}))')
        ws[f"D{rr}"] = f"=IF(ISNUMBER(C{rr}),C{rr},0)"
        ws[f"E{rr}"] = f"={C}B{k}"
        ws[f"F{rr}"] = f"={C}N{k}"
        ws[f"G{rr}"] = f"={C}O{k}"
        ws[f"H{rr}"] = f"={C}P{k}"
        ws[f"I{rr}"] = f"={C}Q{k}"
        ws[f"J{rr}"] = f"={C}R{k}"
        ws[f"K{rr}"] = (f'=IF({C}B{k}=0,"Sin actividades",IF(NOT({C}E{k}),"Pesos no válidos",'
                        f'IF({C}L{k},"Completado",IF({C}F{k}>0,"En curso","Pendiente"))))')
        ws[f"L{rr}"] = (f'=IF({C}B{k}=0,"⚠ Registrar actividades",IF(NOT({C}E{k}),"⚠ Corregir pesos no válidos",'
                        f'IF({C}M{k},"¡Enhorabuena, completado anticipadamente!",IF({C}L{k},"Completado","En ejecución"))))')
        ws.merge_cells(f"L{rr}:M{rr}")
        for col in "BCDEFGHIJKLM":
            c = ws[f"{col}{rr}"]
            c.border = BORDER
            c.font = font(10, col == "B", TXT)
            c.alignment = LEFT if col == "B" else CENTER
        ws[f"B{rr}"].fill = fill(AREAS[i][3]); ws[f"B{rr}"].font = font(10, True, "FFFFFF")
        ws[f"C{rr}"].number_format = "0%"
        ws[f"D{rr}"].number_format = '0%;;""'
        ws[f"D{rr}"].font = font(9, color="FFFFFF")
        ws.row_dimensions[rr].height = 28
    rt = r + 6
    ws[f"B{rt}"] = "Total actividades"
    ws[f"C{rt}"] = f'=IF({C}$E$8,{C}$F$8,"No definitivo")'
    ws[f"D{rt}"] = f"=IF(ISNUMBER(C{rt}),C{rt},0)"
    for col, src in zip("EFGHIJ", ["B8", "N8", "O8", "P8", "Q8", "R8"]):
        ws[f"{col}{rt}"] = f"={C}{src}"
    ws[f"E{rt}"] = f"={C}B11"
    ws[f"K{rt}"] = f'=IF({C}$E$8,"Promedio 4 puestos","No definitivo")'
    ws[f"L{rt}"] = (f'=IF(NOT({C}$E$8),"⚠ Completar configuración de todos los puestos",'
                    f'IF({C}$M$8,"¡Enhorabuena, completado anticipadamente!",IF({C}$L$8,"Completado","En ejecución")))')
    ws.merge_cells(f"L{rt}:M{rt}")
    for col in "BCDEFGHIJKLM":
        c = ws[f"{col}{rt}"]
        c.border = BORDER; c.font = font(10, True, NAVY); c.fill = fill(LIGHT)
        c.alignment = LEFT if col == "B" else CENTER
    ws[f"C{rt}"].number_format = "0%"
    ws[f"D{rt}"].number_format = '0%;;""'
    ws[f"D{rt}"].font = font(9, color=LIGHT)
    ws.row_dimensions[rt].height = 28

    a, b = r + 2, rt
    cf.add(f"D{a}:D{b}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1,
                                     color="2E75B6", showValue=False))
    cf.add(f"C{a}:C{b}", FormulaRule(formula=[f'AND(ISNUMBER(C{a}),C{a}>=1)'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(f"C{a}:C{b}", FormulaRule(formula=[f'AND(ISNUMBER(C{a}),C{a}>0,C{a}<1)'], fill=fill(BLUE_F), font=Font(color=BLUE_T, bold=True)))
    cf.add(f"C{a}:C{b}", FormulaRule(formula=[f'OR(NOT(ISNUMBER(C{a})),C{a}=0)'], fill=fill(RED_F), font=Font(color=RED_T, bold=True)))
    cf.add(f"K{a}:K{b}", FormulaRule(formula=[f'K{a}="Completado"'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(f"K{a}:K{b}", FormulaRule(formula=[f'K{a}="En curso"'], fill=fill(BLUE_F), font=Font(color=BLUE_T, bold=True)))
    cf.add(f"K{a}:K{b}", FormulaRule(formula=[f'OR(K{a}="Sin actividades",K{a}="Pesos no válidos",K{a}="No definitivo")'],
                                     fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))
    cf.add(f"K{a}:K{b}", FormulaRule(formula=[f'K{a}="Pendiente"'], fill=fill(RED_F), font=Font(color=RED_T)))
    cf.add(f"I{a}:J{b}", FormulaRule(formula=[f'I{a}>0'], fill=fill(ORANGE_F), font=Font(color="833C0B", bold=True)))
    msg_cf(cf, f"L{a}:M{b}", f"$L{a}")

    # --- Seguimiento semanal ---
    r = 22
    merge_set(ws, f"B{r}:M{r}", "Seguimiento semanal: avance acumulado ponderado e incremento semanal",
              font=font(12, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[r].height = 24
    merge_set(ws, f"C{r + 1}:F{r + 1}", "Avance ACUMULADO ponderado al cierre de cada semana",
              font=font(9, True, "FFFFFF"), fill=fill(MID), alignment=CENTER, border=BORDER)
    merge_set(ws, f"H{r + 1}:K{r + 1}", "INCREMENTO respecto a la semana anterior (puntos %)",
              font=font(9, True, "FFFFFF"), fill=fill(MID), alignment=CENTER, border=BORDER)
    merge_set(ws, f"L{r + 1}:M{r + 1}", "Brecha vs meta 100%",
              font=font(9, True, "FFFFFF"), fill=fill(MID), alignment=CENTER, border=BORDER)
    hdr_r = r + 2
    ws[f"B{hdr_r}"] = "Puesto / Indicador"
    for w in range(4):
        ws[f"{'CDEF'[w]}{hdr_r}"] = f"Semana {w + 1}"
        ws[f"{'HIJK'[w]}{hdr_r}"] = f"Semana {w + 1}"
    ws[f"L{hdr_r}"] = "Pendiente para la meta"
    ws.merge_cells(f"L{hdr_r}:M{hdr_r}")
    for col in "BCDEFHIJKLM":
        c = ws[f"{col}{hdr_r}"]
        c.font = font(9, True, NAVY); c.fill = fill(LIGHT); c.alignment = CENTER; c.border = BORDER
    labels = PUESTOS + ["Avance general de las actividades", "Plan de referencia (supuesto)", "Meta final"]
    for i, lbl in enumerate(labels):
        rr = hdr_r + 1 + i
        k = 4 + i
        if i < 4:
            ws[f"B{rr}"] = f"={C}A{k}"
        else:
            ws[f"B{rr}"] = lbl
        for w in range(4):
            ccol = "CDEF"[w]; icol = "HIJK"[w]; kcol = "GHIJ"[w]
            if i < 4:
                ws[f"{ccol}{rr}"] = f'=IF(NOT({C}{kcol}$10),"",IF({C}$E{k},{C}{kcol}{k},"–"))'
            elif i == 4:
                ws[f"{ccol}{rr}"] = f'=IF(NOT({C}{kcol}$10),"",IF({C}$E$8,{C}{kcol}8,"n/d"))'
            elif i == 5:
                ws[f"{ccol}{rr}"] = f"=INDEX(PlanRef,{w + 1})"
            else:
                ws[f"{ccol}{rr}"] = "=MetaFinal"
            if i <= 4:
                if w == 0:
                    ws[f"{icol}{rr}"] = f'=IF(ISNUMBER(C{rr}),C{rr},"")'
                else:
                    prev = "CDEF"[w - 1]
                    ws[f"{icol}{rr}"] = f'=IF(AND(ISNUMBER({ccol}{rr}),ISNUMBER({prev}{rr})),{ccol}{rr}-{prev}{rr},"")'
            ws[f"{ccol}{rr}"].number_format = "0%"
            ws[f"{icol}{rr}"].number_format = '+0%;-0%;0%'
        if i <= 4:
            ws[f"L{rr}"] = (f'=IFERROR(MetaFinal-LOOKUP(2,1/ISNUMBER(C{rr}:F{rr}),C{rr}:F{rr}),"")')
            ws[f"L{rr}"].number_format = "0%"
        ws.merge_cells(f"L{rr}:M{rr}")
        for col in "BCDEFHIJKLM":
            c = ws[f"{col}{rr}"]
            c.border = BORDER
            c.alignment = LEFT if col == "B" else CENTER
            c.font = font(10, i >= 4, NAVY if i >= 4 else TXT)
            if i >= 4:
                c.fill = fill(LIGHT if i == 4 else "FFFFFF")
        if i < 4:
            ws[f"B{rr}"].fill = fill(AREAS[i][3]); ws[f"B{rr}"].font = font(10, True, "FFFFFF")
        if i >= 5:
            for col in "BCDEF":
                ws[f"{col}{rr}"].font = font(9, False, "404040", italic=(i == 5))
    first_w, last_w = hdr_r + 1, hdr_r + 5
    cf.add(f"C{first_w}:F{last_w}", FormulaRule(formula=[f'AND(ISNUMBER(C{first_w}),C{first_w}>=1)'],
                                               fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(f"C{first_w}:F{last_w}", FormulaRule(formula=[f'AND(ISNUMBER(C{first_w}),C{first_w}<INDEX(PlanRef,COLUMN(C{first_w})-2))'],
                                               fill=fill(AMBER_F), font=Font(color=AMBER_T)))
    cf.add(f"H{first_w}:K{last_w}", FormulaRule(formula=[f'AND(ISNUMBER(H{first_w}),H{first_w}<0)'],
                                               fill=fill(RED_F), font=Font(color=RED_T, bold=True)))
    note_r = hdr_r + 9
    merge_set(ws, f"B{note_r}:M{note_r}",
              "Lectura: el acumulado NO se suma entre semanas; el incremento es la diferencia entre semanas consecutivas. "
              "«–» = puesto sin configuración válida; «n/d» = avance general no definitivo; celda vacía = semana aún no alcanzada. "
              "Ámbar = acumulado por debajo del plan de referencia.",
              font=font(8, italic=True, color="404040"), alignment=LEFT)
    ws.row_dimensions[note_r].height = 26

    # --- Gráficos (a todo el ancho, uno debajo del otro, sin amontonar) ---
    ch_r = note_r + 2
    # Nombres cortos para los ejes (columna auxiliar oculta P)
    ws.column_dimensions["P"].hidden = True
    for i, short in enumerate(SHORT):
        ws[f"P{16 + i}"] = short
        ws[f"P{25 + i}"] = short
    ws["P30"] = "Plan de referencia"

    bar = BarChart()
    bar.type = "bar"
    bar.title = "Avance ponderado de cada puesto en el mes"
    bar.add_data(Reference(ws, min_col=4, min_row=16, max_row=19), titles_from_data=False)
    bar.set_categories(Reference(ws, min_col=16, min_row=16, max_row=19))
    bar.y_axis.scaling.min = 0; bar.y_axis.scaling.max = 1
    bar.y_axis.number_format = "0%"; bar.y_axis.majorUnit = 0.25
    bar.y_axis.majorGridlines = None
    bar.x_axis.scaling.orientation = "maxMin"
    bar.gapWidth = 60
    bar.legend = None
    for i, a in enumerate(AREAS):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = a[3]
        pt.graphicalProperties.line.solidFill = a[3]
        bar.series[0].dPt.append(pt)
    bar.dataLabels = DataLabelList(); bar.dataLabels.showVal = True; bar.dataLabels.numFmt = "0%"
    bar.dataLabels.showSerName = False; bar.dataLabels.showCatName = False
    bar.dataLabels.showLegendKey = False; bar.dataLabels.showPercent = False
    bar.dataLabels.position = "outEnd"
    bar.visible_cells_only = False
    bar.x_axis.delete = False; bar.y_axis.delete = False
    bar.height = 7.5; bar.width = 31
    ws.add_chart(bar, f"B{ch_r}")

    col = BarChart()
    col.type = "col"
    col.grouping = "clustered"
    col.title = "Avance acumulado por semana de cada puesto vs plan de referencia"
    for i in range(4):
        col.add_data(Reference(ws, min_col=3, max_col=6, min_row=25 + i), titles_from_data=False, from_rows=True)
    col.set_categories(Reference(ws, min_col=3, max_col=6, min_row=hdr_r))
    for i, a in enumerate(AREAS):
        col.series[i].tx = SeriesLabel(v=SHORT[i])
        col.series[i].graphicalProperties.solidFill = a[3]
        col.series[i].graphicalProperties.line.solidFill = a[3]
    col.gapWidth = 80
    col.overlap = -10
    col.y_axis.scaling.min = 0; col.y_axis.scaling.max = 1
    col.y_axis.number_format = "0%"; col.y_axis.majorUnit = 0.25
    plan = LineChart()
    plan.add_data(Reference(ws, min_col=3, max_col=6, min_row=hdr_r + 6), titles_from_data=False, from_rows=True)
    plan.series[0].tx = SeriesLabel(v="Plan de referencia")
    plan.series[0].graphicalProperties.line.solidFill = "7F7F7F"
    plan.series[0].graphicalProperties.line.dashStyle = "dash"
    plan.series[0].graphicalProperties.line.width = 22000
    plan.series[0].marker.symbol = "circle"
    plan.series[0].marker.size = 6
    col += plan
    col.legend.position = "b"
    col.visible_cells_only = False
    col.x_axis.delete = False; col.y_axis.delete = False
    col.height = 8; col.width = 31
    ws.add_chart(col, f"B{ch_r + 16}")

    # --- Control de calidad ---
    r = ch_r + 33
    merge_set(ws, f"B{r}:M{r}", "Control de calidad de datos",
              font=font(12, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[r].height = 24
    hr = r + 1
    merge_set(ws, f"B{hr}:F{hr}", "Verificación", font=font(10, True, "FFFFFF"), fill=fill(MID), alignment=CENTER, border=BORDER)
    ws[f"G{hr}"] = "Casos"
    ws[f"H{hr}"] = "Resultado"
    for cc in "GH":
        c = ws[f"{cc}{hr}"]; c.font = font(10, True, "FFFFFF"); c.fill = fill(MID); c.alignment = CENTER; c.border = BORDER
    merge_set(ws, f"I{hr}:M{hr}", "Acción sugerida", font=font(10, True, "FFFFFF"), fill=fill(MID), alignment=CENTER, border=BORDER)
    qc = [
        ("Puestos sin actividades registradas", f"=COUNTIF({C}B4:B7,0)", "crit",
         "Registrar al menos una macroactividad por puesto."),
        ("Puestos con pesos no válidos (fuera de 0–100%)", f"=COUNTIFS({C}B4:B7,\">0\",{C}E4:E7,FALSE)", "crit",
         "Corregir el peso (los pesos se ajustan solos a 100%)."),
        ("Actividades con peso vacío o fuera de rango", count_msgs("Peso vacío", "Peso fuera de rango"), "crit",
         "Ingresar un peso entre 0% y 100% (mayor que 0)."),
        ("Actividades sin fecha límite", count_msgs("Sin fecha límite"), "crit", "Registrar la fecha límite."),
        ("Fechas incoherentes o no válidas (incluye sin fecha de inicio)",
         count_msgs("Inicio posterior", "Fecha no válida", "Sin fecha de inicio"), "crit",
         "La fecha de inicio debe ser anterior o igual a la fecha límite."),
        ("Porcentajes de avance > 100%, < 0% o no numéricos", count_msgs("Avance fuera de rango"), "crit",
         "Corregir: los avances deben estar entre 0% y 100%."),
        ("Avance semanal menor que el de una semana anterior", count_msgs("decreciente"), "crit",
         "Revisar: el avance es acumulado y no debería disminuir."),
        ("Actividades vencidas no completadas", f"={C}Q8", "crit",
         "Actualizar avance o reprogramar con justificación en Observaciones."),
        ("Actividades próximas a vencer (no completadas)", f"={C}B13", "warn",
         "Alerta preventiva: priorizar seguimiento."),
        ("Filas con datos pero sin nombre de macroactividad", count_msgs("Macroactividad vacía"), "crit",
         "Escribir el nombre de la actividad."),
        ("Actividades vencidas sin observación que explique el incumplimiento", count_msgs("sin observación"), "crit",
         "Escribir en Observaciones por qué no se cumplió y la nueva fecha prevista."),
        ("Avance general no definitivo (puestos no válidos)", f"=4-{C}C8", "crit",
         "El avance general solo es definitivo con los 4 puestos válidos."),
    ]
    for i, (lbl, f, kind, act) in enumerate(qc):
        rr = hr + 1 + i
        merge_set(ws, f"B{rr}:F{rr}", lbl, font=font(10), alignment=LEFT, border=BORDER)
        ws[f"G{rr}"] = f
        ws[f"G{rr}"].number_format = "0"
        ws[f"G{rr}"].font = font(10, True); ws[f"G{rr}"].alignment = CENTER; ws[f"G{rr}"].border = BORDER
        ws[f"H{rr}"] = f'=IF(G{rr}=0,"✔ OK",IF("{kind}"="warn","● Atención","⚠ Revisar"))'
        ws[f"H{rr}"].font = font(10, True); ws[f"H{rr}"].alignment = CENTER; ws[f"H{rr}"].border = BORDER
        merge_set(ws, f"I{rr}:M{rr}", act, font=font(9), alignment=LEFT, border=BORDER)
        ws.row_dimensions[rr].height = 20
    q1, q2 = hr + 1, hr + len(qc)
    cf.add(f"H{q1}:H{q2}", FormulaRule(formula=[f'LEFT(H{q1},1)="✔"'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(f"H{q1}:H{q2}", FormulaRule(formula=[f'LEFT(H{q1},1)="●"'], fill=fill(AMBER_F), font=Font(color=AMBER_T, bold=True)))
    cf.add(f"H{q1}:H{q2}", FormulaRule(formula=[f'LEFT(H{q1},1)="⚠"'], fill=fill(RED_F), font=Font(color=RED_T, bold=True)))
    note = q2 + 1
    merge_set(ws, f"B{note}:M{note}",
              "Detalle por actividad: columna «Revisión automática (alertas)» de la hoja «Cronograma …» de cada puesto (filtrar por «⚠»). "
              "Pendiente normal = sin avance y dentro de plazo; Vencida = fecha límite anterior a la fecha de corte sin llegar a 100%.",
              font=font(8, italic=True, color="404040"), alignment=LEFT)
    ws.row_dimensions[note].height = 26

    # --- Hitos ---
    r = note + 2
    merge_set(ws, f"B{r}:M{r}", "Entregas por semana y alertas de plazo (macroactividades completadas / programadas)",
              font=font(12, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[r].height = 24
    hr = r + 1
    heads = {"B": "Puesto", "C": "Semana 1", "D": "Semana 2", "E": "Semana 3", "F": "Semana 4",
             "G": "Total", "H": "Vencidas", "I": "Por vencer"}
    for colh, t in heads.items():
        c = ws[f"{colh}{hr}"]; c.value = t; c.font = font(10, True, "FFFFFF"); c.fill = fill(MID)
        c.alignment = CENTER; c.border = BORDER
    merge_set(ws, f"J{hr}:M{hr}", "Alerta de plazo (actividades y subactividades del mes)", font=font(10, True, "FFFFFF"), fill=fill(MID), alignment=CENTER, border=BORDER)
    for i in range(4):
        rr = hr + 1 + i
        ws[f"B{rr}"] = f"={C}A{4 + i}"
        hc = AREAS[i][2]
        for w in range(4):
            colw = "CDEF"[w]
            ws[f"{colw}{rr}"] = (f'=COUNTIFS(hSemana_{hc},"Semana {w + 1}",hEstado_{hc},"Completado",hTipo_{hc},"Macro")&" / "'
                                 f'&COUNTIFS(hSemana_{hc},"Semana {w + 1}",hTipo_{hc},"Macro")')
        ws[f"G{rr}"] = (f'=COUNTIFS(hEstado_{hc},"Completado",hTipo_{hc},"Macro")&" / "'
                        f'&COUNTIFS(hEstado_{hc},"?*",hTipo_{hc},"Macro")')
        ws[f"H{rr}"] = f'=COUNTIFS(rMes_{hc},MesCiclo,rAlerta_{hc},"Vencida*")'
        ws[f"I{rr}"] = f'=COUNTIFS(rMes_{hc},MesCiclo,rAlerta_{hc},"Vence*")'
        ws[f"B{rr}"].hyperlink = f"#{q(AREAS[i][1])}!A1"
        merge_set(ws, f"J{rr}:M{rr}",
                  f'=IF(H{rr}+I{rr}=0,"✔ Sin vencidas ni por vencer",IF(H{rr}>0,"⚠ "&H{rr}&" vencida(s) sin cumplir","")'
                  f'&IF(AND(H{rr}>0,I{rr}>0),"  ·  ","")&IF(I{rr}>0,I{rr}&" por vencer (≤ "&DiasAlerta&" días)",""))',
                  font=font(9, True), alignment=LEFT, border=BORDER)
        for colw in "BCDEFGHI":
            c = ws[f"{colw}{rr}"]; c.border = BORDER; c.font = font(10, colw == "B")
            c.alignment = LEFT if colw == "B" else CENTER
        ws[f"B{rr}"].fill = fill(AREAS[i][3]); ws[f"B{rr}"].font = font(10, True, "FFFFFF")
    # Fila de la jefatura (hoja Administración)
    rr = hr + 5
    ws[f"B{rr}"] = ADMIN[0]
    ws[f"B{rr}"].hyperlink = f"#{q(ADMIN[1])}!A1"
    for colw in "CDEF":
        ws[f"{colw}{rr}"] = "—"
    ws[f"G{rr}"] = (f'=COUNTIFS(rMes_ADM,MesCiclo,rTipo_ADM,"Macro",rEstado_ADM,"Completado")&" / "'
                    f'&COUNTIFS(rMes_ADM,MesCiclo,rTipo_ADM,"Macro")')
    ws[f"H{rr}"] = '=COUNTIFS(rMes_ADM,MesCiclo,rAlerta_ADM,"Vencida*")'
    ws[f"I{rr}"] = '=COUNTIFS(rMes_ADM,MesCiclo,rAlerta_ADM,"Vence*")'
    merge_set(ws, f"J{rr}:M{rr}",
              f'=IF(H{rr}+I{rr}=0,"✔ Sin vencidas ni por vencer",IF(H{rr}>0,"⚠ "&H{rr}&" vencida(s) sin cumplir","")'
              f'&IF(AND(H{rr}>0,I{rr}>0),"  ·  ","")&IF(I{rr}>0,I{rr}&" por vencer (≤ "&DiasAlerta&" días)",""))',
              font=font(9, True), alignment=LEFT, border=BORDER)
    for colw in "BCDEFGHI":
        c = ws[f"{colw}{rr}"]; c.border = BORDER; c.font = font(10, colw == "B")
        c.alignment = LEFT if colw == "B" else CENTER
    ws[f"B{rr}"].fill = fill(ADMIN[3]); ws[f"B{rr}"].font = font(10, True, "FFFFFF")
    h1, h2 = hr + 1, hr + 5
    cf.add(f"H{h1}:H{h2}", FormulaRule(formula=[f'H{h1}>0'], fill=fill("C00000"), font=Font(color="FFFFFF", bold=True)))
    cf.add(f"I{h1}:I{h2}", FormulaRule(formula=[f'I{h1}>0'], fill=fill("FFC000"), font=Font(color="3F2F00", bold=True)))
    cf.add(f"J{h1}:M{h2}", FormulaRule(formula=[f'$H{h1}>0'], stopIfTrue=True, fill=fill("C00000"), font=Font(color="FFFFFF", bold=True)))
    cf.add(f"J{h1}:M{h2}", FormulaRule(formula=[f'$I{h1}>0'], stopIfTrue=True, fill=fill("FFC000"), font=Font(color="3F2F00", bold=True)))
    cf.add(f"J{h1}:M{h2}", FormulaRule(formula=[f'LEFT($J{h1},1)="✔"'], fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))

    ws.freeze_panes = "A3"
    ws.protection.sheet = True
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:N{h2 + 1}"
    ws.print_title_rows = "1:2"
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.row_breaks.append(__import__("openpyxl").worksheet.pagebreak.Break(id=note_r + 1))
    ws.row_breaks.append(__import__("openpyxl").worksheet.pagebreak.Break(id=ch_r + 31))
    ws.oddFooter.center.text = "Página &P de &N"


def msg_cf(cf, rng, first):
    cf.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("Enhorabuena",{first}))'], stopIfTrue=True,
                            fill=fill(DKGREEN), font=Font(color="FFFFFF", bold=True)))
    cf.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("ompletad",{first}))'], stopIfTrue=True,
                            fill=fill(GREEN_F), font=Font(color=GREEN_T, bold=True)))
    cf.add(rng, FormulaRule(formula=[f'LEFT({first},1)="⚠"'], stopIfTrue=True,
                            fill=fill(RED_F), font=Font(color=RED_T, bold=True)))
    cf.add(rng, FormulaRule(formula=[f'ISNUMBER(SEARCH("ejecución",{first}))'],
                            fill=fill(BLUE_F), font=Font(color=BLUE_T, bold=True)))


# --------------------------------------------------------------------------
def build_instrucciones(ws):
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    merge_set(ws, "B1:B1", "Instrucciones – cómo usar este archivo",
              font=font(14, True, "FFFFFF"), fill=fill(NAVY), alignment=LEFT)
    ws.row_dimensions[1].height = 28
    sections = [
        ("¿Para qué sirve este archivo?", [
            "Sirve para planificar y medir, MES A MES, las actividades de los cuatro puestos del área: Riesgo Operacional, Riesgo Normativo, Riesgo Financiero y Asistente Técnico de Riesgo. Cada mes se divide en 4 semanas y la meta es que las actividades lleguen al 100% al terminar la Semana 4. Todo lo de meses anteriores queda guardado como historial.",
        ]),
        ("Colores: qué se llena y qué no", [
            "• GRIS SUAVE = celda que usted llena (desbloqueada).",
            "• AZUL MUY CLARO = celda automática (bloqueada: no se puede escribir en ella).",
            "• Cada puesto tiene su color: NARANJA = Riesgo Operacional · MORADO = Riesgo Normativo · VERDE = Riesgo Financiero · AZUL = Asistente Técnico de Riesgo.",
            "• Colores de avisos: verde = completado · azul claro = en curso · rojo suave = pendiente o dato faltante · ámbar = por vencer o mes actual · naranja = vencida.",
        ]),
        ("Las hojas del archivo", [
            "• Dashboard - Consolidado: resumen del MES ACTUAL con dos gráficas (avance de cada puesto y avance por semana vs plan). Se actualiza solo; no se llena.",
            "• Historial mensual: tabla con el avance de cada puesto mes a mes desde octubre 2026 y gráfica de los últimos 6 meses. Se actualiza sola.",
            "• Cronograma (una por puesto): la ÚNICA hoja donde cada responsable escribe sus actividades, pesos, avances y observaciones.",
            "• Actividades (una por puesto, al lado de su cronograma): muestra sola las actividades del mes, semana por semana. Está bloqueada; no se escribe nada.",
            "• Configuración: mes, semana y fecha de corte. Todo se calcula solo con la fecha de hoy.",
        ]),
        ("Cómo llenar su hoja «Cronograma» (paso a paso)", [
            "1) Abra la pestaña «Cronograma» de SU puesto y escriba en la primera fila vacía (celdas gris suave).",
            "2) En «Tipo» elija «Macroactividad» (si lo deja vacío, también cuenta como macroactividad).",
            "3) Escriba el nombre en «Actividad», la Fecha de inicio y la Fecha límite. El N° (1, 2, 3… por mes) y el Mes se llenan solos.",
            "4) Peso (solo macroactividades, OPCIONAL): si lo deja vacío, todas las macroactividades del mes valen lo mismo. Si escribe pesos, se ajustan solos para sumar 100% (ej. si escribe 60% y 20%, quedan 75% y 25%). No hace falta que sumen 100% ni que esté todo el mes completo.",
            "5) CUANDO TERMINE una actividad, escriba la FECHA DE CUMPLIMIENTO: queda al 100%, «Completado», y la alerta dice «Cumplida» (ya no aparece vencida).",
            "6) Avances parciales (opcional): si una macroactividad sin subactividades va a medias, puede escribir en la columna de la semana el % total logrado hasta esa semana (ej. Semana 1 = 30%, Semana 2 = 70%).",
            "7) El «% de Avance Total Actual», el «Estado» y la «Alerta de plazo» se calculan solos.",
            "8) Observaciones: si una actividad se atrasa o no se cumplió, escriba aquí el motivo y la nueva fecha.",
        ]),
        ("Subactividades (puede agregar todas las que necesite)", [
            "1) Debajo de la macroactividad, en las filas siguientes, elija en «Tipo» la opción «Subactividad» y escriba su nombre en «Actividad», con sus fechas. Se verá con sangría (↳) justo debajo de su macroactividad.",
            "2) El N° NO se escribe: sale solo 1.1, 1.2, 1.3… (la columna está bloqueada).",
            "3) Las subactividades NO llevan peso: todas valen lo mismo dentro de su macroactividad.",
            "4) Cuando termine una subactividad, escriba su FECHA DE CUMPLIMIENTO (queda al 100%). También puede anotar avances parciales por semana.",
            "5) El % de la macroactividad se calcula SOLO: es el promedio de sus subactividades. Ejemplo con 4 subactividades: si cumple 1 → 25%; si cumple 2 → 50%; si cumple las 4 → 100%. En la fila de la macroactividad deje VACÍAS las semanas (se ven en azul claro).",
            "6) Para agregar otra subactividad más adelante, escríbala debajo de las demás subactividades de su macroactividad: siempre pertenece a la macroactividad que esté más arriba.",
        ]),
        ("Alerta de plazo (vencidas y por vencer)", [
            "La columna «Alerta de plazo» (en el Cronograma, la hoja Actividades y Administración) dice en cada actividad y subactividad:",
            "• «Cumplida» (verde): tiene fecha de cumplimiento o llegó al 100%. Si se cumplió después de la fecha límite dice «Cumplida con X días de retraso».",
            "• «Vencida hace X días» (rojo): pasó la fecha límite y no tiene fecha de cumplimiento. Escriba en Observaciones por qué no se cumplió.",
            "• «Vence hoy» o «Vence en X días» (amarillo): la fecha límite está cerca (el número de días se cambia en Configuración).",
            "• «En plazo» (verde claro): todavía hay tiempo.",
            "Arriba de cada hoja aparece el total del mes («X vencida(s) · Y por vencer») y el Dashboard muestra lo mismo por puesto y para la jefatura.",
        ]),
        ("Hoja «Administración» (jefatura) y cómo DELEGAR actividades", [
            "Es la lista de actividades de la jefatura, con el mismo formato del Cronograma: Tipo (macroactividad / subactividad), actividad, fechas, fecha de cumplimiento, avance por semana, estado, alerta de plazo y observaciones.",
            "DELEGAR: 1) La jefatura escribe la macroactividad y en «Delegar a» elige el puesto (si lo deja vacío, la actividad es de la jefatura).",
            "2) Al puesto le aparece arriba de su Cronograma un aviso amarillo: «La jefatura le delegó X actividad(es)… pendiente(s) de agregar».",
            "3) El puesto, en una fila nueva de su Cronograma, abre la lista de la columna «Actividad» y elige la actividad (sale como «[Jefatura] nombre»). Luego pone sus fechas, su avance o fecha de cumplimiento y, si quiere, subactividades debajo. No debe cambiar el texto elegido.",
            "4) En la hoja Administración, las columnas «% de Avance», «Estado», «Alerta de plazo» y «Seguimiento» de esa actividad se llenan solas con lo que registra el puesto. «Seguimiento» dice «✔ Recibida por …» o «⚠ … aún no la agrega».",
            "Nota: si la jefatura cambia el nombre de una actividad ya delegada, el puesto debe volver a elegirla en la lista.",
            "Las vencidas y por vencer de la jefatura también aparecen en el Dashboard, en la tabla de alertas de plazo.",
        ]),
        ("¿Qué es la columna «Revisión automática (alertas)»?", [
            "Es un revisor automático: revisa cada fila y avisa si algo falta o está mal. «✔ OK» = todo bien. «⚠» = hay algo que corregir y dice qué es, por ejemplo: «Sin fecha límite», «Pesos del mes suman 90%», «Pesos de sus subactividades suman 50% (deben sumar 60%)», «Avance semanal decreciente» o «Vencida».",
            "No hay que escribir nada en esa columna. Corrija lo que indica y el aviso desaparece.",
        ]),
        ("La hoja «Actividades» (automática y bloqueada)", [
            "1) Muestra solo las actividades del mes actual de su Cronograma, en orden: cada macroactividad (resaltada) seguida de sus subactividades (con «↳»), con su alerta de plazo. No se escribe nada en ella.",
            "2) «Semana de entrega» indica en qué semana del mes cae la fecha límite (la semana actual se resalta en amarillo).",
            "3) Las columnas Semana 1 a 4 muestran una barra del color del puesto en las semanas en que la actividad está programada (de la fecha de inicio a la fecha límite) y, dentro, el avance anotado esa semana. Si la actividad está vencida, la barra se pone naranja.",
            "4) Para cambiar algo, corríjalo en el Cronograma: esta hoja se actualiza al instante.",
        ]),
        ("Cada mes nuevo: NO se borra nada", [
            "1) Al empezar el mes, cada responsable escribe las actividades del nuevo mes en las filas vacías de su Cronograma, DEBAJO de las del mes anterior. No borre nada: los meses pasados quedan como historial.",
            "2) El mes de cada actividad se asigna solo según su fecha de inicio. El Dashboard y las hojas Actividades muestran automáticamente el mes actual, con su Semana 1 a 4.",
            "3) Si una actividad se repite cada mes, cópiela en una fila nueva y cambie sus fechas al nuevo mes (y borre los avances de la copia).",
            "4) Para revisar un mes anterior: en Configuración, en «Mes que se muestra», escriba cualquier fecha de ese mes (ej. 01/09/2026). Dashboard y Actividades mostrarán ese mes. Para volver al mes actual escriba =FECHA(AÑO(HOY());MES(HOY());1).",
            "5) La hoja «Historial mensual» muestra el avance de cada puesto mes a mes, desde octubre 2026 (24 meses).",
            "6) Cada Cronograma tiene 300 filas (por ejemplo, 25 actividades al mes durante un año). Si se llena, vea «Necesito más filas» o guarde una copia anual del archivo (Archivo > Guardar una copia).",
        ]),
        ("El Dashboard (resumen automático del mes)", [
            "• Muestra el avance de cada puesto y el general, cuántas actividades están completadas, en curso, pendientes o vencidas, y el avance de cada semana del mes.",
            "• Avance del puesto = suma de (peso × avance) de sus actividades del mes. Avance general = promedio de los 4 puestos (todos valen igual).",
            "• Si un puesto llega al 100% antes de la Semana 4 aparece «¡Enhorabuena, completado anticipadamente!». Si llega en la Semana 4: «Completado».",
            "• Si un puesto no tiene actividades en el mes o sus pesos no suman 100%, el avance general aparece como «No definitivo» hasta que se corrija.",
            "• Haga clic en el nombre de un puesto para ir a su hoja.",
        ]),
        ("Compartir el archivo con el equipo", [
            "1) Suba el archivo a OneDrive o SharePoint del área.",
            "2) Botón Compartir > invite a los cuatro responsables con permiso de edición, y a la jefatura con permiso de solo lectura.",
            "3) Todos pueden trabajar al mismo tiempo, cada quien en su Cronograma (Excel en la Web o Microsoft 365, con Autoguardado activado).",
            "4) Si alguien borra algo por error: Archivo > Información > Historial de versiones permite recuperar una versión anterior y ver quién cambió qué.",
        ]),
        ("Protección", [
            "Las celdas automáticas (azul muy claro) están bloqueadas y las de llenar (gris suave) desbloqueadas. Las hojas Dashboard, Actividades, Historial y Configuración (salvo sus celdas gris suave) están completamente bloqueadas. La protección no tiene contraseña.",
            "Para poner contraseña: Revisar > Desproteger hoja y luego Revisar > Proteger hoja; marque «Seleccionar celdas desbloqueadas», «Usar Autofiltro» y «Ordenar», y escriba la contraseña.",
            "Importante: en un archivo compartido, cualquier persona con permiso de edición podría escribir en el Cronograma de otro puesto. Si se necesita impedirlo por completo, se usan cuatro archivos (uno por puesto) y este archivo los une con Power Query (Datos > Obtener datos > Desde SharePoint/carpeta, una consulta por tabla tblAct_XX, con «Actualizar al abrir»).",
        ]),
        ("Necesito más filas", [
            "Revisar > Desproteger hoja, colóquese en la última celda de la tabla del Cronograma y presione Tab; las fórmulas se copian solas. Luego vuelva a proteger la hoja. El Dashboard lee hasta la fila 1000.",
        ]),
        ("Notas técnicas", [
            "• Archivo .xlsx sin macros, para Microsoft 365 en español (escritorio y Web). Las fórmulas se ven en español (BUSCAR, SI.ERROR, SUMAPRODUCTO, HOY…).",
            "• Fórmula del avance actual: =SI.ERROR(BUSCAR(2;1/(G5:J5<>\"\");G5:J5);0) → toma el último avance semanal escrito.",
            "• Mes de la actividad = mes de la fecha de inicio (si no hay, de la fecha límite). Semana actual = ENTERO((fecha de corte − primer lunes del mes) / 7) + 1, entre 1 y 4. Puede forzarse con «Semana manual» en Configuración.",
        ]),
    ]
    r = 3
    for title, lines in sections:
        c = ws.cell(row=r, column=2, value=title)
        c.font = font(11, True, "FFFFFF"); c.fill = fill(MID); c.alignment = LEFT
        ws.row_dimensions[r].height = 20
        r += 1
        for ln in lines:
            c = ws.cell(row=r, column=2, value=ln)
            c.font = font(10); c.alignment = LEFT_TOP
            ws.row_dimensions[r].height = max(15, 15 * (len(ln) // 115 + 1))
            r += 1
        r += 1
    ws.protection.sheet = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "Cronograma_Dashboard_Riesgos.xlsx"
    build(out)
    print("Generado:", out)
