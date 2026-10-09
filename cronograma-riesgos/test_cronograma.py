# -*- coding: utf-8 -*-
"""
Pruebas funcionales del libro: genera copias temporales con datos de prueba,
las recalcula con LibreOffice y verifica cálculos, mensajes y alertas.

Uso:  python test_cronograma.py <ruta_recalc.py> <directorio_temporal>
Los archivos de prueba se generan en el directorio temporal; el entregable
(sin datos de prueba) se genera aparte con build_cronograma.py.
"""
import datetime as dt
import os
import shutil
import subprocess
import sys
import json

from openpyxl import load_workbook

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_cronograma as B  # noqa: E402

RECALC, TMP = sys.argv[1], sys.argv[2]
D = dt.date
RO, RLN, RF, ATR = B.PUESTOS
FAILS = []


def act(i, area, peso, s=(None, None, None, None), fin=D(2026, 11, 6), ini=D(2026, 10, 12),
        resp="Resp. prueba", obs="Observación de prueba", **kw):
    s = tuple(s) + (None,) * (4 - len(s))
    d = dict(area=area, macro=f"Actividad de prueba {i}", resp=resp, ini=ini,
             fin=fin, peso=peso, s1=s[0], s2=s[1], s3=s[2], s4=s[3], obs=obs)
    d.update(kw)
    return d


def run(name, rows, **kw):
    kw.setdefault("mes", D(2026, 10, 1))
    path = os.path.join(TMP, f"test_{name}.xlsx")
    B.build(path, test_rows=rows, **kw)
    out = subprocess.run([sys.executable, RECALC, path, "90"], capture_output=True, text=True)
    res = json.loads(out.stdout[out.stdout.index("{"):])
    assert res.get("status") == "success", (name, res)
    wb = load_workbook(path, data_only=True)
    return wb


def check(name, cond, detail=""):
    print(("  OK   " if cond else "  FAIL ") + name + ("" if cond else f"  -> {detail}"))
    if not cond:
        FAILS.append(name)


def dash_puesto(wb, i):
    d = wb[B.S_DASH]
    r = 16 + i
    return {k: d[f"{c}{r}"].value for k, c in
            dict(avance="C", total="E", comp="F", curso="G", pend="H", venc="I", alert="J",
                 estado="K", msg="L").items()}


def cron(wb, r, area):
    c = wb[B.AREAS[B.PUESTOS.index(area)][0]]
    return c[f"K{r}"].value, c[f"L{r}"].value, c[f"N{r}"].value


# ---------------------------------------------------------------------------
print("Escenario A: Semana actual 2, fecha de corte 23/10/2026")
rows = [
    # RO: completa todo en Semana 2 (escenarios 3 y 4)
    act(1, RO, 0.40, (0.5, 1.0)),
    act(2, RO, 0.35, (1.0, None)),                    # 100% en S1, S2 vacía
    act(3, RO, 0.25, (0.3, 1.0)),
    # RLN: sin avance (1), 50% (2); pesos suman 90% (7)
    act(4, RLN, 0.50, (None, None)),
    act(5, RLN, 0.40, (0.5, None)),
    # RF: sin actividades (6)
    # ATR: avance decreciente (8), vencida (9), semanas intermedias vacías
    act(6, ATR, 0.50, (0.6, 0.4)),
    act(7, ATR, 0.30, (0.2, None, 0.5), fin=D(2026, 10, 20)),
    act(8, ATR, 0.20, (None, None), resp=None, fin=None),
]
wb = run("A", rows, semana=2, fecha_ref=D(2026, 10, 23), inicio_ciclo=D(2026, 10, 12))
L, M, P = cron(wb, 5, RO); check("Act. 1 (50%→100% en S2) = 100% Completado", (L, M) == (1, "Completado"), (L, M, P))
L, M, P = cron(wb, 6, RO); check("Act. 2 (100% en S1, S2 vacía) = 100%", L == 1, (L, M))
L, M, P = cron(wb, 5, RLN); check("Escenario 1: sin avances → 0% Pendiente", (L, M) == (0, "Pendiente"), (L, M))
L, M, P = cron(wb, 6, RLN); check("Escenario 2: 50% → En curso", (L, M) == (0.5, "En curso"), (L, M))
check("Escenario 7: alerta pesos ≠100% en fila", "Pesos del mes suman 90%" in P, P)
L, M, P = cron(wb, 5, ATR); check("Escenario 8: decreciente detectado y toma último valor (40%)",
                              L == 0.4 and "decreciente" in P, (L, P))
L, M, P = cron(wb, 6, ATR); check("Semana intermedia vacía: toma S3 (50%)", L == 0.5, L)
check("Escenario 9: vencida no completada", "Vencida" in P, P)
L, M, P = cron(wb, 7, ATR); check("Sin responsable y sin fecha límite señalados",
                              "Sin responsable" in P and "Sin fecha límite" in P, P)
L, M, P = cron(wb, 30, RF); check("Fila vacía → celdas calculadas vacías", (L, M, P) == (None, None, None), (L, M, P))

p = dash_puesto(wb, 0)
check("Escenario 4: RO 100% anticipado", p["avance"] == 1 and p["msg"] == "¡Enhorabuena, completado anticipadamente!", p)
p = dash_puesto(wb, 1)
check("RLN: pesos 90% → 'Revisar pesos', sin mensaje de completado",
      p["avance"] == "Revisar pesos" and p["msg"].startswith("⚠") and "90%" in p["estado"], p)
p = dash_puesto(wb, 2)
check("Escenario 6: RF sin actividades → no se marca completado",
      p["avance"] == "Sin actividades" and p["msg"] == "⚠ Registrar actividades" and p["total"] == 0, p)
p = dash_puesto(wb, 3)
exp = 0.5 * 0.4 + 0.3 * 0.5 + 0.2 * 0
check("ATR ponderado = 35%", abs(p["avance"] - exp) < 1e-9 and p["msg"] == "En ejecución", p)
check("ATR vencidas = 1", p["venc"] == 1, p)
d = wb[B.S_DASH]
check("Avance general 'No definitivo' (faltan puestos)", d["B5"].value == "No definitivo", d["B5"].value)
check("Mensaje de proyecto advierte 2 puestos", "NO definitivo: 2 puesto" in d["G12"].value, d["G12"].value)
check("KPI total actividades = 8", d["H5"].value == 8, d["H5"].value)
check("KPI completadas = 3", d["K5"].value == 3, d["K5"].value)
x = wb[B.S_CALC]
check("Completadas/En curso/Pendientes = 3/3/2",
      (x["N8"].value, x["O8"].value, x["P8"].value) == (3, 3, 2), (x["N8"].value, x["O8"].value, x["P8"].value))
check("Puestos al 100% = 1", x["B12"].value == 1, x["B12"].value)
# Seguimiento semanal RO: S1 = .4*.5+.35*1+.25*.3 = .625 ; S2 = 1
check("RO acumulado S1 = 62.5%, S2 = 100%",
      abs(d["C25"].value - 0.625) < 1e-9 and abs(d["D25"].value - 1) < 1e-9, (d["C25"].value, d["D25"].value))
check("RO incremento S2 = 37.5 p.p.", abs(d["I25"].value - 0.375) < 1e-9, d["I25"].value)
check("Semana 4 no visible (semana actual 2, sin datos en S4)", d["F25"].value in (None, ""), d["F25"].value)
# ATR semana 3 visible porque hay datos registrados en S3
check("Semana 3 visible si hay avances registrados", d["E28"].value not in (None, ""), d["E28"].value)
v = wb[B.AREAS[3][1]]   # hoja Actividades del Asistente Técnico (vista automática)
check("Actividades se alimenta del Cronograma (N°, nombre, responsable)",
      (v["A5"].value, v["B5"].value, v["C5"].value) == (1, "Actividad de prueba 6", "Resp. prueba"),
      (v["A5"].value, v["B5"].value, v["C5"].value))
check("Actividades: semana de entrega según fecha límite (20/10 → Semana 2; 06/11 → Semana 4)",
      (v["F6"].value, v["F5"].value) == ("Semana 2", "Semana 4"), (v["F6"].value, v["F5"].value))
check("Actividades: avance por semana copiado (S1 20%, S3 50%) y % actual 50%",
      (v["G6"].value, v["H6"].value, v["I6"].value, v["K6"].value) == (0.2, None, 0.5, 0.5),
      (v["G6"].value, v["H6"].value, v["I6"].value, v["K6"].value))
check("Actividades: estado Vencida / En curso / Pendiente", (v["L6"].value, v["L5"].value, v["L7"].value) == ("Vencida", "En curso", "Pendiente"),
      (v["L6"].value, v["L5"].value, v["L7"].value))
check("Actividades: fila sin datos en Cronograma queda vacía", (v["A8"].value, v["B8"].value, v["L8"].value) == (None, None, None),
      (v["A8"].value, v["B8"].value))
check("Actividades: semanas del ciclo (12/10, 19/10, 26/10, 02/11)",
      [v.cell(row=3, column=c).value.date() for c in range(7, 11)] == [D(2026, 10, 12), D(2026, 10, 19), D(2026, 10, 26), D(2026, 11, 2)],
      [v.cell(row=3, column=c).value for c in range(7, 11)])
ent = {d[f"B{r}"].value: (d[f"F{r}"].value, d[f"G{r}"].value, d[f"H{r}"].value) for r in range(60, 80) if d[f"B{r}"].value in B.PUESTOS}
check("Dashboard entregas: RO 3/3 completadas en Semana 4; ATR 1 vencida",
      ent.get(RO) == ("3 / 3", "3 / 3", 0) and ent.get(ATR)[2] == 1, ent)
qc = {d[f"B{r}"].value: d[f"G{r}"].value for r in range(50, 72) if d[f"B{r}"].value}
print("   Control de calidad:", qc)
check("QC: 1 puesto sin actividades", qc.get("Puestos sin actividades registradas") == 1, qc)
check("QC: 1 puesto con pesos ≠100% (ATR y RO suman 100%)",
      qc.get("Puestos cuya ponderación no suma 100% (o con pesos vacíos)") == 1, qc)
check("QC: decreciente = 1", qc.get("Avance semanal menor que el de una semana anterior") == 1, qc)
check("QC: vencidas = 1", qc.get("Actividades vencidas no completadas") == 1, qc)
check("QC: sin responsable = 1", qc.get("Actividades sin responsable") == 1, qc)

# ---------------------------------------------------------------------------
print("Escenario B: proyecto llega a 100% en la Semana 4 (semana actual 4)")
rowsB = [
    act(1, RO, 0.6, (0.5, 0.7, 0.9, 1.0)),
    act(2, RO, 0.4, (0.2, 0.5, 0.8, 1.0)),
    act(3, RLN, 1.0, (0.25, 0.5, 0.75, 1.0)),
    act(4, RF, 0.5, (0.3, 0.6, 0.9, 1.0)),
    act(5, RF, 0.5, (0.1, 0.4, 0.8, 1.0)),
    act(6, ATR, 0.25, (1.0, None, None, None)),   # completa en S1
    act(7, ATR, 0.75, (0.5, 0.8, 0.9, 1.0)),
]
wb = run("B", rowsB, semana=4, fecha_ref=D(2026, 11, 6))
d = wb[B.S_DASH]
check("Escenario 5: avance general = 100%", d["B5"].value == 1, d["B5"].value)
check("Proyecto completado (no anticipado) en S4", d["G12"].value == "Actividades del mes completadas: meta del 100% alcanzada.", d["G12"].value)
for i in range(4):
    p = dash_puesto(wb, i)
    check(f"{B.PUESTOS[i]}: 'Completado' sin felicitación", p["msg"] == "Completado", p)
check("Incremento general S4 calculado", isinstance(d["K29"].value, (int, float)), d["K29"].value)
check("Brecha vs meta = 0%", abs(d["L29"].value) < 1e-9, d["L29"].value)

# ---------------------------------------------------------------------------
print("Escenario C: proyecto completo en Semana 3 (semana actual 3) → anticipado")
rowsC = [dict(r) for r in rowsB]
for r in rowsC:
    r["s4"] = None
    if r["s3"] is not None:
        r["s3"] = 1.0
wb = run("C", rowsC, semana=3, fecha_ref=D(2026, 10, 30))
d = wb[B.S_DASH]
check("Proyecto 100% anticipado", d["G12"].value == "Actividades del mes: ¡Enhorabuena, completado anticipadamente!", d["G12"].value)
check("Total proyecto: mensaje de felicitación", d["L20"].value == "¡Enhorabuena, completado anticipadamente!", d["L20"].value)
# Si la condición deja de cumplirse, el mensaje desaparece
rowsC2 = [dict(r) for r in rowsC]
rowsC2[0]["s3"] = 0.9
wb = run("C2", rowsC2, semana=3, fecha_ref=D(2026, 10, 30))
d = wb[B.S_DASH]
check("Al bajar de 100% el mensaje cambia a 'En ejecución'",
      dash_puesto(wb, 0)["msg"] == "En ejecución" and "ejecución" in d["G12"].value, (dash_puesto(wb, 0), d["G12"].value))

# ---------------------------------------------------------------------------
print("Escenario D: nueva fila agregada a la tabla (fila 105, fuera de las 100 preformateadas)")
NEW = B.FIRST + B.N_ACT_ROWS   # primera fila fuera de las preformateadas
rowsD = [act(1, RO, 0.5, (1.0,), _row=5), act(2, RO, 0.5, (0.4,), _row=NEW)]
wb = run("D", rowsD, semana=1, fecha_ref=D(2026, 10, 16), extra_rows=1)
L, M, P = cron(wb, NEW, RO)
check("Fila nueva calcula avance/estado", (L, M) == (0.4, "En curso"), (L, M, P))
p = dash_puesto(wb, 0)
check("Dashboard incluye la fila nueva (RO = 70%)", abs(p["avance"] - 0.7) < 1e-9 and p["total"] == 2, p)
# ID duplicado y evidencia faltante
rowsE = [act(1, RO, 0.5, (0.6,), fin=D(2026, 10, 14), obs=None),   # vencida sin observación
         act(2, RO, 0.5, (0.2,), ini=D(2026, 10, 20), fin=D(2026, 10, 15))]
wb = run("E", rowsE, semana=1, fecha_ref=D(2026, 10, 16))
L, M, P = cron(wb, 5, RO); check("Vencida sin observación detectada", "Vencida sin observación" in P, P)
L, M, P = cron(wb, 6, RO); check("Fecha inicio posterior a fecha límite", "Inicio posterior" in P, P)
check("Puesto 100%? No (50%+20%·0.5)", dash_puesto(wb, 0)["msg"] == "En ejecución", dash_puesto(wb, 0))

# ---------------------------------------------------------------------------
ws_ro = wb[B.AREAS[0][0]]
check("ID automático 1, 2 y vacío en fila sin datos", (ws_ro["A5"].value, ws_ro["A6"].value, ws_ro["A7"].value) == (1, 2, None),
      (ws_ro["A5"].value, ws_ro["A6"].value, ws_ro["A7"].value))

print("Escenario F: libro vacío (entregable)")
wb = run("F", [], semana=1, inicio_ciclo=D(2026, 10, 12), fecha_ref=D(2026, 10, 19))
d = wb[B.S_DASH]
check("Sin actividades: no hay mensajes de completado",
      all(not str(dash_puesto(wb, i)["msg"]).startswith(("¡", "Completado")) for i in range(4)),
      [dash_puesto(wb, i)["msg"] for i in range(4)])
check("Mensaje de proyecto: no hay actividades", d["G12"].value.startswith("⚠ No hay actividades"), d["G12"].value)
h = wb[B.AREAS[0][1]]
check("Hoja Actividades vacía cuando el Cronograma está vacío", h["A5"].value is None and h["L5"].value is None, h["A5"].value)
hit = [d[f"G{r}"].value for r in range(60, 80) if isinstance(d[f"G{r}"].value, str) and "/" in d[f"G{r}"].value]
check("Resumen de entregas en Dashboard = 0 / 0 por puesto", hit == ["0 / 0"] * 4, hit)

# ---------------------------------------------------------------------------
print("Escenario G: configuración automática (mes nov-2026, corte 18/11/2026, sin semana manual)")
wb = run("G", [], mes=D(2026, 11, 1), fecha_ref=D(2026, 11, 18))
k = wb[B.S_CONF]
check("Inicio automático = primer lunes (02/11/2026)", k["C5"].value.date() == D(2026, 11, 2), k["C5"].value)
check("Semana actual automática = 3", k["C7"].value == 3, k["C7"].value)
h = wb[B.AREAS[1][1]]
check("Actividades: semana del ciclo calculada desde el inicio automático (02/11)", h["G3"].value.date() == D(2026, 11, 2), h["G3"].value)
check("Hoja renombrada a Riesgo Normativo", B.PUESTOS[1] == "Riesgo Normativo" and wb[B.S_DASH]["B17"].value == "Riesgo Normativo",
      wb[B.S_DASH]["B17"].value)

# ---------------------------------------------------------------------------
print("Escenario H: historial – octubre completo, noviembre en curso (se muestra noviembre)")
rowsH = [
    act(1, RO, 0.6, (0.5, 1.0)), act(2, RO, 0.4, (1.0,)),                                  # octubre (100%)
    act(3, RO, 0.5, (0.4,), ini=D(2026, 11, 2), fin=D(2026, 11, 20)),                   # noviembre
    act(4, RO, 0.5, (None,), ini=D(2026, 11, 9), fin=D(2026, 11, 27)),
    act(5, RLN, 1.0, (1.0,)),                                                             # octubre
]
wb = run("H", rowsH, mes=D(2026, 11, 1), fecha_ref=D(2026, 11, 6))
c = wb[B.AREAS[0][0]]
check("Columna Mes automática (oct/nov) y N° reinicia por mes",
      [(c[f"A{r}"].value, c[f"O{r}"].value.month) for r in range(5, 9)] == [(1, 10), (2, 10), (1, 11), (2, 11)],
      [(c[f"A{r}"].value, c[f"O{r}"].value) for r in range(5, 9)])
p = dash_puesto(wb, 0)
check("Dashboard de noviembre: RO = 20% con 2 actividades (octubre no se mezcla)",
      abs(p["avance"] - 0.2) < 1e-9 and p["total"] == 2, p)
check("Dashboard de noviembre: Riesgo Normativo sin actividades en el mes", dash_puesto(wb, 1)["total"] == 0, dash_puesto(wb, 1))
v = wb[B.AREAS[0][1]]
check("Actividades muestra solo noviembre, sin huecos",
      (v["B5"].value, v["B6"].value, v["B7"].value) == ("Actividad de prueba 3", "Actividad de prueba 4", None),
      (v["B5"].value, v["B6"].value, v["B7"].value))
hst = wb[B.S_HIST]
row = {(hst[f"B{r}"].value.year, hst[f"B{r}"].value.month): r for r in range(6, 30)}
check("Historial empieza en octubre 2026 (24 meses, hasta septiembre 2028)",
      hst["B6"].value.date() == D(2026, 10, 1) and hst["B29"].value.date() == D(2028, 9, 1), (hst["B6"].value, hst["B29"].value))
row = {m: row[(2026, m)] for m in (10, 11, 12)}
check("Historial: octubre RO 100% (2 / 2), RN 100%; noviembre RO 20%",
      hst[f"C{row[10]}"].value == 1 and hst[f"D{row[10]}"].value == "2 / 2" and hst[f"E{row[10]}"].value == 1
      and abs(hst[f"C{row[11]}"].value - 0.2) < 1e-9,
      (hst[f"C{row[10]}"].value, hst[f"D{row[10]}"].value, hst[f"E{row[10]}"].value, hst[f"C{row[11]}"].value))
check("Historial: diciembre sin datos", hst[f"C{row[12]}"].value == "Sin datos", hst[f"C{row[12]}"].value)
wb = run("H2", rowsH, mes=D(2026, 10, 15), fecha_ref=D(2026, 10, 30))
check("Revisar mes anterior (escribiendo una fecha de octubre): RO 100% con 2 actividades",
      dash_puesto(wb, 0)["avance"] == 1 and dash_puesto(wb, 0)["total"] == 2, dash_puesto(wb, 0))

print()
print("FALLAS:" if FAILS else "TODAS LAS PRUEBAS PASARON", FAILS if FAILS else "")
sys.exit(1 if FAILS else 0)
