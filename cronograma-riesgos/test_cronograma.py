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
        resp="Resp. prueba", evid="Enlace evidencia", **kw):
    s = tuple(s) + (None,) * (4 - len(s))
    d = dict(id=f"T-{i:02d}", area=area, macro=f"Actividad de prueba {i}", resp=resp, ini=ini,
             fin=fin, peso=peso, s1=s[0], s2=s[1], s3=s[2], s4=s[3], evid=evid)
    d.update(kw)
    return d


def run(name, rows, **kw):
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
    return c[f"L{r}"].value, c[f"M{r}"].value, c[f"P{r}"].value


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
wb = run("A", rows, semana=2, fecha_ref=D(2026, 10, 23))
L, M, P = cron(wb, 5, RO); check("Act. 1 (50%→100% en S2) = 100% Completado", (L, M) == (1, "Completado"), (L, M, P))
L, M, P = cron(wb, 6, RO); check("Act. 2 (100% en S1, S2 vacía) = 100%", L == 1, (L, M))
L, M, P = cron(wb, 5, RLN); check("Escenario 1: sin avances → 0% Pendiente", (L, M) == (0, "Pendiente"), (L, M))
L, M, P = cron(wb, 6, RLN); check("Escenario 2: 50% → En curso", (L, M) == (0.5, "En curso"), (L, M))
check("Escenario 7: alerta pesos ≠100% en fila", "Pesos del puesto suman 90%" in P, P)
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
check("Proyecto completado (no anticipado) en S4", d["G12"].value == "Proyecto completado: meta del 100% alcanzada.", d["G12"].value)
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
check("Proyecto 100% anticipado", d["G12"].value == "Proyecto: ¡Enhorabuena, completado anticipadamente!", d["G12"].value)
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
rowsD = [act(1, RO, 0.5, (1.0,), _row=5), act(2, RO, 0.5, (0.4,), _row=105)]
wb = run("D", rowsD, semana=1, fecha_ref=D(2026, 10, 16), extra_rows=1)
L, M, P = cron(wb, 105, RO)
check("Fila nueva calcula avance/estado", (L, M) == (0.4, "En curso"), (L, M, P))
p = dash_puesto(wb, 0)
check("Dashboard incluye la fila nueva (RO = 70%)", abs(p["avance"] - 0.7) < 1e-9 and p["total"] == 2, p)
# ID duplicado y evidencia faltante
rowsE = [act(1, RO, 0.5, (1.0,), evid=None), act(2, RO, 0.5, (0.2,), ini=D(2026, 11, 10)),
         act(1, RF, 1.0, (0.2,))]   # ID T-01 repetido en la hoja de otro puesto
wb = run("E", rowsE, semana=1, fecha_ref=D(2026, 10, 16))
L, M, P = cron(wb, 5, RO); check("Completada sin evidencia + ID duplicado", "sin evidencia" in P and "ID duplicado" in P, P)
L, M, P = cron(wb, 6, RO); check("Fecha inicio posterior a fecha límite", "Inicio posterior" in P, P)
check("Puesto 100%? No (50%+20%·0.5)", dash_puesto(wb, 0)["msg"] == "En ejecución", dash_puesto(wb, 0))

# ---------------------------------------------------------------------------
L, M, P = cron(wb, 5, RF); check("ID duplicado detectado entre hojas de distintos puestos", "ID duplicado" in P, P)
ws_rf = wb[B.AREAS[2][0]]
check("Columna Área / Puesto se llena sola", ws_rf["B5"].value == RF and ws_rf["B6"].value is None, ws_rf["B5"].value)

print("Escenario F: libro vacío (entregable)")
wb = run("F", [], semana=1)
d = wb[B.S_DASH]
check("Sin actividades: no hay mensajes de completado",
      all(not str(dash_puesto(wb, i)["msg"]).startswith(("¡", "Completado")) for i in range(4)),
      [dash_puesto(wb, i)["msg"] for i in range(4)])
check("Mensaje de proyecto: no hay actividades", d["G12"].value.startswith("⚠ No hay actividades"), d["G12"].value)
h = wb[B.AREAS[0][1]]
check("Fecha objetivo hito S1 = 16/10/2026", h["C5"].value.date() == D(2026, 10, 16), h["C5"].value)
check("Hitos RO: área automática y 4 hitos propios", h["D5"].value == RO and h["A8"].value == "H-RO-04" and h["D9"].value is None,
      (h["D5"].value, h["A8"].value, h["D9"].value))
hit = [d[f"G{r}"].value for r in range(60, 80) if isinstance(d[f"G{r}"].value, str) and "/" in d[f"G{r}"].value]
check("Resumen de hitos en Dashboard = 0 / 4 por puesto", hit == ["0 / 4"] * 4, hit)

print()
print("FALLAS:" if FAILS else "TODAS LAS PRUEBAS PASARON", FAILS if FAILS else "")
sys.exit(1 if FAILS else 0)
