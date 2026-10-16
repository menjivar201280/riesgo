"""Carga el «Cronograma de actividades – Informes semanal/mensual» de Riesgo Financiero (octubre 2026)
en la hoja «Cronograma R. Financiero». Cada actividad es una macroactividad; su % va en la Semana 2
(12/10 – 18/10). Uso: python datos/cargar_riesgo_financiero.py <salida.xlsx>"""
import sys, datetime as dt, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import build_cronograma as B

RF = B.PUESTOS[2]
INICIO = dt.date(2026, 10, 1)   # mantiene las actividades en octubre (el documento no trae fechas)
GRUPOS = {
    "Riesgo Crediticio": [
        ("Elaboración de informe para el Comité de Riesgos sobre modelos de medición del riesgo crediticio: "
         "pérdidas esperadas, matrices de transición, modelo de cosechas, etc.", 0.60, None),
    ],
    "Riesgo de Liquidez y Mercado": [
        ("Elaboración de informe para el Comité de Riesgos sobre modelos de medición del riesgo de mercado y de "
         "liquidez: volatilidad de las captaciones y tasas de interés, etc.", 0.60, None),
        ("Envío a la SSF del anexo 03 de la normativa NRP-05", 0.05, None),
        ("Elaboración del informe sobre escenarios regulatorios del riesgo de liquidez", 0.05, None),
    ],
    "Activos de Riesgo": [
        ("Envíos regulatorios de la normativa NCB-022", 1.0, None),
        ("Envío de información a la Central de Riesgo (NRP-41)", 1.0, None),
        ("Envío a la SSF del archivo CTRI-RECR (recuperaciones de activos de riesgo)", 1.0, None),
        ("Envío a la SSF del detalle de los créditos relacionados", 1.0, None),
        ("Revisión, cuadratura y envío a Contabilidad de la cartera de tarjeta de crédito y de la reserva de saneamiento", 1.0, None),
        ("Elaboración del informe de evaluación de los activos de riesgo", 1.0, None),
        ("Elaboración del informe de evaluación de los créditos relacionados", 1.0, None),
    ],
    "Otros": [
        ("Elaboración del informe de segmentación de los productos de COMEDICA", 1.0, None),
        ("Monitoreo semanal de la reserva de saneamiento", 0.50, None),
        ("Proyección mensual de los incrementos en la reserva de saneamiento", 1.0, None),
        ("Consulta en la Central de Riesgos de los aspirantes a empleados", 0.50, None),
        ("Registro permanente de los traslados de deuda", 0.50, None),
    ],
    "Proyectos a largo plazo": [
        ("Adecuación del sistema de créditos para dar cumplimiento a las reformas de la normativa NRP-41", 0.10, None),
        ("Revisión de expedientes de tarjetahabientes de alto riesgo, a solicitud del Consejo de Admón.", 1.0,
         "En el documento: 100% con estado «Proceso»"),
        ("Actualización de Políticas de Tarjeta de Crédito", 0.80, None),
    ],
}
rows = []
for grupo, acts in GRUPOS.items():
    for nombre, pct, obs in acts:
        rows.append(dict(area=RF, macro=f"{grupo} – {nombre}", ini=INICIO, fin=None, peso=None,
                         s1=None, s2=pct, s3=None, s4=None, obs=obs, resp=None))
B.build(sys.argv[1], test_rows=rows, extra_rows=0)
print(f"{len(rows)} actividades cargadas en {B.AREAS[2][0]}")
