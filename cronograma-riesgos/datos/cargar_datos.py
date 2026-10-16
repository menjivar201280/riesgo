"""Carga las actividades reales de octubre 2026 enviadas por cada puesto:
- Riesgo Financiero: «Cronograma de actividades – Informes semanal/mensual» (Word).
- Riesgo Normativo: lista enviada por mensaje.
Cada actividad es una macroactividad (el grupo va al inicio del nombre); su % va en la Semana 2
(12/10 – 18/10); sin % = vacío (Pendiente). Uso: python datos/cargar_datos.py <salida.xlsx>"""
import sys, datetime as dt, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import build_cronograma as B

RF = B.PUESTOS[2]
RN = B.PUESTOS[1]
INICIO = dt.date(2026, 10, 1)   # mantiene las actividades en octubre (el documento no trae fechas)
GRUPOS_RF = {
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
GRUPOS_RN = {
    "Requerimiento nota DS-SABAO-BCS-BC-22741": [
        ("Elaboración de matriz de requerimiento", None, None),
        ("Envío de requerimiento para dar respuesta a observaciones de visita SABAO-BCS-BC-18432", None,
         "Se envió correo electrónico a todas las áreas involucradas: Gerencia Financiera, Gerencia de Cumplimiento, "
         "Talento Humano, Depto. de Procesos, Administración de Riesgos, Depto. de Educación y Responsabilidad Social."),
        ("Envío de requerimiento para dar respuesta a plan de solución según nota DS-SABAO-BCS-BC-22741 "
         "(seguimiento a observaciones de visitas anteriores)", None,
         "Se envió correo a Administración de Riesgos, Auditoría Interna, Servicios Generales y Talento Humano."),
        ("Seguimiento al envío de las pruebas de descargo para dar respuesta a nota", None, None),
    ],
    "Requerimiento nota DS-SABAO-BCS-BC-23486": [
        ("Seguimiento al envío de pruebas de descargo para dar respuesta a nota", None, None),
    ],
    "Control de avalúos": [
        ("Descargo de avalúo anterior y documento mutuo hipotecario", None, None),
        ("Asignación a perito valuador", None, None),
        ("Gestiones para asignación de avalúo por medios electrónicos sin utilizar dispositivo extraíble", 0.30, None),
        ("Notas por medio de correo masivo para asociados, informando sobre actualización de avalúos", None, None),
    ],
    None: [
        ("Elaboración de acta de sesión de fecha 29/08/2026", 0.20, None),
    ],
}
rows = []
for area, grupos in ((RF, GRUPOS_RF), (RN, GRUPOS_RN)):
    for grupo, acts in grupos.items():
        for nombre, pct, obs in acts:
            rows.append(dict(area=area, macro=f"{grupo} – {nombre}" if grupo else nombre, ini=INICIO, fin=None,
                             peso=None, s1=None, s2=pct, s3=None, s4=None, obs=obs, resp=None))
B.build(sys.argv[1], test_rows=rows, extra_rows=0)
for area in (RF, RN):
    print(f"{sum(r['area'] == area for r in rows)} actividades cargadas en {area}")
