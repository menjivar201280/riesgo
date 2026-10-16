"""Carga las actividades reales de octubre 2026 enviadas por cada puesto:
- Riesgo Financiero: «Cronograma de actividades – Informes semanal/mensual» (Word).
- Riesgo Normativo: lista enviada por mensaje.
- Riesgo Operacional: «Cronograma de actividades – Informe semanal 05 al 17 octubre» de Silvia (Word).
- Asistente Técnico de Riesgo: «Cronograma de actividades – Informe semanal 13 al 17 octubre» de Héctor (Word).
Cada grupo es una MACROACTIVIDAD y sus actividades son SUBACTIVIDADES (todas valen igual dentro de su
macroactividad). El % de cada subactividad va en la Semana 2 (12/10 – 18/10); sin % = vacío (0%).
La macroactividad se llena sola con sus subactividades. Uso: python datos/cargar_datos.py <salida.xlsx>"""
import sys, datetime as dt, pathlib
sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
import build_cronograma as B

RF = B.PUESTOS[2]
RN = B.PUESTOS[1]
RO = B.PUESTOS[0]
ATR = B.PUESTOS[3]
INICIO = dt.date(2026, 10, 1)   # mantiene las actividades en octubre (los documentos no traen fechas)

# (macroactividad, % propio si no tiene subactividades, [(subactividad, %, observación), ...])
RIESGO_FINANCIERO = [
    ("Riesgo Crediticio", None, [
        ("Elaboración de informe para el Comité de Riesgos sobre modelos de medición del riesgo crediticio: "
         "pérdidas esperadas, matrices de transición, modelo de cosechas, etc.", 0.60, None),
    ]),
    ("Riesgo de Liquidez y Mercado", None, [
        ("Elaboración de informe para el Comité de Riesgos sobre modelos de medición del riesgo de mercado y de "
         "liquidez: volatilidad de las captaciones y tasas de interés, etc.", 0.60, None),
        ("Envío a la SSF del anexo 03 de la normativa NRP-05", 0.05, None),
        ("Elaboración del informe sobre escenarios regulatorios del riesgo de liquidez", 0.05, None),
    ]),
    ("Activos de Riesgo", None, [
        ("Envíos regulatorios de la normativa NCB-022", 1.0, None),
        ("Envío de información a la Central de Riesgo (NRP-41)", 1.0, None),
        ("Envío a la SSF del archivo CTRI-RECR (recuperaciones de activos de riesgo)", 1.0, None),
        ("Envío a la SSF del detalle de los créditos relacionados", 1.0, None),
        ("Revisión, cuadratura y envío a Contabilidad de la cartera de tarjeta de crédito y de la reserva de saneamiento", 1.0, None),
        ("Elaboración del informe de evaluación de los activos de riesgo", 1.0, None),
        ("Elaboración del informe de evaluación de los créditos relacionados", 1.0, None),
    ]),
    ("Otros", None, [
        ("Elaboración del informe de segmentación de los productos de COMEDICA", 1.0, None),
        ("Monitoreo semanal de la reserva de saneamiento", 0.50, None),
        ("Proyección mensual de los incrementos en la reserva de saneamiento", 1.0, None),
        ("Consulta en la Central de Riesgos de los aspirantes a empleados", 0.50, None),
        ("Registro permanente de los traslados de deuda", 0.50, None),
    ]),
    ("Proyectos a largo plazo", None, [
        ("Adecuación del sistema de créditos para dar cumplimiento a las reformas de la normativa NRP-41", 0.10, None),
        ("Revisión de expedientes de tarjetahabientes de alto riesgo, a solicitud del Consejo de Admón.", 1.0,
         "En el documento: 100% con estado «Proceso»"),
        ("Actualización de Políticas de Tarjeta de Crédito", 0.80, None),
    ]),
]
RIESGO_NORMATIVO = [
    ("Requerimiento según nota DS-SABAO-BCS-BC-22741", None, [
        ("Elaboración de matriz de requerimiento", None, None),
        ("Envío de requerimiento para dar respuesta a observaciones de visita SABAO-BCS-BC-18432", None,
         "Se envió correo electrónico a todas las áreas involucradas: Gerencia Financiera, Gerencia de Cumplimiento, "
         "Talento Humano, Depto. de Procesos, Administración de Riesgos, Depto. de Educación y Responsabilidad Social."),
        ("Envío de requerimiento para dar respuesta a plan de solución según nota DS-SABAO-BCS-BC-22741, "
         "en seguimiento a observaciones de visitas anteriores", None,
         "Se envió correo a Administración de Riesgos, Auditoría Interna, Servicios Generales y Talento Humano."),
        ("Seguimiento al envío de las pruebas de descargo para dar respuesta a nota", None, None),
    ]),
    ("Requerimiento según nota DS-SABAO-BCS-BC-23486", None, [
        ("Seguimiento al envío de pruebas de descargo para dar respuesta a nota", None, None),
    ]),
    ("Control de avalúos", None, [
        ("Descargo de avalúo anterior y documento mutuo hipotecario", None, None),
        ("Asignación a perito valuador", None, None),
        ("Gestiones para asignación de avalúo por medios electrónicos sin utilizar dispositivo extraíble", 0.30, None),
        ("Notas por medio de correo masivo para asociados, informando sobre actualización de avalúos", None, None),
    ]),
    ("Elaboración de acta de sesión de fecha 29/08/2026", 0.20, []),
]

RIESGO_OPERACIONAL = [   # informe semanal del 05 al 17 de octubre: sin subactividades
    ("Opinión caso asociado Jonathan Jacobo Marroquín", 1.0, []),
    ("Capacitación de GIR 2026", 0.90, []),
    ("Preparación de información de Auditoría Externa", 1.0, []),
    ("Reunión proyecto unificación de formulario Actualización de Datos (Atención al Asociado) "
     "y Conozca tu Cliente (Oficialía de Cumplimiento)", 0.35, []),
    ("Informe de segunda prueba de Continuidad del Negocio realizada el 01.10.2026", 0.75, []),
    ("Informe y presentación Acuerdos Semestrales del Comité de Riesgos para la Junta de Vigilancia", 1.0, []),
    ("Seguimiento planes de acción Matriz de Riesgos Institucional 2026", 1.0, []),
    ("Remisión de información a Auditoría Interna – NRP-24", 1.0, []),
    ("Remisión Matriz Administración de Sistemas Informáticos al departamento de Procesos", 1.0, []),
    ("Flujo de fábrica de créditos – Conclusión Administración de Riesgos", 1.0, []),
    ("Ventana de mantenimiento sistemas", 1.0, []),
    ("Solicitud cambio de estructura Departamento de Administración de Riesgos al departamento de Procesos", 0.50, []),
    ("Seguimiento eventos de riesgos por caídas de TC 25.09.2026 y Orion 26.09.2026", 0.50, []),
    ("Actualización página web – Atención Asociados", 1.0, []),
]
INICIO_RO = dt.date(2026, 10, 5)   # el informe cubre del 05 al 17 de octubre

ASISTENTE_TECNICO = [   # informe semanal del 13 al 17 de octubre: sin subactividades
    ("Realizar informe de Jonathan Jacobo Marroquín", 1.0, []),
    ("Informe de seguimiento de examen de gestión integral de riesgos", 0.90, []),
    ("Presentación resultados preliminares de Gestión de Riesgo Específico", 1.0, []),
    ("Comparativo de política de servicios tercerizados", 0.90, []),
    ("Comparativo de política de gobierno corporativo", 0.90, []),
    ("Presentación visita de supervisión gobierno corporativo", 1.0, []),
    ("Elaboración del cuestionario para la entrevista de BIA", 1.0, []),
]
INICIO_ATR = dt.date(2026, 10, 13)   # el informe cubre del 13 al 17 de octubre

rows = []
for area, macros, ini in ((RF, RIESGO_FINANCIERO, INICIO), (RN, RIESGO_NORMATIVO, INICIO),
                         (RO, RIESGO_OPERACIONAL, INICIO_RO), (ATR, ASISTENTE_TECNICO, INICIO_ATR)):
    for macro, pct, subs in macros:
        rows.append(dict(area=area, macro=macro, ini=ini, s2=pct))
        for nombre, spct, obs in subs:
            rows.append(dict(area=area, sub=nombre, ini=ini, s2=spct, obs=obs))
B.build(sys.argv[1], test_rows=rows, extra_rows=0)
for area in (RO, RN, RF, ATR):
    print(f"{sum(r['area'] == area for r in rows)} filas cargadas en {area}")
