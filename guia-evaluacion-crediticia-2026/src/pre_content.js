// Contenido: Metodología de Precalificación y Preevaluación Crediticia.
const { H1, H1nb, H2, H3, P, B, N, T, F, EX, LEGAL, GOOD, ALERT, NOTE, ficha } = require("./lib");

const presentacion = () => [
  H1nb("Presentación"),
  P("Este documento describe una metodología para precalificar y preevaluar solicitudes de crédito. Es decir, para decidir **en minutos y con poca información** si vale la pena que un solicitante avance a la evaluación completa, y con qué monto, plazo y cuota aproximados. Se dirige a oficiales de negocio, asesores de agencia, centros de contacto, canales digitales, analistas de crédito y áreas de riesgo de bancos, bancos cooperativos, sociedades de ahorro y crédito, cooperativas e instituciones de microfinanzas en El Salvador."),
  P("Es un complemento de la *Guía Integral de Evaluación Crediticia en El Salvador — Actualización 2026*. Allí se desarrolla la evaluación completa; aquí se desarrollan las **etapas previas**, que son el filtro de entrada del proceso de crédito."),
  H2("Objetivos"),
  B([
    "Dar una respuesta rápida, consistente y transparente al solicitante.",
    "Evitar que se armen expedientes que no tienen posibilidad de aprobarse, lo que ahorra tiempo al cliente y a la institución.",
    "Estimar con prudencia una **cuota máxima** y un **monto máximo indicativo** para orientar la solicitud desde el inicio.",
    "Detectar a tiempo las señales que impiden continuar (historial grave, sobreendeudamiento, falta de requisitos).",
    "Aplicar los **mismos criterios** a todos los solicitantes comparables, en cualquier canal.",
  ]),
  H2("Advertencias de uso"),
  ALERT([
    "**La precalificación no es una aprobación** ni un compromiso de otorgar el crédito. Solo indica si la solicitud puede avanzar y en qué rango aproximado. La decisión final depende de la evaluación integral con información verificada y de la instancia facultada.",
    "**Todos los parámetros numéricos** de este documento (factores de reconocimiento de ingresos, gasto mínimo, relación deuda-ingreso máxima, factor de prudencia, tolerancias, vigencias) son **propuestas metodológicas orientativas**. No son límites legales. Cada institución debe validarlos con su propia información histórica y aprobarlos formalmente antes de usarlos (Anexo G).",
    "**Los nombres, cifras y casos de ejemplo son ficticios.** Los montos están en dólares de los Estados Unidos (US$).",
  ], "Advertencias importantes"),
];

const cap1 = () => [
  H1("1. Qué es precalificar y qué es preevaluar"),
  H2("1.1 Definiciones en lenguaje sencillo"),
  T(["Etapa", "Qué es", "Con qué información", "Resultado", "Tiempo de referencia"], [
    ["Precalificación", "Primer filtro: comprueba si el solicitante cumple los requisitos básicos y si, con los datos que declara y su historial crediticio, podría pagar una cuota razonable.", "Datos declarados por el solicitante y consulta del historial crediticio, con su autorización.", "Precalificado / precalificado con ajuste / precalificado con condiciones / requiere análisis / no precalifica, más un monto máximo indicativo.", "Minutos (en el mismo contacto)."],
    ["Preevaluación", "Segundo filtro: verifica con documentos básicos los datos que determinaron la precalificación y vuelve a hacer el cálculo.", "Identificación, comprobantes de ingresos y estados de cuenta básicos, además del historial.", "Pasa a evaluación integral / pasa con ajuste / no continúa.", "Horas o pocos días."],
    ["Evaluación integral", "Análisis completo de capacidad, historial, garantías, riesgos y documentación.", "Expediente completo y verificado.", "Recomendación técnica e informe.", "Según el producto y la política."],
  ], [14, 30, 22, 22, 12], { size: 17, caption: "Tabla 1.1. Etapas del proceso de crédito" }),
  NOTE(["En la práctica, algunas instituciones llaman «precalificación» a las dos primeras etapas juntas, y otras llaman «preevaluación» a la primera. Lo importante no es el nombre, sino que cada etapa tenga **reglas escritas**, use **datos definidos** y produzca un **resultado registrado**. En este documento, la precalificación usa datos declarados y la preevaluación usa datos verificados."], "Sobre los nombres"),
  H2("1.2 Qué no es la precalificación"),
  B([
    "**No es una aprobación ni una oferta en firme.** No se debe comunicar como «crédito aprobado».",
    "**No sustituye la verificación.** Los datos declarados pueden contener errores o exageraciones.",
    "**No es un análisis de garantías.** La garantía se evalúa después y no corrige la falta de capacidad de pago.",
    "**No es un motivo para pedir documentos de más** desde el primer contacto: se pide lo mínimo necesario.",
  ]),
  H2("1.3 Por qué funciona"),
  P("La mayoría de las solicitudes que se rechazan después de una evaluación completa podían descartarse desde el inicio por tres razones: **historial crediticio grave**, **endeudamiento excesivo** o **incumplimiento de requisitos básicos**. La precalificación detecta esas tres situaciones en minutos. Además, la mayoría de solicitudes que se aprueban con un monto menor al pedido pueden orientarse desde el principio con un **monto indicativo**, lo que evita expectativas equivocadas."),
];

const cap2 = () => [
  H1("2. Principios y marco normativo de referencia"),
  H2("2.1 Principios de la metodología"),
  N([
    "**Prudencia:** cuando un dato es declarado y no verificado, se reconoce solo una parte (factor de reconocimiento) y se aplican gastos mínimos realistas.",
    "**Lo mínimo necesario:** se pide solo la información indispensable para la etapa.",
    "**Consistencia:** las mismas reglas valen en agencia, teléfono y canales digitales.",
    "**Transparencia:** se explica al solicitante qué significa su resultado, qué sigue y que no es una aprobación.",
    "**Trazabilidad:** cada precalificación queda registrada con fecha, datos usados, resultado y responsable.",
    "**Criterios objetivos:** las reglas se basan en la capacidad de pago, el comportamiento crediticio y los requisitos del producto, no en características personales ajenas al riesgo.",
    "**Revisión periódica:** los parámetros se recalibran con los resultados reales.",
  ]),
  H2("2.2 Disposiciones que inciden en la precalificación"),
  P("El detalle de cada disposición (identificación, autoridad, vigencia comprobada y enlace) está en la sección P.5 de la Guía Integral. Esta tabla resume **cómo afecta cada una a esta etapa**."),
  T(["Disposición", "Efecto en la precalificación", "Tipo"], [
    ["Ley de Regulación de los Servicios de Información sobre el Historial de Crédito de las Personas (D.L. 695, reformada por D.L. 128 de 2021)", "La consulta del historial debe hacerse con la autorización del solicitante y para la finalidad informada. Se debe guardar evidencia de la autorización.", "Obligación legal"],
    ["Ley para la Protección de Datos Personales (D.L. 144, 2024)", "Se recogen solo los datos necesarios y con finalidad informada. Se protege la información de quienes no continúan y se conserva solo el tiempo que fije la política.", "Obligación legal (su alcance concreto debe confirmarlo la asesoría legal)"],
    ["Ley de Protección al Consumidor (D.L. 776, 2005) y normas de transparencia (NCM-02)", "Se informa con claridad y sin inducir a error. El resultado se presenta como indicativo, no como aprobación.", "Obligación legal / norma técnica para sus sujetos obligados"],
    ["Ley Especial contra el Lavado de Activos, FT y FPADM (D.L. 426, vigente desde el 17/10/2025) y normativa técnica (NRP-36)", "La identificación básica y la revisión de alertas de cumplimiento forman parte de los filtros de elegibilidad. La debida diligencia completa se hace en etapas posteriores, según la política de cumplimiento.", "Obligación legal / norma técnica"],
    ["Ley contra la Usura y tasas máximas publicadas por el BCR", "La tasa que se use para calcular la cuota indicativa no puede superar la tasa máxima legal vigente para el segmento y monto.", "Obligación legal"],
    ["NPB4-49 (gestión del riesgo crediticio)", "Las reglas de precalificación forman parte de las políticas de otorgamiento, que deben estar aprobadas por la Junta Directiva en las entidades sujetas a la norma.", "Norma prudencial para sus sujetos obligados"],
  ], [32, 48, 20], { size: 17, caption: "Tabla 2.1. Marco normativo aplicado a la precalificación (consulta: 27/10/2026)" }),
  NOTE(["En la revisión realizada no se identificó ninguna disposición que fije un porcentaje máximo de relación deuda-ingreso aplicable a todas las entidades. Los límites que se usan en esta metodología son parámetros institucionales."]),
];

const cap3 = () => [
  H1("3. Flujo general del proceso"),
  H2("3.1 Etapas y resultado de cada una"),
  T(["Paso", "Etapa", "Qué se hace", "Si no se cumple"], [
    ["1", "Contacto y autorización", "Se explica el proceso, se obtiene la autorización para consultar el historial y se registran los datos mínimos.", "No se puede precalificar."],
    ["2", "Filtros de elegibilidad", "Se comprueban los requisitos básicos del producto y del solicitante (Capítulo 5).", "No precalifica (se explica el motivo)."],
    ["3", "Semáforo de historial", "Se clasifica el historial crediticio en verde, amarillo o rojo (Capítulo 6).", "Rojo: no precalifica."],
    ["4", "Capacidad estimada", "Se calculan la cuota máxima, la cuota solicitada, la holgura y el monto máximo indicativo (Capítulo 7).", "Sin capacidad: no precalifica."],
    ["5", "Matriz de decisión", "Se combinan el semáforo y la holgura para obtener el resultado (Capítulo 8).", "—"],
    ["6", "Comunicación y registro", "Se informa el resultado, sus condiciones y su vigencia, y se registra (Capítulo 9).", "—"],
    ["7", "Preevaluación", "Se verifican los datos clave con documentos y se vuelve a calcular (Capítulo 10).", "Se ajusta o no continúa."],
    ["8", "Evaluación integral", "Análisis completo según la Guía Integral.", "—"],
  ], [7, 20, 48, 25], { caption: "Tabla 3.1. Flujo de precalificación y preevaluación" }),
  H2("3.2 Diagrama de decisión"),
  T(null, [
    ["¿Autorización y datos mínimos completos?", "NO → no se puede precalificar", "SÍ ↓"],
    ["¿Cumple todos los filtros de elegibilidad?", "NO → No precalifica", "SÍ ↓"],
    ["¿Semáforo de historial rojo?", "SÍ → No precalifica", "NO ↓"],
    ["¿Cuota máxima mayor que cero?", "NO → No precalifica (sin capacidad)", "SÍ ↓"],
    ["Holgura = cuota solicitada ÷ cuota máxima", "Se aplica la matriz de decisión (Tabla 8.1)", "→ Resultado"],
  ], [44, 32, 24], { firstBold: true, zebra: true, caption: "Tabla 3.2. Secuencia de decisión (las reglas se aplican en este orden)" }),
  GOOD(["Aplicar las reglas **en este orden**. Si un filtro de elegibilidad no se cumple, no tiene sentido consultar el historial ni calcular la capacidad. Así se ahorra tiempo y se evitan consultas innecesarias al historial del solicitante."]),
];

const cap4 = () => [
  H1("4. Etapa 1: datos mínimos y autorización"),
  H2("4.1 Datos mínimos por tipo de solicitante"),
  T(["Dato", "Asalariado", "Independiente / profesional", "Negocio propio", "Por qué se necesita"], [
    ["Identificación (DUI u otro documento aplicable)", "Sí", "Sí", "Sí", "Identificar al solicitante y consultar su historial."],
    ["Edad", "Sí", "Sí", "Sí", "Comprobar la edad mínima y máxima del producto o de su seguro."],
    ["Actividad y antigüedad", "Empleador y antigüedad", "Profesión y años de ejercicio", "Giro y años de operación", "Estabilidad del ingreso."],
    ["Tipo de contrato / formalidad", "Indefinido, plazo fijo, servicios", "Inscripción, facturación", "Formal o informal", "Factor de reconocimiento del ingreso."],
    ["Ingreso mensual", "Neto según boleta", "Promedio mensual neto de gastos de la actividad", "Utilidad mensual estimada (ventas − costos − gastos)", "Base de la capacidad."],
    ["Otras fuentes de ingreso", "Sí", "Sí", "Sí", "Se reconocen con factores distintos."],
    ["Tamaño del hogar y gastos aproximados", "Sí", "Sí", "Sí", "Gasto mínimo de referencia."],
    ["Cuotas de deudas actuales", "Sí", "Sí", "Sí", "Se cruzan con el historial."],
    ["Producto, monto, plazo y destino", "Sí", "Sí", "Sí", "Cálculo de la cuota y admisibilidad del destino."],
  ], [22, 16, 20, 20, 22], { size: 17, caption: "Tabla 4.1. Datos mínimos de la precalificación" }),
  H2("4.2 Autorización de consulta"),
  B([
    "Antes de consultar el historial, se explica para qué se usará la información y se obtiene la **autorización del solicitante**, en el formato que defina la política (firma, aceptación digital con registro o grabación, según la ley aplicable y los medios autorizados).",
    "Se registran la fecha, la hora, el canal y la persona que obtuvo la autorización.",
    "Si el solicitante no autoriza la consulta, se le explica que no es posible precalificarlo con esta metodología.",
  ]),
  H2("4.3 Guion breve de entrevista"),
  P("El **Anexo B** contiene un guion de preguntas que sirve para cualquier canal. Las preguntas son neutrales, se entienden fácilmente y siempre siguen el mismo orden, para que todos los solicitantes reciban el mismo trato."),
];

const cap5 = () => [
  H1("5. Etapa 2: filtros de elegibilidad"),
  P("Los filtros de elegibilidad son **reglas de sí o no**. Si un solicitante no cumple alguno, no precalifica, sin importar cuál sea su capacidad. Cada filtro debe tener su **fuente** (ley, norma o política del producto) y un **motivo** que se pueda explicar al solicitante."),
  T(["Filtro", "Regla (parámetro institucional)", "Fuente de la regla", "Si no se cumple"], [
    ["Identidad", "Documento de identidad vigente y datos coincidentes.", "Política de identificación / cumplimiento.", "No precalifica hasta presentar documento válido."],
    ["Edad", "Dentro del rango del producto y de la cobertura del seguro de deuda, si se exige (por ejemplo, de 18 a 70 años al vencimiento).", "Política del producto y del seguro.", "No precalifica o se ajusta el plazo."],
    ["Alertas de cumplimiento", "Sin coincidencias en listas de control que defina la política de prevención de LA/FT.", "D.L. 426 y política de cumplimiento.", "Se escala a cumplimiento; no se continúa hasta resolver."],
    ["Antigüedad mínima de la actividad", "Por ejemplo: asalariado 6 meses; independiente o negocio 12 meses.", "Política del producto.", "No precalifica o requiere análisis."],
    ["Ingreso mínimo", "Ingreso considerado igual o mayor que el mínimo del producto.", "Política del producto.", "No precalifica."],
    ["Monto y plazo", "Dentro de los rangos del producto.", "Ficha del producto.", "Se ofrece ajustar al rango."],
    ["Destino", "Destino lícito y admisible para el producto.", "Política del producto.", "No precalifica para ese producto; se puede orientar a otro."],
    ["Tasa", "La tasa del producto no supera la tasa máxima legal vigente.", "Ley contra la Usura / publicación del BCR.", "No se puede ofrecer el producto en esas condiciones."],
    ["Solicitudes recientes rechazadas", "Sin rechazo de la misma institución en un período definido, salvo que haya cambiado la situación.", "Política institucional.", "Requiere análisis."],
  ], [17, 35, 23, 25], { size: 17, caption: "Tabla 5.1. Filtros de elegibilidad (valores ilustrativos)" }),
  ALERT(["Los filtros deben ser **objetivos y estar relacionados con el riesgo o con el producto**. No se deben crear filtros basados en características personales que no tengan relación demostrada con la capacidad o la voluntad de pago."], "Criterio de diseño"),
];

const cap6 = () => [
  H1("6. Etapa 3: semáforo del historial crediticio"),
  P("El semáforo resume el historial crediticio en **tres colores**. Así, todos aplican el mismo criterio y no se depende solo de una puntuación externa. Si se usa una puntuación de una agencia de información de datos, puede servir como dato adicional, pero las reglas del semáforo se aplican siempre."),
  T(["Color", "Criterios (cualquiera de ellos determina el color; parámetros orientativos)", "Efecto"], [
    ["VERDE", "Sin mora vigente. Sin atrasos mayores de 30 días en los últimos 12 meses. Sin créditos castigados ni en cobro judicial. Historial de al menos 12 meses. Como máximo 2 consultas de otras entidades en los últimos 3 meses.", "Se continúa con el cálculo de capacidad."],
    ["AMARILLO", "Sin historial o con historial de menos de 12 meses. Atrasos de 1 a 30 días en los últimos 6 meses. Un atraso de 31 a 60 días hace más de 6 meses, ya regularizado. Una reestructuración o refinanciamiento en los últimos 24 meses. Tres o más consultas en 3 meses. Uso de tarjetas de crédito mayor del 80 % del límite.", "Se continúa, pero el resultado queda condicionado (matriz de decisión)."],
    ["ROJO", "Mora vigente mayor de 30 días. Atrasos mayores de 60 días en los últimos 12 meses. Créditos castigados o en cobro judicial sin cancelar. Indicios de fraude o suplantación en el historial.", "No precalifica."],
  ], [12, 64, 24], { caption: "Tabla 6.1. Semáforo del historial crediticio" }),
  NOTE([
    "**No tener historial no es lo mismo que tener mal historial.** Se clasifica en amarillo y puede compensarse con evidencias alternativas en la preevaluación (pago puntual de servicios o alquiler, historial de ahorro, antigüedad laboral).",
    "**Deudas no declaradas:** si el historial muestra cuotas que el solicitante no mencionó, se usan las del historial y se registra una alerta de veracidad.",
  ]),
  H2("6.1 Cuotas que se toman del historial"),
  P("Para el cálculo de capacidad se suman las cuotas mensuales de **todas** las obligaciones vigentes que aparecen en el historial, además de las declaradas que no aparezcan en él. Para las tarjetas de crédito se aplica el criterio que defina la política (por ejemplo, el pago mínimo informado o un porcentaje del saldo). Las fianzas otorgadas a terceros se registran como alerta y, si el deudor principal tiene atrasos, se suman como obligación."),
];

const cap7 = () => [
  H1("7. Etapa 4: capacidad de pago estimada"),
  P("Esta es la parte central de la precalificación. Con datos todavía no verificados se calcula, de manera prudente, **cuánto podría pagar el solicitante cada mes sin comprometer su presupuesto** y, a partir de esa cuota, **qué monto podría solicitar como máximo**."),
  H2("7.1 Ingreso considerado: factores de reconocimiento"),
  P("Como el ingreso es declarado y todavía no está verificado, se reconoce solo un porcentaje según el tipo de fuente. El porcentaje es mayor cuanto más estable y fácil de verificar sea la fuente."),
  T(["Fuente de ingreso declarada", "Factor de reconocimiento propuesto", "Razón"], [
    ["Salario formal (con boleta o constancia), neto", "100 %", "Es estable y fácil de verificar."],
    ["Pensión", "100 %", "Es estable."],
    ["Salario sin constancia o informal", "60 %", "Es difícil de verificar y tiene mayor riesgo de discontinuidad."],
    ["Honorarios / trabajo independiente (neto de gastos de la actividad)", "70 %", "Es variable; se compensa la variabilidad."],
    ["Utilidad de negocio propio", "60 %", "Es variable y suele mezclarse con las finanzas del hogar."],
    ["Alquileres con contrato vigente", "80 %", "El inmueble puede quedar desocupado y hay gastos de mantenimiento."],
    ["Remesas familiares", "50 %", "Dependen de terceros; la fuente puede desaparecer."],
    ["Horas extra, comisiones o bonos no garantizados", "50 % del promedio", "No son recurrentes."],
    ["Ingresos ocasionales", "0 %", "No son recurrentes."],
  ], [42, 22, 36], { caption: "Tabla 7.1. Factores de reconocimiento (parámetros orientativos, sujetos a validación)" }),
  F("ingreso considerado", [
    "Ingreso considerado = Σ (Ingreso neto declarado de cada fuente × Factor de reconocimiento)",
  ]),
  H2("7.2 Gasto considerado: gasto mínimo de referencia"),
  T(["Personas en el hogar", "1", "2", "3", "4", "5 o más"], [
    ["Gasto mínimo de referencia mensual (US$)", "300", "400", "480", "560", "640 + 70 por persona adicional"],
  ], [36, 10, 10, 10, 10, 24], { caption: "Tabla 7.2. Gasto mínimo de referencia (valores ilustrativos: cada institución debe fijarlos con estadísticas oficiales de costo de vida o con su propia base histórica)" }),
  F("gasto considerado", [
    "Gasto considerado = el mayor entre (Gastos declarados) y (Gasto mínimo de referencia)",
    "Las cuotas de préstamos NO forman parte del gasto: se restan aparte, para no contarlas dos veces.",
  ]),
  H2("7.3 Disponible mensual"),
  ...ficha({
    def: "Dinero que queda al solicitante cada mes después de cubrir sus gastos y las deudas que ya tiene.",
    uso: "Muestra el espacio real que existe para una nueva cuota.",
    aplica: "Se usan el ingreso considerado, el gasto considerado y las cuotas del historial más las declaradas.",
    formula: "Disponible = Ingreso considerado − Gasto considerado − Cuotas de deudas actuales",
    ejemplo: "1,100 − 480 − 120 = **US$500**.",
    interp: "El solicitante tiene US$500 al mes para una nueva cuota y para imprevistos.",
    errores: "Usar el ingreso bruto. Aceptar gastos irrealmente bajos. Olvidar deudas que aparecen en el historial.",
    decision: "Si el disponible es cero o negativo, el solicitante no precalifica.",
  }),
  H2("7.4 Cuota máxima: el menor de dos límites"),
  P("La cuota máxima se calcula con **dos límites**, y se toma el menor. Así se protege tanto el presupuesto del hogar como el nivel total de endeudamiento."),
  F("cuota máxima de precalificación", [
    "Límite por prudencia = Disponible × Factor de prudencia (propuesto: 70 %)",
    "Límite por endeudamiento = Ingreso considerado × Relación deuda-ingreso máxima (propuesta: 40 %) − Cuotas de deudas actuales",
    "Cuota máxima = el menor de los dos límites (si el resultado es negativo, se toma como cero)",
  ]),
  T(["Límite", "Qué protege", "Ejemplo"], [
    ["Por prudencia (70 % del disponible)", "Deja un margen del 30 % del disponible para imprevistos, como gastos médicos, alzas de precios o meses de menor ingreso.", "500 × 0.70 = US$350"],
    ["Por endeudamiento (40 % del ingreso)", "Evita que la suma de todas las deudas supere una proporción prudente del ingreso, aunque el hogar declare gastos bajos.", "1,100 × 0.40 − 120 = 440 − 120 = US$320"],
    ["Cuota máxima", "Se toma el menor de los dos.", "**US$320**"],
  ], [28, 50, 22], { caption: "Tabla 7.3. Por qué se usan dos límites" }),
  NOTE(["El factor de prudencia del 70 % y la relación deuda-ingreso máxima del 40 % son **parámetros propuestos para la precalificación**. Son algo más conservadores que los rangos orientativos de la Guía Integral (zona de atención de 30 % a 45 %), porque en esta etapa los datos todavía no se han verificado. La institución puede fijar valores distintos por producto, una vez validados."]),
  H2("7.5 Cuota solicitada y holgura"),
  F("cuota solicitada (cuota fija)", [
    "Cuota = Monto × i ÷ [1 − (1 + i)^(−n)]     ·     i = tasa nominal anual ÷ 12     ·     n = número de cuotas",
  ]),
  F("holgura", [
    "Holgura = Cuota solicitada ÷ Cuota máxima",
    "Hasta 0.80: holgada · Más de 0.80 hasta 1.00: ajustada · Más de 1.00: excede la capacidad estimada",
  ]),
  P("La holgura indica **qué parte de la cuota máxima ocuparía el crédito solicitado**. Una holgura de 0.56 significa que la cuota pedida ocupa el 56 % de lo que el solicitante podría pagar como máximo. Una holgura de 1.13 significa que la cuota pedida supera en 13 % la cuota máxima."),
  H2("7.6 Monto máximo indicativo"),
  F("monto máximo indicativo (valor presente de la cuota máxima)", [
    "Monto máximo = Cuota máxima × [1 − (1 + i)^(−n)] ÷ i",
    "Monto indicativo que se comunica = el menor entre (Monto máximo), (Monto solicitado) y (Límite del producto)",
    "En hoja de cálculo: =VA(tasa/12; plazo; −cuota máxima)",
  ]),
  EX("cálculo completo paso a paso", [
    "Datos declarados: salario formal neto US$1,100; hogar de 3 personas; gastos declarados US$450; cuotas en el historial US$120 (tarjeta US$50 + préstamo US$70). Solicita US$5,000 a 36 meses y 17 % anual.",
    "Paso 1. Ingreso considerado = 1,100 × 100 % = **US$1,100**.",
    "Paso 2. Gasto considerado = mayor entre 450 y 480 = **US$480**.",
    "Paso 3. Disponible = 1,100 − 480 − 120 = **US$500**.",
    "Paso 4. Límite por prudencia = 500 × 0.70 = US$350. Límite por endeudamiento = 1,100 × 0.40 − 120 = US$320. **Cuota máxima = US$320**.",
    "Paso 5. Cuota solicitada: i = 0.17 ÷ 12 = 0.014167; (1.014167)^(−36) = 0.602648; cuota = 5,000 × 0.014167 ÷ (1 − 0.602648) = **US$178.26**.",
    "Paso 6. Holgura = 178.26 ÷ 320 = **0.56** (holgada).",
    "Paso 7. Monto máximo = 320 × (1 − 0.602648) ÷ 0.014167 = 320 × 28.0483 ≈ **US$8,975.47**. El monto solicitado de US$5,000 está por debajo de ese máximo.",
  ]),
  H2("7.7 Adaptaciones por tipo de solicitante"),
  T(["Tipo", "Ajuste en la precalificación"], [
    ["Asalariado con descuentos por planilla", "Las cuotas que ya se descuentan por planilla se cuentan como deudas actuales, aunque el solicitante declare su salario neto después de esos descuentos. Se debe aclarar en la entrevista para no contarlas dos veces ni omitirlas."],
    ["Independiente", "Se pide el promedio de los últimos meses y también el monto del mes más bajo. Si el mes más bajo es menor que el 60 % del promedio, se usa como ingreso declarado el promedio entre ambos."],
    ["Negocio propio", "La utilidad se estima como ventas − costo de lo vendido − gastos del negocio, y luego se resta el retiro para el hogar si los gastos del hogar no se cuentan aparte. Se recomienda una visita en la preevaluación."],
    ["Varias fuentes", "Cada fuente lleva su propio factor. Se recomienda comprobar que el solicitante seguiría precalificando sin la fuente menos estable."],
    ["Crédito con codeudor", "En la precalificación solo se considera el ingreso del solicitante principal. El codeudor se evalúa por separado con la misma metodología."],
  ], [26, 74], { caption: "Tabla 7.4. Adaptaciones" }),
];

const cap8 = () => [
  H1("8. Etapa 5: matriz de decisión"),
  P("El resultado de la precalificación se obtiene combinando **el color del semáforo** con **la holgura**, siempre que se hayan cumplido los filtros de elegibilidad y que la cuota máxima sea mayor que cero."),
  T(["Semáforo \\ Holgura", "Holgada (≤ 0.80)", "Ajustada (> 0.80 a 1.00)", "Excede (> 1.00)"], [
    ["VERDE", "PRECALIFICADO", "PRECALIFICADO", "PRECALIFICADO CON AJUSTE (se ofrece el monto o plazo indicativo)"],
    ["AMARILLO", "PRECALIFICADO CON CONDICIONES (verificación reforzada en la preevaluación)", "REQUIERE ANÁLISIS", "REQUIERE ANÁLISIS con monto ajustado"],
    ["ROJO", "NO PRECALIFICA", "NO PRECALIFICA", "NO PRECALIFICA"],
  ], [18, 27, 27, 28], { firstBold: true, caption: "Tabla 8.1. Matriz de decisión de precalificación (propuesta orientativa)" }),
  T(["Resultado", "Significado", "Qué sigue"], [
    ["Precalificado", "Con los datos declarados, el solicitante podría pagar la cuota solicitada.", "Preevaluación estándar."],
    ["Precalificado con ajuste", "Podría pagar una cuota menor que la solicitada.", "Se ofrece el monto máximo indicativo o un plazo mayor dentro del producto, y luego preevaluación."],
    ["Precalificado con condiciones", "Tiene capacidad, pero su historial está en amarillo.", "Preevaluación con verificación reforzada y evidencias alternativas; posible codeudor según la política."],
    ["Requiere análisis", "La situación no permite una conclusión rápida.", "Lo revisa un analista antes de pedir el expediente completo."],
    ["No precalifica", "No cumple un filtro, su historial está en rojo o no tiene capacidad.", "Se explica el motivo y, cuando sea posible, qué podría hacer para volver a solicitar."],
  ], [24, 40, 36], { caption: "Tabla 8.2. Resultados posibles" }),
  H2("8.1 Alertas que obligan a revisar el caso, aunque el resultado sea favorable"),
  B([
    "El ingreso declarado es muy alto para la actividad o la antigüedad indicadas.",
    "El historial muestra deudas que el solicitante no declaró.",
    "El destino no es claro o cambia durante la entrevista.",
    "El solicitante insiste en un monto muy superior al indicativo.",
    "Los datos de contacto coinciden con los de otros solicitantes no relacionados.",
    "Una tercera persona interviene de forma insistente en la solicitud.",
  ]),
  H2("8.2 Vigencia del resultado"),
  P("El resultado de la precalificación tiene una **vigencia limitada**, por ejemplo de 30 días calendario (parámetro institucional). Si pasa ese plazo, o si cambian los datos (monto, plazo, ingreso o deudas), se vuelve a precalificar."),
];

const cap9 = () => [
  H1("9. Etapa 6: comunicación al solicitante y registro"),
  H2("9.1 Cómo comunicar el resultado"),
  B([
    "Se indica que es una **precalificación**: un resultado **indicativo** basado en datos declarados y en el historial crediticio.",
    "Se informan el monto, el plazo y la cuota **aproximados**, y la tasa de referencia, sin presentarlos como condiciones aprobadas.",
    "Se explica qué documentos hacen falta para la siguiente etapa y la vigencia del resultado.",
    "Si no precalifica, se explica el motivo general de forma comprensible y respetuosa (por ejemplo, «el nivel actual de sus deudas no permite una nueva cuota»). Si el motivo está en el historial, se orienta al solicitante sobre su derecho a obtener una copia de su historial en la agencia de información de datos.",
  ]),
  P("El **Anexo F** contiene textos modelo para cada resultado."),
  H2("9.2 Qué se registra"),
  T(["Campo", "Contenido"], [
    ["Identificación", "Número de precalificación, fecha y hora, canal y responsable."],
    ["Autorización", "Evidencia de la autorización para consultar el historial."],
    ["Datos usados", "Ingresos por fuente, factores aplicados, hogar, gastos, deudas, producto, monto, plazo y tasa."],
    ["Resultados intermedios", "Semáforo, disponible, los dos límites, cuota máxima, cuota solicitada, holgura y monto máximo."],
    ["Resultado", "Categoría de la matriz y motivo."],
    ["Alertas", "Alertas detectadas y a quién se escalaron."],
    ["Comunicación", "Fecha y medio en que se informó al solicitante."],
    ["Vigencia", "Fecha de vencimiento del resultado."],
  ], [26, 74], { caption: "Tabla 9.1. Registro mínimo de la precalificación (Anexo E)" }),
];

const cap10 = () => [
  H1("10. Preevaluación: verificación básica y nuevo cálculo"),
  H2("10.1 Objetivo"),
  P("La preevaluación **comprueba con documentos los datos que decidieron la precalificación**. Es más rápida y más barata que la evaluación integral porque se concentra en los pocos datos que más influyen en el resultado: la identidad, el ingreso, las deudas y la estabilidad de la actividad."),
  H2("10.2 Documentos básicos por tipo de solicitante"),
  T(["Tipo", "Documentos básicos de preevaluación"], [
    ["Asalariado", "Documento de identidad; constancia de salario o boleta reciente; estado de cuenta donde se deposita el salario (según la política, por ejemplo de los últimos 3 meses)."],
    ["Independiente / profesional", "Documento de identidad; estados de cuenta del período que exija la política; última declaración de impuestos o comprobantes de retención."],
    ["Negocio propio", "Documento de identidad; estados de cuenta o registros de ventas; declaración de impuestos si existe; visita o fotografías del negocio según la política."],
    ["Fuentes adicionales", "Contrato de alquiler, comprobante de pensión o evidencia de recepción de remesas."],
  ], [26, 74], { caption: "Tabla 10.1. Documentos de preevaluación (la lista final la define cada producto)" }),
  H2("10.3 Tolerancias entre lo declarado y lo verificado"),
  T(["Diferencia entre el ingreso verificado y el declarado", "Acción (parámetros orientativos)"], [
    ["El verificado es igual, mayor o menor en no más de 10 %", "Se mantiene el resultado y se recalcula con el dato verificado."],
    ["El verificado es menor entre 10 % y 25 %", "Se recalcula con el dato verificado; puede cambiar el monto indicativo."],
    ["El verificado es menor en más de 25 %", "Se recalcula y se registra una **alerta de veracidad**. El caso pasa a revisión del analista."],
    ["Aparecen deudas no declaradas", "Se recalcula con las deudas del historial y se registra una alerta."],
    ["Documentos con indicios de alteración", "Se detiene el proceso y se escala según la política de prevención de fraude."],
  ], [42, 58], { caption: "Tabla 10.2. Tolerancias de la preevaluación" }),
  F("desviación del ingreso", [
    "Desviación (%) = (Ingreso verificado − Ingreso declarado) ÷ Ingreso declarado × 100",
  ]),
  H2("10.4 Nuevo cálculo con datos verificados"),
  P("En la preevaluación se repiten los cálculos del Capítulo 7, pero con el **ingreso neto verificado**. Sobre ese ingreso ya no se aplica el factor de reconocimiento de la Tabla 7.1, o se aplica un factor mayor que defina la política, porque el dato ya está comprobado. También se agrega **un escenario adverso simple** (por ejemplo, ingreso −10 % y gastos +10 %) para comprobar que la cuota sigue siendo pagable."),
  EX("preevaluación del ejemplo del Capítulo 7", [
    "Declarado: US$1,100. Verificado en boletas y abonos: US$1,060. Desviación = (1,060 − 1,100) ÷ 1,100 × 100 = **−3.6 %** (dentro de la tolerancia).",
    "Disponible = 1,060 − 480 − 120 = US$460. Límite por prudencia = 460 × 0.70 = US$322. Límite por endeudamiento = 1,060 × 0.40 − 120 = US$304. **Cuota máxima = US$304**.",
    "Holgura = 178.26 ÷ 304 = **0.59** (holgada).",
    "Escenario adverso: ingreso 1,060 × 0.90 = US$954; gastos 480 × 1.10 = US$528. Capacidad residual = 954 − 528 − 120 − 178.26 = **US$127.74** (positiva).",
    "**Resultado:** pasa a la evaluación integral con las condiciones solicitadas.",
  ]),
  H2("10.5 Resultado de la preevaluación"),
  T(["Resultado", "Criterio"], [
    ["Pasa a evaluación integral", "Datos dentro de la tolerancia; holgura de hasta 1.00; escenario adverso positivo; sin alertas pendientes."],
    ["Pasa con ajuste", "Holgura mayor que 1.00 con los datos verificados: se ajusta el monto o el plazo con el cliente antes de continuar."],
    ["Revisión del analista", "Alerta de veracidad, historial en amarillo con holgura ajustada, o escenario adverso negativo."],
    ["No continúa", "Cuota máxima igual a cero con los datos verificados, historial en rojo o indicios de fraude."],
  ], [28, 72], { caption: "Tabla 10.3. Resultados de la preevaluación" }),
];

const cap11 = () => [
  H1("11. Casos prácticos"),
  NOTE(["Casos **ficticios**. Se usan los parámetros propuestos: factores de la Tabla 7.1, gasto mínimo de la Tabla 7.2, factor de prudencia del 70 % y relación deuda-ingreso máxima del 40 %."], "Aviso"),
  H2("11.1 Caso A: asalariado con historial verde"),
  P("Corresponde al ejemplo del Capítulo 7: ingreso considerado US$1,100; gasto considerado US$480; deudas US$120; disponible US$500; cuota máxima US$320; solicita US$5,000 a 36 meses y 17 % (cuota US$178.26); holgura 0.56; monto máximo indicativo US$8,975.47. Historial: sin atrasos en 24 meses → **verde**."),
  P("**Resultado: PRECALIFICADO** por US$5,000. Pasa a la preevaluación estándar."),

  H2("11.2 Caso B: profesional independiente que pide más de lo que puede pagar"),
  T(["Paso", "Cálculo", "Resultado"], [
    ["Ingreso declarado", "Honorarios netos de gastos de la actividad", "US$2,400"],
    ["Ingreso considerado", "2,400 × 70 %", "US$1,680.00"],
    ["Gasto considerado", "Hogar de 4: mayor entre 500 declarado y 560 mínimo", "US$560.00"],
    ["Deudas actuales", "Historial: préstamo de vehículo", "US$200.00"],
    ["Disponible", "1,680 − 560 − 200", "US$920.00"],
    ["Límite por prudencia", "920 × 0.70", "US$644.00"],
    ["Límite por endeudamiento", "1,680 × 0.40 − 200 = 672 − 200", "US$472.00"],
    ["Cuota máxima", "Menor de los dos", "US$472.00"],
    ["Cuota solicitada", "US$22,000 a 60 meses, 16 % (i = 0.013333; factor 41.1217)", "US$535.00"],
    ["Holgura", "535.00 ÷ 472.00", "1.13 (excede)"],
    ["Monto máximo indicativo a 60 meses", "472 × 41.1217", "≈ US$19,409.45"],
    ["Alternativa a 72 meses, si el producto lo permite", "472 × 46.1003", "≈ US$21,759.33"],
  ], [32, 46, 22], { size: 18, totalRows: [7, 9] }),
  P("Historial: ocho años con pagos puntuales → **verde**. **Resultado: PRECALIFICADO CON AJUSTE.** Se le comunica que, con la información declarada, el monto indicativo es de hasta US$19,409.45 a 60 meses (cuota aproximada de US$472). Si el producto lo permite, también se le puede presentar la opción de hasta US$21,759.33 a 72 meses, explicándole que un plazo mayor aumenta el costo total. Con la cuota solicitada, la relación deuda-ingreso sería (200 + 535) ÷ 1,680 × 100 = 43.75 %, que supera el máximo propuesto del 40 %."),

  H2("11.3 Caso C: solicitante con varias deudas"),
  T(["Paso", "Cálculo", "Resultado"], [
    ["Ingreso considerado", "Salario formal neto 1,300 × 100 %", "US$1,300.00"],
    ["Gasto considerado", "Hogar de 4: mayor entre 600 declarado y 560 mínimo", "US$600.00"],
    ["Deudas actuales", "Préstamo personal 210 + tarjeta A 180 + tarjeta B 130", "US$520.00"],
    ["Disponible", "1,300 − 600 − 520", "US$180.00"],
    ["Límite por prudencia", "180 × 0.70", "US$126.00"],
    ["Límite por endeudamiento", "1,300 × 0.40 − 520 = 520 − 520", "US$0.00"],
    ["Cuota máxima", "Menor de los dos", "US$0.00"],
    ["Cuota solicitada", "US$6,000 a 48 meses, 18 %", "US$176.25"],
  ], [32, 46, 22], { size: 18, totalRows: [6] }),
  P("Historial: al día, pero con uso de tarjetas del 90 % del límite → **amarillo**. **Resultado: NO PRECALIFICA** por falta de capacidad, porque sus deudas actuales ya ocupan el 40 % de su ingreso. Si se aprobara, quedarían apenas US$3.75 de disponible al mes (180 − 176.25) y la relación deuda-ingreso subiría a (520 + 176.25) ÷ 1,300 × 100 = 53.56 %. **Orientación:** se le puede explicar la posibilidad de solicitar una consolidación de deudas, que sería una operación distinta y sujeta a su propia evaluación (véase el Caso 3 de la Guía Integral)."),

  H2("11.4 Caso D: joven sin historial crediticio"),
  T(["Paso", "Cálculo", "Resultado"], [
    ["Ingreso considerado", "Primer empleo formal (14 meses de antigüedad), neto 650 × 100 %", "US$650.00"],
    ["Gasto considerado", "Hogar de 1: mayor entre 280 declarado y 300 mínimo", "US$300.00"],
    ["Deudas actuales", "Ninguna", "US$0.00"],
    ["Disponible", "650 − 300 − 0", "US$350.00"],
    ["Límite por prudencia", "350 × 0.70", "US$245.00"],
    ["Límite por endeudamiento", "650 × 0.40 − 0", "US$260.00"],
    ["Cuota máxima", "Menor de los dos", "US$245.00"],
    ["Cuota solicitada", "US$2,500 a 24 meses, 20 % (i = 0.016667; factor 19.6480)", "US$127.24"],
    ["Holgura", "127.24 ÷ 245.00", "0.52 (holgada)"],
    ["Monto máximo indicativo", "245 × 19.6480", "≈ US$4,813.76"],
  ], [32, 46, 22], { size: 18, totalRows: [6] }),
  P("Historial: no tiene → **amarillo**. **Resultado: PRECALIFICADO CON CONDICIONES.** En la preevaluación se verifican la constancia de salario y los abonos en cuenta, y se piden evidencias alternativas de pago puntual (servicios, alquiler). Según la política del producto, también podría pedirse un codeudor. Si el salario fuera informal y sin constancia, el ingreso considerado sería 650 × 60 % = US$390, el disponible 390 − 300 = US$90 y la cuota máxima 90 × 0.70 = US$63. Con ese ingreso, la cuota de US$127.24 excedería la capacidad estimada."),

  H2("11.5 Caso E: buena capacidad, historial en rojo"),
  P("Ingreso formal neto de US$1,500; hogar de 4 (gasto considerado US$560); deudas de US$180. Disponible = 1,500 − 560 − 180 = US$760. Cuota máxima = el menor entre 760 × 0.70 = US$532 y 1,500 × 0.40 − 180 = US$420, es decir, **US$420**. Solicita US$8,000 a 48 meses y 16 % (cuota US$226.72; holgura 0.54)."),
  P("Historial: tarjeta de crédito con **mora vigente de 45 días** → **rojo**. **Resultado: NO PRECALIFICA**, aunque su capacidad estimada sea holgada. **Orientación:** que regularice la obligación en mora. Podrá volver a solicitar cuando lo permita la política (por ejemplo, después de un período de pagos puntuales)."),
  H2("11.6 Síntesis"),
  T(["Caso", "Semáforo", "Cuota máxima", "Cuota solicitada", "Holgura", "Monto máximo indicativo", "Resultado"], [
    ["A", "Verde", "320.00", "178.26", "0.56", "8,975.47", "Precalificado"],
    ["B", "Verde", "472.00", "535.00", "1.13", "19,409.45 (60 m)", "Precalificado con ajuste"],
    ["C", "Amarillo", "0.00", "176.25", "—", "0", "No precalifica (capacidad)"],
    ["D", "Amarillo", "245.00", "127.24", "0.52", "4,813.76", "Precalificado con condiciones"],
    ["E", "Rojo", "420.00", "226.72", "0.54", "14,819.90", "No precalifica (historial)"],
  ], [8, 12, 13, 14, 10, 19, 24], { size: 17, caption: "Tabla 11.1. Resultados de los casos (US$)" }),
];

const cap12 = () => [
  H1("12. Controles, riesgos del proceso e indicadores"),
  H2("12.1 Riesgos propios de la precalificación y controles"),
  T(["Riesgo", "Ejemplo", "Control recomendado"], [
    ["Datos declarados inflados", "El solicitante exagera su ingreso para obtener un monto mayor.", "Factores de reconocimiento; tolerancias en la preevaluación; registro de la desviación."],
    ["Manipulación por el personal", "Se cambian datos para que el solicitante precalifique y así cumplir metas.", "Calculadora institucional bloqueada; registro con usuario; muestreo de calidad; incentivos que no premien solo el volumen."],
    ["Criterios desiguales", "Dos asesores dan resultados distintos a casos iguales.", "Reglas escritas, misma herramienta en todos los canales y guion único."],
    ["Expectativa indebida", "El solicitante entiende que su crédito ya está aprobado.", "Textos modelo (Anexo F); la palabra «indicativo» siempre visible."],
    ["Consultas sin autorización", "Se consulta el historial sin autorización registrada.", "El sistema no permite consultar sin registrar la autorización."],
    ["Exceso de datos", "Se piden documentos que no hacen falta en esta etapa.", "Lista de datos mínimos (Tabla 4.1)."],
    ["Fraude de identidad", "Se usa el documento de otra persona.", "Validación de identidad en la preevaluación; alertas por datos de contacto repetidos."],
  ], [22, 36, 42], { caption: "Tabla 12.1. Riesgos y controles del proceso" }),
  H2("12.2 Indicadores para medir si la metodología funciona"),
  T(["Indicador", "Fórmula", "Para qué sirve"], [
    ["Tasa de precalificación", "Precalificados (todas las categorías favorables) ÷ Solicitudes precalificadas × 100", "Indica qué tan estrictos son los filtros."],
    ["Tasa de conversión", "Créditos aprobados ÷ Precalificados × 100", "Si es baja, la precalificación deja pasar casos que luego se rechazan."],
    ["Desviación media del ingreso", "Promedio de (verificado − declarado) ÷ declarado × 100", "Sirve para calibrar los factores de reconocimiento."],
    ["Diferencia de monto", "Monto aprobado ÷ Monto indicativo", "Indica si el monto indicativo es realista."],
    ["Mora temprana de los precalificados", "Créditos con atraso mayor de 30 días a los 6 meses ÷ Créditos desembolsados × 100, por categoría de resultado", "Comprueba si las categorías ordenan bien el riesgo."],
    ["Tiempo de respuesta", "Minutos desde el contacto hasta el resultado", "Mide la eficiencia del proceso."],
  ], [26, 44, 30], { caption: "Tabla 12.2. Indicadores de desempeño" }),
  EX("calibración de un factor de reconocimiento", [
    "En seis meses, 120 independientes precalificados pasaron a preevaluación. Su ingreso verificado fue, en promedio, 78 % del ingreso declarado.",
    "El factor propuesto (70 %) es algo más conservador que lo observado (78 %), así que sigue siendo prudente. Si lo observado hubiera sido 60 %, el factor del 70 % sobreestimaría la capacidad y habría que reducirlo.",
  ], { split: false }),
  H2("12.3 Revisión y validación"),
  B([
    "Revisar los parámetros al menos una vez al año, o antes si los indicadores cambian de manera significativa.",
    "Comparar la mora de los créditos según su categoría de precalificación (los «precalificados» deberían tener menos mora que los «precalificados con condiciones»).",
    "Analizar los casos «no precalifica» que luego se aprobaron por excepción, y cómo se comportaron.",
    "Documentar cada cambio de parámetro con su justificación y su aprobación.",
  ]),
];

const cap13 = () => [
  H1("13. Implementación"),
  N([
    "**Aprobar los parámetros** del Anexo G por la instancia que corresponda.",
    "**Configurar la herramienta de cálculo** (hoja de cálculo protegida o módulo del sistema) con las fórmulas del Capítulo 7 y la matriz del Capítulo 8.",
    "**Integrar la autorización de consulta** y la consulta del historial en el flujo, de modo que no se pueda consultar sin autorización.",
    "**Capacitar al personal** con los casos del Capítulo 11 y con los textos modelo del Anexo F.",
    "**Aplicar una prueba piloto** en algunas agencias o canales durante un período definido, y medir los indicadores del Capítulo 12.",
    "**Ajustar y extender** la metodología al resto de la red.",
  ]),
  GOOD(["Usar la precalificación también para **orientar al solicitante**: si no precalifica, explicarle qué tendría que cambiar (reducir deudas, regularizar una mora, mostrar más antigüedad) contribuye a una inclusión financiera responsable."]),
];

const anexos = () => [
  H1("Anexo A. Ficha de precalificación"),
  T(null, [
    ["N.º de precalificación / fecha / canal", " "],
    ["Responsable", " "],
    ["Solicitante / documento de identidad", " "],
    ["Autorización de consulta (fecha, hora, medio)", " "],
    ["Producto / monto / plazo / tasa / destino", " "],
  ], [40, 60], { firstBold: true, caption: "1. Datos generales" }),
  T(["Filtro", "Cumple", "Observación"], [
    ["Identidad", "☐ Sí  ☐ No", " "], ["Edad dentro del rango", "☐ Sí  ☐ No", " "], ["Alertas de cumplimiento", "☐ Sin alertas  ☐ Escalado", " "],
    ["Antigüedad mínima", "☐ Sí  ☐ No", " "], ["Ingreso mínimo", "☐ Sí  ☐ No", " "], ["Monto y plazo dentro del producto", "☐ Sí  ☐ No", " "],
    ["Destino admisible", "☐ Sí  ☐ No", " "], ["Tasa dentro de la tasa máxima legal", "☐ Sí  ☐ No", " "],
  ], [40, 25, 35], { caption: "2. Filtros de elegibilidad", size: 17 }),
  T(["Fuente de ingreso", "Monto declarado", "Factor", "Ingreso considerado"], [
    [" ", " ", " ", " "], [" ", " ", " ", " "], [" ", " ", " ", " "], ["Total", "", "", " "],
  ], [40, 20, 15, 25], { caption: "3. Ingresos", size: 17, totalRows: [3] }),
  T(["Concepto", "Valor"], [
    ["Personas en el hogar / gasto declarado / gasto mínimo", " "],
    ["Gasto considerado (el mayor)", " "],
    ["Cuotas de deudas actuales (historial + declaradas)", " "],
    ["Semáforo del historial (verde / amarillo / rojo) y motivo", " "],
    ["Disponible", " "],
    ["Límite por prudencia (disponible × factor)", " "],
    ["Límite por endeudamiento (ingreso × DTI máx. − deudas)", " "],
    ["Cuota máxima", " "],
    ["Cuota solicitada", " "],
    ["Holgura", " "],
    ["Monto máximo indicativo", " "],
    ["RESULTADO (matriz 8.1)", " "],
    ["Alertas", " "],
    ["Vigencia hasta", " "],
  ], [55, 45], { caption: "4. Cálculo y resultado", size: 17 }),

  H1("Anexo B. Guion de entrevista de precalificación"),
  N([
    "«Le explicaremos cómo funciona esta precalificación: es una estimación inicial y no es una aprobación. ¿Nos autoriza a consultar su historial crediticio para este fin?»",
    "«¿Qué producto le interesa, qué monto necesita, a qué plazo y para qué lo va a usar?»",
    "«¿A qué se dedica? ¿Desde cuándo?»",
    "«¿Cuál es su ingreso mensual neto, es decir, lo que efectivamente recibe? ¿Tiene otros ingresos? ¿De qué tipo?»",
    "(Independientes) «¿Cuánto recibe en un mes normal y cuánto en el mes más bajo del año? ¿Cuánto gasta en su actividad?»",
    "«¿Cuántas personas viven en su hogar? ¿Cuánto gastan aproximadamente al mes en vivienda, alimentación, servicios, transporte, salud y educación?»",
    "«¿Qué préstamos, tarjetas o fianzas tiene actualmente y cuánto paga al mes por ellos? ¿Le descuentan alguna cuota por planilla?»",
    "«Con esta información, el resultado indicativo es ___. Para continuar, necesitaremos los siguientes documentos: ___. Este resultado es válido hasta ___.»",
  ]),

  H1("Anexo C. Hoja de cálculo de precalificación"),
  T(["Celda", "Concepto", "Fórmula sugerida"], [
    ["B1", "Ingreso considerado", "=SUMAPRODUCTO(montos; factores)"],
    ["B2", "Gasto considerado", "=MAX(gasto_declarado; gasto_minimo)"],
    ["B3", "Cuotas actuales", "Dato (historial + declaradas)"],
    ["B4", "Disponible", "=B1−B2−B3"],
    ["B5", "Límite por prudencia", "=MAX(0; B4*factor_prudencia)"],
    ["B6", "Límite por endeudamiento", "=MAX(0; B1*dti_max−B3)"],
    ["B7", "Cuota máxima", "=MIN(B5; B6)"],
    ["B8", "Cuota solicitada", "=PAGO(tasa/12; plazo; −monto)"],
    ["B9", "Holgura", "=SI(B7>0; B8/B7; \"sin capacidad\")"],
    ["B10", "Monto máximo indicativo", "=SI(B7>0; VA(tasa/12; plazo; −B7); 0)"],
    ["B11", "Monto indicativo comunicado", "=MIN(B10; monto; limite_producto)"],
    ["B12", "Resultado", "Fórmula de la matriz 8.1 según semáforo y B9"],
  ], [10, 34, 56], { size: 17 }),
  NOTE(["En la versión de Excel en español con separador de punto y coma. Las celdas de parámetros (factores, gasto mínimo, factor de prudencia, DTI máximo) deben estar en una hoja separada y protegida, con su fecha de aprobación."]),

  H1("Anexo D. Lista de verificación de la preevaluación"),
  T(["Verificación", "Cumple", "Observación"], [
    ...["Identidad validada", "Ingreso verificado con documentos", "Desviación del ingreso dentro de la tolerancia", "Deudas del historial conciliadas con las declaradas", "Antigüedad de la actividad comprobada", "Nuevo cálculo con datos verificados", "Escenario adverso positivo", "Alertas resueltas o escaladas", "Resultado de la preevaluación registrado"].map((x) => [x, "☐ Sí  ☐ No", " "]),
  ], [50, 20, 30], { size: 17 }),

  H1("Anexo E. Registro de precalificaciones"),
  T(["N.º", "Fecha", "Canal", "Responsable", "Semáforo", "Cuota máx.", "Holgura", "Monto indicativo", "Resultado", "Vigencia"], Array.from({ length: 12 }, () => Array(10).fill(" ")), [7, 10, 9, 12, 10, 10, 9, 12, 12, 9], { size: 16 }),

  H1("Anexo F. Textos modelo de comunicación"),
  T(["Resultado", "Texto sugerido"], [
    ["Precalificado", "«Con la información que nos proporcionó, su solicitud cumple los criterios iniciales para un crédito de hasta US$___ a ___ meses, con una cuota aproximada de US$___. Esto no es una aprobación: para continuar necesitamos verificar los documentos que le indicamos. Este resultado es válido hasta el ___.»"],
    ["Precalificado con ajuste", "«Con la información proporcionada, el monto que podríamos analizar es de hasta US$___ a ___ meses (cuota aproximada de US$___), menor al que usted solicitó. Si está de acuerdo, continuamos con la verificación de documentos.»"],
    ["Precalificado con condiciones", "«Su solicitud puede continuar, pero necesitaremos información adicional: ___. Una vez la recibamos, haremos la verificación correspondiente.»"],
    ["Requiere análisis", "«Su solicitud necesita una revisión adicional de un analista. Le contactaremos a más tardar el ___ para informarle cómo continuar.»"],
    ["No precalifica", "«Por ahora no es posible continuar con esta solicitud porque ___ (por ejemplo, el nivel actual de sus deudas no permite una nueva cuota). Si desea, podemos orientarle sobre las opciones disponibles. Usted tiene derecho a solicitar una copia de su historial crediticio en la agencia de información de datos.»"],
  ], [24, 76], { size: 17 }),

  H1("Anexo G. Parámetros que la institución debe aprobar"),
  T(["Parámetro", "Valor propuesto en este documento", "Valor aprobado", "Fecha / instancia"], [
    ["Factores de reconocimiento por fuente", "Tabla 7.1", " ", " "],
    ["Gasto mínimo de referencia por tamaño de hogar", "Tabla 7.2", " ", " "],
    ["Factor de prudencia", "70 % del disponible", " ", " "],
    ["Relación deuda-ingreso máxima", "40 % del ingreso considerado", " ", " "],
    ["Criterio de cuota para tarjetas de crédito", "Pago mínimo informado o % del saldo", " ", " "],
    ["Bandas de holgura", "≤ 0.80 / 0.80–1.00 / > 1.00", " ", " "],
    ["Criterios del semáforo", "Tabla 6.1", " ", " "],
    ["Filtros de elegibilidad por producto", "Tabla 5.1", " ", " "],
    ["Tolerancias de la preevaluación", "10 % / 25 %", " ", " "],
    ["Escenario adverso de la preevaluación", "Ingreso −10 %, gastos +10 %", " ", " "],
    ["Vigencia de la precalificación", "30 días", " ", " "],
    ["Período de espera tras un rechazo", "Según política", " ", " "],
  ], [34, 30, 18, 18], { size: 17 }),

  H1("Anexo H. Glosario breve"),
  T(["Término", "Definición"], [
    ["Precalificación", "Filtro inicial con datos declarados y el historial crediticio, que da un resultado indicativo."],
    ["Preevaluación", "Verificación básica con documentos de los datos que decidieron la precalificación."],
    ["Factor de reconocimiento", "Porcentaje del ingreso declarado que se acepta según el tipo de fuente."],
    ["Gasto mínimo de referencia", "Gasto mensual mínimo que se supone para un hogar según su tamaño."],
    ["Disponible", "Ingreso considerado menos gasto considerado menos las cuotas actuales."],
    ["Factor de prudencia", "Parte del disponible que se permite usar para la nueva cuota."],
    ["Relación deuda-ingreso máxima", "Proporción máxima del ingreso que pueden ocupar todas las cuotas."],
    ["Cuota máxima", "El menor entre el límite por prudencia y el límite por endeudamiento."],
    ["Holgura", "Cuota solicitada dividida entre la cuota máxima."],
    ["Monto máximo indicativo", "Monto que corresponde a la cuota máxima, según el plazo y la tasa."],
    ["Semáforo del historial", "Clasificación del historial crediticio en verde, amarillo o rojo."],
    ["Tolerancia", "Diferencia aceptada entre el dato declarado y el verificado."],
  ], [30, 70], { size: 17 }),
];

module.exports = () => [presentacion(), cap1(), cap2(), cap3(), cap4(), cap5(), cap6(), cap7(), cap8(), cap9(), cap10(), cap11(), cap12(), cap13(), anexos()];
