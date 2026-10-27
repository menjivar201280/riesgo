---
title: "Evaluación y segmentación de riesgos de la base pre evaluada de crédito"
subtitle: "Hipotecario, vehículo y tarjeta de crédito: respuesta al memorando del 27/10/2026"
---

> **Estado del documento:** propuesta técnica para el área de Riesgos. Ningún rango, nivel o parámetro de este informe está aprobado. Los umbrales son **estadísticos**, salen de esta base y **no son límites regulatorios**. La segmentación es descriptiva y no sustituye la decisión crediticia.

Archivos analizados (solo los de esta ejecución):

- **Memorando:** "Validación de base pre evaluada de Crédito Hipotecario, vehículo y Tarjeta de Crédito". Fecha 27/10/2026. Lo dirigen Créditos, Comercial TC y Mercadeo al Jefe de Riesgos.
- **Base:** `Base_de_Asociados_pre_evaluada_creditos_y_tc_para_analisis_de_Riesgos.xlsx`, hoja `BD FINAL`.
- **Libro de trabajo entregado:** `Segmentacion_Riesgos_Base_PreEvaluada.xlsx`. Contiene 143,869 fórmulas y 0 errores. Cada número de este informe sale de ese libro.

---

# 1. Respuesta directa: ¿la base aplica con las condiciones del memorando?

**No en su totalidad.** Con los datos del archivo, **528 de 2,262 asociados (23.3%) cumplen a la vez todas las condiciones que se pueden verificar**. Del resto:

- 1,536 (67.9%) tienen información incompleta.
- 198 (8.8%) incumplen al menos un criterio.

Cuatro condiciones del memorando **no se pueden verificar** con la base porque el archivo no tiene esos datos:

- lista negra;
- créditos castigados;
- personas relacionadas o empleados;
- solvencia en Hacienda.

## 1.1 Criterios de la base depurada (memorando, criterios 1 a 8)

| # | Criterio del memorando | Cumplen | No cumplen | Sin dato | ¿Aplica? |
|---|---|---:|---:|---:|---|
| 1 | Poseen ingresos en sistema | 844 | 135 (ingreso = $0) | 1,283 (vacío) | **No.** Es el criterio más incumplido: 62.7% sin ingreso válido |
| 2 | Poseen fondos compensatorios | 2,207 | 55 (fondo = $0) | 0 | **No** en 55 registros |
| 3 | No poseen tarjeta de crédito | 1,296 | 966 | 0 | **No** para la base completa. **Sí** para el producto TC: ningún tarjetahabiente tiene monto de TC; los 966 solo tienen hipotecario o vehículo |
| 4 | Antigüedad como asociado "0 meses" | 1,416 | 0 | 846 (843 vacíos + 3 imposibles) | No verificable en 846. La redacción es ambigua |
| 5 | Categoría A1 y A2 interna | 1,378 | 0 | 884 (sin categoría) | Nadie fuera de A1/A2. Falta definir cómo tratar a quien no tiene historial |
| 6 | Categoría A1 y A2 externa (SSF) | 1,824 | 0 | 438 | Nadie fuera de A1/A2. Faltan 342 sin registro SSF y 96 con la calificación en blanco |
| 7 | Edad de 30 a 55 años cumplidos | 2,254 | **8** | 0 | **No** en 8. Al corte del 30/09/2026 ya tenían 56 años |
| 8 | Sin personas relacionadas ni empleados | — | — | — | **No verificable:** la base no tiene ese campo |

**Hallazgo sobre la edad.** La columna `EDAD` coincide con la fecha de nacimiento si se calcula al **02/09/2026** (2,250 de 2,262 casos), no al corte del memorando (30/09/2026). La base tiene vigencia de 3 meses, así que más asociados pasarán de 55 años antes de que venza.

## 1.2 Otros criterios del memorando

| # | Lo que dice el memorando | Lo que muestra la base | ¿Se confirma? |
|---|---|---|---|
| 1 | Los límites se asignaron según ingresos y fondos compensatorios | Las tablas de TC, vehículo e hipotecario solo se reproducen con `INGRESO`. **1,418 asociados con monto pre evaluado no tienen ingreso válido.** En ellos, el fondo tampoco explica el monto: solo 93 de 743 en vehículo | **Parcial.** El ingreso que se usó para asignar los límites no está en el archivo |
| 2 | Se excluyó a quienes están en lista negra o tienen créditos castigados | No hay columna para verificarlo | **No verificable** |
| 3 | Hay 61 registros sin deuda en la SSF | 342 sin registro SSF y 169 con saldo ≤ $0 | **No coincide.** Ninguna definición razonable da 61 |
| 4 | Hay 54 asociados que no están al día en aportaciones | Hay **186** que no están al día. 54 de ellos tienen monto de TC | Coincide solo para TC. **132** (hipotecario y vehículo) no se mencionan |
| 5 | TC > $4,000 sin solvencia en Hacienda → límite de $4,000 | 298 asociados con límite > $4,000. Con el tope, la exposición en TC baja de $3,986,500 a $3,305,000 | Regla aplicable. La solvencia no está en la base |
| 6 | Hipotecario > $30,000 → verificar solvencia en Hacienda | El monto mínimo es $110,000, así que aplica a los **679** | Todos requieren verificación externa |
| 7 | Vigencia de 3 meses | No hay fecha de aprobación | No verificable |

## 1.3 Consolidado por producto: memorando vs base

| Producto | Memorando | Base | Resultado |
|---|---|---|---|
| Tarjeta de crédito | 1,249 / $3,986,500 | 1,249 / $3,986,500 | Concilia en los 10 tramos de límite |
| Vehículo | 1,340 / $34,185,000 | **1,341 / $34,335,000** | **No concilia:** el asociado 20403 tiene $150,000. Ese monto no existe en la tabla (máximo $40,000) y no aparece en el memorando |
| Hipotecario | 679 / $89,950,000 | 679 / $89,950,000 | Concilia |
| Registros | 2,263 | **2,262** asociados | Diferencia de 1: 2,263 son las filas del archivo contando el encabezado |

---

# 2. Diagnóstico ejecutivo de la base

| Aspecto | Resultado |
|---|---|
| Registros y columnas | 2,262 asociados y 31 columnas con encabezado. Hay además 4 columnas sin encabezado, con 3 valores de prueba |
| Identificadores | `ASOCIADO` y `DUI` sin duplicados: **un registro por asociado** |
| Datos de formato | 3 DUI con longitud distinta de 9. `CONCATENADO` está vacía |
| Ingreso | 56.7% vacío y 6.0% en $0. Solo 844 asociados (37.3%) tienen ingreso > 0 |
| Antigüedad | 37.3% vacía. Hay 3 valores imposibles: −3992, −1 y 1776 años |
| Central de Riesgo (SSF) | Un saldo y **una sola institución** por asociado. 342 sin registro, 166 con saldo $0 y 3 con saldo negativo mínimo |
| Cuotas mensuales de préstamos | **No existen en la base** |
| Registros eliminados | **Ninguno.** Las columnas originales se conservan intactas |

**Conclusión del diagnóstico.** La base sirve para segmentar por fondo, antigüedad, categoría y deuda SSF. **No permite medir la capacidad de pago de la mayoría de asociados**, porque el ingreso falta en 62.7% de los casos y ninguna obligación trae su cuota mensual.

---

# 3. Rangos y niveles propuestos (pendientes de validación)

Cada límite inferior está incluido en su rango. Los porcentajes se calculan sobre los 2,262 asociados.

La hoja **`Variables_Rangos`** del libro presenta todas las variables con el mismo formato de la hoja "Variables" del archivo (rango → nivel). El nivel de cada asociado está en `BD_CLASIFICADA`, en las columnas `NIVEL_ANTIGUEDAD`, `NIVEL_INGRESO`, `NIVEL_DEUDA`, `NIVEL_DEUDA_INGRESO`, `NIVEL_RESPALDO_FONDO`, `NIVEL_COBERTURA_FONDO`, `NIVEL_APORTACIONES`, `NIVEL_MORA_APORTACIONES`, `NIVEL_CAT_INTERNA`, `NIVEL_CAT_EXTERNA` y `TRAMO_EDAD`.

## 3.1 Antigüedad como asociado (`ANTIGUEDAD_ANIOS`, años completos)

Los niveles son los solicitados. La antigüedad mide permanencia, no capacidad de pago.

| Nivel | Rango | Asociados | % |
|---|---|---:|---:|
| Alto | 0 a 1 año | 74 | 3.3% |
| Medio | 2 a 4 años | 212 | 9.4% |
| Bajo | 5 años o más | 1,130 | 50.0% |
| Sin dato | vacío (no se trata como 0) | 843 | 37.3% |
| Inconsistente | valor negativo o mayor que la edad | 3 | 0.1% |

## 3.2 Ingreso mensual (`INGRESO`, USD)

Distribución de los ingresos mayores que cero (n = 844):

- mínimo $218.90;
- P25 $1,484.40;
- mediana $2,416.31;
- promedio $3,718.49;
- P75 $4,156.46;
- P90 $6,660.93;
- máximo $132,222.90.

La base no indica si el ingreso es **bruto o disponible**. Por eso se trata como ingreso registrado.

| Rango | Límites | Nivel (capacidad) | Asociados | % total | % con ingreso > 0 | Cómo se fijó el límite |
|---|---|---|---:|---:|---:|---|
| I | $0.01 – $1,499.99 | Alto | 227 | 10.0% | 26.9% | ≈ P25; coincide con el corte de RANGOS TC |
| II | $1,500 – $2,999.99 | Medio | 237 | 10.5% | 28.1% | P25–P60; $3,000 es el mínimo de la tabla de vehículo |
| III | $3,000 – $4,999.99 | Bajo | 234 | 10.3% | 27.7% | P60–P83; tramo de vehículo $20,000 e hipotecario $110,000 |
| IV | $5,000 o más | Muy bajo | 146 | 6.5% | 17.3% | > P83; tramos superiores de vehículo e hipotecario |
| Ingreso cero | $0 | Revisar | 135 | 6.0% | — | No se calcula ninguna razón |
| Sin dato | vacío | Revisar | 1,283 | 56.7% | — | No se calcula ninguna razón |

Se propone un cuarto rango porque la distribución tiene una cola larga: P95 = $10,561 y P99 = $21,721. También se marcan **29 ingresos atípicos**, los mayores de $12,172.65 (P75 + 3 × rango intercuartílico). Se señalan para verificación y **no se excluyen**.

## 3.3 Saldo de deuda en la Central de Riesgo (`SALDO_SSF`, USD)

Distribución de los saldos mayores que cero (n = 1,751): P25 $571, mediana $3,334, P75 $19,609, P90 $73,940 y máximo $1,579,063.

| Segmento | Límites | Asociados | % |
|---|---|---:|---:|
| D1 Deuda baja | $0.01 – $999.99 (≈ P33) | 584 | 25.8% |
| D2 Deuda intermedia | $1,000 – $9,999.99 (P33–P63) | 536 | 23.7% |
| D3 Deuda elevada | $10,000 – $49,999.99 (P63–P87) | 404 | 17.9% |
| D4 Deuda muy elevada | $50,000 o más (> P87, nivel de saldos hipotecarios) | 227 | 10.0% |
| Sin saldo | saldo $0 o a favor (−$0.02 a −$0.89) | 169 | 7.5% |
| Sin registro SSF | sin saldo ni institución | 342 | 15.1% |

Cómo leer esta tabla:

- **"Sin registro SSF" no equivale a "sin deuda".** No se puede distinguir entre un asociado que no tiene deudas y uno que no fue consultado.
- **Una deuda elevada no implica incumplimiento.**
- Cada asociado trae una sola institución, así que no se sabe si `SALDO_SSF` es el total consolidado. No se sumó nada para no duplicar saldos.

## 3.4 Cuotas, carga financiera y relación cuota-ingreso

**No disponible en la base.** `CUOTAS_VENC_APORT` cuenta cuotas de **aportación** vencidas, no cuotas de préstamos. Por eso no se calcula la carga financiera mensual ni la relación cuota-ingreso. Tampoco se reemplazan las cuotas por saldos.

## 3.5 Relación deuda-ingreso (saldo SSF ÷ ingreso mensual, en **veces** el ingreso)

Este indicador mide el **saldo** acumulado, no la cuota mensual. Solo se calcula cuando el ingreso es mayor que cero y hay dato SSF (n = 769). Mediana 1.21x, P75 7.88x, P90 20.88x.

| Nivel | Rango | Asociados | % |
|---|---|---:|---:|
| Sin deuda | 0x | 56 | 2.5% |
| R1 Bajo | más de 0 y menos de 3x (≈ P60) | 421 | 18.6% |
| R2 Medio | 3x a menos de 12x | 156 | 6.9% |
| R3 Alto | 12x a menos de 24x (1 a 2 años de ingreso) | 73 | 3.2% |
| R4 Muy alto | 24x o más | 63 | 2.8% |
| No calculable | ingreso vacío o $0 | 1,418 | 62.7% |
| No calculable | sin dato SSF | 75 | 3.3% |

## 3.6 Fondo compensatorio (`FONDO_COMPENSATORIO`, USD)

Distribución de los fondos mayores que cero (n = 2,207): mínimo $1, P25 $2,591, mediana $12,692, P75 $32,205, P90 $69,808 y máximo $1,570,712.

| Segmento | Límites | Asociados | % |
|---|---|---:|---:|
| Sin fondo | $0 | 55 | 2.4% |
| F1 Bajo | $0.01 – $2,999.99 (≈ P28) | 609 | 26.9% |
| F2 Medio | $3,000 – $14,999.99 (P28–P53) | 558 | 24.7% |
| F3 Alto | $15,000 – $49,999.99 (P53–P85) | 706 | 31.2% |
| F4 Muy alto | $50,000 o más | 334 | 14.8% |

**Relación fondo ÷ ingreso**, en meses de ingreso (solo con ingreso > 0):

| Nivel | Rango | Asociados |
|---|---|---:|
| M1 | menos de 1 mes | 349 |
| M2 | 1 a menos de 6 meses | 190 |
| M3 | 6 a menos de 12 meses | 86 |
| M4 | 12 meses o más | 164 |
| Sin fondo | — | 55 |

**Validación del campo (opinión de Riesgo, punto 3):**

- La base no dice si el fondo está **disponible, retenido o pignorado**. Por eso se presenta como respaldo **descriptivo** y **no se le atribuye reducción de riesgo**.
- `MONTO_TOTAL_APORTACIONES` es un concepto distinto (mediana $1,443). En 312 asociados el fondo es menor que sus aportaciones, lo que confirma que no son la misma cifra.

## 3.7 Otros criterios con rango y nivel

| Criterio (columna) | Rango → nivel | Asociados |
|---|---|---|
| Cuotas de aportación vencidas (`CUOTAS_VENC_APORT`), **nivel de riesgo** | 0 (al día) → Bajo · 1 a 2 → Medio · 3 o más → Alto | 2,076 · 79 · 107 |
| Categoría interna (`MAX_CAT_RIESGO`), **nivel de riesgo** | A1 → Bajo · A2 → Medio · B a E → Alto · sin categoría → Sin información | 1,356 · 22 · 0 · 884 |
| Calificación SSF (`CALIFICACION_SSF`), **nivel de riesgo** | A1 → Bajo · A2 → Medio · B a E → Alto · sin calificación → Sin información | 1,813 · 11 · 0 · 438 |
| Aportaciones acumuladas (`MONTO_TOTAL_APORTACIONES`), nivel de aportación (no es riesgo) | menos de $500 → Bajo · $500 a $1,999.99 → Medio · $2,000 o más → Alto | 446 · 1,201 · 615 |
| Fondo ÷ ingreso, nivel de respaldo (no es riesgo) | menos de 1 mes → Bajo · 1 a menos de 6 → Medio · 6 a menos de 12 → Alto · 12 o más → Muy alto | 349 · 190 · 86 · 164 |
| Edad al 30/09/2026, tramo descriptivo | 30–39 · 40–49 · 50–55 · mayor de 55 | 574 · 876 · 804 · 8 |
| Tiene crédito vigente / tiene tarjeta | SÍ / NO | 823 / 1,439 · 966 / 1,296 |

Por qué algunos criterios no llevan nivel de riesgo:

- **Edad:** se presenta solo como tramo descriptivo. Asignarle un nivel de riesgo requiere una política institucional.
- **Aportaciones y fondo:** indican vínculo y respaldo con la cooperativa. No se les atribuye reducción de riesgo mientras no se valide si están disponibles.

Cortes usados:

- Aportaciones: $500 (≈ P20) y $2,000 (≈ P73).
- Cuotas vencidas: 1 cuota (cualquier atraso) y 3 cuotas (un trimestre de atraso).

Todos estos cortes se pueden editar en la hoja `Parametros`.

## 3.8 Historial y clasificación

| Variable | Disponible | Categorías encontradas |
|---|---|---|
| Categoría interna (`MAX_CAT_RIESGO`) | Sí | A1: 1,356; A2: 22; sin historial (sin crédito vigente): 866; dato faltante con crédito vigente: 18 |
| Calificación externa (`CALIFICACION_SSF`) | Sí | A1: 1,813; A2: 11; en blanco con saldo: 96; sin registro: 342 |
| Aportaciones al día | Sí | SÍ: 2,076; NO: 186. Las cuotas vencidas van de 1 a 172 |
| Tiene crédito vigente / tiene tarjeta | Sí | Créditos: SÍ 823. Tarjeta: SÍ 966 |
| Días de mora, castigos, listas, número de obligaciones, historial de pago, solvencia tributaria | **No disponible en la base** | Requieren consulta a otra fuente |

---

# 4. Evaluación conjunta y alertas

No se asigna un nivel global ponderado. La base **no contiene resultados históricos de pago**, así que no hay forma de estimar ni validar ponderaciones. En su lugar se cuentan alertas por asociado.

| Alerta de riesgo | Definición | Asociados |
|---|---|---:|
| Fondo bajo con deuda elevada | Fondo $0 o F1, con deuda D3 o D4 | 228 |
| Aportaciones vencidas | No está al día o tiene cuotas vencidas | 186 |
| Deuda alta frente al ingreso | Relación deuda-ingreso de 12x o más | 136 |
| Ingreso bajo con deuda elevada | Rango de ingreso I con deuda D3 o D4 | 49 |
| Categoría A2 | A2 interna o externa | 31 |
| Antigüedad baja sin historial | 0 a 1 año y sin categoría interna | 30 |

| Alerta de calidad de datos | Asociados |
|---|---:|
| Monto pre evaluado sin ingreso válido | **1,418** |
| Monto pre evaluado que no coincide con la tabla (por ejemplo, vehículo con ingreso < $3,000) | 221 |
| Dato inconsistente (antigüedad, edad, DUI, ingreso atípico, saldo negativo) | 198 |

**Estado descriptivo por asociado:**

| Estado | Asociados | % |
|---|---:|---:|
| Sin alertas identificadas | 534 | 23.6% |
| Sin alertas, pero sin ingreso válido (no se pudo evaluar) | 1,258 | 55.6% |
| 1 alerta de riesgo | 337 | 14.9% |
| 2 o más alertas de riesgo | 133 | 5.9% |

Hay que distinguir cuatro niveles de lectura:

1. **Segmentación descriptiva:** los rangos de la sección 3.
2. **Alerta de riesgo:** las combinaciones de esta sección.
3. **Elegibilidad preliminar:** los criterios del memorando (sección 1).
4. **Decisión crediticia:** corresponde al Comité, con consultas externas y análisis individual.

---

# 5. Evaluación por producto

| Indicador | Tarjeta | Vehículo | Hipotecario |
|---|---:|---:|---:|
| Asociados pre evaluados | 1,249 | 1,341 | 679 |
| Exposición pre evaluada | $3,986,500 | $34,335,000 | $89,950,000 |
| Coincide con la tabla según el ingreso | 378 | 379 | 380 |
| No coincide con la tabla | 2 | 219 | 38 |
| Sin ingreso válido | 869 (69.6%) | 743 (55.4%) | 261 (38.4%) |
| Mediana monto ÷ ingreso mensual | 1.6x | 5.9x | 28.1x |
| Cumplen los criterios verificables | 152 | 426 | 301 |
| Con 1 o más alertas de riesgo | 246 | 259 | 179 |

Observaciones por producto:

- **Tarjeta.** La tabla funciona bien cuando hay ingreso. Hay dos diferencias de nombre con la tabla: la base llama "GOLD" al segmento que la tabla llama "DORADA", y rotula como "CLASICA INTERNACIONAL" el límite de $500, que la tabla define como "CLÁSICA LOCAL". Aplica el tope de $4,000 a 298 asociados si no son solventes en Hacienda.
- **Vehículo.** La tabla excluye ingresos menores de $3,000 ("NO APLICA"), pero **219** asociados con ingreso menor recibieron monto. Además está el caso de $150,000.
- **Hipotecario.** El archivo **no trae tabla**. La regla que se deduce de los datos es: $3,000 → $110,000; $5,000 → $150,000; alrededor de $9,000 → $205,000. Hay que confirmarla. Los 679 requieren solvencia en Hacienda. El monto equivale a una mediana de 28 veces el ingreso mensual, pero sin plazo ni tasa no se puede calcular la cuota.
- **Otros créditos de consumo.** No hay montos en la base, así que no se pueden evaluar.
- No se aplican los mismos límites a todos los productos. Para fijar límites por producto se necesitan la cuota, el plazo, la tasa y la garantía.

---

# 6. Metodología, supuestos y restricciones

**Fórmulas principales.** Todas están documentadas en la hoja `Metodologia` y en los comentarios de cada columna de `BD_CLASIFICADA`.

- Edad al corte = año del corte − año de nacimiento − 1, si todavía no había cumplido años en la fecha de corte.
- Relación deuda-ingreso = MAX(SALDO_SSF; 0) ÷ INGRESO, en veces. Solo se calcula si INGRESO > 0.
- Fondo ÷ ingreso = FONDO_COMPENSATORIO ÷ INGRESO, en meses de ingreso.
- Relación cuota-ingreso = Σ cuotas ÷ INGRESO × 100. **No calculada**: no hay cuotas en la base.
- Monto esperado por producto = LOOKUP(INGRESO; tabla de la hoja Parametros), que se compara con el monto PRECA de la base.

**Supuestos (pendientes de validación):**

- La fecha de corte es el 30/09/2026, según el memorando.
- El ingreso es mensual y en USD.
- `SALDO_SSF` es el saldo reportado por la Central de Riesgo, sin confirmar que esté consolidado.
- "Antigüedad 0 meses" se interpreta como "sin antigüedad mínima".
- Las aportaciones no al día no excluyen al asociado; condicionan la TC, como indica el memorando.
- La edad mínima de afiliación es 18 años. Solo se usa para una alerta.

**Controles de calidad realizados:**

1. Los registros, las sumas de ingreso, fondo, SSF, aportaciones y montos pre evaluados concilian al 100% entre la base original y la clasificada.
2. Cada asociado aparece una sola vez en cada variable (control automático en `Consolidado`).
3. Cada regla de clasificación se reprodujo de forma independiente en Python sobre los 2,262 registros: **0 diferencias** en las 14 clasificaciones verificadas.
4. Esto confirma que las fórmulas aplican las reglas como están escritas. No prueba que los límites propuestos sean los correctos para la institución; eso le corresponde decidir a Riesgos.

**Requisitos normativos que necesitan verificación documental.** La gestión del riesgo de crédito y la clasificación de deudores (categorías A1/A2) están reguladas por la normativa prudencial salvadoreña (BCR/SSF). Este análisis **no** verificó ningún límite normativo. Riesgos y Cumplimiento deben confirmar qué norma y qué umbrales aplican a esta cooperativa.

---

# 7. Recomendaciones priorizadas para Riesgos

1. **No validar la base completa todavía.** Pedir a Créditos el ingreso que se usó para asignar los límites de los 1,418 asociados sin ingreso válido. Sin ese dato no se puede evaluar su capacidad de pago.
2. **Excluir o justificar** a los 8 asociados que tienen 56 años al corte, a los 55 con fondo $0 y a los 135 con ingreso $0, porque contradicen los criterios del propio memorando.
3. **Corregir el registro de vehículo de $150,000** (asociado 20403) y conciliar las cifras de 2,263 vs 2,262 registros, 61 sin deuda SSF y 54 vs 186 no al día en aportaciones.
4. **Aclarar el criterio 3** (966 tarjetahabientes en la base) y el criterio 4 (antigüedad "0 meses").
5. **Pedir evidencia de las consultas externas**: lista negra, castigos, personas relacionadas o empleados y solvencia en Hacienda. La solvencia es obligatoria para los 679 hipotecarios y para los 298 con TC > $4,000.
6. **Incorporar las cuotas mensuales** de todas las obligaciones (internas y SSF) para medir la relación cuota-ingreso. Es el indicador clave de capacidad de pago y hoy falta.
7. **Revisar en el Comité los 133 casos con 2 o más alertas** y los 228 con fondo bajo y deuda elevada.
8. **Documentar la naturaleza del fondo compensatorio** antes de considerarlo un mitigante.
9. **Entregar la tabla hipotecaria oficial** y unificar los nombres de los segmentos de TC.
10. **Para un modelo ponderado futuro**, conservar el desempeño de pago de estos créditos (mora a 30, 60 y 90 días) y así calibrar y validar los pesos.

---

# 8. Parámetros pendientes de aprobación institucional

La hoja `Pendientes_Aprobacion` del libro tiene 27 puntos con su responsable sugerido. Los principales son:

- los rangos de ingreso (I–IV), de deuda SSF (D1–D4), de deuda-ingreso (R1–R4) y de fondo (F1–F4 y M1–M4);
- los otros criterios (aportaciones, cuotas de aportación vencidas, niveles de categoría y tramos de edad);
- el umbral para marcar un ingreso como atípico;
- cómo tratar a los asociados sin historial interno y a los que tienen categoría A2;
- la tabla hipotecaria;
- la naturaleza del ingreso y del fondo;
- el límite institucional de la relación cuota-ingreso por producto.

Todos los parámetros están en la hoja `Parametros`, en las celdas amarillas. Si Riesgos los cambia, el libro completo se recalcula.
