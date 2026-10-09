# Cronograma y Dashboard de Gestión de Riesgos (ciclo mensual de 4 semanas)

Entregable: `Cronograma_Dashboard_Riesgos.xlsx` (sin macros y sin datos de prueba).

| Hoja | Contenido |
|---|---|
| Dashboard - Consolidado | Resumen del mes mostrado: 8 indicadores, avance por puesto, mensajes de finalización, seguimiento semanal, 2 gráficos, control de calidad y entregas por semana |
| Historial mensual | Avance y completadas por puesto mes a mes desde octubre 2026, 24 meses (nada se borra al cambiar de mes) |
| Cronograma R. Operacional / R. Normativo / R. Financiero / Asistente Técnico | Una hoja por puesto con su color (Operacional naranja, Normativo morado, Financiero verde, Asistente azul); celdas a llenar en gris suave, tablas `tblAct_RO`, `tblAct_RN`, `tblAct_RF`, `tblAct_ATR` (Tipo Macro/Sub en una sola columna con sangría ↳; N° 1, 1.1… y Mes automáticos; subactividades ilimitadas con peso opcional que completa el de su macroactividad; Alerta de plazo; 300 filas que acumulan los meses; columnas auxiliares ocultas) |
| Actividades R. Operacional / R. Normativo / R. Financiero / Asistente Técnico | Vista AUTOMÁTICA y bloqueada de las actividades del mes (línea de tiempo por semana, semana de entrega, estado), tablas `tblSem_XX` |
| Instrucciones | Uso, colaboración, protección, segregación con Power Query y limitaciones |
| Configuración | Mes mostrado, inicio del ciclo, fecha de corte y semana actual calculados automáticamente; listas y metodología |
| Calculos (oculta) | Cálculos intermedios por puesto |

Regenerar y probar:

```bash
python build_cronograma.py Cronograma_Dashboard_Riesgos.xlsx
python test_cronograma.py <ruta>/recalc.py <directorio_temporal>   # 69 verificaciones con LibreOffice
```
