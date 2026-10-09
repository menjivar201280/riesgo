# Cronograma y Dashboard de Gestión de Riesgos (ciclo mensual de 4 semanas)

Entregable: `Cronograma_Dashboard_Riesgos.xlsx` (sin macros y sin datos de prueba).

| Hoja | Contenido |
|---|---|
| Dashboard - Consolidado | 8 tarjetas de indicadores, avance por puesto, mensajes de finalización, seguimiento semanal, 2 gráficos, control de calidad y actividades semanales |
| Cronograma R. Operacional / R. Normativo / R. Financiero / Asistente Técnico | Una hoja por puesto con su color, tablas `tblAct_RO`, `tblAct_RN`, `tblAct_RF`, `tblAct_ATR` (14 columnas, N° automático, 100 filas) |
| Actividades R. Operacional / R. Normativo / R. Financiero / Asistente Técnico | Actividades semanales por puesto, tablas `tblSem_XX` (4 propuestas editables, estado automático, 20 filas) |
| Instrucciones | Uso, colaboración, protección, segregación con Power Query y limitaciones |
| Configuración | Mes, inicio del ciclo, fecha de corte y semana actual calculados automáticamente; listas y metodología |
| Calculos (oculta) | Cálculos intermedios por puesto |

Regenerar y probar:

```bash
python build_cronograma.py Cronograma_Dashboard_Riesgos.xlsx
python test_cronograma.py <ruta>/recalc.py <directorio_temporal>   # 57 verificaciones con LibreOffice
```
