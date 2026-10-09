# Cronograma y Dashboard de Gestión de Riesgos (ciclo de 4 semanas)

Entregable: `Cronograma_Dashboard_Riesgos.xlsx` (sin macros y sin datos de prueba).

| Hoja | Contenido |
|---|---|
| Dashboard - Consolidado | 8 tarjetas de indicadores, avance por puesto, mensajes de finalización, seguimiento semanal, 2 gráficos, control de calidad e hitos |
| Cronograma R. Operacional / R. Legal y Normativo / R. Financiero / Asistente Técnico | Una hoja por puesto, tablas `tblAct_RO`, `tblAct_RLN`, `tblAct_RF`, `tblAct_ATR` (16 columnas, 100 filas cada una) |
| Hitos R. Operacional / R. Legal y Normativo / R. Financiero / Asistente Técnico | Una hoja por puesto, tablas `tblHitos_XX` con 4 hitos propuestos (editables) y 20 filas |
| Instrucciones | Uso, colaboración, protección, segregación con Power Query y limitaciones |
| Configuración | Semana actual, fecha de corte, inicio del ciclo, listas y metodología |
| Calculos (oculta) | Cálculos intermedios por puesto |

Regenerar y probar:

```bash
python build_cronograma.py Cronograma_Dashboard_Riesgos.xlsx
python test_cronograma.py <ruta>/recalc.py <directorio_temporal>   # 54 verificaciones con LibreOffice
```
