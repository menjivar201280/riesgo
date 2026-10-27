# Guía Integral de Evaluación Crediticia en El Salvador — Actualización 2026

Entregable: `Guia_Integral_Evaluacion_Crediticia_El_Salvador_2026.docx` (Word editable, 92 páginas).

Regenerar el documento:

```bash
cd src
node build.js                     # genera el .docx (requiere el paquete npm `docx`)
python3 -I toc.py ../Guia_Integral_Evaluacion_Crediticia_El_Salvador_2026.docx <ruta>/soffice.py   # rellena el índice con números de página
```

Todos los casos y cifras de ejemplo son ficticios. Fecha de consulta normativa: 27/10/2026.

## Metodología de precalificación y preevaluación

Entregable: `Metodologia_Precalificacion_Preevaluacion_Crediticia_El_Salvador_2026.docx` (35 páginas).

```bash
cd src
node pre_build.js
python3 -I toc.py ../Metodologia_Precalificacion_Preevaluacion_Crediticia_El_Salvador_2026.docx <ruta>/soffice.py "Este documento describe una metodología"
```
