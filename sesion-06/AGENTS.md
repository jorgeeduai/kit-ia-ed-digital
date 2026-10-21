# AGENTS.md

Carpeta de práctica con datos ficticios. Imita el flujo de una persona del equipo que cada lunes y cada miércoles limpia el reporte de inscritos y manda una tabla a un grupo de Teams.

## Quién soy y en qué área trabajo

Llevo el seguimiento de las peticiones de inscripción del periodo en Educación Digital. Cada lunes y cada miércoles el sistema me da el reporte de inscritos; lo limpio, armo una tabla de peticiones por campus y estatus y la mando al grupo de Teams del equipo. Trabajo con Excel y Teams; no programo.

## Qué hay en esta carpeta

- `entrada/`: el reporte que toca procesar, tal como lo exporta el sistema (`reporte-inscritos-AAAA-MM-DD.xlsx`). Viene sucio: el encabezado no está en la fila 1, trae filas de sección y subtotales, duplicados, co-titulares que repiten la petición y fechas en texto.
- `siguiente/`: el reporte del miércoles, guardado aparte para la segunda corrida.
- `salida/`: donde van la tabla limpia y el resumen de tres líneas. Se crea la primera vez que corre la habilidad.
- `skill-de-referencia/limpiar-reporte/`: la habilidad (skill) `limpiar-reporte`, con su `SKILL.md` y su script `limpiar_reporte.py`.
- `generador/`: el script que generó estos reportes ficticios y sus notas. No es fuente de cifras: no tomes cifras de ahí; las cifras salen de leer el reporte.
- `LEEME.md`: para qué sirve cada archivo.

## Qué no tocar

- No modifiques ni muevas los reportes de `entrada/` y `siguiente/`. Lo nuevo va a `salida/`.
- Pregunta antes de borrar o mover cualquier archivo.
- Los reportes traen nombres, matrículas y correos ficticios. Aun así, en las respuestas solo van cifras.

## Cómo quiero las respuestas

- En español, claro y breve.
- Las cifras con separador de miles y sin decimales.
- Cuando digas una cifra, di de qué archivo salió.
- Para limpiar el reporte usa la habilidad `limpiar-reporte`; si no la tienes, dímelo antes de improvisar otra limpieza. Su script está en `skill-de-referencia/limpiar-reporte/limpiar_reporte.py`.
- En esta carpeta no se manda nada por Teams ni por correo: muestra la tabla de peticiones por campus y estatus y las tres líneas del resumen, y ahí termina.
