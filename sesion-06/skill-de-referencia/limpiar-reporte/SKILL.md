---
name: limpiar-reporte
description: >-
  Usa este skill cuando la persona pida limpiar el reporte de inscritos, armar la tabla de
  inscritos, preparar el corte del lunes o del miércoles, o mandar la tabla al grupo de Teams.
  Toma el reporte más reciente que exporta el sistema (Excel o CSV, sucio: encabezado abajo,
  filas de sección, duplicados, mayúsculas y fechas en texto), lo limpia con un script ya probado
  y deja una tabla de peticiones por campus y estatus en Excel y un resumen de tres líneas para Teams.
---

# Limpiar el reporte de inscritos

Cada lunes y cada miércoles llega el reporte de inscritos del sistema. Este skill lo limpia siempre igual, con el script `limpiar_reporte.py` que está en la misma carpeta que este archivo, y deja la tabla lista para el grupo de Teams. El script no necesita instalar nada: corre con cualquier `python3`.

## Cuándo aplica

- La persona dice "limpia el reporte", "arma la tabla de inscritos", "el corte de hoy", "el reporte del lunes" o "del miércoles".
- Una tarea programada pide limpiar el reporte nuevo y mandar el resumen a Teams.
- Llegó un archivo nuevo con "inscritos" en el nombre a la carpeta `entrada/`.

## Pasos

1. **Encuentra el reporte.** Si la persona da la ruta, usa esa. Si no, busca en la carpeta de trabajo, primero en `entrada/`, el archivo `.xlsx` o `.csv` con "inscritos" en el nombre y la fecha más reciente (por ejemplo `reporte-inscritos-2026-10-21.xlsx`). Si no encuentras ninguno, pregunta dónde está; no inventes uno.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/limpiar_reporte.py RUTA_DEL_REPORTE_O_CARPETA`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo. El script deja en `salida/` (junto a `entrada/`) dos archivos: `tabla-inscritos-AAAA-MM-DD.xlsx` (hojas Tabla, Peticiones y Control) y `resumen-teams-AAAA-MM-DD.txt` (tres líneas). La fecha es la del corte que dice el propio reporte.

3. **Revisa el control.** El script imprime cuántos renglones leyó, cuántos duplicados quitó, cuántos renglones de co-titular unió y cuántas peticiones quedaron. Comprueba que el total de la tabla sea igual a "Peticiones únicas". Si el script avisa de campus, estatus o fechas que no reconoció, díselo a la persona antes de mandar nada.

4. **Entrega.** Muestra las tres líneas del resumen y la ruta de la tabla. Si la instrucción pide mandarlo a Teams, manda exactamente las tres líneas de `resumen-teams-AAAA-MM-DD.txt`, sin agregar ni quitar nada, al chat o grupo que la instrucción nombra, y adjunta la tabla si se puede.

5. **Si no puedes ejecutar Python,** dilo y aplica las reglas de limpieza de abajo con el código que sí puedas ejecutar. Nunca cuentes a ojo ni sobre una parte del archivo: la tabla tiene que salir de todo el reporte.

## Reglas de limpieza (lo que hace el script)

1. El encabezado es el primer renglón que trae `Matrícula` y `CRN`; lo de arriba es título (de ahí salen la fecha y la hora de corte).
2. Solo cuentan los renglones con matrícula y CRN. Los renglones `CAMPUS: ...`, `Subtotal ...` y `Total general` son del sistema y no se cuentan.
3. Nombres de columna sin espacios de más.
4. Se quitan los renglones repetidos tal cual (duplicados exactos).
5. La unidad es la **petición**: un alumno pidiendo un grupo, es decir `Matrícula alumno` + `CRN`. Los grupos con co-titular repiten la petición, una vez por profesor; se deja una sola.
6. `Campus` y `Estatus de la petición` se comparan sin mayúsculas, sin acentos y sin espacios de más, y se escriben en su forma oficial: Monterrey, Guadalajara, Ciudad de México, Querétaro, Puebla, Estado de México; Nueva, En proceso, Resuelta, Cancelada.
7. `Fecha de solicitud` viene en texto con varios formatos (`14/10/2026`, `2026-10-14`, `14-oct-2026`, `14/10/2026 09:32`). El día va antes que el mes.
8. La tabla cuenta peticiones únicas por campus y por **estatus de la petición** (no por `Estatus inscripción`).
9. Corte anterior: el lunes o miércoles más reciente antes de la fecha de corte. "Desde el corte anterior" son las peticiones con fecha de solicitud igual o posterior a ese día.

## Reglas

- **No toques el original.** El reporte de `entrada/` no se modifica ni se mueve; todo lo nuevo va a `salida/`.
- **Sin datos personales en el mensaje.** A Teams solo van cifras: nada de nombres, matrículas ni correos. La hoja Peticiones sí los trae; esa hoja no se pega en el chat.
- **Solo al destino nombrado.** Si la instrucción no dice a qué chat o grupo mandar, no mandes: entrega el resumen y pregunta.
- **Mismo formato siempre.** Mismas tres líneas, mismos nombres de archivo, para que el grupo reconozca el mensaje de cada lunes y miércoles.

## Para adaptarlo a otro reporte

El script se apoya en cinco columnas: `Campus`, `Matrícula alumno`, `CRN`, `Estatus de la petición` y `Fecha de solicitud`. Si tu reporte las llama distinto o cuenta otra cosa, pide a tu agente que adapte `limpiar_reporte.py` con tus encabezados, que lo pruebe con una copia ficticia y que actualice estas reglas en el mismo cambio.
