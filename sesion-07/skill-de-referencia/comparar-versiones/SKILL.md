---
name: comparar-versiones
description: >-
  Usa este skill cuando la persona pida comparar dos versiones del mismo archivo, saber qué
  cambió entre la versión de la mañana y la de la tarde, validar si hubo cambios en la oferta de
  grupos o revisar qué grupos son nuevos, cuáles ya no aparecen y qué celdas cambiaron. Toma las
  dos versiones (Excel o CSV), las compara por una clave, como el CRN, con un script ya probado y
  deja un Excel con los cambios y un resumen de tres líneas.
---

# Comparar dos versiones de un archivo

El archivo de oferta de grupos llega varias veces al día y cada vez hay que validar qué cambió desde la versión anterior. Este skill lo compara siempre igual, con el script `comparar_versiones.py` que está en la misma carpeta que este archivo. Comparar no necesita un modelo: lo hace el script, igual cada vez. El script no necesita instalar nada: corre con cualquier `python3`.

## Cuándo aplica

- La persona dice "compara las dos versiones", "qué cambió entre la de las 9:00 y la de las 15:00", "valida si hubo cambios", "qué grupos son nuevos o ya no están".
- Llegó una versión nueva de un archivo que ya se tenía y hay que decir qué cambió.
- Una tarea programada pide comparar la versión más reciente con la anterior.

## Qué archivos espera

- Dos versiones del mismo archivo, `.xlsx` o `.csv`: la anterior y la nueva. Del Excel se lee la primera hoja.
- Una **clave**: la columna que identifica cada renglón y no se repite. En la oferta de grupos es `CRN`. Si un renglón se identifica con dos columnas o más (por ejemplo `Clave de materia` y `Grupo`), la clave es esa combinación.
- El encabezado (la fila con los nombres de las columnas) en alguno de los primeros 10 renglones. Arriba puede haber títulos.
- En la oferta de grupos, las columnas son `CRN`, `Clave de materia`, `Materia`, `Grupo`, `Campus`, `Modalidad`, `Horario`, `Salón`, `Profesor`, `Cupo` y `Estatus del grupo`. Otro archivo sirve igual, con su propia clave.

## Pasos

1. **Encuentra las dos versiones.** Si la persona da las rutas, usa esas. Si no, busca en la carpeta de trabajo, primero en `versiones/`, dos archivos con el mismo nombre y distinta hora o fecha (por ejemplo `oferta-grupos-2026-10-21-0900.xlsx` y `oferta-grupos-2026-10-21-1500.xlsx`); la anterior es la de hora o fecha más temprana. Si no sabes cuál es cuál, pregunta; no inventes una.

2. **Confirma la clave.** Si la persona no la dice y el archivo trae `CRN`, usa `CRN`. Si no trae `CRN`, pregunta qué columna identifica cada renglón.

3. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/comparar_versiones.py ANTERIOR NUEVA --clave CRN`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). Con una clave de varias columnas: `--clave "Clave de materia,Grupo"`. El script deja en `salida/` dos archivos: `cambios-NOMBRE.xlsx` (hojas Resumen, Nuevos, Ya no aparecen y Cambios) y `resumen-cambios-NOMBRE.txt` (tres líneas), donde NOMBRE es el nombre de la versión nueva sin extensión. La carpeta `salida/` queda junto a la carpeta `versiones/` (o `entrada/`) donde está la versión nueva; si la versión nueva está en otra carpeta, `salida/` queda dentro de esa carpeta. Con `--salida CARPETA` va donde se indique.

4. **Revisa el control.** El script imprime cuántos renglones con clave leyó en cada versión, cuántos son nuevos, cuántos ya no aparecen, cuántos cambiaron y cuántos siguen igual. Comprueba que los de la anterior, menos los que ya no aparecen, sean los que cambiaron más los que siguen igual. Si el script avisa de una columna que está en una versión y no en la otra, díselo a la persona.

5. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la ruta del Excel de cambios. Si la instrucción pide mandarlo a alguien, manda exactamente las tres líneas de `resumen-cambios-NOMBRE.txt`, sin agregar ni quitar nada, al destino que la instrucción nombra, y adjunta el Excel si se puede.

6. **Si no puedes ejecutar Python,** dilo y aplica las reglas de comparación de abajo con el código que sí puedas ejecutar. Nunca compares a ojo ni sobre una parte del archivo: el resultado tiene que salir de las dos versiones completas.

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Falta la clave:** dice qué columna falta, en qué archivo, y qué columnas sí hay. Pregunta a la persona cuál identifica cada renglón y vuelve a correr con esa.
- **La clave se repite** dentro de una versión: dice qué valores se repiten y en qué renglones, y sugiere una combinación de columnas que no se repite. Confírmala con la persona antes de usarla.
- **Una columna está en una versión y no en la otra:** el script no se detiene; la avisa y compara las demás. Díselo a la persona junto con el resumen.
- **El archivo llegó vacío** (pesa 0 bytes o no trae ningún renglón) o solo trae el encabezado: se detiene. Pide la versión correcta.
- **El archivo no existe, está dañado o no es .xlsx ni .csv:** se detiene. Pide la ruta correcta, o que se abra en Excel y se guarde otra vez como .xlsx o .csv.

## Reglas de comparación (lo que hace el script)

1. Compara por la clave, nunca por posición: el orden de los renglones no importa.
2. **Nuevo:** la clave está solo en la versión nueva. **Ya no aparece:** la clave está solo en la anterior. **Con cambios:** la clave está en las dos y al menos una celda cambió. Cada celda que cambió es un renglón de la hoja Cambios, con la clave, la materia (si el archivo trae esa columna), la columna, el valor de antes y el de después.
3. Los valores se comparan sin espacios de más, y los números como números: `35` y `35.0` son iguales. Un cambio de mayúsculas o de acentos sí cuenta como cambio.
4. Los nombres de las columnas se reconocen sin distinguir mayúsculas, acentos ni espacios de más.
5. Columnas con más cambios: por número de celdas que cambiaron, de mayor a menor; en empate, en el orden del archivo. La tercera línea del resumen trae las tres primeras.
6. Los renglones vacíos no cuentan. Los renglones sin clave no se comparan, y el control dice cuántos hubo.

## Reglas

- **No toques las versiones.** Los dos archivos no se modifican ni se mueven; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre.** Mismo formato y mismos nombres de archivo, para que el equipo reconozca el aviso de cada versión.
- **Sin datos personales en el mensaje.** Las tres líneas solo traen cifras y nombres de columnas. La hoja Cambios puede traer nombres de profesores: se adjunta, no se pega en un chat.
- **Solo al destino nombrado.** Si la instrucción no dice a quién mandar, no mandes: entrega el resumen y pregunta.

## Para adaptarlo a otro archivo

El script sirve para cualquier archivo que tenga una clave: pásala con `--clave`. Si tu regla lleva un paso más, por ejemplo limpiar el archivo antes de comparar (como hace la habilidad `limpiar-reporte` de la Sesión 6), pide a tu agente que lo agregue al script, que lo pruebe con una copia ficticia hecha por un script y que actualice estas reglas en el mismo cambio.
