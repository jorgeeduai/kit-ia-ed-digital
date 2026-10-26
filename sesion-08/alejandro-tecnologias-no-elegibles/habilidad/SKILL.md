---
name: tecnologias-no-elegibles
description: >-
  Usa este skill cuando la persona pida revisar qué actividades y estrategias de TeachingHub
  mencionan alguna tecnología no elegible, hacer el corte de tecnologías no elegibles, o validar
  el listado de Tecnologías no elegibles contra cada pestaña del archivo matriz TeachingHub.
  Toma el archivo matriz (Excel con varias pestañas, o un CSV por pestaña) y el listado, busca
  pestaña por pestaña con un script ya probado y deja un Excel con las fichas encontradas y un
  resumen de tres líneas.
---

# Tecnologías no elegibles en TeachingHub

Cada vez que se actualiza el listado de Tecnologías no elegibles (el corte es cada tres días), hay que revisar todas las pestañas del archivo matriz TeachingHub y sacar las actividades y estrategias que mencionan alguna de esas tecnologías, para que el equipo de TeachingHub ajuste el portal. Este skill lo hace siempre igual, con el script `buscar_no_elegibles.py` que está en la misma carpeta que este archivo. Buscar no necesita un modelo: lo hace el script, pestaña por pestaña, y dice en cuál pestaña y en cuál renglón encontró cada mención. El script no necesita instalar nada: corre con cualquier `python3`.

## Cuándo aplica

- La persona dice "haz el corte de tecnologías no elegibles", "qué fichas mencionan tecnologías no elegibles", "revisa TeachingHub contra el listado".
- Llegó una versión nueva del listado de Tecnologías no elegibles o del archivo matriz.
- Una tarea programada pide el corte de cada tres días.

## Qué archivos espera

- **El archivo matriz TeachingHub-AppOTecnologias**, en una de dos formas:
  - El `.xlsx` completo, con todas sus pestañas. Se revisan todas.
  - Una carpeta con un `.csv` por pestaña. El nombre del archivo es el nombre de la pestaña; un número al inicio (`1-Actividades.csv`) solo fija el orden.
- En cada pestaña, las columnas `Ficha` y `AppOTecnologia`. Si una pestaña no trae `AppOTecnologia`, el script busca en todas las celdas del renglón y lo avisa.
- **El listado de Tecnologías no elegibles**, `.xlsx` (primera hoja) o `.csv`, con las columnas `Tecnología` y `Motivo`. `Motivo` es opcional.
- **Opcional, el listado del corte anterior**, con las mismas columnas: con él, el resumen dice qué tecnologías son nuevas en este corte.
- El encabezado (la fila con los nombres de las columnas) en alguno de los primeros 10 renglones de cada archivo. Arriba puede haber títulos.

## Pasos

1. **Encuentra los archivos.** Si la persona da las rutas, usa esas. Si no, busca en la carpeta de trabajo el archivo matriz (`TeachingHub-AppOTecnologias.xlsx` o una carpeta con ese nombre) y el listado (`Tecnologias-no-elegibles`). Si hay dos listados, el anterior es el de fecha más temprana. Si no sabes cuál es cuál, pregunta; no inventes uno.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/buscar_no_elegibles.py MATRIZ LISTADO --anterior LISTADO_ANTERIOR`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). Sin listado anterior, quita `--anterior` y lo que sigue. El script deja en `salida/`, junto al archivo matriz, dos archivos: `no-elegibles-NOMBRE.xlsx` (hojas Resumen, Coincidencias, Por pestaña y Por tecnología) y `resumen-no-elegibles-NOMBRE.txt` (tres líneas), donde NOMBRE es el nombre del archivo matriz sin extensión. Con `--salida CARPETA` va donde se indique.

3. **Revisa el control.** El script imprime cuántas tecnologías trae el listado, cuántas pestañas y fichas revisó, cuántas fichas mencionan alguna no elegible y cuántas menciones hubo. Comprueba que la suma de la segunda línea por pestaña sea el total de fichas encontradas. Si el script avisa de una pestaña sin `AppOTecnologia`, o vacía, díselo a la persona.

4. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la ruta del Excel. Si la instrucción pide mandarlo a alguien, manda exactamente las tres líneas de `resumen-no-elegibles-NOMBRE.txt`, sin agregar ni quitar nada, al destino que la instrucción nombra, y adjunta el Excel si se puede.

5. **Si no puedes ejecutar Python,** dilo y aplica las reglas de búsqueda de abajo con el código que sí puedas ejecutar. Nunca busques a ojo ni sobre una parte del archivo: el resultado tiene que salir de todas las pestañas completas.

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Falta la columna `Tecnología` en el listado:** dice en qué archivo y qué columnas sí hay. Pregunta a la persona cómo se llama la columna con el nombre de la tecnología.
- **Una pestaña no trae `AppOTecnologia`:** el script no se detiene; busca en todas las celdas de ese renglón y lo avisa. Díselo a la persona junto con el resumen.
- **Una pestaña está vacía:** no la revisa y lo avisa.
- **El archivo llegó vacío** (pesa 0 bytes o no trae ningún renglón) o el listado solo trae el encabezado: se detiene. Pide el archivo correcto.
- **El archivo no existe, está dañado o no es .xlsx ni .csv:** se detiene. Pide la ruta correcta, o que se abra en Excel y se guarde otra vez como .xlsx, o cada pestaña como .csv.

## Reglas de búsqueda (lo que hace el script)

1. Revisa todas las pestañas del archivo matriz, una por una, y en cada coincidencia anota la pestaña y el renglón de Excel. Nunca mezcla valores de una pestaña con otra.
2. Busca cada tecnología del listado en la celda `AppOTecnologia` de cada renglón (o en todo el renglón, menos `Ficha`, si la pestaña no trae esa columna o si se corre con `--en-toda-la-fila`).
3. La búsqueda no distingue mayúsculas, acentos ni espacios de más: `QUIZRAYO` y ` quizrayo ` cuentan como `QuizRayo`.
4. La tecnología cuenta solo como palabra suelta: antes y después hay un espacio, una coma, un punto y coma, una diagonal, un paréntesis, o el inicio o el fin de la celda. Así `Lienzo` no se confunde con `LienzoEdu`.
5. Una celda con dos tecnologías no elegibles da dos menciones y una sola ficha.
6. Los renglones vacíos no cuentan. Una tecnología que aparece dos veces en el listado se cuenta una vez y se avisa.
7. Nueva en este corte: la tecnología está en el listado de hoy y no en el del corte anterior.

## Reglas

- **No toques los archivos de entrada.** El archivo matriz y los listados no se modifican ni se mueven; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre.** Mismo formato y mismos nombres de archivo, para que el equipo de TeachingHub reconozca el aviso de cada corte.
- **Solo al destino nombrado.** Si la instrucción no dice a quién mandar, no mandes: entrega el resumen y pregunta.
