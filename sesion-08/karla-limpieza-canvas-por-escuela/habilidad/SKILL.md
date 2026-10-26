---
name: limpieza-canvas-por-escuela
description: >-
  Usa este skill cuando la persona pida limpiar los archivos de actividad de Canvas que se
  descargan uno por escuela, dejar solo los grupos digitales, o preparar los seis Excel por
  escuela para la plataforma de seguimiento de profesores. Toma la carpeta con las descargas,
  quita los renglones de estudiantes y los cursos que no son digitales con un script ya
  probado, y deja un Excel limpio por escuela, con el nombre de la escuela, más un conteo de
  control y un resumen de tres líneas.
---

# Limpieza de los archivos de Canvas por escuela

Cada descarga de actividad de Canvas trae, por escuela, a todos los usuarios activos: estudiantes, profesores, observadores y otros roles, de todos los cursos. Para la plataforma de seguimiento de profesores solo sirven los grupos digitales. Este skill limpia los seis archivos siempre igual, con el script `limpiar_por_escuela.py` que está en la misma carpeta que este archivo. Limpiar no necesita un modelo: lo hace el script, renglón por renglón, y cuenta cuántos quitó y por qué. El script no necesita instalar nada: corre con cualquier `python3`.

Esta es una propuesta armada con supuestos (ver `LEEME.md`, "Lo que supuse"). Los nombres de las columnas y de las hojas se confirman con la persona antes de usarla con archivos reales.

## Cuándo aplica

- La persona dice "limpia los archivos de Canvas", "deja solo los grupos digitales", "prepara los seis archivos por escuela".
- Llegó una descarga nueva de actividad de Canvas de una o varias escuelas.

## Qué archivos espera

- **Una carpeta con las descargas**, un `.xlsx` por escuela (o los archivos uno por uno).
- En cada archivo, **la primera hoja** trae una celda que dice `Escuela` y, a su derecha, el nombre de la escuela.
- **Alguna hoja con las columnas `Rol` y `SIS ID del curso`** (la de usuarios por actividad). Si se llaman de otra forma, se corre con `--col-rol "NOMBRE"` y `--col-codigo "NOMBRE"`.
- Otras hojas pueden traer `SIS ID del curso` sin `Rol` (la información general de la materia): en esas solo se aplica la regla del código. Las hojas sin ninguna de las dos columnas se copian igual.
- El encabezado (la fila con los nombres de las columnas) en alguno de los primeros 10 renglones de cada hoja.

## Pasos

1. **Encuentra los archivos.** Si la persona da la carpeta, usa esa. Si no, busca en la carpeta de trabajo los `.xlsx` de la descarga. Si no sabes cuáles son, pregunta; no inventes.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/limpiar_por_escuela.py CARPETA_DE_DESCARGAS`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). El script deja en `limpios/`, junto a la carpeta de descargas, un Excel por escuela (`NOMBRE DE LA ESCUELA.xlsx`, con las mismas hojas ya limpias), `control-limpieza.xlsx` y `resumen-limpieza.txt`. Con `--salida CARPETA` va donde se indique.

3. **Si ya existe el Excel de una escuela,** el script no lo toca y lo avisa. Pregunta a la persona si quiere reemplazarlo; solo con su sí, corre otra vez con `--reemplazar`.

4. **Revisa el control.** Por cada archivo, el script imprime: renglones leídos = fuera por rol + fuera por código + quedan, y cuántos quedan de cada rol. La tercera línea del resumen dice si la suma cuadra en todos. Si no cuadra, o si hay un problema con algún archivo, díselo a la persona antes de entregar.

5. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la lista de Excel limpios. Si la instrucción pide dejarlos en una carpeta, cópialos ahí sin cambiarles el nombre. Subirlos a la plataforma de seguimiento lo hace la persona.

6. **Si no puedes ejecutar Python,** dilo y aplica las reglas de limpieza de abajo con el código que sí puedas ejecutar. Nunca limpies a ojo ni sobre una muestra: el resultado tiene que salir de todos los renglones.

## Si algo falla

El script explica qué pasa. Un archivo con problema no detiene a los demás: se limpian los que sí se pueden y el problema se reporta con el nombre del archivo.

- **No encuentra el nombre de la escuela:** dice en qué archivo. Pregunta en qué hoja y en qué celda viene la escuela.
- **Ninguna hoja trae `Rol` o `SIS ID del curso`:** dice en qué archivo y qué columnas sí hay en cada hoja. Pregunta cómo se llaman y corre con `--col-rol` o `--col-codigo`.
- **Dos archivos son de la misma escuela:** limpia el primero y avisa del segundo, para no encimar uno sobre otro.
- **El archivo llegó vacío, está dañado o no es .xlsx:** lo reporta. Pide descargarlo otra vez.

## Reglas de limpieza (lo que hace el script)

1. En toda hoja con la columna `Rol`, se quitan los renglones cuyo rol es `Estudiante`, sin distinguir mayúsculas, acentos ni espacios de más. Con `--solo-profesores` se queda solo el rol `Profesor`.
2. En toda hoja con la columna `SIS ID del curso`, se quedan solo los renglones cuyo código empieza con `Nacional`, sin distinguir mayúsculas ni espacios al inicio. `Internacional...` no empieza con `Nacional` y se quita.
3. Primero se aplica la regla del rol y después la del código; cada renglón quitado cuenta una sola vez, en la primera regla que lo quita.
4. Los renglones vacíos no cuentan y no pasan al archivo limpio. Los renglones de arriba del encabezado se copian igual.
5. El nombre del archivo limpio es el nombre de la escuela sin espacios de más.

## Reglas

- **No toques los archivos de entrada.** Las descargas no se modifican ni se mueven; todo lo nuevo va a `limpios/`.
- **No reemplaces sin permiso.** Un Excel limpio que ya existe solo se reemplaza con `--reemplazar` y con el sí de la persona.
- **Mismas tres líneas siempre.** Mismo formato y mismos nombres de archivo en cada corrida.
