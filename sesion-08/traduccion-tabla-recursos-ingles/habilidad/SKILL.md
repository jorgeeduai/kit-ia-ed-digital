---
name: tabla-recursos-ingles
description: >-
  Usa este skill cuando la persona pida pasar la tabla "Recursos que requieren una versión en
  inglés" del Reporte_Traducción_[CLAVE]_[NombreUF] al documento ProfesorRevisor_PL2026_UF, o
  copiar las dos primeras columnas de esa tabla de un reporte al otro. Toma los dos documentos
  de Word, saca las dos primeras columnas con un script ya probado, junto con sus ligas y sus
  capturas de pantalla, y deja la tabla lista para pegar, un conteo de control y un resumen de
  tres líneas. El pegado al documento del profesor se hace después, con la instrucción de la
  persona.
---

# Tabla de recursos en inglés: del reporte de traducción al reporte del profesor

Al terminar una traducción, la tabla "Recursos que requieren una versión en inglés" del reporte de traducción ya está aprobada y sus dos primeras columnas tienen que pasar a las dos primeras columnas de la misma tabla en el reporte para el profesor revisor. Este skill hace la parte de copiar siempre igual, con el script `copiar_tabla_recursos.py` que está en la misma carpeta que este archivo: encuentra la tabla por su título, toma las dos primeras columnas renglón por renglón, guarda aparte las ligas y las capturas de pantalla de cada celda, y comprueba que la tabla del profesor esté lista para recibirlas. El script no necesita instalar nada: corre con cualquier `python3`.

Esta es una propuesta armada con supuestos (ver `LEEME.md`, "Lo que supuse"). Los nombres de las columnas y el formato de los documentos se confirman con el equipo antes de usarla con documentos reales.

## Cuándo aplica

- La persona dice "pasa la tabla de recursos al reporte del profesor", "copia las dos primeras columnas de recursos en inglés", "llena el ProfesorRevisor".
- Se aprobó un reporte de traducción y hay que preparar el reporte para el profesor.

## Qué archivos espera

- **El reporte de traducción**, `Reporte_Traducción_[CLAVE]_[NombreUF].docx`.
- **El reporte del profesor**, `ProfesorRevisor_PL2026_UF.docx`.
- Los dos en Word (`.docx`). Si alguno es un documento de Google, se descarga antes como Word: Archivo, Descargar, Microsoft Word (.docx).
- En los dos, una tabla con el título `Recursos que requieren una versión en inglés`, escrito justo encima de la tabla o en su primer renglón. Si se llama de otra forma, se corre con `--titulo "TÍTULO"`.
- El primer renglón de cada tabla es el de los encabezados.

## Pasos

1. **Encuentra los archivos.** Si la persona da las rutas, usa esas. Si no, busca en la carpeta de trabajo los dos documentos por su nombre. Si hay más de un reporte de traducción, pregunta cuál; no adivines.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/copiar_tabla_recursos.py REPORTE PROFESOR`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). El script deja en `salida/`, junto al reporte: `tabla-para-profesor-CLAVE_NombreUF.csv` y `.xlsx` (la tabla lista para pegar, con las columnas del profesor, las ligas y las capturas de cada renglón), la carpeta `imagenes/` con cada captura nombrada por su renglón y su columna, y `resumen-tabla-recursos.txt`. Con `--salida CARPETA` va donde se indique.

3. **Revisa el control.** El script imprime cuántos renglones tenía la tabla, cuántos copió, cuántos vacíos omitió, cuántas ligas y cuántas capturas encontró, y cuántas columnas y renglones llenos tiene la tabla del profesor. Si avisa que la tabla del profesor ya trae renglones llenos, pregunta a la persona si los nuevos van abajo o los reemplazan antes de pegar.

4. **Muestra el resumen.** Las tres líneas, tal cual, y la tabla para pegar.

5. **Pega, solo si la persona lo pide.** Abre una copia del reporte del profesor y, en la tabla `Recursos que requieren una versión en inglés`, llena las dos primeras columnas con la tabla para pegar, en el mismo orden, un renglón de la tabla por cada renglón de la tabla para pegar: el texto como está, cada liga como hipervínculo sobre su texto, y cada captura de `imagenes/` en la celda que dice su nombre. No cambies las otras dos columnas ni las otras tablas. Guarda el documento completo en `salida/` con el mismo nombre del reporte del profesor; nunca sobre el original. Si no puedes editar el Word con código, dilo y entrega la tabla para pegar.

6. **Si no puedes ejecutar Python,** dilo. No copies la tabla a ojo: las capturas y las ligas se pierden.

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Un documento no es .docx:** dice cuál y cómo descargarlo como Word.
- **No encuentra la tabla:** dice en cuál documento y lista las tablas que sí hay, con el texto que tienen encima. Pregunta a la persona cuál es y corre con `--titulo`.
- **La tabla tiene menos de dos columnas:** se detiene y lo dice.
- **Una captura es una liga a una imagen y no está dentro del documento:** sigue y lo avisa con el renglón y la columna.
- **El documento está dañado o vacío:** se detiene. Pide abrirlo en Word y guardarlo otra vez como .docx.

## Reglas de copiado (lo que hace el script)

1. Usa la tabla cuyo título contiene `Recursos que requieren una versión en inglés` (sin distinguir mayúsculas ni acentos). Las demás tablas no se tocan.
2. Copia solo las columnas 1 y 2, en ese orden, a las columnas 1 y 2 de la tabla del profesor. La columna 3 del reporte no se copia.
3. Un renglón con las dos primeras celdas vacías no se copia y se cuenta como omitido.
4. El texto de cada celda se copia completo; si la celda tiene varios párrafos, van en renglones separados dentro de la misma celda.
5. Las ligas de cada celda (hipervínculos y ligas escritas como texto) se anotan aparte, con su columna, para que no se pierdan al pegar.
6. Las capturas de pantalla se guardan en `imagenes/` como `renglon-NN-col-C-K.png`: NN es el renglón de la tabla del reporte (el encabezado es el 1), C la columna y K el número de captura dentro de la celda.

## Reglas

- **No toques los documentos de entrada.** El reporte de traducción y el del profesor no se modifican ni se mueven; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre.** Mismo formato y mismos nombres de archivo en cada corrida.
- **El documento del profesor sale de la carpeta solo con el sí de la persona.** Guardarlo en la carpeta de la unidad de formación lo hace la persona o una instrucción que nombre la carpeta.
