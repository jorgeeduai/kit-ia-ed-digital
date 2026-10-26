---
name: corregir-candidatos-ecoas
description: >-
  Usa este skill cuando la persona pida corregir los alumnos candidatos de las ECOAs de
  tutores, revisar que el número de candidatos de cada tutor sea el que se mandó a escolar o
  quitar la suma duplicada que aparece cuando un CRN tiene varios tutores. Cruza el archivo de
  ECOAs de tutores con el archivo que se mandó a escolar (Excel o CSV) por CRN y nómina, con un
  script ya probado, y deja un Excel con el número correcto por tutor y un resumen de tres
  líneas.
---

# Corregir los alumnos candidatos de las ECOAs de tutores

Cuando se descarga el archivo de ECOAs de tutores, un CRN con varios tutores trae en cada tutor la suma de los alumnos de todo el CRN. Para analizar los promedios, cada tutor necesita solo sus alumnos: los que se mandaron antes a escolar. Este skill hace esa corrección siempre igual, con el script `corregir_candidatos.py` que está en la misma carpeta que este archivo. Corregir no necesita un modelo: lo hace el script, igual cada vez. El script no necesita instalar nada: corre con cualquier `python3`.

Es una **propuesta**: el archivo que se mandó a escolar está supuesto (un renglón por alumno, con su CRN y la nómina de su tutor).

## Cuándo aplica

- La persona dice "corrige los candidatos de las ECOAs", "a cada tutor le sale la suma del grupo", "revisa que los candidatos sean los que mandé a escolar".
- Se descargó un archivo de ECOAs de tutores nuevo y hay que dejarlo listo para el análisis de promedios.

## Qué archivos espera

- **El archivo de ECOAs de tutores**, `.xlsx` o `.csv`, con `CRN`, `Clave`, `Grupo`, `Alumnos Candidatos`, `Nómina del Profesor` y `Nombre del profesor`. Un renglón por tutor de cada CRN.
- **El archivo que se mandó a escolar**, `.xlsx` o `.csv`, con `CRN` y `Nómina del Profesor` (la nómina del tutor). Si trae un renglón por alumno, el script cuenta los renglones (sin repetir `Matrícula` si esa columna está). Si ya trae el número por tutor en una columna como `Alumnos asignados`, la suma.
- Si la nómina se llama distinto en un archivo, se indica con `--nomina-ecoas` o `--nomina-escolar`; si el CRN se llama distinto, con `--crn`.
- El encabezado en alguno de los primeros 10 renglones. Del Excel se lee la primera hoja.

## Pasos

1. **Encuentra los dos archivos.** Si la persona da las rutas, usa esas. Si no, busca en la carpeta de trabajo (primero en `ficticio/` o `entrada/`) el de ECOAs y el que se mandó a escolar. Si no sabes cuál es cuál, pregunta.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/corregir_candidatos.py ECOAS ESCOLAR`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). Primero va el de ECOAs y después el de escolar. Con la copia ficticia: `python3 -B habilidad/corregir_candidatos.py ficticio/ecoas-tutores-ficticio.xlsx ficticio/enviado-a-escolar-ficticio.xlsx`. El script deja en `salida/` dos archivos: `candidatos-corregidos-NOMBRE.xlsx` y `resumen-candidatos-corregidos-NOMBRE.txt` (tres líneas), donde NOMBRE es el nombre del archivo de ECOAs sin extensión. Con `--salida CARPETA` van donde se indique.

3. **Revisa el control.** El script imprime cuántos tutores y CRN leyó, cómo contó los alumnos de escolar, cuántos corrigió, cuántos ya estaban bien y cuáles quedan por revisar. Los que quedan por revisar los imprime uno por uno: díselos a la persona.

4. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la ruta del Excel. El Excel trae las hojas Resumen, ECOAs corregidas (el archivo completo con el número correcto, el número de antes y qué pasó en cada renglón), Cambios y Por revisar.

5. **Si no puedes ejecutar Python,** dilo y aplica las reglas de abajo con el código que sí puedas ejecutar. Nunca corrijas a ojo ni sobre una parte de los archivos.

## Reglas de la corrección (lo que hace el script)

1. La llave es CRN más nómina del tutor. La nómina se compara sin espacios y sin distinguir mayúsculas; el CRN, como número.
2. Para cada tutor del archivo de ECOAs, el número correcto es el de sus alumnos en el archivo de escolar. Un alumno repetido (mismo CRN, tutor y matrícula) cuenta una vez.
3. Si el número ya coincide: "Ya estaba bien". Si traía la suma de todo el CRN: "Corregido: traía la suma del CRN". Si traía cualquier otro número: "Corregido: traía otro número".
4. Un tutor que está en ECOAs y no en escolar **no se toca**: queda "Por revisar". Un tutor que está en escolar y no en ECOAs también sale en Por revisar.
5. El orden y las columnas del archivo de ECOAs se conservan; solo cambia `Alumnos Candidatos` y se agregan dos columnas al final.

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Falta CRN, la nómina o `Alumnos Candidatos`:** dice cuál falta, en qué archivo y qué columnas sí hay. Pregunta cómo se llama y vuelve a correr con la opción que indica.
- **Un tutor aparece dos veces en el mismo CRN** en ECOAs: el segundo renglón queda por revisar y se avisa.
- **El archivo llegó vacío, no existe o no es .xlsx ni .csv:** se detiene. Pide el archivo correcto o que se guarde como .xlsx o .csv.

## Reglas

- **No toques los archivos.** Los dos archivos de entrada no se modifican ni se mueven; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre**, con el mismo formato y los mismos nombres de archivo.
- **Sin datos personales en un mensaje.** Las tres líneas solo traen cifras. El Excel trae nombres y nóminas: se queda con la persona.
- **No mandes nada.** Si la instrucción pide mandarlo a directoras o coordinadoras, pregunta antes.

## Para adaptarlo

Si el archivo de escolar real tiene otras columnas, pide a tu agente que cambie las constantes de arriba del script, que lo pruebe con una copia ficticia hecha por `generar_ficticio.py` y que actualice este archivo en el mismo cambio.
