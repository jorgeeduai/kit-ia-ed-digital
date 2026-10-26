---
name: revisar-programacion-ecoas
description: >-
  Usa este skill cuando la persona pida revisar inconsistencias en la programación de un
  periodo o cruzar la programación con los resultados de ECOAs: renglones repetidos, datos del
  grupo distintos en un mismo CRN, porcentajes de responsabilidad que no suman 100, nóminas mal
  escritas, choques de horario de un profesor, modalidades fuera del catálogo, grupos
  programados sin ECOA y ECOAs de grupos que no están programados. Usa un script ya probado y
  deja un Excel con una hoja por tipo, cada renglón empezando con el CRN, y un resumen de tres
  líneas.
---

# Inconsistencias de la programación y su cruce con ECOAs

Antes de analizar las ECOAs de un periodo conviene saber si la programación está limpia y si cada grupo programado tiene su ECOA. Este skill hace esa revisión siempre igual, con el script `revisar_programacion.py` que está en la misma carpeta que este archivo. Revisar no necesita un modelo: lo hace el script, igual cada vez. El script no necesita instalar nada: corre con cualquier `python3`.

Es una **propuesta con todo supuesto**: los encabezados de los dos archivos, qué cuenta como inconsistencia, el catálogo de modalidades y el formato de la nómina. Están escritos arriba del script (constantes y tabla `TIPOS`) para cambiarlos en un solo lugar.

## Cuándo aplica

- La persona dice "revisa la programación", "busca inconsistencias en la programación", "qué grupos no tienen ECOA", "qué profesores tienen choque de horario".
- Llegó la programación de un periodo o los resultados de ECOAs y hay que revisarlos antes del análisis.

## Qué archivos espera

- **El archivo de programación del periodo**, `.xlsx` o `.csv`: un renglón por profesor de cada grupo, con `CRN` y `Nómina del profesor`. Usa además, si están: `Clave`, `Nombre de la materia`, `Grupo`, `Modalidad`, `Nombre del profesor`, `% de responsabilidad`, `Días`, `Horario inicio` y `Horario fin`. Si falta una, avisa y no revisa lo que depende de ella.
- **Los resultados de ECOAs del mismo periodo** (opcional), `.xlsx` o `.csv`, con `CRN` y `Nómina del profesor`. Sin este archivo, el script solo revisa la programación.
- Si el CRN o la nómina se llaman distinto, se indican con `--crn` y `--nomina` (el mismo nombre en los dos archivos).
- El encabezado en alguno de los primeros 10 renglones. Del Excel se lee la primera hoja.

## Pasos

1. **Encuentra los archivos.** Si la persona da las rutas, usa esas. Si no, busca en la carpeta de trabajo (primero en `ficticio/` o `entrada/`) el de programación y el de ECOAs. Si no sabes cuál es cuál, pregunta.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/revisar_programacion.py PROGRAMACION ECOAS`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). Primero va la programación y después las ECOAs. Con la copia ficticia: `python3 -B habilidad/revisar_programacion.py ficticio/programacion-periodo-ficticia.xlsx ficticio/resultados-ecoas-ficticio.xlsx`. El script deja en `salida/` dos archivos: `inconsistencias-NOMBRE.xlsx` y `resumen-inconsistencias-NOMBRE.txt` (tres líneas), donde NOMBRE es el nombre del archivo de programación sin extensión. Con `--salida CARPETA` van donde se indique.

3. **Revisa el control.** El script imprime cuántos renglones, CRN y profesores leyó, cuántos CRN y profesores tienen ECOA y cuántas inconsistencias hay de cada tipo. Si avisa de una columna que no encontró, díselo a la persona.

4. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la ruta del Excel. El Excel trae la hoja Resumen, la hoja Todas (ordenada por CRN) y una hoja por cada tipo que haya; en todas, la primera columna es el CRN.

5. **Si no puedes ejecutar Python,** dilo y aplica las reglas de abajo con el código que sí puedas ejecutar. Nunca revises a ojo ni sobre una parte de los archivos.

## Tipos de inconsistencia (supuestos)

| Tipo | Cuándo |
|---|---|
| Renglón repetido | El mismo CRN y la misma nómina aparecen dos veces |
| Datos del grupo distintos en el mismo CRN | Dos renglones del mismo CRN traen distinta Clave, materia, Grupo, Modalidad, Días u horario |
| % de responsabilidad que no suma 100 | Los porcentajes de los profesores de un CRN no suman 100 |
| Nómina vacía o con otro formato | Sin nómina, o una nómina que no es L y ocho dígitos |
| Una nómina con dos nombres | La misma nómina aparece con nombres distintos |
| Choque de horario del profesor | El mismo profesor tiene dos CRN con un día en común y horas que se enciman |
| Modalidad fuera del catálogo | La modalidad no es Presencial, Digital, Híbrida ni Flexible digital |
| Programado sin ECOA | El CRN y el profesor están en la programación y no en ECOAs |
| ECOA sin programación | El CRN y el profesor están en ECOAs y no en la programación |

## Reglas de la revisión (lo que hace el script)

1. Cruza por CRN y nómina, nunca por posición. La nómina se compara sin espacios y sin distinguir mayúsculas.
2. Lo que es solo formato no cuenta: `50%` y `50`; `07:00` y `0700`; `Digital` y `DIGITAL`; días en otro orden.
3. Un renglón sin nómina solo cuenta como "Nómina vacía"; no entra a choques ni al cruce con ECOAs (por eso su ECOA sale como "ECOA sin programación").
4. Tipos con más casos: de mayor a menor; en empate, en el orden de la tabla.

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Falta la columna CRN o la de nómina:** dice cuál falta, en qué archivo y qué columnas sí hay. Pregunta cómo se llama y vuelve a correr con `--crn` o `--nomina`.
- **Falta otra columna:** no se detiene; avisa qué no pudo revisar.
- **El archivo llegó vacío, no existe o no es .xlsx ni .csv** (por ejemplo, un `.xls` antiguo): se detiene. Pide el archivo correcto o que se guarde como .xlsx o .csv.

## Reglas

- **No toques los archivos.** Las entradas no se modifican ni se mueven; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre**, con el mismo formato y los mismos nombres de archivo.
- **Sin datos personales en un mensaje.** Las tres líneas solo traen cifras y tipos. El Excel trae nombres y nóminas: se queda con la persona.
- **No mandes nada.** Si la instrucción pide mandar ajustes a alguien, pregunta antes.

## Para adaptarlo

Para quitar o agregar un tipo, cambiar el catálogo de modalidades o el formato de la nómina, pide a tu agente que cambie las constantes de arriba del script, que lo pruebe con una copia ficticia hecha por `generar_ficticio.py` y que actualice este archivo en el mismo cambio.
