---
name: validar-programacion-iris
description: >-
  Usa este skill cuando la persona pida validar la programación contra IRIS, revisar que lo que
  dice el archivo de programación coincida con el reporte de grupos de IRIS, o encontrar
  discrepancias por CRN entre los dos. Cruza los dos archivos (Excel o CSV) por CRN y nómina
  del profesor con un script ya probado y deja un Excel con las discrepancias por tipo, cada
  renglón empezando con el CRN, y un resumen de tres líneas.
---

# Validar la programación contra el reporte de grupos de IRIS

Cada periodo hay que comprobar que lo que se capturó en el archivo de programación es lo que quedó en IRIS, el sistema de inscripción: CRN, clave, materia, grupo, PMT, días, horarios, fechas del periodo, profesor, su porcentaje de responsabilidad y los atributos del curso. Este skill hace ese cruce siempre igual, con el script `validar_programacion.py` que está en la misma carpeta que este archivo. Cruzar no necesita un modelo: lo hace el script, igual cada vez. El script no necesita instalar nada: corre con cualquier `python3`.

## Cuándo aplica

- La persona dice "valida la programación contra IRIS", "revisa si lo de programación coincide con el sistema", "qué CRN tienen discrepancias", "qué no cuadra entre programación e IRIS".
- Llegó un reporte de grupos de IRIS nuevo y hay que compararlo con el archivo de programación del mismo periodo.
- Una tarea programada pide repetir la validación.

## Qué archivos espera

- **El archivo de programación** y **el reporte de grupos de IRIS** del mismo periodo, `.xlsx` o `.csv`. Del Excel se lee la primera hoja.
- En los dos, una columna `CRN` y una columna `Nómina del profesor`. Cada renglón es un profesor de un grupo: un grupo con dos profesores trae dos renglones con el mismo CRN. Si en un archivo se llaman distinto, se indica con `--crn` y `--nomina`.
- Las demás columnas que se comparan, con el mismo nombre en los dos archivos: `Clave`, `Nombre de la materia`, `Grupo`, `PMT`, `Día inicio`, `Día fin`, `Horario inicio`, `Horario fin`, `Fecha de inicio de periodo`, `Fecha de fin de periodo`, `Nombre del profesor`, `% de responsabilidad asignada al grupo` y `Atributos del curso`. El orden de las columnas no importa, ni las mayúsculas, los acentos o los espacios de más en el nombre.
- El encabezado (la fila con los nombres de las columnas) en alguno de los primeros 10 renglones. Arriba puede haber títulos, como los que trae el reporte de IRIS.

## Pasos

1. **Encuentra los dos archivos.** Si la persona da las rutas, usa esas. Si no, busca en la carpeta de trabajo un archivo de programación y un reporte de IRIS. Si no sabes cuál es cuál, pregunta; no inventes uno.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/validar_programacion.py PROGRAMACION IRIS`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). Primero va el archivo de programación y después el de IRIS. Si una columna clave se llama distinto: `--crn "NOMBRE"` o `--nomina "NOMBRE"`. El script deja en `salida/` dos archivos: `discrepancias-NOMBRE.xlsx` y `resumen-discrepancias-NOMBRE.txt` (tres líneas), donde NOMBRE es el nombre del archivo de programación sin extensión. La carpeta `salida/` queda junto a la carpeta `entrada/` o `ficticio/` donde está el archivo de programación; si está en otra carpeta, `salida/` queda dentro de esa carpeta. Con `--salida CARPETA` va donde se indique.

3. **Revisa el control.** El script imprime cuántos renglones y cuántos CRN leyó en cada archivo, cuántos CRN están en los dos, cuántos tienen incidencia, cuántas celdas tienen datos distintos y cuántas solo cambian de formato. Si el script avisa de una columna que está en un archivo y no en el otro, díselo a la persona: esa columna no se comparó.

4. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la ruta del Excel de discrepancias. El Excel trae la hoja Resumen, la hoja Todas y una hoja por cada tipo de discrepancia que haya; en todas, la primera columna es el CRN.

5. **Si no puedes ejecutar Python,** dilo y aplica las reglas de abajo con el código que sí puedas ejecutar. Nunca compares a ojo ni sobre una parte de los archivos: el resultado tiene que salir de los dos archivos completos.

## Tipos de discrepancia

| Tipo | Cuándo |
|---|---|
| CRN solo en programación | El CRN está en programación y no aparece en IRIS |
| CRN solo en IRIS | El CRN está en IRIS y no aparece en programación |
| Profesor asignado | El CRN está en los dos, pero con otra nómina de profesor |
| Materia y grupo | Cambia Clave, Nombre de la materia, Grupo o PMT |
| Días y horarios | Cambia Día inicio, Día fin, Horario inicio u Horario fin |
| Fechas del periodo | Cambia Fecha de inicio de periodo o Fecha de fin de periodo |
| Nombre o % del profesor | Misma nómina, pero cambia el nombre o el % de responsabilidad |
| Atributos del curso | Cambian los atributos |
| Otras columnas | Cambia cualquier otra columna que esté en los dos archivos |

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Falta la columna CRN o la de nómina** en un archivo: dice cuál falta, en qué archivo y qué columnas sí hay. Pregunta a la persona cómo se llama en ese archivo y vuelve a correr con `--crn` o `--nomina`.
- **Un renglón se repite** (mismo CRN y misma nómina dos veces en un archivo): dice cuáles y en qué renglones. Pregunta si son duplicados.
- **Una columna está en un archivo y no en el otro:** el script no se detiene; la avisa y compara las demás. Díselo a la persona junto con el resumen.
- **Los renglones de un mismo CRN no traen los mismos datos del grupo** dentro de un archivo: lo avisa y compara el primero.
- **El archivo llegó vacío** o solo trae el encabezado: se detiene. Pide el archivo correcto.
- **El archivo no existe, está dañado o no es .xlsx ni .csv** (por ejemplo, un `.xls` antiguo): se detiene. Pide que se abra en Excel y se guarde como .xlsx o .csv.

## Reglas del cruce (lo que hace el script)

1. Cruza por CRN y nómina del profesor, nunca por posición: el orden de los renglones y de las columnas no importa.
2. Los datos del grupo (todo menos nómina, nombre y % del profesor) se comparan una vez por CRN. Los datos del profesor se comparan por CRN y nómina.
3. Los valores se comparan sin espacios de más, y los números como números: `3` y `3.0` son iguales. Un cambio de mayúsculas o de acentos sí cuenta.
4. Las diferencias solo de formato no cuentan: fechas como `08/02/2027`, `2027-02-08` o fecha de Excel; horas como `7:30`, `07:30` o `0730`; porcentajes como `50`, `50%` o `0.5`. El control dice cuántas celdas había así.
5. Tipos con más discrepancias: por número de renglones del detalle, de mayor a menor; en empate, en el orden de la tabla de tipos.
6. Los renglones vacíos no cuentan. Los renglones sin CRN no se comparan, y el control dice cuántos hubo.

## Reglas

- **No toques los archivos.** Programación e IRIS no se modifican ni se mueven; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre.** Mismo formato y mismos nombres de archivo, para reconocer el resultado de cada validación.
- **Sin datos personales en un mensaje.** Las tres líneas solo traen cifras y tipos. El Excel trae nombres y nóminas de profesores: se queda con la persona.
- **No mandes nada.** El resultado es para la persona que corre la validación; si la instrucción pide mandarlo a alguien, pregunta antes.

## Para adaptarlo

Si en el archivo real una columna se llama distinto en programación y en IRIS, pide a tu agente que agregue al script la equivalencia de nombres, que lo pruebe con una copia ficticia hecha por un script y que actualice este archivo en el mismo cambio.
