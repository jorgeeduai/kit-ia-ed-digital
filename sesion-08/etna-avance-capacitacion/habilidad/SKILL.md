---
name: avance-capacitacion
description: >-
  Usa este skill cuando la persona pida el avance de capacitación de los profesores
  programados, el % de avance por profesor, quitar de la programación los horarios que no son
  TCM o armar la lista nueva de profesores para revisar su capacitación. Toma el archivo de
  programación diaria y el histórico de capacitación (Excel o CSV), deja solo los renglones
  TCM, arma un profesor por nómina, lo cruza con el histórico con un script ya probado y deja
  un Excel con el detalle y el % de avance por profesor, y un resumen de tres líneas.
---

# Avance de capacitación de los profesores programados

Cada día llega una programación nueva y cada día cambia el histórico de capacitación. Este skill hace el mismo cálculo cada vez, con el script `avance_capacitacion.py` que está en la misma carpeta que este archivo: quita los horarios que no son TCM, deja un profesor por nómina y le calcula el avance con lo que dice el histórico. Calcular no necesita un modelo: lo hace el script, igual cada vez. El script no necesita instalar nada: corre con cualquier `python3`.

Es una **propuesta**: los encabezados del archivo de programación y la regla del % están supuestos. Están escritos en la parte de arriba del script (tablas `CURSOS` y `CURSADO`) para cambiarlos en un solo lugar.

## Cuándo aplica

- La persona dice "calcula el avance de capacitación", "qué profesores programados ya tienen su capacitación", "deja solo TCM y arma la lista de profesores", "qué % de avance trae cada profesor".
- Llegó la programación del día o se actualizó el histórico y hay que rehacer la lista para el equipo de coordinadoras.
- Una tarea programada pide repetir el cálculo.

## Qué archivos espera

- **El archivo de programación** y **el histórico de capacitación**, `.xlsx` o `.csv`. Del Excel se lee la primera hoja.
- En la programación (encabezados supuestos): `Nómina`, `Nombre del profesor`, `UDF`, `Modalidad`, `Idioma`, `Plan` y `Zona horaria`, la columna que dice si el renglón es TCM, TCH o TCP. Si la columna del horario o la de nómina se llaman distinto, se indican con `--horario` y `--nomina-programacion`. Si falta `UDF`, `Modalidad`, `Idioma` o `Plan`, el script avisa y sigue.
- En el histórico: `Nómina`, `Nombre` y las columnas de cursos (`Sensibilización a la microenseñanza`, `Canvas: Conceptos básicos`, `Fundamentos del Modelo Tec21`, `Evaluación Competencias TEC21`, `Aula Virtual`, `AA mediante HD`, `Turnitin`, `Respondus Monitor`, `Inglés Acreditado`). El orden de las columnas no importa, ni las mayúsculas, los acentos o los espacios de más en el nombre.
- El encabezado (la fila con los nombres de las columnas) en alguno de los primeros 10 renglones. Arriba puede haber títulos.

## Pasos

1. **Encuentra los dos archivos.** Si la persona da las rutas, usa esas. Si no, busca en la carpeta de trabajo (primero en `ficticio/` o `entrada/`) un archivo de programación y un histórico de capacitación. Si no sabes cuál es cuál, pregunta; no inventes uno.

2. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/avance_capacitacion.py PROGRAMACION HISTORICO`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). Primero va la programación y después el histórico. Con la copia ficticia: `python3 -B habilidad/avance_capacitacion.py ficticio/programacion-diaria-ficticia.xlsx ficticio/historico-capacitacion-ficticio.xlsx`. El script deja en `salida/` dos archivos: `avance-capacitacion-NOMBRE.xlsx` y `resumen-avance-capacitacion-NOMBRE.txt` (tres líneas), donde NOMBRE es el nombre del archivo de programación sin extensión. La carpeta `salida/` queda junto a la carpeta `ficticio/` o `entrada/`; con `--salida CARPETA` va donde se indique.

3. **Revisa el control.** El script imprime cuántos renglones trae la programación por horario, cuántos se quedan en TCM, cuántos profesores hay en la lista nueva, cuántos no están en el histórico, cuántos traen otro nombre en el histórico y cuántos valores no reconoció. Si avisa de una columna que no encontró, díselo a la persona.

4. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la ruta del Excel. El Excel trae las hojas Resumen, Avance por profesor, Detalle (un renglón por profesor y curso), Sin histórico y Control TCM.

5. **Si no puedes ejecutar Python,** dilo y aplica las reglas de abajo con el código que sí puedas ejecutar. Nunca calcules a ojo ni sobre una parte de los archivos.

## Reglas del cálculo (lo que hace el script)

1. **Solo TCM.** Un renglón se queda si la palabra TCM aparece en la columna del horario, en mayúsculas o minúsculas. Los demás (TCH, TCP) se quitan.
2. **Un profesor por nómina.** Con todas sus UDF en TCM. La nómina se compara sin espacios y sin distinguir mayúsculas.
3. **Qué cursos se piden** (supuesto): `Canvas: Conceptos básicos` a todos; `Sensibilización a la microenseñanza`, `Fundamentos del Modelo Tec21` y `Evaluación Competencias TEC21` a quien imparte una UDF del plan 2019; `Aula Virtual`, `AA mediante HD`, `Turnitin` y `Respondus Monitor` a quien imparte una UDF que no es presencial; `Inglés Acreditado` a quien imparte una UDF en inglés. Las columnas de Prepa y las de Estatus no cuentan.
4. **Qué cuenta como cursado** (supuesto): `Acreditado`, `Sí`, `X`, `Completo`, `Aprobado` o una fecha. Vacío, `Pendiente`, `En curso` o `No` cuentan como falta. Cualquier otro valor cuenta como falta y el control dice cuál fue.
5. **% de avance** = cursos cursados entre cursos que se piden, con un decimal.
6. Un profesor programado que no está en el histórico sale en la hoja Sin histórico, con todos sus cursos como faltantes.

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Falta la columna de nómina o la del horario:** dice cuál falta, en qué archivo y qué columnas sí hay. Pregunta cómo se llama y vuelve a correr con `--nomina-programacion`, `--nomina-historico` o `--horario`.
- **Ningún renglón es TCM:** dice qué valores sí trae la columna del horario. Pregunta cómo se marca TCM y vuelve a correr con `--valor`.
- **El histórico no trae ninguna columna de cursos:** se detiene. Revisa que el segundo archivo sea el histórico.
- **Una nómina se repite en el histórico:** avisa y usa el primer renglón.
- **El archivo llegó vacío, no existe o no es .xlsx ni .csv:** se detiene. Pide el archivo correcto o que se guarde como .xlsx o .csv.

## Reglas

- **No toques los archivos.** La programación y el histórico no se modifican ni se mueven; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre**, con el mismo formato y los mismos nombres de archivo.
- **Sin datos personales en un mensaje.** Las tres líneas solo traen cifras y nombres de cursos. El Excel trae nombres y nóminas: se queda con la persona.
- **No mandes nada.** Si la instrucción pide mandarlo al equipo de coordinadoras, pregunta antes.

## Para adaptarlo

Si la regla real es otra (otros cursos, otro valor para "cursado", otra forma de marcar TCM), pide a tu agente que cambie las tablas de arriba del script, que lo pruebe con una copia ficticia hecha por `generar_ficticio.py` y que actualice este archivo en el mismo cambio.
