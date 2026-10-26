# Avance de capacitación de los profesores programados · propuesta de Jorge para Etna

Esta es la propuesta de Jorge para tu flujo, armada con lo que escribiste en el muro de la Sesión 7 y lo que contaste en la clase. Corre sobre datos ficticios y ya se probó afuera. La idea es que la compares con lo que tú haces y nos digas qué le falta o qué está mal supuesto.

## Lo que supuse

Tu archivo de programación todavía no tiene encabezados en el muro, así que varias cosas son supuestos. Estos son todos:

1. Supuse que el archivo de programación trae `Periodo`, `CRN`, `Clave UDF`, `UDF`, `Grupo`, `Modalidad`, `Idioma`, `Plan`, `Nómina`, `Nombre del profesor`, `Días`, `Horario` y `Zona horaria`; el script solo necesita la nómina y el horario, y si se llaman distinto se indica con `--nomina-programacion` y `--horario`.
2. Supuse que cuentan para el % nueve cursos del histórico y que "ya lo cursó" se escribe `Acreditado`, `Sí`, `X`, `Completo`, `Aprobado`, `Cursado`, `Concluido` o una fecha; vacío, `Pendiente` o `En curso` cuentan como falta.
3. Supuse que el TCM se ve en una columna `Zona horaria`. TCM es tu palabra (del muro); TCH y TCP los inventé yo como los otros valores de esa columna, así que dime qué significa TCM y qué otros valores trae de verdad; si viene dentro del horario (por ejemplo, "LuJu 07:00-08:30 TCM"), se corre con `--horario Horario`.
4. Supuse qué cursos se le piden a cada profesor: `Canvas: Conceptos básicos` a todos, los tres de Tec21 solo al plan 2019, los cuatro de Educación Digital a quien imparte una UDF no presencial e `Inglés Acreditado` a quien imparte en inglés; un profesor con varias UDF junta los requisitos de todas.
5. Supuse que el bloque de Prepa y las columnas `Estatus Tec21`, `Estatus Prepa` y `Estatus ED` no cuentan; si `Estatus` ya trae el dato que usas, el % podría salir de ahí.
6. Supuse que "revisar las actualizaciones de nómina, nombre, UDF, modalidad, idioma y plan" quiere decir que esas columnas pasan a la lista nueva; si quiere decir ver qué cambió desde ayer, eso es `comparar-versiones` de la Sesión 7, corrido antes.
7. Supuse que, si el nombre del profesor está escrito distinto en el histórico, el script lo marca y no lo corrige; cruza solo por nómina.
8. Supuse que un profesor programado sin histórico sale con todos sus cursos como faltantes, en la hoja Sin histórico, y no entra al promedio.
9. Supuse que un valor que no se reconoce (por ejemplo, `Revisar`) cuenta como falta, y el control dice cuál fue y cuántas veces.
10. Supuse que el Excel para las coordinadoras es la hoja Avance por profesor, con un renglón por profesor; si lo quieren por escuela o por coordinadora, es un cambio chico.
11. Supuse que la habilidad no manda nada al equipo de coordinadoras; mandarlo es un paso aparte.

Los encabezados del histórico de capacitación son los que tú escribiste en el muro, con todo y la columna sin nombre que trae tu lista.

## Tu flujo, como lo escribiste

- **Llega:** el archivo de Excel con la información de programación, actualizado diariamente, y el archivo de Excel con el histórico de capacitación, actualizado diariamente.
- **Regla:** revisar las actualizaciones de nómina, nombre de profesor, UDF que imparte, modalidad de la UDF, idioma de la UDF y plan (2019, 2016); eliminar todos los horarios diferentes a TCM; hacer una nueva lista con todos los profesores y revisar en ella el avance y detalle de capacitación del histórico.
- **Sale:** un Excel con la información por profesor del detalle de avance de capacitación y el % de avance.
- **Se entrega:** al equipo de coordinadoras cada vez que se actualiza la información.

## Qué hace

La habilidad `avance-capacitacion` hace tus tres pasos en orden:

1. Quita de la programación los renglones que no son TCM.
2. Arma la lista nueva: un profesor por nómina, con las UDF que imparte en TCM, su modalidad, idioma y plan.
3. Busca a cada profesor en el histórico por nómina y le calcula qué cursos se le piden, cuáles ya cursó, cuáles le faltan y su % de avance.

Deja en `salida/`:

- `avance-capacitacion-NOMBRE.xlsx` con cinco hojas: Resumen, Avance por profesor (un renglón por profesor, con su % y lo que le falta), Detalle (un renglón por profesor y curso, con el valor que trae el histórico), Sin histórico (profesores programados que no aparecen en el histórico) y Control TCM (cuántos renglones había por horario).
- `resumen-avance-capacitacion-NOMBRE.txt`: tres líneas con las cifras.

| Archivo o carpeta | Qué es |
|---|---|
| `generar_ficticio.py` | El programa que inventa la copia ficticia de tus dos archivos y anota qué sembró |
| `ficticio/` | La copia ficticia: `programacion-diaria-ficticia.xlsx`, `historico-capacitacion-ficticio.xlsx` y `NOTAS-ficticio.txt` con la respuesta correcta |
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivos espera, los pasos y qué hacer si algo falla |
| `habilidad/avance_capacitacion.py` | El script que hace el cálculo. Solo biblioteca estándar de Python |

## Cómo se prueba afuera

Abre el agente que prefieras (Claude Code, Codex o Antigravity CLI) en esta carpeta, `etna-avance-capacitacion`, y manda:

```text
Corre generar_ficticio.py con python3 -B y dime qué archivos creó y qué dice NOTAS-ficticio.txt.
```

Después:

```text
Lee habilidad/SKILL.md y corre la habilidad sobre los archivos ficticios que acabas de generar. Dime las tres líneas de resumen y si coinciden con lo sembrado en NOTAS-ficticio.txt.
```

Sin agente, desde PowerShell o Terminal en esta carpeta (en Windows, `python` en lugar de `python3`):

```text
python3 -B generar_ficticio.py
python3 -B habilidad/avance_capacitacion.py ficticio/programacion-diaria-ficticia.xlsx ficticio/historico-capacitacion-ficticio.xlsx
```

## Qué siembra la copia ficticia

- 224 renglones de programación: 40 profesores con grupos en TCM, cada grupo repetido en TCH y casi siempre en TCP. Un renglón trae ` tcm ` en minúsculas y con espacios, y sí debe contar.
- 2 profesores que solo aparecen en TCH o TCP: no entran a la lista.
- 3 profesores programados que no están en el histórico.
- 20 profesores en el histórico que no están programados: no salen.
- Un profesor con todos sus cursos y otro con ninguno; 2 profesores con un solo apellido en el histórico; un valor `Revisar` que el script no reconoce; una nómina repetida en el histórico; 3 nóminas escritas en minúsculas y con espacios.
- Lo cursado escrito de tres formas: `Acreditado`, una fecha como texto y una fecha de Excel.

## Lo que debe salir

La salida real de la prueba del 22 de octubre de 2026, con la semilla 22:

```text
Aviso: 1 nómina aparece más de una vez en historico-capacitacion-ficticio.xlsx (L01357666); uso el primer renglón.

Resumen:
Avance de capacitación de 40 profesores programados en horario TCM (79 de 224 renglones de programacion-diaria-ficticia.xlsx), cruzado con historico-capacitacion-ficticio.xlsx por nómina.
Promedio de avance: 52.7 %. Con 100 %: 2; de 50 a 99 %: 23; menos de 50 %: 12; sin histórico: 3.
Cursos que más faltan: Canvas: Conceptos básicos (25), Turnitin (19) y Sensibilización a la microenseñanza (17). Detalle: avance-capacitacion-programacion-diaria-ficticia.xlsx.

Control:
  Archivo de programación: programacion-diaria-ficticia.xlsx
  Histórico de capacitación: historico-capacitacion-ficticio.xlsx
  Renglón del encabezado (programación, histórico): 1, 3
  Renglones en programación: 224
  Renglones por horario: TCH 84, TCM 78, TCP 61, tcm 1
  Renglones TCM que se quedan: 79
  Profesores en programación (cualquier horario): 42
  Profesores en TCM (lista nueva): 40
  Profesores que solo aparecen fuera de TCM: 2
  Profesores en el histórico: 59
  Programados con histórico: 37
  Programados sin histórico: 3
  Nombre distinto en el histórico: 2
  Cursos que cuentan: 9
  Valores no reconocidos (cuentan como falta): 1
  Valores no reconocidos: 'Revisar' 1

Archivos:
  salida/avance-capacitacion-programacion-diaria-ficticia.xlsx
  salida/resumen-avance-capacitacion-programacion-diaria-ficticia.txt
```

Las tres líneas son idénticas a las que anota `ficticio/NOTAS-ficticio.txt`, que el generador calcula desde lo que sembró y sin leer los archivos. También se probó con otras tres semillas (7, 101 y 2026), con los dos archivos guardados como CSV separado por punto y coma (mismas cifras) y que falle bien: con una columna de horario que no existe, con un archivo vacío, con el histórico en el lugar equivocado y con un valor TCM que no aparece, el script se detiene y dice qué falta.

## Cómo se instala en TECgpt

Solo cuando la propuesta esté confirmada contigo y probada con tu archivo real.

1. Copia la carpeta `habilidad` en TECgpt: Configuración, Habilidades, carpeta de habilidades, y cámbiale el nombre a `avance-capacitacion`. Al abrirla deben verse `SKILL.md` y `avance_capacitacion.py` en la primera capa.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube tus dos archivos reales al banco del proyecto o adjúntalos con el clip.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad avance-capacitacion con el archivo de programación [NOMBRE] y el histórico de capacitación [NOMBRE] y procésalos con código. Muéstrame las tres líneas del resumen y dame el Excel de avance. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

El script lee `.xlsx` y `.csv` con la biblioteca estándar de Python: no necesita `openpyxl` ni en tu laptop ni en TECgpt. Si un archivo llega como `.xls` antiguo, ábrelo en Excel y guárdalo como `.xlsx`.

## Qué revisar

- Que el número de renglones por horario del control sea el que ves en tu archivo.
- Que el número de profesores de la lista nueva sea el que esperas.
- En la hoja Avance por profesor, tres profesores que conozcas: que los cursos que se les piden y su % sean los correctos.
- La hoja Sin histórico: si alguien ahí sí tiene capacitación, la nómina se escribe distinto en un archivo y en otro.

## Qué contestar

Para dejar la propuesta lista con tus archivos reales, contesta estas tres preguntas (sin datos, solo nombres de columnas y valores de ejemplo):

1. ¿Nos pasas los encabezados del archivo de programación, sin datos?
2. En el histórico, ¿qué columnas cuentan para el % de avance de cada profesor y qué valor aparece cuando ya lo cursó?
3. ¿En qué columna de la programación ves que el horario es TCM?

**Todo es ficticio.** Los dos archivos de `ficticio/` los generó un programa con datos inventados: ningún nombre ni nómina corresponde a una persona real. Tus archivos reales no entran a esta carpeta ni a tu agente: se trabajan solo en TECgpt Desktop.
