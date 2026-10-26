# Tecnologías no elegibles en TeachingHub · borrador de habilidad

Borrador de la habilidad de tu flujo, armado con lo que escribiste en el muro de la Sesión 7. Está construido y probado afuera con una copia ficticia; adentro, en TECgpt Desktop, corre con tus dos archivos reales.

## Lo que supuse

No tenemos tus archivos, así que el borrador se apoya en estos supuestos. Si alguno no es así, dínoslo en la Sesión 8.

1. Supuse que `Ficha` y `AppOTecnologia` son columnas de cada pestaña de TeachingHub, y que `Tecnología` y `Motivo` son del listado de no elegibles; si es al revés, solo se cambian los nombres al inicio del script.
2. Supuse que las herramientas de cada ficha vienen en la celda `AppOTecnologia`, varias separadas con coma, punto y coma o diagonal; si vienen en otra columna o en el texto de la ficha, se corre con `--en-toda-la-fila`.
3. Supuse que todas las pestañas tienen las mismas columnas, que cada pestaña es una colección y que cada renglón es una actividad o estrategia; si una pestaña trae otras columnas, el script te avisa y busca en todo el renglón.
4. Supuse que una tecnología cuenta como mencionada solo si aparece como palabra suelta, sin importar mayúsculas ni acentos; si el portal escribe nombres pegados o abreviados, hace falta una columna de sinónimos en el listado.
5. Supuse que saber cuáles tecnologías se agregan como no elegibles se resuelve comparando el listado de hoy con el del corte anterior (`--anterior`); si tu listado ya trae una fecha de alta, conviene usarla.
6. Supuse que la salida es un Excel con las hojas Resumen, Coincidencias, Por pestaña y Por tecnología, más tres líneas de resumen; se ajusta a lo que pida el equipo de TeachingHub.
7. Supuse que el corte cada 3 días lo disparas tú o una tarea programada de TECgpt; el script solo aplica la regla y no lleva calendario.
8. Supuse que partes de los dos Excel ya descargados; sacarlos del portal queda fuera.

## Tu flujo, como lo escribiste

- **Llega:** el archivo de Excel TeachingHub-AppOTecnologias y el archivo de Excel Tecnologías no elegibles.
- **Regla:** identificar las herramientas del listado Tecnologías no elegibles y validarlas en cada una de las pestañas del archivo matriz TeachingHub, con un corte cada 3 días.
- **Sale:** un Excel con las actividades y estrategias de las colecciones que mencionan alguna de las herramientas del listado.
- **Se entrega:** al equipo de TeachingHub, cada vez que se actualizan las tecnologías, informando cuáles se agregan como no elegibles.

Encabezados que escribiste: `Ficha`, `AppOTecnologia`, `Tecnología`, `Motivo`. El borrador supone que `Ficha` y `AppOTecnologia` son de cada pestaña de TeachingHub, y que `Tecnología` y `Motivo` son del listado de no elegibles. Si no es así, dínoslo en la Sesión 8 (abajo, en "Qué contestar").

## Qué hace

Lee el listado de tecnologías no elegibles y revisa, pestaña por pestaña, la columna `AppOTecnologia` del archivo matriz. Cada vez que una ficha menciona una tecnología del listado, la anota con la pestaña y el renglón donde la encontró, así nunca se mezclan valores de una pestaña con otra. Si le das también el listado del corte anterior, dice cuáles tecnologías son nuevas en este corte, que es lo que se informa al equipo de TeachingHub.

Deja un Excel con cuatro hojas (Resumen, Coincidencias, Por pestaña y Por tecnología) y un resumen de tres líneas.

| Archivo o carpeta | Qué es |
|---|---|
| `generar_ficticio.py` | El programa que hace la copia ficticia de tus dos archivos, con tus encabezados y datos inventados, y anota en `NOTAS-ficticio.txt` qué sembró |
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivos espera, los pasos y qué hacer si algo falla |
| `habilidad/buscar_no_elegibles.py` | El script que hace la búsqueda |

## Excel con varias pestañas: en la laptop y en TECgpt

Tu archivo matriz tiene varias pestañas, y eso cambia un poco cómo se lee según dónde corra:

- **En la laptop** (afuera, con tu agente), los scripts del curso corren sin instalar nada. Por eso el generador deja la copia ficticia de dos formas: una carpeta con **un CSV por pestaña** (`TeachingHub-AppOTecnologias/1-Actividades.csv`, `2-Estrategias.csv`...) y el mismo contenido en un solo `.xlsx`. El script lee cualquiera de las dos: la carpeta de CSV o el `.xlsx`, que abre con la biblioteca estándar de Python.
- **En TECgpt Desktop** (adentro, con tu archivo real), Python ya trae `openpyxl`, una herramienta para leer Excel. El script la usa si la encuentra y lee directo el `.xlsx` con todas sus pestañas; no hace falta exportar nada a CSV.

En el control, la línea `Leído con` dice cómo leyó el script el archivo: `CSV, uno por pestaña`, `biblioteca estándar` u `openpyxl`.

## Cómo se prueba afuera

Abre tu agente (Claude Code, Codex o Antigravity CLI) en esta carpeta y manda:

```text
Corre python3 -B generar_ficticio.py y después python3 -B habilidad/buscar_no_elegibles.py ficticio/TeachingHub-AppOTecnologias ficticio/Tecnologias-no-elegibles.csv --anterior ficticio/Tecnologias-no-elegibles-corte-anterior.csv. Compara lo que sale con ficticio/NOTAS-ficticio.txt y dime si encontró todo lo sembrado.
```

Para correrlo sin agente, desde Terminal o PowerShell en esta carpeta, son los mismos dos comandos (en Windows, `python` en lugar de `python3`). Para probar el `.xlsx` en lugar de la carpeta de CSV, cambia `ficticio/TeachingHub-AppOTecnologias` por `ficticio/TeachingHub-AppOTecnologias.xlsx`.

## Qué siembra la copia ficticia

Cuatro pestañas (Actividades, Estrategias, Evaluación y Trabajo colaborativo), 174 fichas en total y un listado de 8 tecnologías inventadas, 2 de ellas nuevas respecto al corte anterior. Lo sembrado:

- 9 fichas que mencionan una tecnología no elegible, escritas de formas distintas: normal, en mayúsculas, en minúsculas con espacios de más, separada con punto y coma, con diagonal, con un comentario entre paréntesis, y una celda con dos tecnologías no elegibles a la vez (10 menciones en total).
- Una tecnología del listado que no aparece en ninguna pestaña.
- Una trampa: 3 fichas con `LienzoEdu`, una tecnología permitida cuyo nombre contiene el de una no elegible (`Lienzo`). No deben salir.
- En cada pestaña, un renglón vacío y una ficha sin tecnología; en el listado, un título arriba del encabezado.

## Lo que debe salir

Salida real del script sobre la copia ficticia (carpeta de CSV, semilla 22, corrida el 22 de octubre de 2026):

```text
Resumen:
Revisión de TeachingHub-AppOTecnologias contra Tecnologias-no-elegibles.csv (8 tecnologías), en 4 pestañas y 174 fichas.
Hay 9 fichas que mencionan alguna tecnología no elegible (10 menciones): Actividades 4, Estrategias 3, Evaluación 2 y Trabajo colaborativo 0.
Tecnologías encontradas: 7 de 8; nuevas en este corte: TriviaMax y DibujaYa (2 fichas). Detalle: no-elegibles-TeachingHub-AppOTecnologias.xlsx.

Control de la revisión:
  Archivo matriz: TeachingHub-AppOTecnologias
  Leído con: CSV, uno por pestaña
  Listado de no elegibles: Tecnologias-no-elegibles.csv
  Listado del corte anterior: Tecnologias-no-elegibles-corte-anterior.csv
  Tecnologías en el listado: 8
  Pestañas revisadas: 4
  Fichas revisadas: 174
  Fichas con alguna no elegible: 9
  Menciones (ficha y tecnología): 10
  Tecnologías encontradas: 7

Fichas encontradas:
  Actividades, renglón 3: ACT-002 Aprendizaje basado en retos -> QuizRayo
  Actividades, renglón 10: ACT-009 Video reflexivo -> QuizRayo
  Actividades, renglón 17: ACT-016 Cuestionario rápido -> NotaVoz
  Actividades, renglón 61: ACT-059 Video reflexivo -> Lienzo
  Estrategias, renglón 12: EST-011 Aprendizaje basado en retos -> MuroLibre
  Estrategias, renglón 12: EST-011 Aprendizaje basado en retos -> ChatAula
  Estrategias, renglón 25: EST-023 Video reflexivo -> ChatAula
  Estrategias, renglón 36: EST-034 Línea del tiempo -> NotaVoz
  Evaluación, renglón 15: EVA-014 Exposición breve -> TriviaMax
  Evaluación, renglón 17: EVA-016 Glosario colaborativo -> DibujaYa

Archivos:
  ficticio/salida/no-elegibles-TeachingHub-AppOTecnologias.xlsx
  ficticio/salida/resumen-no-elegibles-TeachingHub-AppOTecnologias.txt
```

Coincide con `NOTAS-ficticio.txt`: las 9 fichas y 10 menciones sembradas, 0 en Trabajo colaborativo (las 3 trampas con `LienzoEdu` no salen), `FotoCollage` sin ninguna mención y `TriviaMax` y `DibujaYa` como nuevas. Con el `.xlsx` en lugar de la carpeta de CSV salen las mismas fichas, en los mismos renglones; se probó leyendo con `openpyxl` y con la biblioteca estándar.

Si falta una columna, el script se detiene y explica. Por ejemplo, con un listado sin la columna `Tecnología`:

```text
No pude revisar: Falta la columna "Tecnología" en el listado de tecnologías no elegibles (listado-sin-tec.csv): no la encontré en los primeros 10 renglones. Columnas que sí hay: Motivo. Revisa cómo se llama la columna con el nombre de la tecnología.
```

Y si una pestaña no trae `AppOTecnologia`, sigue, busca en todo el renglón y lo avisa:

```text
Aviso: La pestaña Estrategias no trae la columna AppOTecnologia; busqué en todas las celdas del renglón.
```

## Cómo se instala en TECgpt Desktop

1. Copia la carpeta `habilidad/` en TECgpt: Configuración, Habilidades, carpeta de habilidades, con el nombre `tecnologias-no-elegibles`. Al abrirla deben verse `SKILL.md` y `buscar_no_elegibles.py` en la primera capa. Si queda una carpeta dentro de otra, la habilidad no aparece.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube al banco del proyecto, o adjunta con el clip, tus dos archivos reales: el archivo matriz TeachingHub y el listado de Tecnologías no elegibles. Si tienes el listado del corte anterior, súbelo también.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad tecnologias-no-elegibles con el archivo matriz TeachingHub y el listado de Tecnologías no elegibles (y el listado del corte anterior, si lo subí) y procésalos con código. Muéstrame las tres líneas del resumen y el Excel. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

Los archivos reales se trabajan solo en TECgpt; no van a la carpeta de tu agente. Sacar los datos directo del portal no está en este borrador: pide credenciales, así que el punto de partida son los dos Excel ya descargados.

## Qué revisar

- Que el generador diga qué sembró y que el script lo encuentre todo: 9 fichas, 10 menciones, ninguna de `LienzoEdu`.
- Con tu archivo real, en TECgpt: que el control diga el número de pestañas que tiene tu archivo y que no aparezca el aviso de una pestaña sin `AppOTecnologia`. Si aparece, esa pestaña trae otra columna con las herramientas: dinos cómo se llama.
- Que una ficha que tú sabes que menciona una tecnología no elegible salga en la hoja Coincidencias, con su pestaña y su renglón.

## Qué contestar en la Sesión 8

1. De los encabezados que escribiste (`Ficha`, `AppOTecnologia`, `Tecnología`, `Motivo`), ¿cuáles son del archivo de TeachingHub y cuáles del listado de no elegibles?
2. ¿Todas las pestañas de TeachingHub tienen las mismas columnas? ¿Las herramientas vienen siempre en `AppOTecnologia`, separadas con coma?
3. ¿Qué le mandas al equipo de TeachingHub en cada corte: el Excel completo, solo las fichas de las tecnologías nuevas, o las tres líneas?

Todo lo de esta carpeta es ficticio: las tecnologías, las fichas y los motivos los inventó un programa y ninguno corresponde a una herramienta real ni a una ficha del portal.
