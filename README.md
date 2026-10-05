# Kit del participante · archivos de práctica

Archivos **ficticios** que imitan los que maneja Ed Digital. Sirven para practicar con el agente cuando no traes archivos propios, o cuando los tuyos contienen datos personales de profesores o alumnos. Los tres primeros vienen de la Sesión 1; los otros cuatro y la carpeta `plan-b-csv` se agregaron para la Sesión 2. **Desde la Sesión 4, cada sesión agrega su propia carpeta `sesion-NN`** con los archivos de sus ejemplos y su `LEEME.md`; lo que ya tenías en el kit no cambia.

| Archivo | Qué imita | Preguntas de arranque |
|---|---|---|
| `padron-cursos-2026.xlsx` | Padrón de cursos con sede, instructor y estatus | "¿Cuántos cursos hay por sede?" · "¿Qué instructores tienen más de 2 cursos?" |
| `lineamientos-ed-digital.docx` | Documento de lineamientos internos | "Resúmelo en 5 puntos" · "¿Qué dice sobre la evaluación de cursos?" |
| `evaluaciones-curso-DI-101.csv` | Exportación de evaluaciones de un curso | "¿Cuál es el promedio?" · "¿Hay comentarios negativos? Agrúpalos" |
| `programacion-ago-dic-2026.xlsx` | Programación de agosto a diciembre de 2026 que llega de otro sistema: sus columnas se llaman distinto y las fechas vienen como día/mes/año | "¿Qué columnas de este archivo corresponden a las del padrón?" · "¿Qué no coincide entre la programación y el padrón?" |
| `comparacion-dos-agentes.xlsx` | Hoja para anotar la misma tarea resuelta en los dos agentes | Se llena a mano durante la práctica; las instrucciones están en su segunda hoja |
| `AGENTS-plantilla.md` | Plantilla en blanco del AGENTS.md, con sus cuatro secciones | "Usa AGENTS-plantilla.md como guía, hazme cinco preguntas y muéstrame el resultado antes de guardarlo como AGENTS.md. No modifiques AGENTS-plantilla.md ni crees un CLAUDE.md." |
| `AGENTS-ejemplo.md` | AGENTS.md lleno para la coordinación de programación de cursos | "Lee AGENTS-ejemplo.md y dime qué le cambiarías para que sirva en mi trabajo" |
| `plan-b-csv/` | Copias en CSV de las hojas Padron Cursos 2026 e Instructores y de la programación, con las mismas filas que el Excel; instrucciones en su `LEEME.md` | "Trabaja con plan-b-csv\padron-cursos-2026.csv en lugar del Excel" |
| `sesion-04/` | Archivos de la Sesión 4: la asistencia a los webinars de septiembre (`asistencia-webinars-septiembre.xlsx`), las solicitudes de constancia (`solicitudes-constancias.csv`), el protocolo de constancias (`protocolo-constancias.docx`) y un skill de documentar ya escrito (`skill-de-referencia/documentar/`); instrucciones en su `LEEME.md` | Ver **Sesión 4 · los skills**, abajo |
| `sesion-04/protocolo-constancias.docx` | Protocolo interno de constancias del área: requisitos, plazo de entrega, seguimiento y reposición; habla del mismo registro que `solicitudes-constancias.csv` | Sesión 4, ejemplo 2: la fuente del cuaderno "Protocolo de constancias" en Gemini Notebook · "¿En cuántos días debe entregarse una constancia y qué excepción hay?" |

**Cómo tenerlo en tu computadora:** este kit vive en el repositorio público `https://github.com/jorgeeduai/kit-ia-ed-digital`. Abre el agente en `Documents\claude-proyectos` y pídele: "Clona el repositorio https://github.com/jorgeeduai/kit-ia-ed-digital en una carpeta llamada kit y otra vez en una carpeta llamada kit-agy". Quedan dos carpetas con los mismos archivos, una para cada agente. Si prefieres hacerlo tú, en PowerShell, desde `claude-proyectos`: `git clone https://github.com/jorgeeduai/kit-ia-ed-digital kit` (y lo mismo con `kit-agy`).

**Cómo actualizarlo en la siguiente sesión:** abre el agente en `claude-proyectos` y pídele "Actualiza las carpetas kit y kit-agy con git pull" (desde la Sesión 3, también `kit-codex`). Tus archivos nuevos (tu AGENTS.md, los Excel, Word y PowerPoint que creó el agente) no se pierden. Tu hoja de comparación llena vive en `kit`: si `git pull` avisa de cambios locales, no descartes nada. Déjalo como está y sigue; el kit que ya tienes basta para la práctica.

**Sesión 3 · la carpeta de Codex:** Codex, el tercer agente, trabaja en su propia carpeta, `kit-codex`, otra copia de este mismo kit con tu mismo AGENTS.md. Se prepara con el agente abierto en `claude-proyectos`: "Actualiza las carpetas kit y kit-agy con git pull. Después clona https://github.com/jorgeeduai/kit-ia-ed-digital en una carpeta llamada kit-codex y copia a ella el AGENTS.md de kit." Codex se abre en `kit-codex` y lee ese AGENTS.md; Claude Code sigue en `kit`. El Word y el PowerPoint que salen de la práctica (`reporte-cursos-en-riesgo.docx` y `cursos-en-riesgo-por-sede.pptx`) quedan en la carpeta del agente que los creó, y las filas nuevas de la comparación van en tu misma `comparacion-dos-agentes.xlsx` de `kit`.

**Sesión 4 · los skills:** un skill es una receta escrita que el agente sigue cuando la tarea lo pide: una carpeta con un archivo `SKILL.md` que empieza con `name` (el nombre) y `description` (cuándo usarlo), y debajo los pasos en texto. El AGENTS.md describe la carpeta y el agente lo lee siempre; el skill describe cómo se hace una tarea y el agente lo lee cuando la tarea lo pide. Los tres agentes aceptan el mismo archivo, pero lo buscan en carpetas distintas:

| Agente | Carpeta de usuario (sirve en todas las carpetas) | Carpeta del proyecto (solo donde se abre) |
|---|---|---|
| Claude Code | `.claude\skills\` en la carpeta de usuario | `.claude\skills\` en la carpeta de trabajo |
| Codex | `.agents\skills\` en la carpeta de usuario | `.agents\skills\` en la carpeta de trabajo |
| Antigravity CLI | ninguna | `.agents\skills\` en la carpeta de trabajo |

Por eso el mismo skill en los tres agentes son dos copias en la carpeta de trabajo: `.claude\skills\documentar\` y `.agents\skills\documentar\`, con el mismo texto. En la sesión, cada quien crea en `kit` su skill de documentar con el agente. Se usa al cerrar una tarea sobre archivos, escribiendo solo `documenta`: el agente guarda el script en `scripts\`, escribe un ADR numerado en `documentacion\adr\` y agrega una línea a `documentacion\BITACORA.md`.

**La carpeta de la Sesión 4:** llega a tu `kit` con `git pull`. Con Claude Code abierto en `kit`, al empezar la sesión: "Actualiza esta carpeta con git pull y dime qué carpeta nueva llegó." Debe llegar `sesion-04`. Dos de sus archivos son las tareas del ejemplo 1: `asistencia-webinars-septiembre.xlsx` con Claude Code y `solicitudes-constancias.csv` con un segundo agente, las dos terminadas con `documenta`. El tercero, `protocolo-constancias.docx`, es la fuente del cuaderno de Gemini Notebook del ejemplo 2.

**Si tu skill de documentar no salió:** `sesion-04` trae uno de referencia en `skill-de-referencia\documentar\`, con su plantilla de ADR en `resources\`. Vive en una carpeta aparte para no chocar con el que creas en la sesión, y ahí ningún agente lo carga: está guardado para copiarlo. Con el agente abierto en `kit`, pídele:

"Copia la carpeta sesion-04/skill-de-referencia/documentar a .claude/skills/documentar y a .agents/skills/documentar, y confirma que existe el SKILL.md en las dos."

Si después, al escribir `documenta`, el agente no lo sigue, sal y vuelve a abrirlo en `kit`.

**Sin git:** descarga el .zip de la página del curso y deja sus archivos directamente dentro de `Documents\claude-proyectos\kit`, sin la carpeta `kit-participante`. Si ya tienes `kit` y solo te falta la carpeta de una sesión, copia del zip esa carpeta (por ejemplo, `sesion-04`) dentro de `kit`.

**Si el agente no puede abrir el Excel** o propone instalar un programa para leerlo: responde No y pídele que use los CSV de `plan-b-csv`. Ahí mismo, `LEEME.md` explica cómo guardar un Excel propio como CSV.

**Sobre los dos AGENTS:** se llaman `AGENTS-plantilla.md` y `AGENTS-ejemplo.md` para que el agente no los cargue como instrucciones al abrirse en la carpeta del kit. Para usar uno, guarda una copia con el nombre `AGENTS.md` en la carpeta donde abres el agente, o pídele al agente que lo haga. Si lo guardas desde el Bloc de notas, elige "Todos los archivos" en el tipo de archivo, para que el nombre no termine en `.txt`.

Ningún nombre ni dato de este kit corresponde a una persona real.
