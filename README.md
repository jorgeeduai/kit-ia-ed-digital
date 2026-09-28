# Kit del participante · archivos de práctica

Archivos **ficticios** que imitan los que maneja Ed Digital. Sirven para practicar con el agente cuando no traes archivos propios, o cuando los tuyos contienen datos personales de profesores o alumnos. Los tres primeros vienen de la Sesión 1; los otros cuatro y la carpeta `plan-b-csv` se agregaron para la Sesión 2.

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

**Cómo tenerlo en tu computadora:** este kit vive en el repositorio público `https://github.com/jorgeeduai/kit-ia-ed-digital`. Abre el agente en `Documents\claude-proyectos` y pídele: "Clona el repositorio https://github.com/jorgeeduai/kit-ia-ed-digital en una carpeta llamada kit y otra vez en una carpeta llamada kit-agy". Quedan dos carpetas con los mismos archivos, una para cada agente. Si prefieres hacerlo tú, en PowerShell, desde `claude-proyectos`: `git clone https://github.com/jorgeeduai/kit-ia-ed-digital kit` (y lo mismo con `kit-agy`).

**Cómo actualizarlo en la siguiente sesión:** abre el agente en `claude-proyectos` y pídele "Actualiza las carpetas kit y kit-agy con git pull". Tus archivos nuevos (tu AGENTS.md, los Excel que creó el agente) no se pierden. Si el agente avisa que cambiaste un archivo del kit, pídele que descarte esos cambios y vuelva a actualizar.

**Sin git:** descarga el .zip de la página del curso y deja sus archivos directamente dentro de `Documents\claude-proyectos\kit`, sin la carpeta `kit-participante`.

**Si el agente no puede abrir el Excel** o propone instalar un programa para leerlo: responde No y pídele que use los CSV de `plan-b-csv`. Ahí mismo, `LEEME.md` explica cómo guardar un Excel propio como CSV.

**Sobre los dos AGENTS:** se llaman `AGENTS-plantilla.md` y `AGENTS-ejemplo.md` para que el agente no los cargue como instrucciones al abrirse en la carpeta del kit. Para usar uno, guarda una copia con el nombre `AGENTS.md` en la carpeta donde abres el agente, o pídele al agente que lo haga. Si lo guardas desde el Bloc de notas, elige "Todos los archivos" en el tipo de archivo, para que el nombre no termine en `.txt`.

Ningún nombre ni dato de este kit corresponde a una persona real.
