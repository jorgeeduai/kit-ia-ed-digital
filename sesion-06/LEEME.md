# Sesión 6 · El flujo de Cecilia, en vivo: tu flujo en cuatro renglones

Carpeta de trabajo de la Sesión 6. Llega a tu `kit` con `git pull`, o en el .zip de la página de la sesión; nada de lo que ya tenías cambia y nada de esta carpeta sobrescribe archivos de sesiones anteriores.

| Archivo o carpeta | Qué es | Para qué se usa hoy |
|---|---|---|
| `entrada/reporte-inscritos-2026-10-19.xlsx` | El reporte de inscritos del lunes 19 de octubre, como lo exportaría el sistema: peticiones de inscripción del periodo Febrero-Junio 2027 con campus, materia, grupo, profesor, alumno, fecha de solicitud, estatus de la petición y canal. 3,524 renglones en la hoja | Ahora tú: la habilidad lo limpia y da la tabla de peticiones por campus y estatus y las tres líneas para Teams |
| `siguiente/reporte-inscritos-2026-10-21.xlsx` | El mismo reporte el miércoles 21 de octubre: trae las peticiones del lunes con su estatus al día y suma las que llegaron después | La segunda corrida, con la misma instrucción. Si quieres seguir, a tu ritmo |
| `skill-de-referencia/limpiar-reporte/` | La habilidad `limpiar-reporte` (así llama TECgpt a los skills): `SKILL.md`, con cuándo aplica, los pasos y las reglas de limpieza, y `limpiar_reporte.py`, el script que limpia el reporte y arma la tabla | Se instala en TECgpt Desktop por carpeta (abajo) |
| `AGENTS.md` | La personalidad del proyecto "Flujo de Cecilia" y las instrucciones de esta carpeta para el agente: qué hay en cada carpeta, qué no se toca y cómo se quieren las respuestas | Si armas tu propio proyecto en TECgpt con esta carpeta, pega su texto como personalidad |
| `generador/generar_reporte_inscritos.py` | El programa que generó los dos reportes con datos inventados | Adelanto de la Sesión 7, donde cada quien hace la copia ficticia de su propio archivo |
| `generador/NOTAS-generacion.md` | Cómo se generaron los reportes y las cifras correctas de cada uno | Para comprobar lo que responde TECgpt o tu agente |

Los reportes traen a propósito lo que trae un reporte real del sistema: el encabezado en el renglón 5, filas de sección y subtotales por campus, columnas con espacios de más, estatus y campus escritos de varias formas, renglones duplicados, co-titulares que repiten la petición y fechas en cuatro formatos. La habilidad los limpia siempre igual. Del reporte del lunes salen **3,143 peticiones de 1,061 alumnos**; del miércoles, **3,651 peticiones de 1,243 alumnos**.

**Para instalar la habilidad en TECgpt Desktop:** se instala por carpeta. La carpeta es `skill-de-referencia/limpiar-reporte`, completa y descomprimida: al abrirla deben verse `SKILL.md` y `limpiar_reporte.py` en la primera capa, sin otra carpeta en medio. En TECgpt, entra a la sección de habilidades de la app y coloca ahí la carpeta `limpiar-reporte`, donde la app lo indica. Si la habilidad no aparece, revisa que la carpeta esté descomprimida y que `SKILL.md` esté en la primera capa.

**Para correr el script sin TECgpt** (tu agente lo puede ejecutar, o tú desde PowerShell): desde esta carpeta, `python skill-de-referencia/limpiar-reporte/limpiar_reporte.py entrada` (en Mac, `python3`). Lee el reporte de `entrada/`, deja en `salida/` la tabla (`tabla-inscritos-2026-10-19.xlsx`) y el resumen (`resumen-teams-2026-10-19.txt`) y muestra las tres líneas y el control de la limpieza. El reporte original no cambia.

**Todo es ficticio.** Los dos reportes los generó un programa con datos inventados y la forma de los reportes del curso: ningún nombre, matrícula, nómina ni correo de estos archivos corresponde a una persona real. Por eso esta carpeta se puede trabajar con cualquier agente (Claude Code, Antigravity o Codex) y también en TECgpt.

**Los archivos reales** (tu reporte de verdad, los archivos de cargas de Etna) no vienen en este kit ni van a la carpeta de tu agente: se trabajan solo en TECgpt Desktop, que es la herramienta del Tec. Los archivos de cargas llegan por el proyecto compartido "Excel de Etna" o por el Teams del equipo.
