# Sesión 5 · TECgpt Desktop: tus archivos con información institucional

Carpeta de trabajo de la Sesión 5. Llega a tu `kit` con `git pull`; nada de lo que ya tenías cambia y nada de esta carpeta sobrescribe archivos de sesiones anteriores.

| Archivo o carpeta | Qué imita | Para qué se usa hoy |
|---|---|---|
| `csv/inscripciones-*.csv` (8 archivos) | Inscripciones de alumnos por grupo en ocho periodos, de agosto-diciembre 2022 a febrero-junio 2026: campus, materia, grupo, profesor, alumno, estatus y calificaciones. 50,220 renglones en total | Práctica 1: preguntas que se responden contando con código (P3 y P4). Práctica 2: el script que cuenta alumnos por campus |
| `csv/cursos-profesores-*.csv` (8 archivos) | Cursos que imparte cada profesor por periodo: grupo, horario, salón, titular o co-titular y alumnos inscritos | Parte de la carpeta que se consulta como una base; sirve para tus propias preguntas |
| `csv/NOTAS-generacion.md` | Cómo se generaron los CSV y las cifras correctas de cada periodo | Para comprobar lo que responde el agente o TECgpt |
| `protocolo-constancias.docx` | Protocolo interno de constancias del área (el mismo de la Sesión 4) | Práctica 1: una pregunta con cita (P2) |
| `generar_csv_ed_digital.py` | El programa que generó los CSV | Si quieres seguir: genera en tu laptop la versión de un millón de renglones |
| `script-de-referencia/alumnos_por_campus.py` | Un script (programa corto) ya escrito y probado que cuenta alumnos distintos por campus y por periodo y guarda `alumnos-por-campus.xlsx` con una gráfica | Práctica 2, y respaldo si el script que escribe tu agente no sale |

Los CSV traen a propósito algunos problemas que aparecen en los archivos reales: renglones repetidos, grupos con co-titular que repiten a sus alumnos, campus escritos de dos formas (en mayúsculas o con un espacio al final), correos mal escritos. Por eso las preguntas de hoy piden contar con código y no a ojo.

**Para usarlos en TECgpt Desktop:** agrega esta carpeta, `sesion-05`, a TECgpt como se indica en la página de la sesión. Es una carpeta de datos ficticios; con tus archivos reales se trabaja igual, y ahí sí pueden llevar información institucional, porque TECgpt Desktop es la herramienta del Tec.

**Para correr el script de referencia** (TECgpt o tu agente lo pueden ejecutar): `python script-de-referencia/alumnos_por_campus.py` desde esta carpeta. Lee los CSV de `csv/` y deja `alumnos-por-campus.xlsx` en esta carpeta. Para otra carpeta, se agrega su ruta al final: `python script-de-referencia/alumnos_por_campus.py RUTA`.

**Los archivos reales de la sesión** (las cargas de 2025 del equipo) no vienen en este kit: se reciben en la sesión, por TECgpt Desktop o por el Teams del equipo, y se trabajan solo en TECgpt.

Ningún nombre, matrícula, nómina ni correo de estos archivos corresponde a una persona real.
