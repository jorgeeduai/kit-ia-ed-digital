# Limpieza de los archivos de Canvas por escuela · propuesta de Jorge

Propuesta de habilidad para tu flujo, armada con lo que escribiste en el muro y lo que mostraste al final de la Sesión 7. Tú ya tienes un script que hace esta limpieza; esta propuesta sirve para compararlo: qué hace el tuyo que esta no hace, qué hace esta que el tuyo no hacía y qué está mal supuesto. Está construida y probada afuera con una copia ficticia; adentro, en TECgpt Desktop, correría con tus seis archivos reales.

## Lo que supuse

No tenemos los encabezados de tus archivos, así que la propuesta se apoya en estos supuestos. Cada uno está en una constante al inicio del script, para cambiarlo en cuanto lo confirmes.

| # | Supuesto | De dónde salió |
|---|---|---|
| 1 | Cada descarga es un `.xlsx` con cuatro hojas: `Escuela`, `Departamento` (no se ocupa), `Materias` (información general de la materia) y `Actividades` (usuarios activos por actividad) | Mostraste cuatro hojas: la escuela y el departamento, una que no ocupas, la información general de la materia y los usuarios activos de cada actividad |
| 2 | El nombre de la escuela está en la primera hoja, en la celda a la derecha de la que dice `Escuela` | Tu script "lee de qué escuela es, en la primera hoja" |
| 3 | La columna del rol se llama `Rol` y los roles son `Estudiante`, `Profesor`, `Observador` y `Sistema` | Dijiste que la columna C identifica el rol y nombraste estudiante, observador, profesor y uno más que en la grabación se oye como "sistema" |
| 4 | La columna del código se llama `SIS ID del curso` | En la grabación se oye "los códigos del SIS ID curso" |
| 5 | Grupo digital = el código empieza con `Nacional`, sin distinguir mayúsculas | "Selecciono solamente lo que empieza con Nacional, que son los grupos de nosotros" |
| 6 | Del rol solo se quita `Estudiante`; observadores y otros roles se quedan | Dijiste "identifique todos los que tienen rol de estudiante y que elimine esos registros". Si solo quieres profesores, la opción `--solo-profesores` ya está probada |
| 7 | La regla del código se aplica también a la hoja de materias; las hojas sin código se copian igual | "Ese mismo criterio de los grupos nacionales lo repito" en otra hoja |
| 8 | El archivo limpio conserva las cuatro hojas y se llama como la escuela: `Escuela Aurora.xlsx` | "6 archivos de Excel limpios y con el nombre de la escuela" |
| 9 | Si el archivo limpio de una escuela ya existe, no se toca; se reemplaza solo si lo pides | Tu script pregunta si quieres sustituirlo. En TECgpt no hay a quién preguntarle a media corrida, así que se pide al correrlo |
| 10 | Después de filtrar quedan grupos `MFD`, `MFDN` y `MCBL`, y se aceptan | Dijiste que se queda información de esos grupos, "pero pues ya se reduce bastante el tamaño" |
| 11 | El archivo limpio conserva todas las columnas de cada hoja | No dijiste que quitaras columnas |
| 12 | Un archivo con problema no detiene a los otros cinco | Son seis archivos independientes |
| 13 | Procesar solo lo nuevo de cada descarga queda fuera | Lo dejaste para cuando la limpieza completa esté probada |

## Tu flujo, como lo escribiste

- **Llega:** 6 archivos con los datos de las actividades realizadas por todos los profesores en Canvas, uno por escuela.
- **Regla:** reviso que los datos de los profesores correspondan solo a los grupos digitales y limpio cada archivo para dejar solo esta información.
- **Sale:** 6 archivos de Excel limpios y con el nombre de la escuela.
- **Se entrega:** a la plataforma que muestra el seguimiento de profesores.

## Qué hace

Lee los seis archivos de la carpeta de descargas. En cada uno toma el nombre de la escuela de la primera hoja; en la hoja de actividades quita a los estudiantes y deja solo los cursos cuyo código empieza con `Nacional`; en la hoja de materias deja solo esos mismos cursos; las otras hojas pasan igual. Guarda un Excel limpio por escuela y un conteo de control: por cada archivo, cuántos renglones leyó, cuántos quitó por rol, cuántos por código y cuántos quedan, y comprueba que la suma cuadre.

| Archivo o carpeta | Qué es |
|---|---|
| `generar_ficticio.py` | El programa que hace la copia ficticia de tus seis descargas, con las cuatro hojas supuestas y datos inventados, y anota en `NOTAS-ficticio.txt` qué sembró |
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivos espera, los pasos y qué hacer si algo falla |
| `habilidad/limpiar_por_escuela.py` | El script que hace la limpieza |

## En la laptop y en TECgpt

- **En la laptop** (afuera, con tu agente), el script lee y escribe los `.xlsx` con la biblioteca estándar de Python, sin instalar nada.
- **En TECgpt Desktop** (adentro, con tus archivos reales), Python ya trae `openpyxl`, una herramienta para leer Excel. El script la usa si la encuentra, en modo de solo lectura, que es el más ligero para archivos grandes.

Tus archivos reales son pesados (tu script tarda unos 15 minutos en los seis). Cuánto tarda esta propuesta con ellos: sin dato hasta probarla con un archivo real.

## Cómo se prueba afuera

Abre tu agente (Claude Code, Codex o Antigravity CLI) en esta carpeta y manda:

```text
Corre python3 -B generar_ficticio.py y después python3 -B habilidad/limpiar_por_escuela.py ficticio/descargas. Compara lo que sale con ficticio/NOTAS-ficticio.txt y dime si quitó y dejó exactamente lo sembrado.
```

Para correrlo sin agente, desde Terminal o PowerShell en esta carpeta, son los mismos dos comandos (en Windows, `python` en lugar de `python3`).

## Qué siembra la copia ficticia

Seis archivos, `actividad-canvas-01.xlsx` a `actividad-canvas-06.xlsx`, de seis escuelas inventadas, con 580 renglones de actividad en total. En cada archivo:

- Estudiantes, profesores, observadores y el rol `Sistema`, en cursos con código `Nacional...` y en cursos con códigos que no son digitales (`Regional...`, `Sede...`, `Presencial...`).
- Cinco trampas: un rol escrito `  estudiante ` (minúsculas y espacios) y otro `ESTUDIANTE`, que se deben quitar; un curso `Internacional...`, que contiene "nacional" pero no empieza con eso y se debe quitar; y dos códigos escritos `  nacional...` y `NACIONAL...`, que se deben quedar.
- Dos renglones vacíos a media hoja, que no cuentan.
- En la hoja de materias, un curso `Internacional...` que se debe quitar.
- En el archivo 06, el nombre de la escuela con espacios de más: el limpio se debe llamar `Escuela Faro.xlsx`.

## Lo que debe salir

Salida real del script sobre la copia ficticia (semilla 22, corrida el 22 de octubre de 2026 en el Studio):

```text
Resumen:
Limpieza de 6 archivos de actividad de Canvas: 6 escuelas en limpios/.
Actividades: 580 renglones leídos; quedan 87 de grupos digitales (Observador 18, Profesor 57, Sistema 12); fuera 410 por rol Estudiante y 83 por código.
Control: la suma cuadra en los 6 archivos; sin avisos. Detalle: control-limpieza.xlsx.

Control por archivo (hoja de actividades):
  actividad-canvas-01.xlsx -> Escuela Aurora.xlsx: 77 leídos = 53 fuera por rol + 9 fuera por código + 15 quedan (Observador 4, Profesor 8, Sistema 3)
  actividad-canvas-02.xlsx -> Escuela Bosque.xlsx: 107 leídos = 78 fuera por rol + 16 fuera por código + 13 quedan (Observador 2, Profesor 9, Sistema 2)
  actividad-canvas-03.xlsx -> Escuela Cumbre.xlsx: 95 leídos = 68 fuera por rol + 9 fuera por código + 18 quedan (Observador 3, Profesor 13, Sistema 2)
  actividad-canvas-04.xlsx -> Escuela Delta.xlsx: 83 leídos = 53 fuera por rol + 16 fuera por código + 14 quedan (Observador 4, Profesor 10)
  actividad-canvas-05.xlsx -> Escuela Estrella.xlsx: 103 leídos = 71 fuera por rol + 18 fuera por código + 14 quedan (Observador 1, Profesor 10, Sistema 3)
  actividad-canvas-06.xlsx -> Escuela Faro.xlsx: 115 leídos = 87 fuera por rol + 15 fuera por código + 13 quedan (Observador 4, Profesor 7, Sistema 2)
  Materias (solo código): 78 leídas, 42 fuera por código, 36 quedan
  Leído con: biblioteca estándar

Archivos en ficticio/limpios:
  Escuela Aurora.xlsx (nuevo)
  Escuela Bosque.xlsx (nuevo)
  Escuela Cumbre.xlsx (nuevo)
  Escuela Delta.xlsx (nuevo)
  Escuela Estrella.xlsx (nuevo)
  Escuela Faro.xlsx (nuevo)
  control-limpieza.xlsx
  resumen-limpieza.txt
```

Coincide con `NOTAS-ficticio.txt` archivo por archivo: 580 = 410 + 83 + 87, y en materias 78 = 42 + 36. En los Excel limpios no queda ningún estudiante ni ningún curso `Internacional`, y sí quedan los dos códigos escritos `  nacional...` y `NACIONAL...`. Con `openpyxl` (como en TECgpt) sale exactamente lo mismo; solo cambia la línea `Leído con`.

Si lo corres otra vez, no reemplaza nada y lo avisa:

```text
Aviso: Ya existe Escuela Aurora.xlsx; no lo toqué. Si quieres generarlo de nuevo, corre con --reemplazar.
```

Si un archivo no trae la columna del rol, se detiene en ese archivo y dice qué columnas sí hay en cada hoja:

```text
No pude limpiar: ningún archivo quedó limpio. En sin-rol.xlsx ninguna hoja trae la columna "Rol" junto a "SIS ID del curso", así que no sé quién es estudiante. Columnas que sí hay: Escuela: Escuela | Escuela Bosque; Departamento: Departamento | Coordinación | Sede; Materias: SIS ID del curso | Curso | Tipo de materia | Modalidad; Actividades: Usuario | Nómina | SIS ID del curso | Curso | Actividad | Entregas | Calificadas | Sin calificar | Días para calificar.
```

## Cómo compararla con tu script

Corre los dos sobre el mismo archivo (en TECgpt, con uno real; o el tuyo sobre la copia ficticia) y fíjate en:

- **El número que queda por escuela.** Si el tuyo deja menos renglones, probablemente quita también observadores u otros roles (supuesto 6), o filtra por algo más que `Nacional`.
- **Las hojas.** Si el tuyo deja menos hojas o quita columnas, dinos cuáles: la propuesta conserva las cuatro hojas y todas las columnas.
- **El nombre del archivo** que sube a la plataforma.

En la grabación quedó una idea para después: procesar solo lo nuevo de cada descarga, porque el archivo crece con el semestre. Esta propuesta no lo hace todavía; primero se confirma que la limpieza completa da lo mismo que la tuya.

## Cómo se instala en TECgpt Desktop

1. Copia la carpeta `habilidad/` en TECgpt: Configuración, Habilidades, carpeta de habilidades, con el nombre `limpieza-canvas-por-escuela`. Al abrirla deben verse `SKILL.md` y `limpiar_por_escuela.py` en la primera capa. Si queda una carpeta dentro de otra, la habilidad no aparece.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube al banco del proyecto, o adjunta con el clip, las seis descargas.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad limpieza-canvas-por-escuela con los seis archivos de actividad de Canvas que subí y procésalos con código. Muéstrame las tres líneas del resumen y la lista de Excel limpios. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

Los archivos reales se trabajan solo en TECgpt; no van a la carpeta de tu agente.

## Qué revisar

- Que el generador diga qué sembró y que el script lo encuentre todo: 87 renglones quedan, 410 fuera por rol, 83 fuera por código.
- Con un archivo real, en TECgpt: que el control diga la escuela correcta y que la suma cuadre. Si el script dice que no encuentra una columna, te muestra las que sí hay: con eso nos dices cómo se llaman.
- Que el número de renglones que quedan se parezca al que deja tu script en la misma escuela.

## Qué contestar

1. ¿Nos compartes los encabezados de las cuatro hojas, sin datos? Con eso cambiamos los supuestos 1, 3 y 4.
2. ¿El filtro es solo quitar `Estudiante` y quedarse con lo que empieza con `Nacional`, o también quitas observadores u otros roles?
3. ¿Qué hoja no ocupas: la quitas del archivo limpio o la dejas? ¿Y cuál de las cuatro es la que cambia de una descarga a otra?

Todo lo de esta carpeta es ficticio: las escuelas, las materias, las claves, las nóminas y las personas los inventó un programa.
