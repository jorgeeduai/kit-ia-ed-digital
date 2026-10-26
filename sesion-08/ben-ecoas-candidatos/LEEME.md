# Alumnos candidatos de las ECOAs de tutores · propuesta de Jorge para Ben

Esta es la propuesta de Jorge para tu flujo, armada con lo que escribiste en el muro de la Sesión 7 y lo que contaste en la clase. Corre sobre datos ficticios y ya se probó afuera. La idea es que la compares con lo que tú haces y nos digas qué le falta o qué está mal supuesto.

## Lo que supuse

Los encabezados del archivo de ECOAs son los tuyos. El segundo archivo, el que mandas antes a escolar, no lo conocemos, así que hay supuestos. Estos son todos:

1. Supuse que el archivo que mandas a escolar trae un renglón por alumno con `CRN`, `Matrícula`, `Nómina del Profesor` y `Nombre del profesor`; si ya trae el número por tutor (por ejemplo, `Alumnos asignados`), el script suma esa columna.
2. Supuse que la llave para el conteo es el CRN más la nómina del tutor; si lo que une a cada alumno con su tutor es otra cosa (el grupo de tutoría o la tutora líder), el cruce se rehace.
3. Supuse que en un CRN con varios tutores el error es siempre la suma del CRN repetida en cada tutor; si con tu archivo real salen muchos "otro número", el error es distinto.
4. Supuse que, cuando ECOAs y escolar no coinciden, el número correcto es el de escolar.
5. Supuse que un tutor que está en ECOAs y no tiene alumnos en escolar no se toca y queda por revisar; podría ser un tutor que se agregó después del envío.
6. Supuse que un tutor que está en escolar y no en ECOAs solo se avisa y no se agrega al archivo de ECOAs.
7. Supuse que un alumno repetido en escolar (mismo CRN, tutor y matrícula) cuenta una vez; si un alumno tiene dos tutores en un CRN, cuenta en los dos.
8. Supuse que, si un tutor sale dos veces en el mismo CRN en ECOAs, el segundo renglón queda por revisar.
9. Supuse que el archivo de ECOAs se lee de la primera hoja y que la salida conserva sus columnas y su orden, con solo `Alumnos Candidatos` cambiado; si las directoras reciben otro formato, hay que saber cuál.

## Tu flujo, como lo escribiste

- **Llega:** el archivo de las ECOAs de los tutores, que descargas.
- **Regla:** revisar que el número de alumnos candidatos a contestar sea igual a la cantidad de alumnos que se enviaron previamente para asignar la ECOA; se cambia la información porque se duplican los candidatos.
- **Sale:** un Excel con el número de alumnos candidatos correcto, para hacer el análisis de los promedios como corresponde.
- **Se entrega:** a las directoras y coordinadoras.

En clase contaste el problema así: "si una materia tiene siete mentores, me pone en cada tutor la suma de los alumnos de cada tutor".

## Qué hace

La habilidad `corregir-candidatos-ecoas` cuenta en el archivo de escolar cuántos alumnos tiene cada tutor en cada CRN, revisa tutor por tutor el archivo de ECOAs y le pone a cada uno su número. Dice en cada renglón qué pasó: ya estaba bien, traía la suma del CRN, traía otro número o hay que revisarlo.

Deja en `salida/`:

- `candidatos-corregidos-NOMBRE.xlsx` con cuatro hojas: Resumen, ECOAs corregidas (tu archivo completo, con el número correcto, el de antes y qué pasó), Cambios (solo los tutores que cambiaron) y Por revisar.
- `resumen-candidatos-corregidos-NOMBRE.txt`: tres líneas con las cifras.

| Archivo o carpeta | Qué es |
|---|---|
| `generar_ficticio.py` | El programa que inventa la copia ficticia de tus dos archivos y anota qué sembró |
| `ficticio/` | La copia ficticia: `ecoas-tutores-ficticio.xlsx`, `enviado-a-escolar-ficticio.xlsx` y `NOTAS-ficticio.txt` con la respuesta correcta |
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivos espera, los pasos y qué hacer si algo falla |
| `habilidad/corregir_candidatos.py` | El script que hace la corrección. Solo biblioteca estándar de Python |

## Cómo se prueba afuera

Abre el agente que prefieras (Claude Code, Codex o Antigravity CLI) en esta carpeta, `ben-ecoas-candidatos`, y manda:

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
python3 -B habilidad/corregir_candidatos.py ficticio/ecoas-tutores-ficticio.xlsx ficticio/enviado-a-escolar-ficticio.xlsx
```

## Qué siembra la copia ficticia

- 40 CRN con 114 tutores. En los CRN con varios tutores, cada tutor trae la suma del CRN (tu caso), salvo en dos CRN que ya vienen bien.
- Un CRN de un solo tutor con 3 alumnos de más (otro número).
- Un tutor que no está en el archivo de escolar y otro que está en escolar y no en ECOAs.
- Dos alumnos repetidos en el archivo de escolar.
- Cosas de formato que no deben contar: el CRN como texto en un archivo y como número en el otro, y nóminas en minúsculas y con un espacio adelante.

## Lo que debe salir

La salida real de la prueba del 22 de octubre de 2026, con la semilla 22:

```text
Aviso: 2 alumnos vienen repetidos en enviado-a-escolar-ficticio.xlsx (mismo CRN, tutor y matrícula): se cuenta una vez.

Resumen:
Corrección de Alumnos Candidatos de ecoas-tutores-ficticio.xlsx con enviado-a-escolar-ficticio.xlsx, por CRN y Nómina del Profesor.
De 114 tutores en 40 CRN: 96 corregidos (95 traían la suma del CRN y 1 otro número), 17 ya estaban bien y 2 por revisar.
Alumnos candidatos en total: 6,007 antes y 1,492 después; en el archivo de escolar, 1,475. Detalle: candidatos-corregidos-ecoas-tutores-ficticio.xlsx.

Control:
  Archivo de ECOAs: ecoas-tutores-ficticio.xlsx
  Archivo de escolar: enviado-a-escolar-ficticio.xlsx
  Llave: CRN y Nómina del Profesor
  Renglón del encabezado (ECOAs, escolar): 1, 3
  Cómo se cuentan los alumnos de escolar: un renglón por alumno, sin repetir matrícula
  Renglones de ECOAs (tutores): 114
  CRN en ECOAs: 40
  Renglones en escolar: 1,477
  Tutores en escolar (CRN y nómina): 114
  Corregidos: traían la suma del CRN: 95
  Corregidos: traían otro número: 1
  Ya estaban bien: 17
  Por revisar: tutor sin alumnos en escolar: 1
  Por revisar: tutor solo en escolar: 1
  Alumnos repetidos en escolar (se cuentan una vez): 2

Por revisar:
  CRN 60040, nómina L04887977: el tutor no está en el archivo de escolar
  CRN 60109, nómina L03888120: está en escolar y no en el archivo de ECOAs

Archivos:
  salida/candidatos-corregidos-ecoas-tutores-ficticio.xlsx
  salida/resumen-candidatos-corregidos-ecoas-tutores-ficticio.txt
```

Las tres líneas son idénticas a las que anota `ficticio/NOTAS-ficticio.txt`, que el generador calcula desde lo que sembró y sin leer los archivos, y los dos tutores por revisar son los dos sembrados. La diferencia entre 1,492 y 1,475 es esperada: el tutor que no está en escolar conserva sus 23 candidatos y el tutor que solo está en escolar (6 alumnos) no entra al archivo de ECOAs. También se probó con otras tres semillas (7, 101 y 2026), con los dos archivos como CSV separado por punto y coma (mismas cifras), con un archivo de escolar que trae el número por tutor en lugar de un renglón por alumno, y que falle bien: sin la columna `Alumnos Candidatos` o con otro nombre para la nómina, el script se detiene y dice qué falta.

## Cómo se instala en TECgpt

Solo cuando la propuesta esté confirmada contigo y probada con tus archivos reales.

1. Copia la carpeta `habilidad` en TECgpt: Configuración, Habilidades, carpeta de habilidades, y cámbiale el nombre a `corregir-candidatos-ecoas`. Al abrirla deben verse `SKILL.md` y `corregir_candidatos.py` en la primera capa.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube tus dos archivos reales al banco del proyecto o adjúntalos con el clip.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad corregir-candidatos-ecoas con el archivo de ECOAs de tutores [NOMBRE] y el archivo que mandé a escolar [NOMBRE] y procésalos con código. Muéstrame las tres líneas del resumen, los tutores por revisar y dame el Excel corregido. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

El script lee `.xlsx` y `.csv` con la biblioteca estándar de Python: no necesita `openpyxl` ni en tu laptop ni en TECgpt.

## Qué revisar

- Tres CRN que conozcas con varios tutores: que el número nuevo de cada tutor sea el correcto.
- Que la suma de los números nuevos de un CRN sea el número que traía antes cada tutor.
- La hoja Por revisar: por qué un tutor no está en el archivo de escolar.

## Qué contestar

Para dejar la propuesta lista con tus archivos reales, contesta estas tres preguntas (sin datos, solo nombres de columnas):

1. ¿Nos pasas los encabezados del archivo que mandas a escolar, sin datos?
2. En ese archivo, ¿qué columna dice a qué tutor le toca cada alumno?
3. ¿Nos das un ejemplo inventado de un CRN con dos tutores: qué número trae hoy cada uno y cuál debería traer?

**Todo es ficticio.** Los dos archivos de `ficticio/` los generó un programa con datos inventados: ningún nombre, nómina ni matrícula corresponde a una persona real (las matrículas empiezan con FIC). Tus archivos reales no entran a esta carpeta ni a tu agente: se trabajan solo en TECgpt Desktop.
