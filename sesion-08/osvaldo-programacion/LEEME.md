# Inconsistencias de programación y ECOAs · propuesta de Jorge para Osvaldo

Esta es la propuesta de Jorge para un flujo tuyo. No escribiste tu flujo en el muro, así que la armé con lo que dijiste en la Sesión 5: que estas herramientas se podían llevar al "tema de ECOAs" y a "todo el tema de la programación, identificar inconsistencias o las relaciones". Corre sobre datos ficticios y ya se probó afuera. La idea es que la compares con lo que tú haces y nos digas si se parece, qué le falta o si tu flujo es otro.

## Lo que supuse

**Todo en esta propuesta es supuesto**: el flujo, los encabezados de los dos archivos y qué cuenta como inconsistencia. Estos son todos:

1. Supuse tu flujo: antes de analizar las ECOAs de un periodo, revisas que la programación esté limpia y que cada grupo programado tenga su ECOA; si tu flujo es otro, escríbelo en el muro `cHLBo2Xa`.
2. Supuse los encabezados de la programación (`Periodo`, `CRN`, `Clave`, `Nombre de la materia`, `Grupo`, `Modalidad`, `Nómina del profesor`, `Nombre del profesor`, `% de responsabilidad`, `Días`, `Horario inicio`, `Horario fin`) y de ECOAs (`Periodo`, `CRN`, `Clave`, `Nómina del profesor`, `Nombre del profesor`, `Alumnos candidatos`, `Respuestas`, `Promedio`); el script solo exige `CRN` y `Nómina del profesor`.
3. Supuse nueve tipos de inconsistencia: renglón repetido, datos del grupo distintos en el mismo CRN, % de responsabilidad que no suma 100, nómina vacía o mal escrita, una nómina con dos nombres, choque de horario del profesor, modalidad fuera del catálogo, grupo programado sin ECOA y ECOA de un grupo que no está programado.
4. Supuse que la nómina es una L y ocho dígitos (`L01234567`).
5. Supuse un catálogo de modalidades inventado: Presencial, Digital, Híbrida y Flexible digital, sin importar mayúsculas ni acentos.
6. Supuse que hay choque de horario cuando el mismo profesor tiene dos grupos con un día en común y horas que se enciman en el mismo periodo; como no se consideran bloques ni PMT, dos grupos de bloques distintos saldrían como choque.
7. Supuse que los días se escriben como `Lu Ju`, `Ma Vi`, `Mi` o `Lunes, Jueves`: cuentan las dos primeras letras de cada día.
8. Supuse que el % de responsabilidad debe sumar 100 por CRN, con tolerancia de medio punto, y que `50%`, `50` y `0.5` son lo mismo.
9. Supuse que el cruce con ECOAs es por CRN y nómina, sin comparar nombre ni clave entre los dos archivos.
10. Supuse que el archivo de ECOAs es opcional: sin él, la habilidad solo revisa la programación.
11. Supuse una hoja por tipo con el CRN en la primera columna; no sabemos a quién le entregas el resultado ni en qué formato.
12. Supuse que la sábana de todos los periodos queda fuera; si tu flujo es juntar periodos, es otra habilidad.

## Lo que dijiste en la Sesión 5

- Preguntaste si conviene juntar los archivos de cada periodo en una sola "sábana" de consulta.
- Invitaste al equipo a llevar las herramientas más allá de lo que hacen hoy: el análisis de las ECOAs y la programación, "identificar inconsistencias o las relaciones".

Esta propuesta toma la segunda idea. La sábana queda como pregunta.

## Qué hace

La habilidad `revisar-programacion-ecoas` lee la programación, revisa cada CRN y cada profesor contra los nueve tipos de la tabla y, si le das también los resultados de ECOAs, cruza los dos archivos por CRN y nómina.

Deja en `salida/`:

- `inconsistencias-NOMBRE.xlsx` con la hoja Resumen, la hoja Todas (ordenada por CRN) y una hoja por tipo. En todas, la primera columna es el CRN, después el tipo, la nómina y el detalle.
- `resumen-inconsistencias-NOMBRE.txt`: tres líneas con las cifras.

| Archivo o carpeta | Qué es |
|---|---|
| `generar_ficticio.py` | El programa que inventa la copia ficticia de los dos archivos y anota qué sembró |
| `ficticio/` | La copia ficticia: `programacion-periodo-ficticia.xlsx`, `resultados-ecoas-ficticio.xlsx` y `NOTAS-ficticio.txt` con la respuesta correcta |
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivos espera, los pasos y qué hacer si algo falla |
| `habilidad/revisar_programacion.py` | El script que hace la revisión. Solo biblioteca estándar de Python |

## Cómo se prueba afuera

Abre el agente que prefieras (Claude Code, Codex o Antigravity CLI) en esta carpeta, `osvaldo-programacion`, y manda:

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
python3 -B habilidad/revisar_programacion.py ficticio/programacion-periodo-ficticia.xlsx ficticio/resultados-ecoas-ficticio.xlsx
```

## Qué siembra la copia ficticia

- 70 CRN y 85 renglones de programación; 83 renglones de ECOAs.
- 18 inconsistencias conocidas en 17 CRN: 1 renglón repetido, 2 CRN con datos del grupo distintos, 2 CRN cuyo % no suma 100, 2 nóminas mal escritas (una vacía), 1 nómina con dos nombres, 2 choques de horario, 2 modalidades fuera del catálogo, 3 grupos sin ECOA y 3 ECOAs sin programación (una de ellas es la del grupo con la nómina vacía).
- Cosas de formato que no deben contar: `50%` y `50`, `07:00` y `0700`, modalidad en minúsculas o mayúsculas, nóminas en minúsculas en ECOAs.

## Lo que debe salir

La salida real de la prueba del 22 de octubre de 2026, con la semilla 22:

```text
Resumen:
Revisión de programacion-periodo-ficticia.xlsx (85 renglones, 70 CRN) y su cruce con resultados-ecoas-ficticio.xlsx, por CRN y nómina del profesor.
Hay 18 inconsistencias en 17 CRN: 12 en la programación y 6 en el cruce con ECOAs.
Tipos con más casos: Programado sin ECOA (3), ECOA sin programación (3) y Datos del grupo distintos en el mismo CRN (2). Detalle: inconsistencias-programacion-periodo-ficticia.xlsx.

Control:
  Archivo de programación: programacion-periodo-ficticia.xlsx
  Resultados de ECOAs: resultados-ecoas-ficticio.xlsx
  Renglón del encabezado (programación, ECOAs): 1, 3
  Renglones en programación: 85
  CRN en programación: 70
  Profesores (nóminas) en programación: 30
  Renglones en ECOAs: 83
  CRN y profesor con ECOA: 80
  Inconsistencias: 18
  CRN con alguna inconsistencia: 17
  Por tipo: Renglón repetido 1, Datos del grupo distintos en el mismo CRN 2, % de responsabilidad que no suma 100 2, Nómina vacía o con otro formato 2, Una nómina con dos nombres 1, Choque de horario del profesor 2, Modalidad fuera del catálogo 2, Programado sin ECOA 3, ECOA sin programación 3

Archivos:
  salida/inconsistencias-programacion-periodo-ficticia.xlsx
  salida/resumen-inconsistencias-programacion-periodo-ficticia.txt
```

Las tres líneas son idénticas a las que anota `ficticio/NOTAS-ficticio.txt`, que el generador calcula desde lo que sembró y sin leer los archivos, y los 18 renglones de la hoja Todas son los 18 sembrados, CRN por CRN. También se probó con otras tres semillas (7, 101 y 2026), con los dos archivos como CSV separado por punto y coma (mismas cifras), solo con la programación (12 inconsistencias, sin cruce) y que falle bien: con un `.xls` antiguo o un archivo vacío, el script se detiene y dice qué pasa.

## Cómo se instala en TECgpt

Solo cuando la propuesta esté confirmada contigo y probada con tus archivos reales.

1. Copia la carpeta `habilidad` en TECgpt: Configuración, Habilidades, carpeta de habilidades, y cámbiale el nombre a `revisar-programacion-ecoas`. Al abrirla deben verse `SKILL.md` y `revisar_programacion.py` en la primera capa.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube tus archivos reales al banco del proyecto o adjúntalos con el clip.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad revisar-programacion-ecoas con el archivo de programación [NOMBRE] y los resultados de ECOAs [NOMBRE] y procésalos con código. Muéstrame las tres líneas del resumen y dame el Excel de inconsistencias. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

El script lee `.xlsx` y `.csv` con la biblioteca estándar de Python: no necesita `openpyxl` ni en tu laptop ni en TECgpt.

## Qué revisar

- Si los nueve tipos son las inconsistencias que de verdad buscas, cuáles sobran y cuáles faltan.
- Que el catálogo de modalidades (Presencial, Digital, Híbrida, Flexible digital) sea el del Tec.
- En la hoja Todas, que cada inconsistencia de la copia ficticia se entienda sin abrir los archivos.

## Qué contestar

Para dejar la propuesta lista, contesta estas tres preguntas (sin datos):

1. ¿Este es tu flujo o es otro? Si es otro, escríbelo en el muro `cHLBo2Xa` en cuatro renglones: qué llega, qué regla aplicas, qué sale y a quién se entrega.
2. ¿Nos pasas los encabezados de tu archivo de programación y del de ECOAs, sin datos?
3. De los nueve tipos de inconsistencia, ¿cuáles buscas hoy a mano y cuál te toma más tiempo?

**Todo es ficticio.** Los dos archivos de `ficticio/` los generó un programa con datos inventados: ningún nombre ni nómina corresponde a una persona real. Tus archivos reales no entran a esta carpeta ni a tu agente: se trabajan solo en TECgpt Desktop.
