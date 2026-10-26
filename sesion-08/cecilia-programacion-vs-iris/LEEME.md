# Validar la programación contra IRIS · borrador de habilidad para Cecilia

Este es un borrador de la habilidad de tu flujo, construido con lo que escribiste en el muro de la Sesión 7. Está hecho con datos ficticios y ya se probó afuera. Antes de usarlo con tus archivos reales, revisa la sección "Lo que supuse" y contesta las preguntas de "Qué contestar".

## Lo que supuse

Con lo que escribiste en el muro quedaron huecos, así que el borrador se apoya en estos supuestos:

1. Supuse que las columnas se llaman igual en los dos archivos, con los nombres que usaste en el muro; si no, el CRN y la nómina se cambian con `--crn` y `--nomina`, y para las demás hay que agregar una tabla de equivalencias.
2. Supuse que la clave de cada renglón es el CRN más la `Nómina del profesor`; si un grupo tiene dos profesores, sale en dos renglones con el mismo CRN.
3. Supuse que, cuando un dato no coincide, ninguno de los dos manda: el Excel muestra los dos valores, En programación y En IRIS; si uno manda, se agrega una columna con el ajuste a enviar.
4. Supuse que "Días (inicio/fin)" son dos fechas, el primer y el último día de clase del grupo, en las columnas `Día inicio` y `Día fin`; si son los días de la semana (por ejemplo, lunes y jueves), solo cambian los nombres de las columnas.
5. Supuse que los horarios van en `Horario inicio` y `Horario fin`, las fechas del periodo en `Fecha de inicio de periodo` y `Fecha de fin de periodo`, y el porcentaje en `% de responsabilidad asignada al grupo`.
6. Supuse que el PMT es un código corto (1, 2, 3); no sabemos qué valores trae tu archivo real.
7. Supuse que los atributos del curso se comparan como texto completo; si un grupo tiene varios y cambian de orden entre archivos ("A; B" contra "B; A"), hoy cuenta como discrepancia.
8. Supuse que las diferencias solo de formato no cuentan (fechas, horas y porcentajes escritos de otra forma, como `50`, `50%` y `0.5`) y que un `1` sin signo de % en la responsabilidad es 100 %.
9. Supuse que mayúsculas y acentos sí cuentan como discrepancia ("Fisica I" contra "Física I"); si IRIS guarda todo sin acentos, habría que ignorarlos.
10. Supuse que los datos del grupo se comparan una vez por CRN, con el primer renglón de cada archivo; si dos renglones del mismo CRN traen datos distintos dentro de un archivo, el script te avisa.
11. Supuse que "empezando siempre con el CRN" quiere decir el CRN en la primera columna de cada hoja, con una hoja por tipo de discrepancia; la hoja Todas ya es un solo listado ordenado por CRN.
12. Supuse que IRIS exporta `.xlsx` o `.csv`; un `.xls` antiguo hay que guardarlo como `.xlsx` antes.
13. Supuse que "se entrega a mí misma" quiere decir que la habilidad no manda nada; mandar los ajustes es un paso aparte que haces tú.

## Tu flujo, como lo escribiste

- **Llega:** el archivo de programación y el reporte de grupos de IRIS (el sistema de inscripción).
- **Regla:** comprobar que los datos del archivo de programación coinciden con lo que refleja el sistema: CRN, clave, nombre de la materia, grupo y PMT; días y horarios (inicio y fin); fechas de inicio y fin de periodo; nómina del profesor, nombre del profesor y % de responsabilidad asignada al grupo; atributos del curso.
- **Sale:** un reporte en Excel con el listado detallado por tipo de discrepancia, empezando siempre con el CRN con incidencia.
- **Se entrega:** a ti misma, para enviar los ajustes cuando hay discrepancias.

## Qué hace

La habilidad `validar-programacion-iris` cruza los dos archivos por CRN y nómina del profesor y deja en `salida/`:

- `discrepancias-NOMBRE.xlsx`: la hoja Resumen, la hoja Todas y una hoja por cada tipo de discrepancia. En todas, la primera columna es el CRN, después el tipo, la columna, la nómina (cuando la diferencia es de un profesor), el valor en programación, el valor en IRIS y el nombre de la materia.
- `resumen-discrepancias-NOMBRE.txt`: tres líneas con las cifras.

Los tipos son: CRN solo en programación, CRN solo en IRIS, profesor asignado (otra nómina), materia y grupo, días y horarios, fechas del periodo, nombre o % del profesor, y atributos del curso. Las diferencias que son solo de formato no cuentan como discrepancia: una fecha escrita como `08/02/2027` en un archivo y guardada como fecha de Excel en el otro, una hora `07:30` contra `0730`, un porcentaje `50%` contra `50`.

| Archivo o carpeta | Qué es |
|---|---|
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivos espera, los pasos y qué hacer si algo falla |
| `habilidad/validar_programacion.py` | El script que hace el cruce. Solo biblioteca estándar de Python |
| `generar_ficticio.py` | El programa que inventa la copia ficticia de tus dos archivos y anota qué sembró |
| `ficticio/` | La copia ficticia: `programacion-ficticia.xlsx`, `reporte-iris-ficticio.xlsx` y `NOTAS-ficticio.txt` con la respuesta correcta |

## Cómo se prueba afuera

Abre el agente que prefieras (Claude Code, Codex o Antigravity CLI) en esta carpeta, `cecilia-programacion-vs-iris`, y manda:

```text
Corre generar_ficticio.py con python3 -B y después el script de la carpeta habilidad con los dos archivos de la carpeta ficticio: primero el de programación y después el de IRIS. Dime si las tres líneas del resumen coinciden con las que anota ficticio/NOTAS-ficticio.txt.
```

Sin agente, desde PowerShell o Terminal en esta carpeta (en Windows, `python` en lugar de `python3`):

```text
python3 -B generar_ficticio.py --semilla 22 --salida ficticio
python3 -B habilidad/validar_programacion.py ficticio/programacion-ficticia.xlsx ficticio/reporte-iris-ficticio.xlsx
```

## Lo que debe salir

La salida real de la prueba del 22 de octubre de 2026, con la semilla 22:

```text
Aviso: La columna Estatus está en IRIS y no en programación: no se compara.

Resumen:
Validación de programacion-ficticia.xlsx contra reporte-iris-ficticio.xlsx, por CRN y Nómina del profesor.
Hay 19 CRN con incidencia: 3 solo en programación, 2 solo en IRIS, 2 con otro profesor asignado y 12 con datos distintos (13 celdas).
Tipos con más discrepancias: Días y horarios (4), CRN solo en programación (3) y Materia y grupo (3). Detalle: discrepancias-programacion-ficticia.xlsx.

Control del cruce:
  Archivo de programación: programacion-ficticia.xlsx
  Reporte de IRIS: reporte-iris-ficticio.xlsx
  Clave: CRN y Nómina del profesor
  Renglón del encabezado (programación, IRIS): 1, 3
  Renglones leídos en programación: 219
  Renglones leídos en IRIS: 218
  CRN en programación: 200
  CRN en IRIS: 199
  CRN en los dos: 197
  CRN con incidencia: 19
  Celdas con datos distintos: 13
  Celdas que solo cambian de formato (no cuentan): 1,389
  Por tipo: CRN solo en programación 3, CRN solo en IRIS 2, Profesor asignado 2, Materia y grupo 3, Días y horarios 4, Fechas del periodo 1, Nombre o % del profesor 3, Atributos del curso 2

Archivos:
  salida/discrepancias-programacion-ficticia.xlsx
  salida/resumen-discrepancias-programacion-ficticia.txt
```

Las tres líneas coinciden con las que anota `ficticio/NOTAS-ficticio.txt`, y los 20 renglones de la hoja Todas son los 20 que sembró el generador, CRN por CRN. El aviso de `Estatus` es esperado: la copia ficticia de IRIS trae esa columna de más para probar que el script la avisa y sigue.

También se probó que falle bien: sin la columna `Nómina del profesor` en IRIS, con un archivo vacío, con un `.xls` antiguo y con un renglón repetido, el script se detiene y dice qué falta. Con el reporte de IRIS guardado como CSV (separado por punto y coma) da las mismas cifras.

## Cómo se instala en TECgpt

1. Copia la carpeta `habilidad` en TECgpt: Configuración, Habilidades, carpeta de habilidades, y cámbiale el nombre a `validar-programacion-iris`. Al abrirla deben verse `SKILL.md` y `validar_programacion.py` en la primera capa. Si queda una carpeta dentro de otra, la habilidad no aparece.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube tus dos archivos reales al banco del proyecto o adjúntalos con el clip.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad validar-programacion-iris con el archivo de programación [NOMBRE] y el reporte de grupos de IRIS [NOMBRE] y procésalos con código. Muéstrame las tres líneas del resumen y dame el Excel de discrepancias. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

El script lee `.xlsx` y `.csv` con la biblioteca estándar de Python: no necesita `openpyxl` ni en tu laptop ni en TECgpt. Si IRIS te entrega un `.xls` antiguo, ábrelo en Excel y guárdalo como `.xlsx` antes de correrlo.

## Qué revisar

- Que las columnas de tus archivos reales se llamen como las de la copia ficticia. Si una se llama distinto, el script avisa que no la comparó o se detiene diciendo cuál falta.
- Que el número de CRN que dice el control sea el que esperas de cada archivo.
- En la hoja Todas, unos cuantos CRN que ya conoces: que la discrepancia sea real y esté en el tipo correcto.
- Que lo que el script trata como "solo formato" (fechas, horas y porcentajes escritos distinto) de verdad no te importe.

## Qué contestar

Para dejar la habilidad lista con tus archivos reales, contesta estas tres preguntas (sin datos, solo los nombres de las columnas):

1. ¿Las columnas se llaman igual en tu archivo de programación y en el reporte de IRIS? Si no, ¿cómo se llama cada una en cada archivo?
2. Cuando un grupo tiene dos profesores, ¿aparece en dos renglones con el mismo CRN?
3. Cuando un dato no coincide, ¿cuál manda: el de programación o el de IRIS?

**Todo es ficticio.** Los dos archivos de `ficticio/` los generó un programa con datos inventados: ningún nombre ni nómina corresponde a una persona real. Tus archivos reales no entran a esta carpeta ni a tu agente: se trabajan solo en TECgpt Desktop.
