# Tabla de recursos en inglés · propuesta de Jorge para el equipo de traducción

Propuesta de habilidad para su flujo, armada con lo que escribieron en el muro y lo que mostraron en la Sesión 7. Ustedes ya tienen un prompt escrito para esta tarea; esta propuesta sirve para compararlo: qué hace su proceso que esta no hace, qué hace esta que ustedes no hacían y qué está mal supuesto. Está construida y probada afuera con dos documentos ficticios; adentro, en TECgpt Desktop, correría con sus reportes reales.

## Lo que supuse

No tenemos los encabezados de sus tablas ni sabemos si los reportes son de Word o de Google, así que la propuesta se apoya en estos supuestos. Cada uno está en una constante al inicio del script o del generador, para cambiarlo en cuanto lo confirmen.

| # | Supuesto | De dónde salió |
|---|---|---|
| 1 | Los dos reportes son documentos de Word (`.docx`). Si son de Google, se descargan como Word antes de usar la habilidad | En el muro no dice el formato; en la sesión se habló de "entrar al documento de Google" y también de "el documento Word" |
| 2 | La tabla se reconoce por su título, `Recursos que requieren una versión en inglés`, escrito encima de la tabla (o en su primer renglón), y es el mismo en los dos reportes | Lo escribieron así en el muro, para los dos documentos |
| 3 | En el reporte de traducción la tabla tiene tres columnas: `Nombre del recurso en español`, `Comentarios y enlace donde aparece en Canvas en inglés` y una tercera que no se copia (`Estatus de la revisión`) | Los nombres de las dos primeras son los que se leyeron en voz; el de la tercera es inventado |
| 4 | En el reporte del profesor la tabla tiene cuatro columnas: `Recurso en español`, `Comentarios y enlace en Canvas`, `Respuesta del profesor` y `Comentarios del profesor`, y llega vacía | Dijeron que tiene cuatro columnas; los nombres no se entendieron en la grabación y son inventados |
| 5 | El primer renglón de cada tabla es el de los encabezados | Es lo usual en sus reportes; sin confirmar |
| 6 | Las celdas traen texto, ligas y capturas de pantalla. Las ligas se anotan aparte y las capturas se guardan en una carpeta, con el renglón y la columna en el nombre, para volver a insertarlas | Preguntaron qué pasa con las capturas; en la sesión se propuso un paso intermedio que guarde las imágenes en una carpeta |
| 7 | El script no escribe en el Word del profesor: deja la tabla lista para pegar en CSV y en Excel, y el pegado lo hace TECgpt con una instrucción | Así no hay que instalar nada en la laptop para editar Word |
| 8 | La tercera tabla ("tabla 3", las alternativas para revisión del profesor) queda fuera de esta propuesta | La mencionaron como algo que quizá se sume después |
| 9 | Se copian solo los renglones con algo en las columnas 1 o 2; los vacíos se omiten y se cuentan | Una tabla de Word suele traer renglones vacíos al final |
| 10 | Los nombres de archivo siempre siguen `Reporte_Traducción_[CLAVE]_[NombreUF]` y `ProfesorRevisor_PL2026_UF`, y solo cambia lo de los corchetes; el script usa el nombre solo para nombrar la salida | "Los nombres de archivo siempre se dan igual, la única variante será lo que viene dentro de cada corchete" |

## Su flujo, como lo escribieron

- **Llega:** el documento `Reporte_Traducción_[CLAVE]_[NombreUF]` y el documento `ProfesorRevisor_PL2026_UF`.
- **Regla:** copiar la información de las primeras dos columnas de la tabla "Recursos que requieren una versión en inglés" del reporte de traducción y pasarla a las primeras dos columnas de la tabla del mismo nombre del documento del profesor.
- **Sale:** la tabla del documento `ProfesorRevisor_PL2026_UF` ya completada.
- **Se entrega:** en la carpeta de escritorio de la unidad de formación.

## Qué hace

Abre los dos documentos, encuentra en cada uno la tabla por su título (aunque haya otras tablas antes o después), toma las dos primeras columnas del reporte de traducción renglón por renglón y deja una tabla lista para pegar con los encabezados de la tabla del profesor. Las ligas de cada celda van en su propia columna para que no se pierdan, y cada captura de pantalla se guarda como imagen en `imagenes/`. Antes de terminar comprueba que la tabla del profesor tenga cuatro columnas y que esté vacía.

| Archivo o carpeta | Qué es |
|---|---|
| `generar_ficticio.py` | El programa que hace los dos documentos de Word ficticios, con ligas y capturas dentro de las celdas, y anota en `NOTAS-ficticio.txt` qué sembró |
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivos espera, los pasos, cómo se pega y qué hacer si algo falla |
| `habilidad/copiar_tabla_recursos.py` | El script que saca la tabla |

## Cómo se prueba afuera

Abre tu agente (Claude Code, Codex o Antigravity CLI) en esta carpeta y manda:

```text
Corre python3 -B generar_ficticio.py y después python3 -B habilidad/copiar_tabla_recursos.py ficticio/Reporte_Traducción_FIC1001_NarrativaFicticia.docx ficticio/ProfesorRevisor_PL2026_UF.docx. Compara lo que sale con ficticio/NOTAS-ficticio.txt y dime si copió exactamente lo sembrado.
```

Para correrlo sin agente, desde Terminal o PowerShell en esta carpeta, son los mismos dos comandos (en Windows, `python` en lugar de `python3`).

## Qué siembra la copia ficticia

- Un reporte de traducción con tres tablas: una antes ("Recursos traducidos sin observaciones") y una después ("Alternativas encontradas para revisión del profesor") que no se deben tocar, y en medio la de recursos en inglés con ocho renglones.
- En esos ocho renglones: dos ligas dentro del texto, una captura de pantalla, un renglón vacío que no se copia, una celda con dos párrafos y una liga, una liga escrita como texto sin hipervínculo, una celda con captura y liga a la vez, y texto con acentos y paréntesis.
- Un reporte del profesor con la misma tabla de cuatro columnas, vacía, entre otras dos tablas.

Todo es inventado: la unidad de formación FIC1001, los recursos, las ligas (`example.org`, un dominio que no existe) y las capturas, que son rectángulos de color.

## Lo que debe salir

Salida real del script sobre la copia ficticia (semilla 22, corrida el 22 de octubre de 2026 en el Studio):

```text
Resumen:
De Reporte_Traducción_FIC1001_NarrativaFicticia.docx se tomaron 7 renglones de la tabla Recursos que requieren una versión en inglés (columnas 1 y 2: Nombre del recurso en español y Comentarios y enlace donde aparece en Canvas en inglés; 1 renglón vacío omitido).
Traen 5 ligas y 2 capturas de pantalla; las capturas quedaron en imagenes/ con el renglón y la columna en el nombre.
La tabla de ProfesorRevisor_PL2026_UF.docx tiene 4 columnas y 0 renglones llenos: lista para pegar en Recurso en español y Comentarios y enlace en Canvas. Detalle: tabla-para-profesor-FIC1001_NarrativaFicticia.xlsx.

Control:
  Reporte de traducción: Reporte_Traducción_FIC1001_NarrativaFicticia.docx
  Reporte del profesor: ProfesorRevisor_PL2026_UF.docx
  Tabla: Recursos que requieren una versión en inglés
  Columnas en el reporte: 3: Nombre del recurso en español | Comentarios y enlace donde aparece en Canvas en inglés | Estatus de la revisión
  Columnas en la tabla del profesor: 4: Recurso en español | Comentarios y enlace en Canvas | Respuesta del profesor | Comentarios del profesor
  Renglones de datos en el reporte: 8
  Renglones copiados: 7
  Renglones vacíos omitidos: 1
  Ligas: 5
  Capturas de pantalla guardadas: 2
  Renglones ya llenos en la tabla del profesor: 0

Renglones para pegar (renglón del reporte: columna 1 | ligas | capturas):
  2: Lectura 1. Orígenes del relato (PDF) | 1 liga(s) | sin captura
  3: Video: cómo se construye un personaje | 1 liga(s) | sin captura
  4: Infografía de la estructura en tres actos | 0 liga(s) | renglon-04-col-2-1.png
  6: Rúbrica del reto final | 1 liga(s) | sin captura
  7: Podcast: voces de la narrativa | 1 liga(s) | sin captura
  8: Mapa conceptual del curso | 1 liga(s) | renglon-08-col-2-1.png
  9: Guía de estilo para ensayos (Word) | 0 liga(s) | sin captura

Archivos:
  ficticio/salida/tabla-para-profesor-FIC1001_NarrativaFicticia.csv
  ficticio/salida/tabla-para-profesor-FIC1001_NarrativaFicticia.xlsx
  ficticio/salida/imagenes/renglon-04-col-2-1.png
  ficticio/salida/imagenes/renglon-08-col-2-1.png
  ficticio/salida/resumen-tabla-recursos.txt
```

Coincide con `NOTAS-ficticio.txt`: 7 renglones copiados, el renglón 5 vacío omitido, 5 ligas (4 hipervínculos y 1 escrita como texto), 2 capturas (renglones 4 y 8), la columna 3 sin copiar y las tablas de antes y de después sin tocar.

Si un documento no es de Word, se detiene y explica:

```text
No pude copiar la tabla: El reporte de traducción (falla.gdoc) no es un documento de Word (.docx). Si es un documento de Google, ábrelo y descárgalo como Word: Archivo, Descargar, Microsoft Word (.docx).
```

Y si no encuentra la tabla, dice cuáles sí hay:

```text
No pude copiar la tabla: No encontré la tabla "Recursos en inglés" en el reporte de traducción (Reporte_Traducción_FIC1001_NarrativaFicticia.docx). Tablas que sí hay, con el texto que tienen encima: Recursos traducidos sin observaciones; Recursos que requieren una versión en inglés; Alternativas encontradas para revisión del profesor. Revisa el nombre de la tabla o corre con --titulo.
```

## El pegado al Word del profesor

El script deja la tabla lista y las capturas en una carpeta; pasarla al documento del profesor lo hace TECgpt, con una instrucción, después de que ustedes revisan la tabla. Así se cuida lo que les preocupaba: que las capturas no se pierdan ni muevan la tabla. Cuánto respeta TECgpt el formato del Word al pegar: sin dato hasta probarlo con un reporte real.

## Cómo se instala en TECgpt Desktop

1. Copia la carpeta `habilidad/` en TECgpt: Configuración, Habilidades, carpeta de habilidades, con el nombre `tabla-recursos-ingles`. Al abrirla deben verse `SKILL.md` y `copiar_tabla_recursos.py` en la primera capa. Si queda una carpeta dentro de otra, la habilidad no aparece.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube al banco del proyecto, o adjunta con el clip, los dos reportes en Word.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad tabla-recursos-ingles con el reporte de traducción y el reporte del profesor revisor que subí, y procésalos con código. Muéstrame las tres líneas del resumen y la tabla para pegar. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

5. Revisen la tabla. Si está bien, manden:

```text
Ahora pega la tabla en una copia del reporte del profesor revisor, en las dos primeras columnas de la tabla Recursos que requieren una versión en inglés, en el mismo orden. Cada liga como hipervínculo sobre su texto y cada captura de la carpeta imagenes en la celda que dice su nombre. No cambies las otras columnas ni las otras tablas. Dame el documento para descargarlo.
```

Los reportes reales se trabajan solo en TECgpt; no van a la carpeta del agente.

## Qué revisar

- Que el generador diga qué sembró y que el script lo encuentre todo: 7 renglones, 5 ligas, 2 capturas.
- Con un reporte real, en TECgpt: que el control diga 3 columnas en el reporte y 4 en la tabla del profesor, y que el número de renglones copiados sea el de la tabla aprobada.
- Que cada captura quede en la celda correcta del documento del profesor y que las ligas sigan funcionando.

## Qué contestar

1. ¿Los dos reportes son de Word o documentos de Google?
2. ¿Cómo se llaman exactamente las columnas de la tabla en cada reporte?
3. ¿Las capturas de pantalla tienen que pasar a la tabla del profesor, o basta la liga?

Todo lo de esta carpeta es ficticio: la unidad de formación, los recursos, las ligas y las capturas los inventó un programa.
