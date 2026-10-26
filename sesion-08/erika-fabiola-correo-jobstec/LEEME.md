# Correo de JobsTec para profesores nuevos · borrador de habilidad para Erika Fabiola

Borrador de la habilidad de tu flujo, armado después de la Sesión 7 con lo que escribiste en el muro. Está hecho con el mismo modelo que `comparar-versiones` de la Sesión 7: una copia ficticia hecha por script, una habilidad con su `SKILL.md` y su script, y la respuesta correcta anotada antes de correrla. Todo lo de esta carpeta es inventado.

## Lo que supuse

Con lo que escribiste en el muro quedaron huecos, así que el borrador se apoya en estos supuestos:

1. Supuse que tu Excel trae las dos columnas que escribiste, `Nombre del profesor` y `Correo electrónico`; si trae también el curso que va a impartir, basta escribir `{Curso}` en la plantilla para usarlo.
2. Supuse que `Correo electrónico` es el correo personal del candidato; si el Excel del Forms trae otra columna con ese mismo nombre (la de quien llena el formulario), el script se detiene y te pide aclarar cuál es.
3. Supuse que "atendido" quiere decir que ya se le preparó el borrador, aunque todavía no se mande; si armas un borrador y no lo mandas, ese profesor no vuelve a salir hasta que borres su renglón de `atendidos-jobstec.csv`.
4. Supuse que lo que evita mandar dos veces es el correo: el mismo correo con mayúsculas o espacios distintos cuenta como uno, y un profesor que use dos correos recibiría dos borradores.
5. Supuse que el Excel del Forms acumula: cada descarga trae todas las respuestas desde el inicio, y por eso hace falta el registro de atendidos.
6. Supuse una plantilla de correo inventada, con asunto, texto, pasos y firma de ejemplo entre corchetes; falta la liga real de JobsTec.
7. Supuse que el script nunca manda: TECgpt te muestra la lista y un borrador, espera tu sí y no manda nada mientras la plantilla sea de ejemplo.
8. Supuse que, mientras el Forms no exista, la habilidad sirve igual con un Excel que llenes a mano con las dos columnas.
9. Supuse que la fase 2 (revisar documentos en JobsTec y escribir a nóminas) queda fuera.

## Tu flujo, como lo escribiste (fase 1)

- **Llega:** "De un forms recibo un excel con datos para solicitar al departamento de nóminas la creación del número de nómina de un profesor que se integrá como nuevo colaborador Tec."
- **Regla:** "Tomo los datos personales del profesor y su correo electrónico personal."
- **Sale:** "Envío al candidato un correo electrónico con las instrucciones neecesarias para crear su perfil en JobsTec."
- **Se entrega:** "Al candidato se entrega correo electrónico con instrucciones para crear el perfil JobsTec."
- **Encabezados:** "Nombre del profesor y correo electrónico."

## Qué hace

Toma el Excel del Forms y arma **un borrador de correo por renglón**, con la misma plantilla de instrucciones para crear el perfil en JobsTec. **No manda nada.** Mandar el correo es la entrega: lo haces tú, o lo hace TECgpt con tu correo después de que le confirmes la lista.

Además lleva un registro, `atendidos-jobstec.csv`, con los correos que ya tienen borrador. Como el Excel de un Forms acumula todas las respuestas, la siguiente vez que lo corras solo arma los borradores de los renglones nuevos y no prepara dos veces el mismo correo. También aparta los renglones que no se pueden atender (sin nombre, sin correo o con un correo mal escrito) y los que repiten un correo dentro del mismo archivo.

| Archivo o carpeta | Qué es |
|---|---|
| `generar_ficticio.py` | El script que inventa dos copias del Excel del Forms, la del lunes 26 y la del miércoles 28 de octubre, con tus dos encabezados, y anota en `NOTAS-ficticio.txt` qué sembró y la respuesta correcta |
| `habilidad/SKILL.md` | La receta: cuándo se usa, qué archivo espera, los pasos, qué hacer si algo falla y, al final, **la plantilla del correo** |
| `habilidad/preparar_correos.py` | El script que arma los borradores. Corre con `python3` sin instalar nada |

**La plantilla del correo es un texto de ejemplo.** Está al final de `habilidad/SKILL.md`, en la sección "Plantilla del correo", y todo lo que va entre corchetes, como `[EJEMPLO]` o `[LIGA DE JOBSTEC]`, es para que lo sustituyas por tu texto de hoy. Mientras quede algo entre corchetes, el script avisa en el resumen que la plantilla es de ejemplo y que esos borradores no se mandan. Lo que va entre llaves, como `{Nombre del profesor}`, se llena solo con la columna de ese nombre: déjalo así.

## Cómo se prueba afuera

Abre el agente que prefieras (Claude Code, Codex o Antigravity CLI) en una copia de esta carpeta y manda:

```text
Corre con python3 -B el script generar_ficticio.py y después corre habilidad/preparar_correos.py dos veces: primero con entrada/respuestas-forms-2026-10-26.xlsx y luego con entrada/respuestas-forms-2026-10-28.xlsx. Compara lo que dio cada corrida con la respuesta correcta de NOTAS-ficticio.txt y dime si coincide. No mandes ningún correo.
```

**Para correrlo sin agente** (Terminal en Mac o PowerShell en Windows, en esta carpeta; en Windows, `python` en lugar de `python3`):

```text
python3 -B generar_ficticio.py
python3 -B habilidad/preparar_correos.py entrada/respuestas-forms-2026-10-26.xlsx
python3 -B habilidad/preparar_correos.py entrada/respuestas-forms-2026-10-28.xlsx
```

El script lee `.xlsx` y `.csv` con la biblioteca estándar de Python: no necesita openpyxl ni ningún paquete, ni en tu laptop ni en TECgpt.

## Lo que debe salir

Salida real de la prueba del 22 de octubre, copiada tal cual.

`NOTAS-ficticio.txt`, la respuesta correcta que anota el generador:

```text
RESPUESTA CORRECTA
  1. Con el del lunes y el registro de atendidos vacío:
     20 renglones: 15 borradores, 2 repetidos en el mismo archivo, 3 por revisar.
  2. Después, con el del miércoles y el registro que dejó la corrida 1:
     26 renglones: 5 borradores, 18 ya atendidos, 3 por revisar.
     (los 18 ya atendidos son los 15 del lunes, sus 2 repeticiones y el que volvió a
     llenar el Forms)
  3. Si se corre otra vez con el del miércoles: 0 borradores, 23 ya atendidos, 3 por revisar.
```

Primera corrida, con el Excel del lunes:

```text
Aviso: la plantilla todavía trae textos de ejemplo entre corchetes: [EJEMPLO], [EJEMPLO: este texto se sustituye por el correo que se manda hoy], [LIGA DE JOBSTEC] y [NOMBRE Y ÁREA DE QUIEN FIRMA].

Resumen:
Borradores de correo de JobsTec a partir de respuestas-forms-2026-10-26.xlsx.
Hay 15 borradores nuevos, 0 renglones ya atendidos, 2 repetidos en el mismo archivo y 3 por revisar (de 20 renglones).
Detalle: borradores-respuestas-forms-2026-10-26.xlsx. Por revisar: renglones 6, 14 y 21. Aviso: la plantilla todavía es de ejemplo; no mandes estos borradores.

Control:
  Archivo del Forms: respuestas-forms-2026-10-26.xlsx
  Renglón del encabezado: 1
  Columnas usadas: Nombre del profesor y Correo electrónico
  Plantilla: SKILL.md, sección Plantilla del correo
  Registro de atendidos: salida/atendidos-jobstec.csv
  Correos en el registro antes de esta corrida: 0
  Renglones con datos: 20
  Borradores nuevos: 15
  Ya atendidos: 0
  Repetidos en el mismo archivo: 2
  Por revisar: 3
  Comprobación: 15 + 0 + 2 + 3 = 20 renglones

Archivos:
  salida/borradores-respuestas-forms-2026-10-26.xlsx
  salida/borradores-respuestas-forms-2026-10-26.txt
  salida/resumen-respuestas-forms-2026-10-26.txt
  salida/atendidos-jobstec.csv

No se mandó ningún correo.
```

Segunda corrida, con el Excel del miércoles (las tres líneas del resumen y la comprobación):

```text
Borradores de correo de JobsTec a partir de respuestas-forms-2026-10-28.xlsx.
Hay 5 borradores nuevos, 18 renglones ya atendidos, 0 repetidos en el mismo archivo y 3 por revisar (de 26 renglones).
Detalle: borradores-respuestas-forms-2026-10-28.xlsx. Por revisar: renglones 6, 14 y 21. Aviso: la plantilla todavía es de ejemplo; no mandes estos borradores.
  Comprobación: 5 + 18 + 0 + 3 = 26 renglones
```

Las dos corridas coinciden con la respuesta correcta. Una tercera corrida con el Excel del miércoles dio 0 borradores nuevos, 23 ya atendidos y 3 por revisar, y guardó sus archivos como `...-2.xlsx` sin sobrescribir los anteriores.

Así se ve un borrador en `borradores-respuestas-forms-2026-10-26.txt` (nombre y correo inventados):

```text
Borrador 1 de 15 (renglón 2)
Para: sergio.barragan@ejemplo.com
Asunto: [EJEMPLO] Instrucciones para crear tu perfil en JobsTec

[EJEMPLO: este texto se sustituye por el correo que se manda hoy]

Hola, Sergio Barragán Escamilla:

Te damos la bienvenida al proceso para integrarte como profesor colaborador del Tec. El siguiente paso es crear tu perfil en JobsTec:

1. Entra a [LIGA DE JOBSTEC].
2. Regístrate con este correo: sergio.barragan@ejemplo.com.
...
```

También se probó que falle bien: sin la columna del correo, con el archivo vacío, con la columna del correo dos veces y con una plantilla que usa `{Curso}` en un archivo sin esa columna. En los cuatro casos el script se detiene, dice qué falta y no deja archivos a medias.

## Cómo se instala en TECgpt Desktop

1. Sustituye la plantilla de ejemplo por tu correo de hoy, en la sección "Plantilla del correo" de `habilidad/SKILL.md`. Deja la primera línea como `Asunto: ...` y deja las llaves.
2. Copia la carpeta `habilidad/` en TECgpt (Configuración, Habilidades, carpeta de habilidades) y ponle de nombre `correo-jobstec`. Al abrirla deben verse el `SKILL.md` y `preparar_correos.py` en la primera capa.
3. Crea un proyecto, habilita ahí la habilidad y sube el Excel del Forms. Si ya la corriste antes, sube también el `atendidos-jobstec.csv` de la vez anterior.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad correo-jobstec con el Excel del Forms y procésalo con código. Si subí atendidos-jobstec.csv, úsalo como registro. Muéstrame el resumen de tres líneas y los renglones por revisar. No mandes ningún correo. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

5. Descarga el `atendidos-jobstec.csv` que deja y guárdalo: es lo que evita preparar dos veces el mismo correo la próxima vez.

Para mandar, revisa primero los borradores. Si quieres que los mande TECgpt, pídeselo en un mensaje aparte; la habilidad le indica mostrarte la lista de destinatarios y un borrador completo y esperar tu sí antes de mandar.

## Qué revisar

- Que el resumen diga los mismos números que `NOTAS-ficticio.txt`.
- Que un borrador del `.txt` se lea como el correo que mandas hoy, una vez que pongas tu plantilla.
- Que en TECgpt la segunda corrida, con el registro, no vuelva a preparar los correos de la primera.

## Qué contestar

Tres preguntas para la Sesión 8:

1. ¿Nos pasas el texto del correo que mandas hoy, sin nombres ni ligas personales?
2. ¿Qué columnas va a tener el Excel donde se junten los datos: nombre, correo personal, curso, coordinadora?
3. ¿Quieres que TECgpt deje los borradores para que tú los mandes, o que los mande él?

La fase 2 (revisar los documentos en JobsTec con el área de credenciales y escribir a nóminas) no está en este borrador.
