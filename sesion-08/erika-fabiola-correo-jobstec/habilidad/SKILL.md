---
name: correo-jobstec
description: >-
  Usa este skill cuando la persona pida preparar los correos de JobsTec para los profesores
  nuevos, armar los borradores de las instrucciones para crear el perfil en JobsTec, o revisar
  quién del Excel del Forms de altas ya tiene su correo. Toma el Excel (o CSV) que llega del
  Forms con el nombre del profesor y su correo, arma con un script ya probado un borrador por
  renglón con la plantilla de este archivo, aparta los renglones que ya se atendieron o que
  vienen incompletos, y deja un Excel con los borradores y un resumen de tres líneas. El script
  no manda nada.
---

# Preparar los correos de JobsTec para profesores nuevos

Cuando un profesor se integra como colaborador, su coordinadora llena un Forms con su nombre y su correo personal. De ese Excel sale, para cada profesor, el mismo correo con las instrucciones para crear su perfil en JobsTec. Este skill arma esos correos siempre igual, con el script `preparar_correos.py` que está en la misma carpeta que este archivo. Armar los borradores no necesita un modelo: lo hace el script, igual cada vez. El script no necesita instalar nada: corre con cualquier `python3`.

**El script no manda ningún correo.** Prepara los borradores. Mandarlos es la entrega: la hace la persona, o la hace el agente solo si la instrucción lo pide y la persona confirma la lista.

## Cuándo aplica

- La persona dice "prepara los correos de JobsTec", "arma los borradores para los profesores nuevos", "a quién le falta su correo de JobsTec".
- Llegó una versión nueva del Excel del Forms y hay que atender solo a los que no se han atendido.
- Una tarea programada pide preparar los correos de los renglones nuevos del Forms.

## Qué archivos espera

- El Excel que llega del Forms, `.xlsx` o `.csv`. Del Excel se lee la primera hoja.
- Dos columnas: `Nombre del profesor` y `Correo electrónico` (el correo personal del profesor). Se reconocen sin distinguir mayúsculas, acentos ni espacios de más. Si el archivo trae más columnas, no estorban. Si en tu archivo se llaman de otra forma: `--columna-nombre "OTRO NOMBRE"` y `--columna-correo "OTRO NOMBRE"`.
- El encabezado (la fila con los nombres de las columnas) en alguno de los primeros 10 renglones.
- La plantilla del correo: la sección **Plantilla del correo**, al final de este archivo.
- El registro de atendidos, `atendidos-jobstec.csv`, en la carpeta `salida/`. La primera vez no existe y el script lo crea.

## Pasos

1. **Encuentra el archivo del Forms.** Si la persona da la ruta, usa esa. Si no, busca en la carpeta de trabajo, primero en `entrada/`, el Excel más reciente del Forms de altas. Si hay dos y no sabes cuál, pregunta; no elijas uno.

2. **Revisa la plantilla.** Si la plantilla de este archivo todavía trae textos entre corchetes, como `[EJEMPLO]` o `[LIGA DE JOBSTEC]`, es el texto de ejemplo: díselo a la persona antes de seguir. Se pueden armar los borradores para revisarlos, pero no se mandan.

3. **Corre el script.** Desde la carpeta de trabajo:
   `python3 RUTA_DEL_SKILL/preparar_correos.py ARCHIVO`
   donde `RUTA_DEL_SKILL` es la carpeta de este archivo (en Windows, `python` en lugar de `python3`). El script deja en `salida/` tres archivos: `borradores-NOMBRE.xlsx` (hojas Resumen, Borradores, Por revisar y Sin borrador), `borradores-NOMBRE.txt` (cada borrador con Para, Asunto y el texto) y `resumen-NOMBRE.txt` (tres líneas), donde NOMBRE es el nombre del archivo del Forms sin extensión. Además agrega al registro `salida/atendidos-jobstec.csv` los correos a los que les preparó borrador. La carpeta `salida/` queda junto a la carpeta `entrada/` donde está el archivo; si el archivo está en otra carpeta, `salida/` queda dentro de esa carpeta. Con `--salida CARPETA` va donde se indique, y con `--registro RUTA` el registro se lee y se escribe donde se indique. Si la persona subió un `atendidos-jobstec.csv` (el registro de la vez anterior), pásalo siempre con `--registro` y, al terminar, dile dónde quedó el registro actualizado para que lo guarde.

4. **Revisa el control.** El script imprime cuántos renglones leyó y cuántos quedaron en cada estado. Comprueba que los borradores, los ya atendidos, los repetidos y los por revisar sumen los renglones leídos (el script imprime esa suma). Si hay renglones por revisar, díselo a la persona con el número de renglón y el motivo.

5. **Entrega.** Muestra las tres líneas del resumen, tal cual, y la ruta del Excel de borradores. Si la persona pide ver un borrador, muéstralo del archivo `borradores-NOMBRE.txt`, sin cambiarle nada.

6. **Mandar, solo si la instrucción lo pide.** Antes de mandar, muestra a la persona la lista de destinatarios (Para) y un borrador completo, y espera a que diga que sí. Después manda cada borrador de la hoja Borradores a su Para, con su Asunto y su Cuerpo exactamente como están, uno por profesor. No mandes nada si el resumen trae el aviso de plantilla de ejemplo. No mandes a los renglones de las hojas Por revisar ni Sin borrador.

7. **Si no puedes ejecutar Python,** dilo y aplica las reglas de abajo con el código que sí puedas ejecutar. Nunca armes los correos a mano ni sobre una parte del archivo: el resultado tiene que salir del archivo completo y del registro.

## Si algo falla

El script se detiene y explica qué pasa; no deja resultados a medias.

- **Falta una columna:** dice cuál falta, en qué archivo y qué columnas sí hay. Pregunta a la persona cómo se llama en su archivo y vuelve a correr con `--columna-nombre` o `--columna-correo`.
- **La columna del correo aparece dos veces** (en los Excel de Forms, una puede ser el correo de quien llenó el formulario): se detiene y lo dice. Pregunta a la persona cuál es el correo personal del profesor.
- **La plantilla usa un nombre entre llaves que no es columna del archivo** (por ejemplo `{Curso}` y el archivo no trae Curso): se detiene y dice cuál. Corrige la plantilla o pide el archivo con esa columna.
- **El archivo llegó vacío** o solo trae el encabezado: se detiene. Pide el archivo correcto.
- **El archivo no existe, está dañado o no es .xlsx ni .csv:** se detiene. Pide la ruta correcta, o que se abra en Excel y se guarde otra vez como .xlsx o .csv.
- **Un profesor dice que no le llegó el correo** y ya está en el registro: borra su renglón de `atendidos-jobstec.csv` (se abre en Excel) y vuelve a correr el script; le arma su borrador otra vez.

## Reglas para armar los borradores (lo que hace el script)

1. Un renglón del Excel, un estado. Se revisa en este orden:
   - **Por revisar:** falta el nombre, falta el correo o el correo no sirve (sin arroba, con más de una arroba, con espacios por dentro o con el dominio incompleto). No lleva borrador.
   - **Ya atendido:** el correo ya está en el registro de atendidos. No lleva borrador.
   - **Repetido en este archivo:** el mismo correo ya apareció en un renglón de arriba del mismo archivo. No lleva borrador; el de arriba sí.
   - **Borrador:** todo lo demás. Lleva un borrador con la plantilla.
2. Los correos se comparan sin espacios en los extremos y sin distinguir mayúsculas, y así se escriben en el Para.
3. En la plantilla, lo que va entre llaves se llena con la columna de ese nombre: `{Nombre del profesor}` con el nombre, `{Correo electrónico}` con el correo.
4. **Atendido quiere decir que ya se le preparó borrador.** El registro se escribe al armar los borradores, no al mandarlos.
5. El script nunca sobrescribe: si `borradores-NOMBRE.xlsx` ya existe, el nuevo se llama `borradores-NOMBRE-2.xlsx`, y así.
6. Los renglones vacíos no cuentan.

## Reglas

- **No toques el archivo del Forms.** No se modifica ni se mueve; todo lo nuevo va a `salida/`.
- **Mismas tres líneas siempre**, con el mismo formato y los mismos nombres de archivo.
- **Sin datos personales en el resumen.** Las tres líneas solo traen cifras y números de renglón. Los nombres y correos van en el Excel de borradores, que no se pega en un chat.
- **Nada se manda sin el sí de la persona**, y nunca a un correo que no esté en la hoja Borradores.

## Plantilla del correo

**Texto de ejemplo: se sustituye por el correo que se manda hoy, sin cambiar la primera línea (Asunto) ni las llaves.** Mientras traiga textos entre corchetes, el script avisa que es de ejemplo y los borradores no se mandan. La primera línea es el asunto. Lo que va entre llaves se llena solo; puedes usar `{Nombre del profesor}`, `{Correo electrónico}` o cualquier otra columna de tu archivo, escrita igual que en el encabezado.

```text
Asunto: [EJEMPLO] Instrucciones para crear tu perfil en JobsTec
[EJEMPLO: este texto se sustituye por el correo que se manda hoy]

Hola, {Nombre del profesor}:

Te damos la bienvenida al proceso para integrarte como profesor colaborador del Tec. El siguiente paso es crear tu perfil en JobsTec:

1. Entra a [LIGA DE JOBSTEC].
2. Regístrate con este correo: {Correo electrónico}.
3. Llena todos los campos de tu perfil y sube los documentos que se piden.
4. Cuando termines, responde a este correo para avisarnos.

Si tienes alguna duda, responde a este mensaje.

Saludos,
[NOMBRE Y ÁREA DE QUIEN FIRMA]
```
