# Sesión 7 · El skill de cada quien: tu regla, construida afuera y corrida en TECgpt

Carpeta de trabajo de la Sesión 7. Llega a tu `kit` con `git pull`, o en el .zip de la página de la sesión; nada de lo que ya tenías cambia y nada de esta carpeta sobrescribe archivos de sesiones anteriores.

La idea de hoy: tu regla se vuelve un skill. El agente lo construye afuera, con una copia ficticia de tu archivo hecha por un script; TECgpt Desktop lo corre con tu archivo real.

| Archivo o carpeta | Qué es | Para qué se usa hoy |
|---|---|---|
| `mi-flujo/` | Tu carpeta de trabajo. Llega vacía, solo con su `LEEME.md` | Ahí deja el agente lo que construyes en los pasos 1 a 3: el script de la copia ficticia, la copia y tu habilidad en `mi-flujo/habilidad/` |
| `versiones/oferta-grupos-2026-10-21-0900.xlsx` | La oferta de grupos del miércoles 21 de octubre a las 9:00, periodo Febrero-Junio 2027: 180 grupos, uno por renglón, con CRN, clave de materia, materia, grupo, campus, modalidad, horario, salón, profesor, cupo y estatus del grupo | El caso de Karla: la versión anterior |
| `versiones/oferta-grupos-2026-10-21-1500.xlsx` | La misma oferta a las 15:00: 182 grupos, con grupos nuevos, grupos que ya no aparecen, cambios de profesor, cupo, estatus, horario y salón, y los renglones en otro orden | El caso de Karla: la versión nueva |
| `skill-de-referencia/comparar-versiones/` | La habilidad `comparar-versiones` (así llama TECgpt a los skills): `SKILL.md`, con cuándo se usa, qué archivos espera, los pasos y qué hacer si algo falla, y `comparar_versiones.py`, el script que compara dos versiones por una clave | El modelo del paso 2. Si tu habilidad no sale, es la que se instala en TECgpt |
| `generador/generar_versiones.py` | El programa que generó las dos versiones con datos inventados | Un ejemplo de copia ficticia hecha por script, como la del paso 1 |
| `generador/NOTAS-generacion.md` | Lo que sembró el generador, celda por celda, y la respuesta correcta | Para comprobar lo que responde tu agente o TECgpt |
| `AGENTS.md` | Las instrucciones de esta carpeta para el agente: qué hay en cada carpeta, qué no se toca y cómo se quieren las respuestas | El agente lo lee al abrirse en `sesion-07` |

**Dónde abrir el agente:** en la carpeta `sesion-07`, dentro de tu kit (`kit\sesion-07` en Windows, `kit/sesion-07` en Mac), con el agente que prefieras: Claude Code, Codex o Antigravity CLI. Todo lo que construyes queda en `sesion-07/mi-flujo/`.

## Tu flujo en cuatro renglones

- **Llega:** ¿qué archivo te llega y de dónde?
- **Regla:** ¿qué le haces siempre igual?
- **Sale:** ¿qué entregas?
- **Se entrega:** ¿a quién y cada cuándo?

Y los encabezados de tu archivo, sin ningún dato. Con eso llenas las instrucciones de abajo.

## Los pasos 1 a 3, en tu agente

Copia cada instrucción, cambia lo que está en MAYÚSCULAS entre corchetes y mándala. Espera a que termine un paso antes de mandar el siguiente.

**Paso 1 · La copia ficticia.** El script inventa los datos con tus encabezados y anota qué sembró: así sabes la respuesta correcta antes de correr el skill.

```text
Trabaja en la carpeta mi-flujo. Voy a construir la habilidad de un flujo de mi trabajo; los datos reales no salen de TECgpt, así que aquí trabajamos con una copia ficticia. Escribe un script de Python, generar_ficticio.py, que cree un archivo de Excel con estas columnas, en este orden: [PEGA AQUÍ TUS ENCABEZADOS]. Que genere 200 renglones con datos inventados pero creíbles para Educación Digital del Tec, sin nombres ni correos de personas reales, y que siembre los problemas que trae mi archivo real: [DESCRIBE LOS PROBLEMAS DE TU ARCHIVO]. El script anota en un archivo de texto qué sembró y cuántos de cada uno. No escribas tú los datos: los genera el script, y el script corre con python3 sin instalar nada. Córrelo y dime qué sembró.
```

**Paso 2 · El skill de tu regla.** La carpeta de la habilidad trae un `SKILL.md` y el script. El script hace el trabajo; el modelo lo ejecuta.

```text
Este es mi flujo en cuatro renglones:
Llega: [QUÉ ARCHIVO TE LLEGA Y DE DÓNDE]
Regla: [QUÉ LE HACES SIEMPRE IGUAL]
Sale: [QUÉ ENTREGAS]
Se entrega: [A QUIÉN Y CADA CUÁNDO]
Crea la habilidad en la carpeta mi-flujo/habilidad/[NOMBRE-CORTO], con un SKILL.md y un script de Python que aplique la regla. Usa como modelo la carpeta skill-de-referencia/comparar-versiones. El script hace todo el trabajo y corre con python3 sin instalar nada; el SKILL.md dice cuándo usarla y cómo correr el script. Pruébala con el archivo ficticio y comprueba que el resultado coincide con lo que sembró el generador.
```

**Paso 3 · Documentarlo y probar que falle bien.** Un buen `SKILL.md` dice cuándo se usa, qué archivo espera y con qué columnas, los pasos, qué entrega y qué hacer si falta algo.

```text
Revisa el SKILL.md de mi-flujo/habilidad/[NOMBRE-CORTO] para que lo entienda otra persona del equipo: cuándo se usa, qué archivo espera y con qué columnas, los pasos, qué entrega y qué hacer si falta una columna o el archivo llega vacío. Después prueba el script con una copia del archivo ficticio a la que le quites una columna y dime qué pasa. Si el script se detiene sin explicar, corrígelo para que diga qué falta.
```

**Qué revisar:** que el generador diga qué sembró y que el skill lo encuentre.

## Paso 4 · En TECgpt Desktop, con tu archivo real

1. Copia la carpeta de tu habilidad, `mi-flujo/habilidad/` y el nombre corto que le diste, en TECgpt: Configuración, Habilidades, carpeta de habilidades. Al abrir la carpeta deben verse el `SKILL.md` y el script en la primera capa. Si queda una carpeta dentro de otra, la habilidad no aparece.
2. Crea un proyecto y habilita ahí la habilidad.
3. Sube tu archivo real al banco del proyecto o adjúntalo con el clip.
4. Elige GPT-5.6 y manda:

```text
Usa la habilidad [NOMBRE-CORTO] con [EL ARCHIVO O LOS ARCHIVOS] y procésalo con código. Muéstrame el resultado. Si no encuentras la habilidad, dímelo antes de hacer otra cosa.
```

## Si no tienes flujo propio: el caso de Karla

Karla, en la Sesión 6: "Ahorita yo reporto un dato y en la tarde tengo que validar si hubo algún cambio, y luego al día siguiente tengo que volver a validar." Con su caso llenas las mismas instrucciones:

- **Tus encabezados** (paso 1): CRN, Clave de materia, Materia, Grupo, Campus, Modalidad, Horario, Salón, Profesor, Cupo, Estatus del grupo
- **Los problemas de tu archivo** (paso 1): el archivo llega varias veces al día, así que genera dos versiones, la de las 9:00 y la de las 15:00; en la segunda, algunos grupos son nuevos, otros ya no están, algunos cambian de profesor, de cupo o de estatus, y los renglones vienen en otro orden
- **Tus cuatro renglones** (paso 2):

```text
Llega: dos versiones del archivo de oferta de grupos del mismo día, la de las 9:00 y la de las 15:00, en Excel.
Regla: compararlas por CRN y decir qué grupos son nuevos, cuáles ya no están y qué celdas cambiaron.
Sale: un Excel con los cambios y tres líneas de resumen.
Se entrega: al equipo de inscripciones, cada vez que llega una versión nueva.
```

- **El nombre corto** (pasos 2 a 4): `comparar-oferta`.

En TECgpt, el archivo real son las dos versiones de `versiones/`. La respuesta correcta está en `generador/NOTAS-generacion.md`: 5 grupos nuevos, 3 que ya no aparecen y 12 grupos con cambios en 15 celdas.

## Para comprobar afuera

Con el agente abierto en `sesion-07`, la habilidad de referencia da la respuesta del caso de Karla:

```text
Ejecuta el script comparar_versiones.py de la carpeta skill-de-referencia/comparar-versiones con las dos versiones de la carpeta versiones, con la clave CRN. Dime cuántos grupos son nuevos, cuántos ya no aparecen y cuántos cambiaron.
```

Tiene que dar estas tres líneas:

```text
Comparación de oferta-grupos-2026-10-21-0900.xlsx contra oferta-grupos-2026-10-21-1500.xlsx, por CRN.
Hay 5 grupos nuevos, 3 grupos que ya no aparecen y 12 grupos con cambios (15 celdas).
Columnas con más cambios: Profesor (4), Cupo (4) y Estatus del grupo (3). Detalle: cambios-oferta-grupos-2026-10-21-1500.xlsx.
```

Si dice que cambiaron casi todos los grupos, comparó por posición y no por CRN: la versión de las 15:00 viene en otro orden.

**Para correr el script sin agente** (desde PowerShell o Terminal, en esta carpeta): `python skill-de-referencia/comparar-versiones/comparar_versiones.py versiones/oferta-grupos-2026-10-21-0900.xlsx versiones/oferta-grupos-2026-10-21-1500.xlsx --clave CRN` (en Mac, `python3`). Deja en `salida/` el Excel de cambios y el resumen de tres líneas. Las versiones no cambian.

**Todo es ficticio.** Las dos versiones las generó un programa con datos inventados: ningún nombre de profesor ni ningún dato de estos archivos corresponde a una persona real. Por eso esta carpeta se puede trabajar con cualquier agente (Claude Code, Codex o Antigravity CLI) y también en TECgpt.

**Los archivos reales** (tu archivo de verdad, el de cada día) no vienen en este kit ni van a la carpeta de tu agente: se trabajan solo en TECgpt Desktop, que es la herramienta del Tec. Afuera se construye con la copia ficticia; adentro se corre con el archivo real.
