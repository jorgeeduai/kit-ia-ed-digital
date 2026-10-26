# Sesión 8 · Mi propuesta de tu flujo

Carpeta de trabajo de la Sesión 8. Llega a tu `kit` con `git pull`, o en el .zip de la página de la sesión; nada de lo que ya tenías cambia y nada de esta carpeta sobrescribe archivos de sesiones anteriores.

La idea de hoy: Jorge recreó el flujo de cada quien a partir de lo que escribió en el muro y lo que contó en clase. Cada propuesta trae un generador de datos ficticios con tus encabezados, una habilidad que aplica tu regla y la lista de lo que se supuso. Tú la corres afuera, la comparas con lo que haces y dices qué le falta para correr con tu archivo real.

## Las carpetas

Una carpeta por persona o equipo. Busca la tuya:

| Carpeta | De quién | Qué hace la propuesta |
|---|---|---|
| `alejandro-tecnologias-no-elegibles/` | Alejandro | Busca las tecnologías no elegibles en cada pestaña del archivo de TeachingHub |
| `ben-ecoas-candidatos/` | Ben | Corrige los alumnos candidatos de las ECOAs de tutores contra lo que se mandó a escolar |
| `cecilia-programacion-vs-iris/` | Cecilia | Cruza la programación con el reporte de IRIS por CRN y nómina y separa las discrepancias por tipo |
| `erika-fabiola-correo-jobstec/` | Erika Fabiola | Prepara un borrador de correo de JobsTec por cada profesor nuevo; el envío lo confirmas tú |
| `etna-avance-capacitacion/` | Etna | Deja solo los horarios TCM de la programación y calcula el avance de capacitación de cada profesor |
| `karla-limpieza-canvas-por-escuela/` | Karla | Deja solo los grupos digitales en los seis archivos de actividad de Canvas, uno por escuela |
| `osvaldo-programacion/` | Osvaldo | Busca inconsistencias en la programación de un periodo y la cruza con los resultados de ECOAs |
| `traduccion-tabla-recursos-ingles/` | Equipo de traducción | Toma las dos primeras columnas de la tabla de recursos en inglés y la deja lista para el documento del profesor revisor |

Diego no tiene carpeta: hoy los datos de su flujo están en una página con usuario y contraseña, y eso no cabe en un skill.

Si no tienes carpeta o tu flujo es otro, escríbelo en el muro de la Sesión 7 en cuatro renglones (qué llega, qué regla aplicas, qué sale y a quién se entrega) y los encabezados de tu archivo, sin datos.

En cada carpeta:

| Archivo o carpeta | Qué es |
|---|---|
| `LEEME.md` | La propuesta para ti: lo que se supuso, tu flujo como lo escribiste, qué hace, qué debe salir y qué contestar |
| `generar_ficticio.py` | El programa que inventa la copia ficticia de tus archivos y anota qué sembró en `NOTAS-ficticio.txt` |
| `ficticio/` | La copia ficticia. No viene en el kit: aparece cuando corres el generador (P1) |
| `habilidad/` | La habilidad: `SKILL.md`, con cuándo se usa y cómo se corre, y el script que aplica la regla |

En la raíz de `sesion-08` está `comparacion-mi-flujo.md`, la hoja que llenas en "Compara y corrige".

**Dónde abrir el agente:** en la carpeta `sesion-08`, dentro de tu kit (`kit\sesion-08` en Windows, `kit/sesion-08` en Mac), con el agente que prefieras: Claude Code, Codex o Antigravity CLI.

## Las instrucciones de hoy

Copia cada instrucción, cambia `<slug>` por el nombre de tu carpeta (por ejemplo, `etna-avance-capacitacion`) y mándala. Espera a que termine una antes de mandar la siguiente.

**P1 · Generar mi copia ficticia.**

```text
Entra a la carpeta <slug> del kit de la sesión 8. Corre generar_ficticio.py con python3 -B y dime qué archivos creó y qué dice NOTAS-ficticio.txt.
```

Cuando termine, escribe "listo 1" en el chat.

**P2 · Correr la propuesta.**

```text
Lee habilidad/SKILL.md de esa carpeta y corre la habilidad sobre los archivos ficticios que acabas de generar. Dime las tres líneas de resumen y si coinciden con lo sembrado en NOTAS-ficticio.txt.
```

**P3 · Mi comparación (opcional, si quieren que el agente ayude).**

```text
Lee LEEME.md de esa carpeta, sobre todo la sección Lo que supuse, y ayúdame a llenar comparacion-mi-flujo.md: pregúntame uno por uno los supuestos y anota mis respuestas.
```

Cuando llenes la hoja, escribe "listo 2" en el chat.

**T1 · En TECgpt (solo quien ya tenga su propuesta confirmada).** Instala la habilidad en la carpeta de habilidades y córrela con el archivo real, como en el paso 4 de la Sesión 7. Una corrida por persona. Los pasos exactos y la instrucción para TECgpt están en la sección "Cómo se instala en TECgpt" del `LEEME.md` de tu carpeta.

## Qué revisar

- Que el generador diga qué sembró y que la habilidad lo encuentre: las tres líneas de resumen tienen que coincidir con `NOTAS-ficticio.txt`.
- Que los encabezados de la copia ficticia se parezcan a los de tu archivo. Si no, eso va en el renglón 3 de tu hoja.

**Todo es ficticio.** Los archivos de cada carpeta los generó un programa con datos inventados: ningún nombre, nómina, matrícula ni correo corresponde a una persona real. Por eso esta carpeta se puede trabajar con cualquier agente.

**Los archivos reales** no vienen en este kit ni van a la carpeta de tu agente: se trabajan solo en TECgpt Desktop. Los que ya mandaste a Jorge los procesa un modelo local en su computadora, sin salir de ahí.
