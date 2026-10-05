---
name: documentar
description: >-
  Usa este skill cuando la persona termine una tarea sobre archivos (un reporte, un análisis,
  una presentación, una limpieza de datos) y pida documentar, dejar registro, guardar el proceso,
  escribir un ADR, o diga "ya terminamos" o "antes de compactar". Guarda el script que produjo el
  resultado, escribe el registro de decisiones (ADR) y una nota de cómo repetir la tarea, todo
  dentro de la carpeta de trabajo, para que otra persona u otro agente lo retome sin esta conversación.
---

# Documentar una tarea terminada

Lo que se queda en la conversación se pierde. Este skill lo baja a archivos en la carpeta de trabajo.

## Cuándo aplica

- La persona pide "documenta", "deja registro", "guarda el proceso", "escribe el ADR", "ya terminamos".
- Terminó una tarea que produjo un archivo (Excel, Word, PowerPoint, CSV) o una cifra que se va a usar.
- Antes de compactar o cerrar la conversación, si hubo decisiones que no están escritas.

## Pasos

1. **Reconstruye qué se hizo.** Revisa la conversación y anota: archivos de entrada, archivos de salida, qué se decidió y con qué criterios, cifras que se obtuvieron y cómo se comprobaron. Si algo no quedó claro, pregúntalo antes de escribir.

2. **Guarda el script.** Si generaste código para producir el resultado (por ejemplo, Python para leer el Excel o armar el Word), guárdalo en `scripts/` con un nombre que diga qué hace, por ejemplo `scripts/reporte-cursos-en-riesgo.py`. Si lo corriste sin guardarlo, reconstrúyelo en un archivo. Al inicio del archivo, en un comentario, pon una línea con qué hace y cómo se corre.

3. **Escribe el ADR.** Crea `documentacion/adr/ADR-NNNN-titulo-corto.md`, con el número consecutivo al último que exista en esa carpeta (si no hay ninguno, `ADR-0001`). Usa la plantilla de `resources/plantilla-adr.md`: fecha, contexto, decisión, criterios, archivos de entrada y salida, cómo se validó, cómo repetirlo, pendientes. Escribe en español, en registro profesional, con frases completas que entienda una persona que no estuvo en la conversación.

4. **Agrega una línea a la bitácora.** En `documentacion/BITACORA.md` (créalo si no existe) agrega una fila con fecha, qué se hizo, qué archivo salió y qué ADR lo describe. Una línea por tarea, la más reciente al final.

5. **Reporta en tres líneas.** Dile a la persona qué archivos creaste y dónde, y recuérdale que ya puede compactar la conversación.

## Reglas

- **Sin datos personales.** Describe las columnas y los criterios, nunca copies filas con nombres, matrículas, correos o calificaciones individuales. Si el archivo de entrada tiene datos sensibles, dilo en el ADR y documenta su estructura, no su contenido.
- **Solo lo que pasó.** No inventes decisiones ni criterios que no se tomaron. Si una cifra no se comprobó, escribe que no se comprobó.
- **No toques los archivos de trabajo.** Este skill solo crea o actualiza archivos dentro de `scripts/` y `documentacion/`.
- **Mismo formato siempre.** Así cualquier agente y cualquier persona del equipo encuentran lo mismo en el mismo lugar.
