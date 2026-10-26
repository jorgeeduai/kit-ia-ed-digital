# AGENTS.md

Carpeta de trabajo de la Sesión 8, con datos ficticios. Trae una propuesta de habilidad por persona o equipo de Educación Digital, cada una en su carpeta. Aquí se corre la propuesta con una copia ficticia y se compara con el flujo real de la persona; la habilidad corre con el archivo real solo en TECgpt Desktop.

## Quién soy y en qué área trabajo

Trabajo en Educación Digital del Tec. Tengo un flujo que se repite: me llega un archivo, le aplico siempre la misma regla y entrego un resultado. Trabajo con Excel, Word y Teams; no programo. Jorge me propuso una habilidad para ese flujo y quiero ver qué hace, compararla con lo que hago y decir qué le falta.

## Qué hay en esta carpeta

- Una carpeta por persona o equipo (la lista está en `LEEME.md`). En cada una: `LEEME.md` con la propuesta y sus supuestos, `generar_ficticio.py`, la carpeta `habilidad/` (`SKILL.md` y el script) y, en algunas, `ficticio/` con la copia ya generada.
- `comparacion-mi-flujo.md`: la hoja de comparación, con cuatro renglones.
- `LEEME.md`: para qué sirve cada carpeta y las instrucciones de la sesión.

## Reglas de trabajo

- Trabaja solo dentro de la carpeta de la persona que te lo pida. Lo que generes (la copia ficticia, la carpeta `salida/`) se queda ahí.
- Todo aquí es ficticio. Los archivos reales no entran a esta carpeta: se usan solo en TECgpt Desktop.
- Nunca escribas datos a mano en un archivo. Los datos ficticios los genera `generar_ficticio.py`, y el script anota qué sembró en `NOTAS-ficticio.txt`.
- Los scripts son de Python y corren con `python3` sin instalar nada: solo la biblioteca estándar, sin `openpyxl` ni `pandas`. No instales paquetes ni propongas instalarlos.
- Corre todo con `python3 -B`, para que no queden carpetas `__pycache__`: la carpeta `habilidad/` se copia a TECgpt y en ella solo van el `SKILL.md` y el script.
- Corre cada script desde la carpeta de la persona, no desde `sesion-08`: así la copia ficticia y la salida quedan donde las espera el `LEEME.md`.

## Qué no tocar

- No modifiques `habilidad/` ni `generar_ficticio.py`. Si la persona quiere cambiar algo, anótalo en `comparacion-mi-flujo.md` y pregúntale antes de tocar un script.
- Pregunta antes de borrar o mover cualquier archivo.
- No mandes nada a nadie: ni correo, ni Teams, ni ningún otro mensaje.

## Cómo quiero las respuestas

- En español, claro y breve.
- Las cifras sin decimales y con separador de miles, salvo los porcentajes; cuando digas una cifra, di de qué archivo salió.
- Al correr la habilidad, copia tal cual las tres líneas de resumen y di si coinciden con `NOTAS-ficticio.txt`.
