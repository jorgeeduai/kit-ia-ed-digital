# AGENTS.md

Carpeta de trabajo de la Sesión 7, con datos ficticios. Aquí se construye, con una copia ficticia, la habilidad de un flujo de trabajo de Educación Digital; después la habilidad se instala en TECgpt Desktop y corre allá con el archivo real.

## Quién soy y en qué área trabajo

Trabajo en Educación Digital del Tec. Tengo un flujo que se repite: me llega un archivo, le aplico siempre la misma regla y entrego un resultado. Trabajo con Excel, Word y Teams; no programo. Quiero que esa regla quede en un script que haga el trabajo igual cada vez, con un `SKILL.md` que diga cuándo y cómo usarlo.

## Qué hay en esta carpeta

- `mi-flujo/`: donde va todo lo que construyas: el script que genera la copia ficticia, la copia ficticia y la carpeta `habilidad/` con la habilidad.
- `versiones/`: las dos versiones de la oferta de grupos del caso de Karla, la de las 9:00 y la de las 15:00 del miércoles 21 de octubre.
- `skill-de-referencia/comparar-versiones/`: una habilidad ya probada (`SKILL.md` y `comparar_versiones.py`). Sirve de modelo: para leer y escribir Excel sin instalar nada, usa sus funciones.
- `generador/`: el script que generó las dos versiones y `NOTAS-generacion.md`, con lo que sembró.
- `LEEME.md`: para qué sirve cada archivo y las instrucciones de la sesión.

## Reglas de trabajo

- Trabaja solo dentro de esta carpeta, y lo nuevo, dentro de `mi-flujo/`, también los archivos de prueba y los temporales.
- Todo aquí es ficticio. Los archivos reales no entran a esta carpeta: se usan solo en TECgpt Desktop.
- Nunca escribas datos a mano en un archivo. Los datos ficticios los genera un script, y el script anota qué sembró y cuántos de cada uno.
- Los scripts son de Python y corren con `python3` sin instalar nada: solo la biblioteca estándar, sin `openpyxl` ni `pandas`. No instales paquetes ni propongas instalarlos.
- Después de escribir un script, córrelo y muestra lo que dio. Córrelo con `python3 -B`, para que no queden carpetas `__pycache__`: la carpeta de la habilidad se copia a TECgpt y en ella solo van el `SKILL.md` y el script.

## Qué no tocar

- No modifiques ni muevas nada de `versiones/`, `skill-de-referencia/` ni `generador/`. Si necesitas partir de la habilidad de referencia, cópiala a `mi-flujo/` y trabaja en la copia.
- Pregunta antes de borrar o mover cualquier archivo.
- No mandes nada a nadie: ni correo, ni Teams, ni ningún otro mensaje.

## Cómo quiero las respuestas

- En español, claro y breve.
- Las cifras sin decimales y con separador de miles; cuando digas una cifra, di de qué archivo salió.
- Al terminar cada paso, dime qué archivos creaste, dónde quedaron y cómo se corre el script.
