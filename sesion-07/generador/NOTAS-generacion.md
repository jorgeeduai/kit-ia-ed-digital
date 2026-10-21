# Notas de generación · oferta de grupos ficticia (caso de Karla, S7)

**Todos los datos de estos archivos son ficticios.** Ningún nombre de profesor corresponde a una persona real; las materias, claves, CRN, horarios y salones son inventados. Este archivo lo escribe `generar_versiones.py`: las cifras salen de lo que el generador sembró, no de leer los archivos, y `comparar_versiones.py` las reproduce por su lado leyendo las dos versiones.

- Semilla: 21
- Periodo: Febrero-Junio 2027
- Fecha de las dos versiones: miércoles 21 de octubre de 2026
- Comando (desde la carpeta `sesion-07`): `python3 generador/generar_versiones.py --semilla 21 --salida versiones --notas generador/NOTAS-generacion.md`

## Las dos versiones

| Archivo | Hora | Grupos (renglones de datos) | Orden de los renglones |
|---|---|---:|---|
| `oferta-grupos-2026-10-21-0900.xlsx` | 09:00 | 180 | por CRN |
| `oferta-grupos-2026-10-21-1500.xlsx` | 15:00 | 182 | por campus, materia y grupo |

Encabezado en la fila 1, una fila por CRN y once columnas, en este orden: `CRN`, `Clave de materia`, `Materia`, `Grupo`, `Campus`, `Modalidad`, `Horario`, `Salón`, `Profesor`, `Cupo`, `Estatus del grupo`. La hoja se llama `Oferta Febrero-Junio 2027`. Seis campus (Monterrey, Guadalajara, Ciudad de México, Querétaro, Puebla, Estado de México); modalidad Presencial, Híbrida o En línea (los grupos en línea dicen `Sin salón` en Salón); estatus `Grupo Confirmado` o `Grupo Cancelado`. Los valores vienen limpios en las dos versiones: sin espacios de más y con CRN, Grupo y Cupo como números.

## Lo que sembró el generador

| Qué | Cuántos |
|---|---:|
| Grupos nuevos (CRN solo en la de las 15:00) | 5 |
| Grupos que ya no aparecen (CRN solo en la de las 09:00) | 3 |
| Grupos con cambios | 12 |
| Celdas que cambiaron | 15 |
| · en Profesor | 4 |
| · en Cupo | 4 |
| · en Estatus del grupo | 3 |
| · en Horario | 2 |
| · en Salón | 2 |
| Grupos con dos cambios | 3 |
| Grupos sin cambios (180 − 3 − 12) | 165 |

Además, los renglones de la versión de las 15:00 vienen en otro orden.

**Grupos nuevos:**

| CRN | Clave de materia | Materia | Campus | Profesor |
|---|---|---|---|---|
| 41464 | AD1148 | Gestión de proyectos | Puebla | Leticia Fuentes Torres |
| 41465 | AD1148 | Gestión de proyectos | Monterrey | Fernanda Padilla Reyes |
| 41471 | MA1811 | Cálculo diferencial | Ciudad de México | Manuel Flores Torres |
| 41481 | TC3428 | Programación orientada a objetos | Querétaro | Silvia Trejo Herrera |
| 41491 | ED1191 | Diseño instruccional | Monterrey | Gustavo Ruiz Reyes |

**Grupos que ya no aparecen:**

| CRN | Clave de materia | Materia | Campus | Profesor |
|---|---|---|---|---|
| 41185 | ED1191 | Diseño instruccional | Monterrey | María Cortés González |
| 41320 | AR1758 | Taller de diseño arquitectónico | Puebla | Óscar Valdez Sánchez |
| 41439 | H2553 | Comunicación oral y escrita | Guadalajara | José Mendoza Acosta |

**Celdas que cambiaron** (lo mismo que trae la hoja Cambios de `comparar_versiones.py`):

| CRN | Materia | Columna | Antes | Después |
|---|---|---|---|---|
| 41003 | Emprendimiento e innovación | Profesor | Andrés Rivera Jiménez | Eduardo Álvarez Vázquez |
| 41027 | Taller de diseño arquitectónico | Profesor | Laura Medina Garza | Rodrigo Salinas Peña |
| 41188 | Biología celular | Estatus del grupo | Grupo Confirmado | Grupo Cancelado |
| 41241 | Mercadotecnia digital | Estatus del grupo | Grupo Confirmado | Grupo Cancelado |
| 41243 | Inteligencia artificial aplicada | Cupo | 40 | 50 |
| 41258 | Taller de diseño arquitectónico | Cupo | 45 | 40 |
| 41275 | Salud pública | Horario | Mar y Vie 14:30-16:00 | Sáb 11:00-14:00 |
| 41275 | Salud pública | Salón | A6-320 | A2-320 |
| 41305 | Salud pública | Cupo | 35 | 40 |
| 41324 | Fundamentos de administración | Estatus del grupo | Grupo Confirmado | Grupo Cancelado |
| 41344 | Contabilidad financiera | Horario | Mié 16:00-19:00 | Mar y Vie 16:00-17:30 |
| 41344 | Contabilidad financiera | Salón | A4-306 | A1-319 |
| 41387 | Inteligencia artificial aplicada | Profesor | Roberto Cantú Silva | Francisco Carrillo Flores |
| 41387 | Inteligencia artificial aplicada | Cupo | 60 | 65 |
| 41397 | Bases de datos | Profesor | Emilio Sánchez García | María Cortés González |

Grupos con dos cambios: 41275 (Horario y Salón), 41344 (Horario y Salón) y 41387 (Profesor y Cupo).

**Los tres cambios fijos** (los que muestra el mazo; no dependen de la semilla):

- CRN 41027, Profesor: `Laura Medina Garza` → `Rodrigo Salinas Peña`
- CRN 41188, Estatus del grupo: `Grupo Confirmado` → `Grupo Cancelado`
- CRN 41305, Cupo: `35` → `40`

## Respuesta correcta

Las tres líneas de resumen (texto exacto que deja `comparar_versiones.py` en `resumen-cambios-oferta-grupos-2026-10-21-1500.txt` con `--clave CRN`):

```
Comparación de oferta-grupos-2026-10-21-0900.xlsx contra oferta-grupos-2026-10-21-1500.xlsx, por CRN.
Hay 5 grupos nuevos, 3 grupos que ya no aparecen y 12 grupos con cambios (15 celdas).
Columnas con más cambios: Profesor (4), Cupo (4) y Estatus del grupo (3). Detalle: cambios-oferta-grupos-2026-10-21-1500.xlsx.
```

## Cómo leer una respuesta equivocada

- Si dice que cambiaron casi todos los grupos, comparó por posición (renglón 2 contra renglón 2) y no por CRN: la versión de las 15:00 viene en otro orden. Comparando por posición, 177 de los 180 primeros renglones salen distintos (1,571 celdas).
- Si dice 2 grupos nuevos, restó renglones (182 − 180) en lugar de comparar claves: hay 5 nuevos y 3 que ya no aparecen.
- Si dice más de 15 celdas, contó como cambio algo que no lo es: espacios de más, o `35` contra `35.0`. Estas dos versiones no traen ninguna de las dos cosas, pero un archivo real puede traerlas, y el script de referencia las cuenta como iguales.
