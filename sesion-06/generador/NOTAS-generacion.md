# Notas de generación · reporte de inscritos ficticio (flujo de Cecilia, S6)

**Todos los datos de estos archivos son ficticios.** Ningún nombre, matrícula, nómina ni correo corresponde a una persona real. Este archivo lo escribe `generar_reporte_inscritos.py`; las cifras salen del modelo con el que se generaron los datos, no de leer el archivo, y `limpiar_reporte.py` las reproduce por su lado leyendo el reporte.

- Semilla: 42
- Formato: xlsx
- Periodo: Febrero-Junio 2027
- Comando: `python3 generar_reporte_inscritos.py --semilla 42 --salida reportes`

## Las dos entregas

| Archivo | Día | Corte | Renglones en la hoja | Renglones de datos | Peticiones únicas |
|---|---|---|---:|---:|---:|
| `reporte-inscritos-2026-10-19.xlsx` | lunes | 19/10/2026 07:00 | 3,524 | 3,506 | 3,143 |
| `reporte-inscritos-2026-10-21.xlsx` | miércoles | 21/10/2026 07:00 | 4,111 | 4,093 | 3,651 |

La entrega del miércoles es acumulada: trae todas las peticiones del lunes con su estatus al día y suma las que llegaron el 19 y el 20 de octubre. Es el mismo reporte dos veces con fecha distinta, para que el flujo se repita.

## Cómo viene sucio

1. El encabezado está en el renglón 5. Arriba hay tres renglones de título (sistema, periodo, fecha y hora de corte) y uno vacío.
2. Filas de sección: el sistema agrupa por campus (en orden alfabético) con un renglón `CAMPUS: NOMBRE` antes de cada grupo, un renglón `Subtotal Campus: N registros` después y un `Total general` al final. Los subtotales cuentan renglones, no peticiones.
3. Columnas con espacios de más: ` Campus`, `Clave materia `, `Nombre  materia`, `Matrícula alumno `, `Fecha de solicitud `, `Estatus de la petición  `, ` Canal`.
4. `Estatus de la petición` con variantes de mayúsculas y espacios (`NUEVA`, `en proceso`, `En Proceso`, `En  proceso`, `Resuelta ` y así).
5. `Campus` escrito de dos formas en algunos renglones: en MAYÚSCULAS o con un espacio al final (como en los CSV de la S5).
6. Duplicados exactos: renglones repetidos tal cual, uno debajo del otro.
7. Co-titulares: en los grupos con dos profesores, cada petición aparece dos veces, una por profesor; cambia la nómina y el nombre del profesor, lo demás es igual.
8. `Fecha de solicitud` en texto con cuatro formatos: `14/10/2026`, `2026-10-14`, `14-oct-2026` y `14/10/2026 09:32`. Siempre día antes que mes.
9. `Calificación parcial 1` y `Asistencia %` vienen vacías (el periodo no ha empezado). `Estatus inscripción` se deriva del estatus de la petición: Resuelta es Inscrito, Cancelada es Baja, Nueva y En proceso son Pendiente.

## Definición de la tabla limpia

- La unidad es la **petición**: un alumno pidiendo un grupo. Se identifica por `Matrícula alumno` + `CRN` (el reporte no trae folio).
- Pasos: encontrar el encabezado; quedarse con los renglones que tienen matrícula y CRN; quitar duplicados exactos; unir los renglones de co-titular por matrícula + CRN; normalizar `Campus` y `Estatus de la petición` (sin espacios de más, sin distinguir mayúsculas); convertir las fechas.
- La tabla cuenta peticiones únicas por campus y por estatus de la petición (Nueva, En proceso, Resuelta, Cancelada).
- Corte anterior: el lunes o miércoles más reciente antes de la fecha de corte. "Peticiones desde el corte anterior" son las que tienen fecha de solicitud igual o posterior a ese día (el corte es a las 07:00 y el reporte llega hasta el día anterior).

## Respuesta correcta por entrega

### Lunes 19 de octubre de 2026 · `reporte-inscritos-2026-10-19.xlsx`

Control de la limpieza (lo mismo que imprime `limpiar_reporte.py`):

| Concepto | Valor |
|---|---:|
| Renglones en la hoja (título, encabezado, secciones, subtotales y datos) | 3,524 |
| Renglón del encabezado | 5 |
| Renglones que no son datos | 18 |
| Renglones de datos (lo que el sistema llama "registros") | 3,506 |
| Duplicados exactos quitados | 47 |
| Renglones repetidos por co-titular unidos | 316 |
| **Peticiones únicas (alumno + CRN)** | **3,143** |
| Alumnos distintos | 1,061 |
| Campus escritos con variante (corregidos) | 75 |
| Estatus escritos con variante (corregidos) | 813 |
| Fechas de solicitud sin reconocer | 0 |
| Corte anterior | miércoles 14 de octubre de 2026 |
| Peticiones desde el corte anterior | 1,040 |

Fechas de solicitud por formato (renglones de datos): dd/mm/aaaa: 1,735, aaaa-mm-dd: 681, d-mmm-aaaa: 533, dd/mm/aaaa hh:mm: 557.

Tabla limpia (peticiones únicas por campus y estatus de la petición):

| Campus | Nueva | En proceso | Resuelta | Cancelada | Total |
|---|---:|---:|---:|---:|---:|
| Monterrey | 47 | 242 | 661 | 157 | 1,107 |
| Guadalajara | 18 | 120 | 388 | 112 | 638 |
| Ciudad de México | 27 | 133 | 360 | 89 | 609 |
| Querétaro | 17 | 53 | 205 | 37 | 312 |
| Puebla | 5 | 58 | 161 | 44 | 268 |
| Estado de México | 3 | 32 | 144 | 30 | 209 |
| **Total** | **117** | **638** | **1,919** | **469** | **3,143** |

Resumen de tres líneas para Teams (texto exacto que produce el skill):

```
Reporte de inscritos, corte del lunes 19 de octubre de 2026 a las 07:00 (Febrero-Junio 2027): 3,143 peticiones de 1,061 alumnos.
Por estatus: 117 nuevas, 638 en proceso, 1,919 resueltas y 469 canceladas.
Desde el corte del miércoles 14 de octubre llegaron 1,040 peticiones; el campus con más peticiones es Monterrey (1,107). Tabla completa: tabla-inscritos-2026-10-19.xlsx.
```

### Miércoles 21 de octubre de 2026 · `reporte-inscritos-2026-10-21.xlsx`

Control de la limpieza (lo mismo que imprime `limpiar_reporte.py`):

| Concepto | Valor |
|---|---:|
| Renglones en la hoja (título, encabezado, secciones, subtotales y datos) | 4,111 |
| Renglón del encabezado | 5 |
| Renglones que no son datos | 18 |
| Renglones de datos (lo que el sistema llama "registros") | 4,093 |
| Duplicados exactos quitados | 73 |
| Renglones repetidos por co-titular unidos | 369 |
| **Peticiones únicas (alumno + CRN)** | **3,651** |
| Alumnos distintos | 1,243 |
| Campus escritos con variante (corregidos) | 81 |
| Estatus escritos con variante (corregidos) | 886 |
| Fechas de solicitud sin reconocer | 0 |
| Corte anterior | lunes 19 de octubre de 2026 |
| Peticiones desde el corte anterior | 508 |

Fechas de solicitud por formato (renglones de datos): dd/mm/aaaa: 2,026, aaaa-mm-dd: 842, d-mmm-aaaa: 638, dd/mm/aaaa hh:mm: 587.

Tabla limpia (peticiones únicas por campus y estatus de la petición):

| Campus | Nueva | En proceso | Resuelta | Cancelada | Total |
|---|---:|---:|---:|---:|---:|
| Monterrey | 78 | 229 | 796 | 180 | 1,283 |
| Guadalajara | 48 | 116 | 449 | 126 | 739 |
| Ciudad de México | 54 | 121 | 428 | 108 | 711 |
| Querétaro | 22 | 52 | 232 | 48 | 354 |
| Puebla | 33 | 47 | 190 | 46 | 316 |
| Estado de México | 14 | 34 | 165 | 35 | 248 |
| **Total** | **249** | **599** | **2,260** | **543** | **3,651** |

Resumen de tres líneas para Teams (texto exacto que produce el skill):

```
Reporte de inscritos, corte del miércoles 21 de octubre de 2026 a las 07:00 (Febrero-Junio 2027): 3,651 peticiones de 1,243 alumnos.
Por estatus: 249 nuevas, 599 en proceso, 2,260 resueltas y 543 canceladas.
Desde el corte del lunes 19 de octubre llegaron 508 peticiones; el campus con más peticiones es Monterrey (1,283). Tabla completa: tabla-inscritos-2026-10-21.xlsx.
```

## Del lunes al miércoles

- Peticiones nuevas en la entrega del miércoles: 508 (fecha de solicitud 19 y 20 de octubre).
- Peticiones del lunes que cambiaron de estatus: 491.
  - Nueva a En proceso: 96
  - Nueva a Resuelta: 12
  - Nueva a Cancelada: 3
  - En proceso a Resuelta: 325
  - En proceso a Cancelada: 55

## Cómo leer una respuesta equivocada

- Lunes: si el total es 3,506, contó renglones de datos o tomó el `Total general` del sistema (no quitó duplicados ni unió co-titulares); si es 3,459, quitó duplicados exactos pero no unió los renglones de co-titular; la respuesta correcta es 3,143.
- Miércoles: si el total es 4,093, contó renglones de datos o tomó el `Total general` del sistema (no quitó duplicados ni unió co-titulares); si es 4,020, quitó duplicados exactos pero no unió los renglones de co-titular; la respuesta correcta es 3,651.
- Si la tabla trae más de seis campus o más de cuatro estatus, no normalizó mayúsculas ni espacios.
- Si "desde el corte anterior" no coincide, revisar cómo leyó las fechas: leer `05/10/2026` como mes/día (5 de mayo) o descartar los formatos con texto o con hora cambia la cuenta.
- Si contó por la columna `Estatus inscripción` y no por `Estatus de la petición`, verá tres categorías (Pendiente, Inscrito, Baja) en vez de cuatro.
