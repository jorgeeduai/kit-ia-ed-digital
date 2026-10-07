# Notas de generación

**Todos los datos de estos archivos son ficticios.** Ningún nombre, matrícula, nómina ni correo corresponde a una persona real.

- Semilla: 42
- Escala: 0.05
- Total de renglones de inscripciones: 50220

## Definiciones

- Inscripciones reales: pares (alumno, grupo) distintos, sin duplicados ni repetición por co-titular. `Alumnos inscritos` en cursos-profesores cuenta estas inscripciones reales de cada grupo (todos los estatus).
- P1: renglones de inscripciones repetidos de forma exacta.
- P2: renglones extra de inscripciones por grupos con co-titular.
- P3: renglones con `Campus` en variante (MAYÚSCULAS o espacio al final).
- P4: en Febrero-Junio 2024, `Asistencia %` viene como texto con signo.
- P5: tres profesores cambian de nómina conservando nombre y correo; la columna cuenta cuántos aparecen con nómina nueva en ese periodo.
- P6: correos de alumno con doble punto o dominio incompleto.

## Tabla por periodo

| Periodo | Renglones cursos | Renglones inscripciones | Inscripciones reales | Profesores distintos | Alumnos distintos | P1 duplicados exactos | P2 renglones por co-titular | P3 campus con variante (inscripciones) | P3 campus con variante (cursos) | P4 asistencia como texto | P5 profesores con nómina nueva | P6 correos de alumno mal formados |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Agosto-Diciembre 2022 | 320 | 5674 | 5245 | 82 | 944 | 83 | 346 | 124 | 6 | 0 | 0 | 20 |
| Febrero-Junio 2023 | 359 | 7447 | 6648 | 85 | 1219 | 105 | 694 | 160 | 5 | 0 | 0 | 31 |
| Agosto-Diciembre 2023 | 320 | 6254 | 5802 | 87 | 1053 | 109 | 343 | 110 | 3 | 0 | 0 | 31 |
| Febrero-Junio 2024 | 322 | 5769 | 5344 | 88 | 965 | 87 | 338 | 103 | 5 | 5769 | 0 | 30 |
| Agosto-Diciembre 2024 | 318 | 5732 | 5324 | 88 | 958 | 81 | 327 | 116 | 2 | 0 | 0 | 44 |
| Febrero-Junio 2025 | 327 | 6270 | 5636 | 83 | 1028 | 73 | 561 | 112 | 12 | 0 | 1 | 32 |
| Agosto-Diciembre 2025 | 351 | 7299 | 6751 | 82 | 1228 | 104 | 444 | 151 | 6 | 0 | 1 | 49 |
| Febrero-Junio 2026 | 333 | 5775 | 5162 | 79 | 942 | 81 | 532 | 105 | 9 | 0 | 1 | 26 |
| Total | 2650 | 50220 | 45912 |  | 3000 | 723 | 3585 | 981 | 48 | 5769 | 3 | 263 |
