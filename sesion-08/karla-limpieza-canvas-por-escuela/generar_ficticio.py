#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera la copia FICTICIA de los seis archivos de actividad de Canvas,
uno por escuela, para la propuesta de limpieza por escuela (curso Ed Digital, propuesta de
skill del grupo, ronda 2).

Ninguno de los encabezados de abajo lo escribio la persona: son SUPUESTOS armados con lo que
describio en la grabacion de la Sesion 7 (cuatro hojas; en la primera, la escuela; una hoja
que no se ocupa; la informacion general de la materia; y los usuarios activos de cada
actividad, con el rol en la columna C y un codigo de curso que empieza con "Nacional" para los
grupos digitales). Estan en las constantes de abajo para cambiarlos en cuanto se confirmen.

Cada archivo trae cuatro hojas:
  1. Escuela        la escuela y el departamento (el nombre de la escuela va junto a "Escuela")
  2. Departamento   una hoja que la limpieza no ocupa (se copia igual)
  3. Materias       informacion general de cada materia, con el SIS ID del curso
  4. Actividades    usuarios activos por actividad: rol, SIS ID del curso, entregas...

Como correrlo (solo biblioteca estandar, Python 3.8 o mas reciente), desde esta carpeta:
    python3 -B generar_ficticio.py
    python3 -B generar_ficticio.py --semilla 22 --salida ficticio

Salen, dentro de --salida (ficticio/ por omision):
    descargas/actividad-canvas-01.xlsx ... actividad-canvas-06.xlsx   las seis descargas
    NOTAS-ficticio.txt                                                lo sembrado y la respuesta

Todo es inventado: escuelas, materias, claves, nominas y personas.
"""
import argparse
import random
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

# ------------------------------------------------------------------ supuestos (cambiar aqui)
HOJA_ESCUELA, HOJA_SIN_USO, HOJA_MATERIAS, HOJA_ACTIVIDADES = (
    "Escuela", "Departamento", "Materias", "Actividades")
ETIQUETA_ESCUELA = "Escuela"
COL_ROL = "Rol"
COL_CODIGO = "SIS ID del curso"
ENC_ACTIVIDADES = ["Usuario", "Nómina", COL_ROL, COL_CODIGO, "Curso", "Actividad",
                   "Entregas", "Calificadas", "Sin calificar", "Días para calificar"]
ENC_MATERIAS = [COL_CODIGO, "Curso", "Tipo de materia", "Modalidad"]

# Escuelas inventadas. La de Faro trae espacios de mas a proposito.
ESCUELAS = ["Escuela Aurora", "Escuela Bosque", "Escuela Cumbre", "Escuela Delta",
            "Escuela Estrella", "  Escuela   Faro "]

# Materias inventadas (clave, nombre)
MATERIAS = [("FIC1001", "Pensamiento crítico aplicado"), ("FIC1010", "Estadística para decidir"),
            ("FIC1020", "Comunicación oral y escrita"), ("FIC2003", "Diseño de experiencias"),
            ("FIC2015", "Ética y ciudadanía"), ("FIC2030", "Modelación matemática"),
            ("FIC3001", "Gestión de proyectos"), ("FIC3012", "Liderazgo colaborativo"),
            ("FIC3020", "Innovación social"), ("FIC4004", "Analítica de datos")]
ACTIVIDADES = ["Foro de bienvenida", "Tarea 1", "Tarea 2", "Quiz de repaso", "Reto parcial",
               "Bitácora semanal", "Entrega de avance", "Autoevaluación"]
NOMBRES = ["Ana", "Luis", "Marta", "Andrés", "Sofía", "Raúl", "Elena", "Iván", "Paula", "Óscar",
           "Lucía", "Tomás", "Irene", "Hugo", "Noemí", "Ramiro"]
APELLIDOS = ["Ficticio", "Inventado", "Prueba", "Ejemplo", "Muestra", "Simulado"]

# Formas del codigo. Digital = empieza con "Nacional" (la regla). Las otras no.
TIPOS_DIGITALES = ["MFD", "MFDN", "MCBL"]


def codigo_nacional(clave, rnd):
    return "Nacional.%s.%s.%02d" % (clave, rnd.choice(TIPOS_DIGITALES), rnd.randint(1, 9))


def codigo_otro(clave, rnd):
    return "%s.%s.%03d" % (rnd.choice(["Regional", "Sede", "Presencial"]), clave,
                           rnd.randint(100, 499))


# ------------------------------------------------------------------ escritura xlsx
def _col_letra(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _hoja_xml(filas):
    partes = []
    for i, fila in enumerate(filas):
        celdas = []
        for j, v in enumerate(fila):
            if v is None or v == "":
                continue
            ref = "%s%d" % (_col_letra(j), i + 1)
            if isinstance(v, int):
                celdas.append('<c r="%s"><v>%d</v></c>' % (ref, v))
            else:
                celdas.append('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                              % (ref, escape(str(v))))
        if celdas:
            partes.append('<row r="%d">%s</row>' % (i + 1, "".join(celdas)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData>%s</sheetData></worksheet>' % "".join(partes))


def escribir_xlsx(ruta, hojas):
    """hojas: lista de (nombre, filas). Un Excel valido sin instalar nada."""
    n = len(hojas)
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    nsr = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    archivos = [
        ("[Content_Types].xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
         '%s</Types>' % "".join(
             '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.'
             'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
             for i in range(n))),
        ("_rels/.rels",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
         '</Relationships>'),
        ("xl/workbook.xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
         % (ns, nsr, "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                             % (escape(h[0]), i + 1, i + 1) for i, h in enumerate(hojas)))),
        ("xl/_rels/workbook.xml.rels",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '%s</Relationships>'
         % "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>'
                   % (i + 1, i + 1) for i in range(n))),
    ]
    for i, (_, filas) in enumerate(hojas):
        archivos.append(("xl/worksheets/sheet%d.xml" % (i + 1), _hoja_xml(filas)))
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in archivos:
            # Fecha fija: misma semilla, mismo archivo byte a byte
            info = zipfile.ZipInfo(nombre, date_time=(2026, 10, 22, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, contenido.encode("utf-8"))


# ------------------------------------------------------------------ generacion
def persona(rnd, k):
    return "%s %s %03d" % (rnd.choice(NOMBRES), rnd.choice(APELLIDOS), k)


def generar_escuela(num, escuela, rnd):
    """Devuelve (hojas, cuenta) de un archivo. cuenta = lo que la limpieza debe dar."""
    cuenta = dict(act_leidos=0, act_quita_rol=0, act_quita_codigo=0, act_quedan=0,
                  quedan_por_rol={}, mat_leidos=0, mat_quitados=0, mat_quedan=0, trampas=[])
    materias = rnd.sample(MATERIAS, 6)
    # Cada materia tiene un curso digital (Nacional) y uno que no lo es
    cursos = []
    for clave, nombre in materias:
        cursos.append((codigo_nacional(clave, rnd), nombre, True))
        cursos.append((codigo_otro(clave, rnd), nombre, False))

    # Hoja Materias
    mat = [list(ENC_MATERIAS)]
    for cod, nombre, digital in cursos:
        mat.append([cod, nombre, rnd.choice(["Disciplinar", "Transversal", "Optativa"]),
                    "Digital" if digital else rnd.choice(["Presencial", "Híbrida"])])
        cuenta["mat_leidos"] += 1
        if digital:
            cuenta["mat_quedan"] += 1
        else:
            cuenta["mat_quitados"] += 1
    # Trampa en Materias: un curso "Internacional" (contiene "nacional", no empieza con eso)
    clave, nombre = materias[0]
    mat.append(["Internacional.%s.%03d" % (clave, rnd.randint(100, 499)), nombre, "Optativa",
                "Presencial"])
    cuenta["mat_leidos"] += 1
    cuenta["mat_quitados"] += 1

    # Hoja Actividades
    act = [list(ENC_ACTIVIDADES)]
    k = 0

    def renglon(rol, cod, nombre_curso):
        nonlocal k
        k += 1
        entregas = rnd.randint(0, 40)
        calif = rnd.randint(0, entregas)
        es_prof = rol.strip().casefold() == "profesor"
        return [persona(rnd, k), "F%07d" % rnd.randint(1, 9999999), rol, cod, nombre_curso,
                rnd.choice(ACTIVIDADES), entregas if es_prof else "",
                calif if es_prof else "", (entregas - calif) if es_prof else "",
                rnd.randint(0, 12) if es_prof else ""]

    def anota(rol_limpio, digital):
        cuenta["act_leidos"] += 1
        if rol_limpio == "estudiante":
            cuenta["act_quita_rol"] += 1
        elif not digital:
            cuenta["act_quita_codigo"] += 1
        else:
            cuenta["act_quedan"] += 1
            r = rol_limpio.capitalize()
            cuenta["quedan_por_rol"][r] = cuenta["quedan_por_rol"].get(r, 0) + 1

    n = rnd.randint(70, 110)
    for _ in range(n):
        cod, nombre_curso, digital = rnd.choice(cursos)
        rol = rnd.choices(["Estudiante", "Profesor", "Observador", "Sistema"],
                          weights=[70, 18, 7, 5])[0]
        act.append(renglon(rol, cod, nombre_curso))
        anota(rol.casefold(), digital)

    # Trampas, distintas en cada escuela pero las mismas cinco formas
    clave, nombre = materias[1]
    trampas = [
        ("  estudiante ", "Nacional.%s.MFD.01" % clave, True,
         "rol en minúsculas y con espacios: se quita por rol"),
        ("ESTUDIANTE", "Nacional.%s.MFDN.02" % clave, True,
         "rol en mayúsculas: se quita por rol"),
        ("Profesor", "Internacional.%s.210" % clave, False,
         "código Internacional (contiene nacional, no empieza con eso): se quita por código"),
        ("Profesor", "  nacional.%s.MCBL.03" % clave, True,
         "código en minúsculas y con espacios al inicio: se queda"),
        ("Profesor", "NACIONAL.%s.MFD.04" % clave, True,
         "código en mayúsculas: se queda"),
    ]
    for rol, cod, digital, que in trampas:
        pos = rnd.randint(1, len(act))
        act.insert(pos, renglon(rol, cod, nombre))
        anota(" ".join(rol.split()).casefold(), digital)
        cuenta["trampas"].append(que)
    # Dos renglones vacios en medio, como en una descarga real (no cuentan)
    for _ in range(2):
        act.insert(rnd.randint(2, len(act) - 1), [""] * len(ENC_ACTIVIDADES))

    esc = [[ETIQUETA_ESCUELA, escuela], ["Departamento", "Departamento de prueba %d" % num],
           ["Fecha de descarga", "21/10/2026"], ["Periodo", "Agosto-Diciembre 2026"]]
    sin_uso = [["Departamento", "Coordinación", "Sede"],
               ["Departamento de prueba %d" % num, "Coordinación ficticia", "Sede ficticia"]]
    hojas = [(HOJA_ESCUELA, esc), (HOJA_SIN_USO, sin_uso), (HOJA_MATERIAS, mat),
             (HOJA_ACTIVIDADES, act)]
    return hojas, cuenta


def principal():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia de los seis archivos de "
                                             "actividad de Canvas, uno por escuela.")
    ap.add_argument("--semilla", type=int, default=22, help="fija el azar (22)")
    ap.add_argument("--salida", default="ficticio", help="carpeta de salida (ficticio)")
    args = ap.parse_args()

    rnd = random.Random(args.semilla)
    out = Path(args.salida)
    descargas = out / "descargas"
    descargas.mkdir(parents=True, exist_ok=True)

    lineas = [
        "NOTAS-ficticio.txt: lo que sembró generar_ficticio.py (semilla %d)" % args.semilla,
        "Todo es inventado: escuelas, materias, claves, nóminas y personas.",
        "Los encabezados son SUPUESTOS (ver LEEME.md, Lo que supuse).",
        "",
        "Seis archivos en descargas/, cuatro hojas cada uno: %s, %s (no se ocupa), %s y %s."
        % (HOJA_ESCUELA, HOJA_SIN_USO, HOJA_MATERIAS, HOJA_ACTIVIDADES),
        "Regla: en %s se quitan los renglones con %s = Estudiante; en %s y %s se quedan solo"
        % (HOJA_ACTIVIDADES, COL_ROL, HOJA_ACTIVIDADES, HOJA_MATERIAS),
        "los renglones cuyo %s empieza con \"Nacional\" (sin distinguir mayúsculas)." % COL_CODIGO,
        "",
        "Trampas sembradas en cada archivo (Actividades):",
    ]
    total = dict(act_leidos=0, act_quita_rol=0, act_quita_codigo=0, act_quedan=0,
                 mat_leidos=0, mat_quitados=0, mat_quedan=0)
    por_rol_total = {}
    detalle = []
    trampas = None
    for i, escuela in enumerate(ESCUELAS, start=1):
        hojas, c = generar_escuela(i, escuela, rnd)
        nombre_archivo = "actividad-canvas-%02d.xlsx" % i
        escribir_xlsx(descargas / nombre_archivo, hojas)
        trampas = c["trampas"]
        for k in total:
            total[k] += c[k]
        for r, v in c["quedan_por_rol"].items():
            por_rol_total[r] = por_rol_total.get(r, 0) + v
        limpio = " ".join(escuela.split())
        detalle.append("  %s -> %s.xlsx | Actividades: %d leídos, %d fuera por rol, %d fuera por "
                       "código, %d quedan (%s) | Materias: %d leídas, %d quedan"
                       % (nombre_archivo, limpio, c["act_leidos"], c["act_quita_rol"],
                          c["act_quita_codigo"], c["act_quedan"],
                          ", ".join("%s %d" % (r, v) for r, v in sorted(c["quedan_por_rol"].items())),
                          c["mat_leidos"], c["mat_quedan"]))
    lineas += ["  - " + t for t in trampas]
    lineas += [
        "  - dos renglones vacíos en medio de la hoja (no cuentan)",
        "Además, en Materias de cada archivo, un curso Internacional que se debe quitar.",
        "En el archivo 06 el nombre de la escuela trae espacios de más (\"  Escuela   Faro \"):",
        "el archivo limpio se debe llamar \"Escuela Faro.xlsx\".",
        "",
        "Respuesta correcta, archivo por archivo:",
    ]
    lineas += detalle
    lineas += [
        "",
        "Totales (los seis archivos):",
        "  Actividades: %d renglones leídos = %d fuera por rol + %d fuera por código + %d quedan"
        % (total["act_leidos"], total["act_quita_rol"], total["act_quita_codigo"],
           total["act_quedan"]),
        "  Quedan por rol: %s" % ", ".join("%s %d" % (r, v)
                                          for r, v in sorted(por_rol_total.items())),
        "  Materias: %d leídas = %d fuera por código + %d quedan"
        % (total["mat_leidos"], total["mat_quitados"], total["mat_quedan"]),
    ]
    (out / "NOTAS-ficticio.txt").write_text("\n".join(lineas) + "\n", encoding="utf-8")

    print("Listo. Archivos en %s/:" % out)
    print("  descargas/actividad-canvas-01.xlsx a actividad-canvas-06.xlsx (cuatro hojas cada uno)")
    print("  NOTAS-ficticio.txt")
    print()
    print("Sembrado: %d renglones en Actividades; deben quedar %d (fuera por rol %d, fuera por "
          "código %d), con 5 trampas por archivo." % (total["act_leidos"], total["act_quedan"],
                                                      total["act_quita_rol"],
                                                      total["act_quita_codigo"]))
    return 0


if __name__ == "__main__":
    sys.exit(principal())
