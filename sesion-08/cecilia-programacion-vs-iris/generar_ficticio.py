#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera la copia FICTICIA de los dos archivos del flujo de Cecilia
(curso Ed Digital, borrador del grupo despues de la Sesion 7): el archivo de programacion y el
reporte de grupos de IRIS del mismo periodo. Entre uno y otro siembra discrepancias conocidas
y las anota en NOTAS-ficticio.txt, con la respuesta correcta.

Como correrlo (solo biblioteca estandar, Python 3.8 o mas reciente), desde la carpeta
cecilia-programacion-vs-iris:
    python3 -B generar_ficticio.py --semilla 22 --salida ficticio

Parametros:
    --semilla  numero que fija el azar (22). Misma semilla, mismos archivos, byte a byte.
    --salida   carpeta donde se escriben los dos archivos y NOTAS-ficticio.txt (ficticio).

Salen:
    programacion-ficticia.xlsx   200 grupos (CRN), algunos con dos profesores; encabezado en la
                                 fila 1; fechas como texto 08/02/2027, horas 07:30, % como 50%
    reporte-iris-ficticio.xlsx   el mismo periodo visto desde IRIS: dos renglones de titulo
                                 arriba del encabezado, columnas en otro orden, una columna mas
                                 (Estatus), renglones en otro orden, fechas como fecha de Excel,
                                 horas como 0730, % como numero
    NOTAS-ficticio.txt           lo que se sembro, CRN por CRN, y las tres lineas que debe dar
                                 validar_programacion.py

Lo sembrado es fijo; el azar solo elige que grupos y que valores:
    3 CRN solo en programacion, 2 CRN solo en IRIS, 2 CRN con otro profesor asignado y
    12 CRN con datos distintos en 13 celdas. Las diferencias de formato (fechas, horas, %,
    espacios de mas) estan en todos los renglones y NO deben contar.
Los encabezados son los que Cecilia nombro en el muro; los datos son inventados: ningun nombre
ni nomina corresponde a una persona real.
"""
import argparse
import datetime
import random
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

COLUMNAS = ["CRN", "Clave", "Nombre de la materia", "Grupo", "PMT", "Día inicio", "Día fin",
            "Horario inicio", "Horario fin", "Fecha de inicio de periodo",
            "Fecha de fin de periodo", "Nómina del profesor", "Nombre del profesor",
            "% de responsabilidad asignada al grupo", "Atributos del curso"]
ANCHOS_PROG = [8, 10, 36, 7, 6, 12, 12, 10, 10, 14, 14, 14, 30, 12, 30]
# IRIS: mismas columnas en otro orden, mas Estatus
COLUMNAS_IRIS = ["CRN", "Nómina del profesor", "Nombre del profesor",
                 "% de responsabilidad asignada al grupo", "Clave", "Nombre de la materia",
                 "Grupo", "PMT", "Fecha de inicio de periodo", "Fecha de fin de periodo",
                 "Día inicio", "Día fin", "Horario inicio", "Horario fin", "Atributos del curso",
                 "Estatus"]
TITULO_IRIS = ["Reporte de grupos IRIS (copia ficticia)", "Periodo Febrero-Junio 2027"]

N_GRUPOS = 200
P_DOS_PROFESORES = 0.12
PERIODO = (datetime.date(2027, 2, 8), datetime.date(2027, 6, 11))
# PMT: parte del periodo, con sus fechas de clase (primer y ultimo dia)
PMT = {"1": (datetime.date(2027, 2, 8), datetime.date(2027, 6, 11)),
       "2": (datetime.date(2027, 2, 8), datetime.date(2027, 4, 2)),
       "3": (datetime.date(2027, 4, 5), datetime.date(2027, 6, 11))}
PMT_PESOS = [0.6, 0.2, 0.2]
INICIOS = [420, 510, 600, 690, 780, 870, 960, 1050, 1140]   # minutos desde las 0:00
DURACIONES = [90, 90, 90, 180]
ATRIBUTOS = ["", "", "Impartido en inglés", "Semana Tec", "Curso de exploración",
             "Impartido en inglés; Semana Tec"]
REPARTOS = [(50, 50), (60, 40), (70, 30)]

# Discrepancias sembradas en IRIS: (columna, cuantos CRN). Una celda por CRN, salvo el %.
SIEMBRA_CELDAS = [("Nombre de la materia", 1), ("Grupo", 1), ("PMT", 1), ("Día fin", 1),
                  ("Horario inicio", 2), ("Horario fin", 1), ("Fecha de fin de periodo", 1),
                  ("Atributos del curso", 2), ("Nombre del profesor", 1)]
N_SOLO_PROG, N_SOLO_IRIS, N_PROF_ASIGNADO = 3, 2, 2
N_REPARTO = 1      # grupos con dos profesores a los que IRIS les cambia el % (2 celdas cada uno)

MATERIAS = [
    ("TC", "Pensamiento computacional"), ("TC", "Programación orientada a objetos"),
    ("TC", "Bases de datos"), ("MA", "Cálculo diferencial"), ("MA", "Álgebra lineal"),
    ("MA", "Probabilidad y estadística"), ("F", "Física I"), ("Q", "Química general"),
    ("AD", "Fundamentos de administración"), ("AD", "Gestión de proyectos"),
    ("CF", "Contabilidad financiera"), ("EC", "Microeconomía"), ("MK", "Mercadotecnia digital"),
    ("H", "Ética, persona y sociedad"), ("H", "Comunicación oral y escrita"),
    ("LE", "Inglés para negocios"), ("DI", "Diseño de experiencias digitales"),
    ("MD", "Ciencia de datos aplicada"), ("ED", "Diseño instruccional"),
    ("PS", "Psicología social"),
]
NOMBRES = ["José", "Luis", "Miguel", "Ricardo", "Fernando", "Javier", "Sergio", "Héctor",
           "Andrés", "Gerardo", "Arturo", "Emilio", "María", "Ana", "Patricia", "Gabriela",
           "Verónica", "Sofía", "Daniela", "Paola", "Lucía", "Elena", "Mónica", "Julieta"]
APELLIDOS = ["Hernández", "García", "López", "Ramírez", "Flores", "Gómez", "Díaz", "Reyes",
             "Morales", "Ortiz", "Gutiérrez", "Ruiz", "Jiménez", "Mendoza", "Álvarez",
             "Castillo", "Vázquez", "Romero", "Herrera", "Aguilar", "Vargas", "Navarro",
             "Cortés", "Soto", "Luna", "Ibarra", "Vega", "Fuentes", "Núñez", "Treviño"]


def miles(n):
    return "{:,}".format(n)


def y_lista(partes):
    partes = list(partes)
    if len(partes) <= 1:
        return "".join(partes)
    return ", ".join(partes[:-1]) + " y " + partes[-1]


# ------------------------------------------------------------------ xlsx (stdlib)
# Las funciones de escritura de generar_versiones.py (Sesion 7), con fecha fija dentro
# del zip para que la misma semilla dé los mismos archivos, byte a byte.
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NORMAL, NEGRITA, FECHA = 0, 1, 2


def _col_letra(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _hoja_xml(filas, anchos, congelar):
    partes = []
    for i, fila in enumerate(filas):
        celdas = []
        for j, celda in enumerate(fila):
            v, estilo = celda if isinstance(celda, tuple) else (celda, NORMAL)
            if v is None or v == "":
                continue
            ref = "%s%d" % (_col_letra(j), i + 1)
            s = ' s="%d"' % estilo if estilo else ""
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                celdas.append('<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(v)))
            else:
                celdas.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t>'
                              '</is></c>' % (ref, s, escape(str(v))))
        if celdas:
            partes.append('<row r="%d">%s</row>' % (i + 1, "".join(celdas)))
    vista = ('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
             'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
             if congelar else '<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (j + 1, j + 1, w)
                   for j, w in enumerate(anchos))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s" xmlns:r="%s">%s%s<sheetData>%s</sheetData></worksheet>'
            % (NS_MAIN, NS_REL, vista, "<cols>%s</cols>" % cols if cols else "",
               "".join(partes)))


def escribir_xlsx(ruta, hojas):
    """hojas: lista de (nombre, filas, anchos, congelar_encabezado)."""
    n = len(hojas)
    tipos = "".join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
                    for i in range(n))
    archivos = [
        ("[Content_Types].xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
         '%s<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
         '</Types>' % tipos),
        ("_rels/.rels",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
         '</Relationships>'),
        ("xl/workbook.xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<workbook xmlns="%s" xmlns:r="%s"><sheets>%s</sheets></workbook>'
         % (NS_MAIN, NS_REL, "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
                                     % (escape(h[0]), i + 1, i + 1)
                                     for i, h in enumerate(hojas)))),
        ("xl/_rels/workbook.xml.rels",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '%s<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '</Relationships>'
         % ("".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
                    'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>'
                    % (i + 1, i + 1) for i in range(n)), n + 1)),
        ("xl/styles.xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<styleSheet xmlns="%s">'
         '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
         '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
         '<fills count="2"><fill><patternFill patternType="none"/></fill>'
         '<fill><patternFill patternType="gray125"/></fill></fills>'
         '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
         '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
         '<cellXfs count="3">'
         '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
         '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
         '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
         '</cellXfs>'
         '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
         '</styleSheet>' % NS_MAIN),
    ]
    for i, (_, filas, anchos, congelar) in enumerate(hojas):
        archivos.append(("xl/worksheets/sheet%d.xml" % (i + 1), _hoja_xml(filas, anchos, congelar)))
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in archivos:
            info = zipfile.ZipInfo(nombre, date_time=(2026, 10, 21, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.create_system = 3
            info.external_attr = 0o644 << 16
            z.writestr(info, contenido.encode("utf-8"))


# ------------------------------------------------------------------ catalogos
# ------------------------------------------------------------------ datos
TIPO_COL = {"Nombre de la materia": "Materia y grupo", "Grupo": "Materia y grupo",
            "PMT": "Materia y grupo", "Día fin": "Días y horarios",
            "Horario inicio": "Días y horarios", "Horario fin": "Días y horarios",
            "Fecha de fin de periodo": "Fechas del periodo",
            "Atributos del curso": "Atributos del curso",
            "Nombre del profesor": "Nombre o % del profesor",
            "% de responsabilidad asignada al grupo": "Nombre o % del profesor"}
TIPOS = ["CRN solo en programación", "CRN solo en IRIS", "Profesor asignado", "Materia y grupo",
         "Días y horarios", "Fechas del periodo", "Nombre o % del profesor", "Atributos del curso"]
# Columnas que en IRIS vienen en otro formato en todos los renglones (no deben contar)
FORMATO_GRUPO = ["Día inicio", "Día fin", "Horario inicio", "Horario fin",
                 "Fecha de inicio de periodo", "Fecha de fin de periodo"]


def construir_profesores(rng, n):
    vistos, profs = set(), []
    while len(profs) < n:
        nombre = "%s %s %s" % (rng.choice(NOMBRES), rng.choice(APELLIDOS), rng.choice(APELLIDOS))
        nomina = "L0%07d" % rng.randint(1000000, 9999999)
        if nombre in vistos or nomina in {p[0] for p in profs}:
            continue
        vistos.add(nombre)
        profs.append((nomina, nombre))
    return profs


def nuevo_grupo(rng, crn, profs, usados):
    clave_area, materia = rng.choice(MATERIAS)
    pmt = rng.choices(list(PMT), weights=PMT_PESOS)[0]
    inicio = rng.choice(INICIOS)
    dur = rng.choice(DURACIONES)
    if inicio + dur > 1320:
        inicio = 1140 - dur + 90
    g = dict(crn=crn, clave="%s%04d" % (clave_area, rng.randint(1000, 3099)), materia=materia,
             grupo=rng.randint(1, 12), pmt=pmt, dia_ini=PMT[pmt][0], dia_fin=PMT[pmt][1],
             h_ini=inicio, h_fin=inicio + dur, f_ini=PERIODO[0], f_fin=PERIODO[1],
             atributos=rng.choice(ATRIBUTOS), profesores=[])
    if rng.random() < P_DOS_PROFESORES:
        a, b = rng.choice(REPARTOS)
        p1, p2 = rng.sample(profs, 2)
        g["profesores"] = [[p1[0], p1[1], a], [p2[0], p2[1], b]]
    else:
        p = rng.choice(profs)
        g["profesores"] = [[p[0], p[1], 100]]
    return g


def copia(g):
    h = dict(g)
    h["profesores"] = [list(p) for p in g["profesores"]]
    return h


def hhmm(m):
    return "%02d:%02d" % (m // 60, m % 60)


def construir(rng):
    profs = construir_profesores(rng, 120)
    crns = sorted(rng.sample(range(50001, 59999), N_GRUPOS + N_SOLO_IRIS))
    prog = [nuevo_grupo(rng, c, profs, None) for c in crns[:N_GRUPOS]]
    extras = [nuevo_grupo(rng, c, profs, None) for c in crns[N_GRUPOS:]]
    iris = {g["crn"]: copia(g) for g in prog}

    uno = [g["crn"] for g in prog if len(g["profesores"]) == 1]
    dos = [g["crn"] for g in prog if len(g["profesores"]) == 2]
    rng.shuffle(uno)
    rng.shuffle(dos)
    siembra = []        # (tipo, crn, columna, nomina, valor en programacion, valor en IRIS)
    por_crn_prog = {g["crn"]: g for g in prog}

    for _ in range(N_SOLO_PROG):
        c = uno.pop()
        del iris[c]
        siembra.append(("CRN solo en programación", c, "(todo el grupo)", "", "está", "no está"))
    for g in extras:
        iris[g["crn"]] = copia(g)
        siembra.append(("CRN solo en IRIS", g["crn"], "(todo el grupo)", "", "no está", "está"))
    for _ in range(N_PROF_ASIGNADO):
        c = uno.pop()
        antes = iris[c]["profesores"][0]
        otro = rng.choice([p for p in profs if p[0] != antes[0]])
        iris[c]["profesores"][0] = [otro[0], otro[1], 100]
        siembra.append(("Profesor asignado", c, "Nómina del profesor", "",
                        "%s (%s)" % (antes[0], antes[1]), "%s (%s)" % (otro[0], otro[1])))
    for col, n in SIEMBRA_CELDAS:
        for _ in range(n):
            c = uno.pop()
            g, h = por_crn_prog[c], iris[c]
            if col == "Nombre de la materia":
                h["materia"] = rng.choice([m for _, m in MATERIAS if m != g["materia"]])
                v = (g["materia"], h["materia"])
            elif col == "Grupo":
                h["grupo"] = g["grupo"] + 1
                v = (g["grupo"], h["grupo"])
            elif col == "PMT":
                h["pmt"] = rng.choice([p for p in PMT if p != g["pmt"]])
                v = (g["pmt"], h["pmt"])
            elif col == "Día fin":
                h["dia_fin"] = g["dia_fin"] - datetime.timedelta(days=7)
                v = (fecha_texto(g["dia_fin"]), fecha_texto(h["dia_fin"]))
            elif col == "Horario inicio":
                h["h_ini"] = g["h_ini"] - 30
                v = (hhmm(g["h_ini"]), hhmm(h["h_ini"]))
            elif col == "Horario fin":
                h["h_fin"] = g["h_fin"] + 30
                v = (hhmm(g["h_fin"]), hhmm(h["h_fin"]))
            elif col == "Fecha de fin de periodo":
                h["f_fin"] = g["f_fin"] + datetime.timedelta(days=7)
                v = (fecha_texto(g["f_fin"]), fecha_texto(h["f_fin"]))
            elif col == "Atributos del curso":
                h["atributos"] = rng.choice([a for a in dict.fromkeys(ATRIBUTOS) if a != g["atributos"]])
                v = (g["atributos"] or "(vacío)", h["atributos"] or "(vacío)")
            else:   # Nombre del profesor: IRIS trae otro segundo apellido
                partes = g["profesores"][0][1].split(" ")
                partes[-1] = rng.choice([a for a in APELLIDOS if a != partes[-1]])
                h["profesores"][0][1] = " ".join(partes)
                v = (g["profesores"][0][1], h["profesores"][0][1])
            nom = g["profesores"][0][0] if col == "Nombre del profesor" else ""
            siembra.append((TIPO_COL[col], c, col, nom, str(v[0]), str(v[1])))
    for _ in range(N_REPARTO):
        c = dos.pop()
        g, h = por_crn_prog[c], iris[c]
        nuevo = rng.choice([r for r in REPARTOS if r[0] != g["profesores"][0][2]])
        for i in range(2):
            h["profesores"][i][2] = nuevo[i]
            siembra.append((TIPO_COL["% de responsabilidad asignada al grupo"], c,
                            "% de responsabilidad asignada al grupo", g["profesores"][i][0],
                            "%d%%" % g["profesores"][i][2], str(nuevo[i])))

    # Cuantas celdas difieren solo en formato (las debe contar el script aparte)
    sembradas = {(s[1], s[2], s[3]) for s in siembra}
    formato = 0
    for g in prog:
        c = g["crn"]
        if c not in iris:
            continue
        formato += sum(1 for col in FORMATO_GRUPO if (c, col, "") not in sembradas)
        noms_i = {p[0] for p in iris[c]["profesores"]}
        for p in g["profesores"]:
            if p[0] in noms_i and (c, "% de responsabilidad asignada al grupo", p[0]) \
                    not in sembradas:
                formato += 1
    return prog, list(iris.values()), siembra, formato


# ------------------------------------------------------------------ hojas
def fecha_texto(d):
    return d.strftime("%d/%m/%Y")


def serial(d):
    return (d - datetime.date(1899, 12, 30)).days


def filas_programacion(prog):
    filas = [[(c, NEGRITA) for c in COLUMNAS]]
    for g in sorted(prog, key=lambda g: g["crn"]):
        for nom, nombre, pct in g["profesores"]:
            filas.append([g["crn"], g["clave"], g["materia"], g["grupo"], g["pmt"],
                          fecha_texto(g["dia_ini"]), fecha_texto(g["dia_fin"]), hhmm(g["h_ini"]),
                          hhmm(g["h_fin"]), fecha_texto(g["f_ini"]), fecha_texto(g["f_fin"]),
                          nom, nombre, "%d%%" % pct, g["atributos"]])
    return filas


def filas_iris(iris, rng):
    filas = [[(TITULO_IRIS[0], NEGRITA)], [TITULO_IRIS[1]],
             [(c, NEGRITA) for c in COLUMNAS_IRIS]]
    for g in sorted(iris, key=lambda g: (g["materia"], g["grupo"], g["crn"])):
        for nom, nombre, pct in g["profesores"]:
            # Ruido de formato: algunos nombres con doble espacio
            if rng.random() < 0.15:
                nombre = nombre.replace(" ", "  ", 1)
            filas.append([g["crn"], nom, nombre, pct, g["clave"], g["materia"], g["grupo"],
                          int(g["pmt"]), (serial(g["f_ini"]), FECHA), (serial(g["f_fin"]), FECHA),
                          (serial(g["dia_ini"]), FECHA), (serial(g["dia_fin"]), FECHA),
                          hhmm(g["h_ini"]).replace(":", ""), hhmm(g["h_fin"]).replace(":", ""),
                          g["atributos"], "Activo"])
    return filas


def esperado(siembra, n_prog, nombre_prog, nombre_iris):
    por_tipo = {}
    for s in siembra:
        por_tipo[s[0]] = por_tipo.get(s[0], 0) + 1
    crn_inc = {s[1] for s in siembra}
    fijos = ("CRN solo en programación", "CRN solo en IRIS", "Profesor asignado")
    con_datos = {s[1] for s in siembra if s[0] not in fijos}
    celdas = sum(1 for s in siembra if s[0] not in fijos)
    top = sorted(por_tipo, key=lambda t: (-por_tipo[t], TIPOS.index(t)))[:3]
    l1 = "Validación de %s contra %s, por CRN y Nómina del profesor." % (nombre_prog, nombre_iris)
    l2 = ("Hay %s CRN con incidencia: %s solo en programación, %s solo en IRIS, %s con otro "
          "profesor asignado y %s con datos distintos (%s celdas)." % (
              miles(len(crn_inc)), por_tipo.get(fijos[0], 0), por_tipo.get(fijos[1], 0),
              por_tipo.get(fijos[2], 0), miles(len(con_datos)), miles(celdas)))
    l3 = "Tipos con más discrepancias: %s. Detalle: discrepancias-%s.xlsx." % (
        y_lista("%s (%d)" % (t, por_tipo[t]) for t in top), nombre_prog.rsplit(".", 1)[0])
    return por_tipo, crn_inc, con_datos, celdas, [l1, l2, l3]


def escribir_notas(ruta, semilla, prog, iris, siembra, formato, nombres):
    nombre_prog, nombre_iris = nombres
    por_tipo, crn_inc, con_datos, celdas, lineas = esperado(siembra, len(prog), *nombres)
    renglones_p = sum(len(g["profesores"]) for g in prog)
    renglones_i = sum(len(g["profesores"]) for g in iris)
    dos = sum(1 for g in prog if len(g["profesores"]) == 2)
    t = []
    t.append("NOTAS DE LA COPIA FICTICIA · programación contra IRIS (flujo de Cecilia)")
    t.append("Todos los datos son inventados por generar_ficticio.py. Ningún nombre ni nómina")
    t.append("corresponde a una persona real. Las cifras salen de lo que se sembró, no de leer")
    t.append("los archivos; validar_programacion.py las debe reproducir por su lado.")
    t.append("")
    t.append("Semilla: %d" % semilla)
    t.append("Comando: python3 -B generar_ficticio.py --semilla %d --salida ficticio" % semilla)
    t.append("")
    t.append("ARCHIVOS")
    t.append("  %s: %s CRN en %s renglones (%s CRN con dos profesores). Encabezado en la fila 1."
             % (nombre_prog, miles(len(prog)), miles(renglones_p), dos))
    t.append("  %s: %s CRN en %s renglones. Dos renglones de título; encabezado en la fila 3;"
             % (nombre_iris, miles(len(iris)), miles(renglones_i)))
    t.append("    columnas en otro orden y una más (Estatus); renglones por materia y grupo.")
    t.append("")
    t.append("LO QUE SE SEMBRÓ (en IRIS)")
    for tipo in TIPOS:
        if tipo in por_tipo:
            t.append("  %s: %d" % (tipo, por_tipo[tipo]))
    t.append("  CRN con incidencia: %d (%d con datos distintos en %d celdas)"
             % (len(crn_inc), len(con_datos), celdas))
    t.append("")
    t.append("DETALLE, CRN POR CRN")
    for s in sorted(siembra, key=lambda s: (s[1], TIPOS.index(s[0]))):
        nom = " · nómina %s" % s[3] if s[3] else ""
        t.append("  CRN %d · %s · %s%s: programación '%s', IRIS '%s'"
                 % (s[1], s[0], s[2], nom, s[4], s[5]))
    t.append("")
    t.append("LO QUE NO DEBE CONTAR (solo formato)")
    t.append("  Fechas: programación 08/02/2027 (texto); IRIS fecha de Excel.")
    t.append("  Horas: programación 07:30; IRIS 0730.")
    t.append("  Porcentaje: programación 50%; IRIS 50.")
    t.append("  Nombres del profesor con doble espacio en IRIS (unos cuantos renglones).")
    t.append("  PMT: programación como texto; IRIS como número.")
    t.append("  Celdas que solo cambian de formato: %s (el script las cuenta aparte)."
             % miles(formato))
    t.append("")
    t.append("LAS TRES LÍNEAS QUE DEBE DAR validar_programacion.py")
    for l in lineas:
        t.append("  " + l)
    ruta.write_text("\n".join(t) + "\n", encoding="utf-8")
    return lineas


def main():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia de programación e IRIS.")
    ap.add_argument("--semilla", type=int, default=22)
    ap.add_argument("--salida", default="ficticio")
    args = ap.parse_args()
    rng = random.Random(args.semilla)
    prog, iris, siembra, formato = construir(rng)
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)
    nombre_prog, nombre_iris = "programacion-ficticia.xlsx", "reporte-iris-ficticio.xlsx"
    escribir_xlsx(out / nombre_prog, [("Programación", filas_programacion(prog), ANCHOS_PROG,
                                       True)])
    escribir_xlsx(out / nombre_iris, [("Grupos IRIS", filas_iris(iris, rng),
                                       [8, 14, 30, 10, 10, 36, 7, 6, 14, 14, 12, 12, 10, 10, 30,
                                        10], False)])
    lineas = escribir_notas(out / "NOTAS-ficticio.txt", args.semilla, prog, iris, siembra,
                            formato, (nombre_prog, nombre_iris))
    print("Listo. En %s/:" % out)
    print("  %s, %s y NOTAS-ficticio.txt" % (nombre_prog, nombre_iris))
    print("Sembrado: %s." % ", ".join("%s %d" % (t, n) for t, n in
                                      sorted(esperado(siembra, 0, nombre_prog, nombre_iris)[0]
                                             .items(), key=lambda x: TIPOS.index(x[0]))))
    print("Celdas que solo cambian de formato (no deben contar): %s." % miles(formato))
    print("Lo que debe salir:")
    for l in lineas:
        print("  " + l)
    return 0


if __name__ == "__main__":
    sys.exit(main())
