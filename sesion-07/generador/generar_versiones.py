#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_versiones.py: genera las dos versiones FICTICIAS de la oferta de grupos del caso de
Karla (curso Ed Digital, Sesion 7). Imita el archivo de oferta que llega varias veces al dia:
la version de las 9:00 y la de las 15:00 del miercoles 21 de octubre de 2026, periodo
Febrero-Junio 2027. Entre una y otra hay grupos nuevos, grupos que ya no aparecen, celdas
que cambiaron y los renglones vienen en otro orden.

Como correrlo (solo biblioteca estandar, Python 3.8 o mas reciente), desde la carpeta sesion-07:
    python3 generador/generar_versiones.py --semilla 21 --salida versiones --notas generador/NOTAS-generacion.md

Parametros:
    --semilla  numero que fija el azar (21). Misma semilla, mismos archivos, byte a byte.
    --salida   carpeta donde se escriben las dos versiones (versiones).
    --notas    ruta de NOTAS-generacion.md (por omision, dentro de --salida).

Salen:
    oferta-grupos-2026-10-21-0900.xlsx   180 grupos, ordenados por CRN
    oferta-grupos-2026-10-21-1500.xlsx   182 grupos, ordenados por campus, materia y grupo
    NOTAS-generacion.md                  lo que se sembro y la respuesta correcta

Lo sembrado es fijo; el azar solo elige que grupos y que valores:
    5 grupos nuevos, 3 que ya no aparecen y 12 grupos con cambios en 15 celdas
    (Profesor 4, Cupo 4, Estatus del grupo 3, Horario 2, Salon 2), mas tres cambios fijos
    que muestra el mazo (CRN 41027, 41188 y 41305).
Todos los datos son inventados; ningun nombre de profesor corresponde a una persona real.
"""
import argparse
import random
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

PERIODO = "Febrero-Junio 2027"
FECHA = "2026-10-21"
VERSIONES = [("0900", "09:00"), ("1500", "15:00")]
COLUMNAS = ["CRN", "Clave de materia", "Materia", "Grupo", "Campus", "Modalidad", "Horario",
            "Salón", "Profesor", "Cupo", "Estatus del grupo"]
ANCHOS = [8, 16, 38, 7, 18, 11, 24, 11, 28, 6, 18]

CAMPUS = ["Monterrey", "Guadalajara", "Ciudad de México", "Querétaro", "Puebla",
          "Estado de México"]
CAMPUS_PESOS = [0.32, 0.20, 0.20, 0.10, 0.10, 0.08]
MODALIDADES = ["Presencial", "Híbrida", "En línea"]
MODALIDAD_PESOS = [0.60, 0.25, 0.15]
CUPOS = {"Presencial": [25, 30, 35, 40, 45], "Híbrida": [30, 35, 40],
         "En línea": [40, 50, 60]}
SIN_SALON = "Sin salón"
CONFIRMADO, CANCELADO = "Grupo Confirmado", "Grupo Cancelado"
P_CANCELADO = 0.08

N_ANTERIOR = 180          # grupos en la version de las 9:00
N_NUEVOS = 5              # CRN solo en la de las 15:00
N_YA_NO = 3               # CRN solo en la de las 9:00
N_PROFESORES = 90

# Los tres cambios fijos que muestra el mazo: (CRN, columna, antes, despues)
FIJOS = [(41027, "Profesor", "Laura Medina Garza", "Rodrigo Salinas Peña"),
         (41188, "Estatus del grupo", CONFIRMADO, CANCELADO),
         (41305, "Cupo", 35, 40)]
# Los otros nueve grupos con cambios, por tipo (cuantos de cada uno)
OTROS_CAMBIOS = [(("Estatus del grupo",), 2), (("Horario", "Salón"), 2),
                 (("Profesor", "Cupo"), 1), (("Profesor",), 2), (("Cupo",), 2)]

MATERIAS = [
    ("TC", "Pensamiento computacional"), ("TC", "Programación orientada a objetos"),
    ("TC", "Bases de datos"), ("TC", "Análisis y diseño de algoritmos"),
    ("MA", "Cálculo diferencial"), ("MA", "Cálculo integral"), ("MA", "Álgebra lineal"),
    ("MA", "Probabilidad y estadística"), ("F", "Física I"), ("Q", "Química general"),
    ("BI", "Biología celular"), ("AD", "Fundamentos de administración"),
    ("AD", "Gestión de proyectos"), ("CF", "Contabilidad financiera"),
    ("FZ", "Finanzas corporativas"), ("EC", "Microeconomía"), ("MK", "Mercadotecnia digital"),
    ("EM", "Emprendimiento e innovación"), ("H", "Ética, persona y sociedad"),
    ("H", "Comunicación oral y escrita"), ("H", "Pensamiento crítico"),
    ("LE", "Inglés para negocios"), ("DI", "Diseño de experiencias digitales"),
    ("AR", "Taller de diseño arquitectónico"), ("DE", "Derecho constitucional"),
    ("PS", "Psicología social"), ("MD", "Ciencia de datos aplicada"),
    ("MD", "Inteligencia artificial aplicada"), ("SA", "Salud pública"),
    ("ED", "Diseño instruccional"),
]
NOMBRES = [
    "José", "Juan", "Luis", "Carlos", "Miguel", "Pedro", "Ricardo", "Fernando", "Roberto",
    "Alejandro", "Daniel", "Eduardo", "Javier", "Manuel", "Francisco", "Sergio", "Raúl",
    "Héctor", "Óscar", "Andrés", "Diego", "Gerardo", "Mario", "Arturo", "Emilio", "Gustavo",
    "María", "Guadalupe", "Ana", "Patricia", "Claudia", "Gabriela", "Verónica", "Sofía",
    "Fernanda", "Alejandra", "Daniela", "Mariana", "Paola", "Lucía", "Elena", "Rosa",
    "Teresa", "Silvia", "Adriana", "Leticia", "Mónica", "Beatriz", "Irene", "Julieta",
]
APELLIDOS = [
    "Hernández", "García", "Martínez", "López", "González", "Rodríguez", "Pérez", "Sánchez",
    "Ramírez", "Flores", "Gómez", "Díaz", "Reyes", "Morales", "Cruz", "Ortiz", "Gutiérrez",
    "Chávez", "Ruiz", "Jiménez", "Mendoza", "Álvarez", "Castillo", "Vázquez", "Romero",
    "Torres", "Herrera", "Aguilar", "Vargas", "Castro", "Ramos", "Moreno", "Guerrero",
    "Domínguez", "Rojas", "Silva", "Navarro", "Cortés", "Soto", "Contreras", "Luna",
    "Estrada", "Delgado", "Ibarra", "Espinoza", "Vega", "Campos", "Cabrera", "Fuentes",
    "Rivera", "Mejía", "Carrillo", "Sandoval", "Núñez", "Zamora", "Trejo", "Montes",
    "Ávila", "Cervantes", "Ríos", "Lara", "Valdez", "Padilla", "Acosta", "Orozco",
    "Villarreal", "Treviño", "Elizondo", "Cantú", "Solís",
]
# Patrones de horario: dias, duracion en minutos, horas de inicio posibles
PATRONES = [("Lun y Jue", 90, [420, 510, 600, 690, 780, 870, 960, 1050, 1140]),
            ("Mar y Vie", 90, [420, 510, 600, 690, 780, 870, 960, 1050, 1140]),
            ("Lun, Mié y Vie", 60, [420, 480, 540, 600, 660, 720, 780, 840, 960, 1020]),
            ("Mié", 180, [420, 600, 960]),
            ("Sáb", 180, [480, 660])]
PATRON_PESOS = [0.32, 0.32, 0.20, 0.10, 0.06]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def fold(s):
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s))
                if not unicodedata.combining(c))
    return " ".join(s.casefold().split())


def miles(n):
    return "{:,}".format(n)


def y_lista(partes):
    partes = list(partes)
    if len(partes) <= 1:
        return "".join(partes)
    return ", ".join(partes[:-1]) + " y " + partes[-1]


# ------------------------------------------------------------------ xlsx (stdlib)
# Las mismas funciones de escritura de limpiar_reporte.py (Sesion 6), con fecha fija dentro
# del zip para que la misma semilla dé los mismos archivos, byte a byte.
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NORMAL, NEGRITA = 0, 1


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
         '<cellXfs count="2">'
         '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
         '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
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
def construir_materias(rng):
    materias, usadas = [], set()
    for prefijo, nombre in MATERIAS:
        while True:
            clave = "%s%d%03d" % (prefijo, rng.randint(1, 3), rng.randrange(1000))
            if clave not in usadas:
                usadas.add(clave)
                break
        materias.append(dict(clave=clave, nombre=nombre))
    return materias


def construir_profesores(rng):
    fijos = ["Laura Medina Garza", "Rodrigo Salinas Peña"]
    vistos, profs = set(fijos), []
    while len(profs) < N_PROFESORES - len(fijos):
        a1, a2 = rng.sample(APELLIDOS, 2)
        nombre = "%s %s %s" % (rng.choice(NOMBRES), a1, a2)
        if nombre not in vistos:
            vistos.add(nombre)
            profs.append(nombre)
    return profs + fijos


def horario(rng):
    dias, dur, inicios = rng.choices(PATRONES, PATRON_PESOS)[0]
    ini = rng.choice(inicios)
    fin = ini + dur
    return "%s %02d:%02d-%02d:%02d" % (dias, ini // 60, ini % 60, fin // 60, fin % 60)


def salon(rng):
    return "A%d-%d%02d" % (rng.randint(1, 8), rng.randint(1, 4), rng.randint(1, 20))


def nuevo_grupo(rng, crn, materias, profs, contador):
    m = rng.randrange(len(materias))
    c = rng.choices(range(len(CAMPUS)), CAMPUS_PESOS)[0]
    contador[(m, c)] = contador.get((m, c), 0) + 1
    mod = rng.choices(MODALIDADES, MODALIDAD_PESOS)[0]
    return {"CRN": crn, "Clave de materia": materias[m]["clave"],
            "Materia": materias[m]["nombre"], "Grupo": contador[(m, c)], "Campus": CAMPUS[c],
            "Modalidad": mod, "Horario": horario(rng),
            "Salón": SIN_SALON if mod == "En línea" else salon(rng),
            "Profesor": rng.choice(profs), "Cupo": rng.choice(CUPOS[mod]),
            "Estatus del grupo": CANCELADO if rng.random() < P_CANCELADO else CONFIRMADO}


def valor_nuevo(rng, g, col, profs):
    """Un valor distinto del actual para la columna col del grupo g."""
    actual = g[col]
    while True:
        if col == "Profesor":
            v = rng.choice(profs)
        elif col == "Cupo":
            v = actual + rng.choice([5, 10, -5])
            if v < 20:
                continue
        elif col == "Horario":
            v = horario(rng)
        elif col == "Salón":
            v = salon(rng)
        elif col == "Estatus del grupo":
            v = CANCELADO
        else:
            raise ValueError(col)
        if v != actual:
            return v


def construir(rng):
    materias = construir_materias(rng)
    profs = construir_profesores(rng)
    crn_fijos = [f[0] for f in FIJOS]
    libres = [c for c in range(41000, 41460) if c not in crn_fijos]
    crns = sorted(rng.sample(libres, N_ANTERIOR - len(crn_fijos)) + crn_fijos)
    contador = {}
    grupos = {crn: nuevo_grupo(rng, crn, materias, profs, contador) for crn in crns}

    # Los tres grupos fijos, con el valor de antes que dice el mazo
    grupos[41027]["Profesor"] = "Laura Medina Garza"
    grupos[41188]["Estatus del grupo"] = CONFIRMADO
    g = grupos[41305]
    if g["Modalidad"] == "En línea":
        g["Modalidad"], g["Salón"] = "Presencial", salon(rng)
    g["Cupo"] = 35

    # Los que ya no aparecen y los que cambian (nunca los fijos)
    otros = [c for c in crns if c not in crn_fijos]
    ya_no = sorted(rng.sample(otros, N_YA_NO))
    candidatos = [c for c in otros if c not in ya_no]
    rng.shuffle(candidatos)
    cambios, usados = [], set(crn_fijos)
    for crn, col, antes, despues in FIJOS:
        cambios.append((crn, col, antes, despues))
    for columnas, cuantos in OTROS_CAMBIOS:
        for _ in range(cuantos):
            for crn in candidatos:
                g = grupos[crn]
                if crn in usados:
                    continue
                if "Estatus del grupo" in columnas and g["Estatus del grupo"] != CONFIRMADO:
                    continue
                if "Salón" in columnas and g["Modalidad"] == "En línea":
                    continue
                usados.add(crn)
                for col in columnas:
                    cambios.append((crn, col, g[col], valor_nuevo(rng, g, col, profs)))
                break
    orden_col = {c: i for i, c in enumerate(COLUMNAS)}
    cambios.sort(key=lambda x: (x[0], orden_col[x[1]]))

    # Version de las 15:00: sin los que ya no aparecen, con los cambios y los nuevos
    despues = {crn: dict(g) for crn, g in grupos.items() if crn not in ya_no}
    for crn, col, antes, valor in cambios:
        assert despues[crn][col] == antes
        despues[crn][col] = valor
    nuevos = sorted(rng.sample(range(41460, 41500), N_NUEVOS))
    for crn in nuevos:
        despues[crn] = nuevo_grupo(rng, crn, materias, profs, contador)

    anterior = [grupos[c] for c in crns]
    nueva = sorted(despues.values(), key=lambda g: (CAMPUS.index(g["Campus"]),
                                                     fold(g["Materia"]), g["Grupo"], g["CRN"]))
    return anterior, nueva, nuevos, ya_no, cambios, grupos, despues


# ------------------------------------------------------------------ notas
def tres_lineas(nombre_ant, nombre_nueva, n_nuevos, n_ya_no, n_grupos, n_celdas, por_col):
    top = sorted([c for c in COLUMNAS if por_col.get(c)],
                 key=lambda c: (-por_col[c], COLUMNAS.index(c)))[:3]
    base = nombre_nueva.rsplit(".", 1)[0]
    l1 = "Comparación de %s contra %s, por CRN." % (nombre_ant, nombre_nueva)
    l2 = ("Hay %s grupos nuevos, %s grupos que ya no aparecen y %s grupos con cambios "
          "(%s celdas)." % (miles(n_nuevos), miles(n_ya_no), miles(n_grupos), miles(n_celdas)))
    l3 = "Columnas con más cambios: %s. Detalle: cambios-%s.xlsx." % (
        y_lista("%s (%s)" % (c, miles(por_col[c])) for c in top), base)
    return [l1, l2, l3]


def escribir_notas(ruta, semilla, nombres, anterior, nueva, nuevos, ya_no, cambios, grupos,
                   despues):
    por_col = {}
    for _, col, _, _ in cambios:
        por_col[col] = por_col.get(col, 0) + 1
    con_cambios = sorted({c[0] for c in cambios})
    dobles = [crn for crn in con_cambios if sum(1 for c in cambios if c[0] == crn) == 2]
    comunes = len(anterior) - len(ya_no)
    sin_cambios = comunes - len(con_cambios)
    lineas = tres_lineas(nombres[0], nombres[1], len(nuevos), len(ya_no), len(con_cambios),
                         len(cambios), por_col)
    # Lo que daria comparar renglon por renglon (por posicion), para leer respuestas equivocadas
    pos_renglones = pos_celdas = 0
    for a, b in zip(anterior, nueva):
        dif = sum(1 for c in COLUMNAS if a[c] != b[c])
        pos_celdas += dif
        pos_renglones += dif > 0

    def fmt(v):
        return "`%s`" % v

    s = []
    s.append("# Notas de generación · oferta de grupos ficticia (caso de Karla, S7)\n")
    s.append("**Todos los datos de estos archivos son ficticios.** Ningún nombre de profesor "
             "corresponde a una persona real; las materias, claves, CRN, horarios y salones son "
             "inventados. Este archivo lo escribe `generar_versiones.py`: las cifras salen de lo "
             "que el generador sembró, no de leer los archivos, y `comparar_versiones.py` las "
             "reproduce por su lado leyendo las dos versiones.\n")
    s.append("- Semilla: %d\n- Periodo: %s\n- Fecha de las dos versiones: miércoles 21 de "
             "octubre de 2026\n- Comando (desde la carpeta `sesion-07`): `python3 "
             "generador/generar_versiones.py --semilla %d --salida versiones --notas "
             "generador/NOTAS-generacion.md`\n" % (semilla, PERIODO, semilla))
    s.append("## Las dos versiones\n")
    s.append("| Archivo | Hora | Grupos (renglones de datos) | Orden de los renglones |\n"
             "|---|---|---:|---|")
    s.append("| `%s` | 09:00 | %s | por CRN |" % (nombres[0], miles(len(anterior))))
    s.append("| `%s` | 15:00 | %s | por campus, materia y grupo |" % (nombres[1],
                                                                      miles(len(nueva))))
    s.append("")
    s.append("Encabezado en la fila 1, una fila por CRN y once columnas, en este orden: %s. "
             "La hoja se llama `Oferta %s`. Seis campus (%s); modalidad Presencial, Híbrida o "
             "En línea (los grupos en línea dicen `%s` en Salón); estatus `%s` o `%s`. Los "
             "valores vienen limpios en las dos versiones: sin espacios de más y con CRN, "
             "Grupo y Cupo como números.\n"
             % (", ".join("`%s`" % c for c in COLUMNAS), PERIODO, ", ".join(CAMPUS),
                SIN_SALON, CONFIRMADO, CANCELADO))
    s.append("## Lo que sembró el generador\n")
    s.append("| Qué | Cuántos |\n|---|---:|")
    s.append("| Grupos nuevos (CRN solo en la de las 15:00) | %d |" % len(nuevos))
    s.append("| Grupos que ya no aparecen (CRN solo en la de las 09:00) | %d |" % len(ya_no))
    s.append("| Grupos con cambios | %d |" % len(con_cambios))
    s.append("| Celdas que cambiaron | %d |" % len(cambios))
    for c in sorted([c for c in COLUMNAS if por_col.get(c)],
                    key=lambda c: (-por_col[c], COLUMNAS.index(c))):
        s.append("| · en %s | %d |" % (c, por_col[c]))
    s.append("| Grupos con dos cambios | %d |" % len(dobles))
    s.append("| Grupos sin cambios (%d − %d − %d) | %d |" % (len(anterior), len(ya_no),
                                                            len(con_cambios), sin_cambios))
    s.append("")
    s.append("Además, los renglones de la versión de las 15:00 vienen en otro orden.\n")
    s.append("**Grupos nuevos:**\n")
    s.append("| CRN | Clave de materia | Materia | Campus | Profesor |\n|---|---|---|---|---|")
    for crn in nuevos:
        g = despues[crn]
        s.append("| %d | %s | %s | %s | %s |" % (crn, g["Clave de materia"], g["Materia"],
                                                g["Campus"], g["Profesor"]))
    s.append("")
    s.append("**Grupos que ya no aparecen:**\n")
    s.append("| CRN | Clave de materia | Materia | Campus | Profesor |\n|---|---|---|---|---|")
    for crn in ya_no:
        g = grupos[crn]
        s.append("| %d | %s | %s | %s | %s |" % (crn, g["Clave de materia"], g["Materia"],
                                                g["Campus"], g["Profesor"]))
    s.append("")
    s.append("**Celdas que cambiaron** (lo mismo que trae la hoja Cambios de "
             "`comparar_versiones.py`):\n")
    s.append("| CRN | Materia | Columna | Antes | Después |\n|---|---|---|---|---|")
    for crn, col, antes, valor in cambios:
        s.append("| %d | %s | %s | %s | %s |" % (crn, grupos[crn]["Materia"], col, antes, valor))
    s.append("")
    s.append("Grupos con dos cambios: %s." % y_lista(
        "%d (%s)" % (crn, " y ".join(c[1] for c in cambios if c[0] == crn)) for crn in dobles))
    s.append("")
    s.append("**Los tres cambios fijos** (los que muestra el mazo; no dependen de la semilla):\n")
    for crn, col, antes, valor in FIJOS:
        s.append("- CRN %d, %s: %s → %s" % (crn, col, fmt(antes), fmt(valor)))
    s.append("")
    s.append("## Respuesta correcta\n")
    s.append("Las tres líneas de resumen (texto exacto que deja `comparar_versiones.py` en "
             "`resumen-cambios-%s.txt` con `--clave CRN`):\n" % nombres[1].rsplit(".", 1)[0])
    s.append("```\n%s\n```\n" % "\n".join(lineas))
    s.append("## Cómo leer una respuesta equivocada\n")
    s.append("- Si dice que cambiaron casi todos los grupos, comparó por posición (renglón 2 "
             "contra renglón 2) y no por CRN: la versión de las 15:00 viene en otro orden. "
             "Comparando por posición, %s de los %s primeros renglones salen distintos (%s "
             "celdas)." % (miles(pos_renglones), miles(min(len(anterior), len(nueva))),
                           miles(pos_celdas)))
    s.append("- Si dice %d grupos nuevos, restó renglones (%d − %d) en lugar de comparar "
             "claves: hay %d nuevos y %d que ya no aparecen."
             % (len(nueva) - len(anterior), len(nueva), len(anterior), len(nuevos), len(ya_no)))
    s.append("- Si dice más de %d celdas, contó como cambio algo que no lo es: espacios de más, "
             "o `35` contra `35.0`. Estas dos versiones no traen ninguna de las dos cosas, pero "
             "un archivo real puede traerlas, y el script de referencia las cuenta como iguales."
             % len(cambios))
    s.append("")
    Path(ruta).write_text("\n".join(s), encoding="utf-8")
    return lineas


# ------------------------------------------------------------------ principal
def main():
    ap = argparse.ArgumentParser(description="Genera las dos versiones ficticias de la oferta "
                                             "de grupos (caso de Karla, S7).")
    ap.add_argument("--semilla", type=int, default=21)
    ap.add_argument("--salida", default="versiones")
    ap.add_argument("--notas", default=None)
    args = ap.parse_args()

    rng = random.Random(args.semilla)
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)
    anterior, nueva, nuevos, ya_no, cambios, grupos, despues = construir(rng)

    nombres = []
    for (sufijo, hora), filas in zip(VERSIONES, (anterior, nueva)):
        nombre = "oferta-grupos-%s-%s.xlsx" % (FECHA, sufijo)
        hoja = [[(c, NEGRITA) for c in COLUMNAS]] + [[g[c] for c in COLUMNAS] for g in filas]
        escribir_xlsx(out / nombre, [("Oferta %s" % PERIODO, hoja, ANCHOS, False)])
        nombres.append(nombre)
        print("%s (%s): %s grupos" % (nombre, hora, miles(len(filas))))

    notas = Path(args.notas) if args.notas else out / "NOTAS-generacion.md"
    notas.parent.mkdir(parents=True, exist_ok=True)
    lineas = escribir_notas(notas, args.semilla, nombres, anterior, nueva, nuevos, ya_no,
                            cambios, grupos, despues)
    print("Sembrado: %d grupos nuevos, %d que ya no aparecen, %d grupos con cambios "
          "(%d celdas)." % (len(nuevos), len(ya_no), len({c[0] for c in cambios}),
                            len(cambios)))
    print("Respuesta correcta (tres líneas):")
    for l in lineas:
        print("  " + l)
    print("Notas: %s" % notas.resolve())
    print("Salida: %s" % out.resolve())
    return 0


if __name__ == "__main__":
    sys.exit(main())
