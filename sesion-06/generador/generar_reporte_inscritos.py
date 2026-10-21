#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_reporte_inscritos.py: genera el "reporte de inscritos" FICTICIO del flujo de
Cecilia (curso Ed Digital, Sesion 6). Imita el archivo que exporta un sistema de
inscripciones cada lunes y cada miercoles con las peticiones del periodo, y lo deja
SUCIO como llega: encabezado que no esta en la fila 1, filas de seccion y subtotales,
columnas con espacios, estatus con variantes de mayusculas, campus escritos de dos
formas, renglones duplicados, co-titulares que repiten la peticion y fechas en texto.

Como correrlo (solo biblioteca estandar, Python 3.8 o mas reciente):
    python3 generar_reporte_inscritos.py --semilla 42 --salida reportes
    python3 generar_reporte_inscritos.py --semilla 42 --salida reportes --formato ambos

Parametros:
    --semilla  numero que fija el azar (42). Misma semilla, mismos archivos y mismas cifras.
    --salida   carpeta donde se escriben los reportes (reportes).
    --formato  xlsx, csv o ambos (xlsx).
    --notas    ruta de NOTAS-generacion.md (por omision, dentro de --salida).

Salen dos entregas del mismo reporte:
    reporte-inscritos-2026-10-19.xlsx  lunes, corte a las 07:00
    reporte-inscritos-2026-10-21.xlsx  miercoles; trae las peticiones del lunes con el
                                       estatus al dia y suma las que llegaron despues
NOTAS-generacion.md trae la respuesta correcta de la tabla limpia de cada entrega.
Todos los datos son inventados; ningun nombre, matricula, nomina ni correo es real.
"""
import argparse
import csv
import random
import sys
import unicodedata
import zipfile
from datetime import date, timedelta
from pathlib import Path
from xml.sax.saxutils import escape

PERIODO = "Febrero-Junio 2027"
INICIO_SOLICITUDES = date(2026, 10, 5)
ULTIMO_DIA_SOLICITUDES = date(2026, 10, 20)
CORTES = [date(2026, 10, 19), date(2026, 10, 21)]   # lunes y miercoles
HORA_CORTE = "07:00"

CAMPUS = ["Monterrey", "Guadalajara", "Ciudad de México", "Querétaro",
          "Puebla", "Estado de México"]
CAMPUS_PESOS = [0.34, 0.20, 0.20, 0.10, 0.09, 0.07]
ESTATUS = ["Nueva", "En proceso", "Resuelta", "Cancelada"]
CANALES = ["Formulario en línea", "Correo", "Ventanilla", "Teams"]
CANAL_PESOS = [0.55, 0.25, 0.12, 0.08]
OBSERVACIONES = ["Solicita cupo adicional", "Choque de horario", "Cambio de grupo",
                 "Requisito pendiente", "Autorización de dirección", "Beca académica"]

N_ALUMNOS = 1250
N_PROFESORES = 160
N_GRUPOS = 260
P_COTITULAR = 0.08
P_DUPLICADO = 0.015
P_CAMPUS_VARIANTE = 0.02
P_ESTATUS_VARIANTE = 0.30
DOMINIO_ALUMNO = "@alumnos.ejemplo.edu.mx"

COLUMNAS = ["Periodo", "Campus", "Clave materia", "Nombre materia", "Grupo", "CRN",
            "Modalidad", "Nómina profesor", "Nombre profesor", "Matrícula alumno",
            "Nombre alumno", "Correo alumno", "Programa", "Semestre alumno",
            "Estatus inscripción", "Tipo de inscripción", "Créditos",
            "Promedio previo", "Calificación parcial 1", "Asistencia %",
            "Observaciones", "Fecha de solicitud", "Estatus de la petición", "Canal"]
# Como las escribe el sistema: con espacios de mas
ENCABEZADO_SUCIO = {"Campus": " Campus", "Clave materia": "Clave materia ",
                    "Nombre materia": "Nombre  materia",
                    "Matrícula alumno": "Matrícula alumno ",
                    "Fecha de solicitud": "Fecha de solicitud ",
                    "Estatus de la petición": "Estatus de la petición  ",
                    "Canal": " Canal"}

ESCUELAS = [
    ("Negocios", "N", 0.24, ["LAE", "LAF", "LCDE", "LMC", "LEM"],
     ["Finanzas Corporativas", "Mercadotecnia Digital", "Contabilidad Financiera",
      "Microeconomía", "Gestión de Proyectos", "Emprendimiento", "Econometría"]),
    ("Ingeniería", "I", 0.30, ["IMT", "ISC", "ITC", "IME", "IC", "IQ", "IID"],
     ["Circuitos Eléctricos", "Programación", "Bases de Datos", "Termodinámica",
      "Estructuras", "Control Automático", "Investigación de Operaciones"]),
    ("Humanidades", "H", 0.14, ["LRI", "LPS", "LLE", "LCC", "LEC"],
     ["Sociología", "Redacción Académica", "Periodismo", "Psicología Social",
      "Ciencia Política", "Lingüística", "Psicometría"]),
    ("Medicina", "M", 0.10, ["MC", "LNB", "LBC", "LFT"],
     ["Anatomía", "Fisiología", "Bioquímica", "Epidemiología", "Farmacología",
      "Bioestadística", "Nutrición"]),
    ("Arquitectura", "A", 0.10, ["LAR", "LUB", "LDI"],
     ["Taller de Diseño Arquitectónico", "Diseño Urbano", "Sistemas Constructivos",
      "Teoría de la Arquitectura", "Paisaje", "Instalaciones", "Movilidad"]),
    ("Derecho", "D", 0.12, ["LD", "LDE", "LCP"],
     ["Derecho Civil", "Derecho Mercantil", "Derecho Constitucional", "Derecho Penal",
      "Derecho Fiscal", "Derechos Humanos", "Comercio Exterior"]),
]
ESC_PESOS = [e[2] for e in ESCUELAS]
NIVELES = ["I", "II", "III", "IV"]

NOMBRES = [
    "José", "Juan", "Luis", "Carlos", "Miguel", "Pedro", "Ricardo", "Fernando",
    "Roberto", "Alejandro", "Daniel", "Eduardo", "Javier", "Manuel", "Francisco",
    "Sergio", "Raúl", "Héctor", "Óscar", "Andrés", "Diego", "Gerardo", "Mario",
    "Arturo", "Ángel", "Iván", "Rodrigo", "Emilio", "Gustavo", "Rafael",
    "María", "Guadalupe", "Ana", "Laura", "Patricia", "Claudia", "Gabriela",
    "Verónica", "Sofía", "Fernanda", "Alejandra", "Daniela", "Mariana", "Paola",
    "Karla", "Lucía", "Elena", "Rosa", "Teresa", "Silvia", "Adriana", "Leticia",
    "Mónica", "Beatriz", "Carmen", "Irene", "Julieta", "Valeria", "Renata", "Andrea",
]
APELLIDOS = [
    "Hernández", "García", "Martínez", "López", "González", "Rodríguez", "Pérez",
    "Sánchez", "Ramírez", "Flores", "Gómez", "Díaz", "Reyes", "Morales", "Cruz",
    "Ortiz", "Gutiérrez", "Chávez", "Ruiz", "Jiménez", "Mendoza", "Álvarez",
    "Castillo", "Vázquez", "Romero", "Torres", "Herrera", "Medina", "Aguilar",
    "Vargas", "Castro", "Ramos", "Moreno", "Guerrero", "Salazar", "Domínguez",
    "Rojas", "Silva", "Navarro", "Cortés", "Soto", "Contreras", "Luna", "Estrada",
    "Delgado", "Ibarra", "Espinoza", "Santiago", "Vega", "Campos", "Cabrera",
    "Fuentes", "Rivera", "Mejía", "Carrillo", "Sandoval", "Núñez", "Zamora",
    "Trejo", "Montes", "Peña", "Ávila", "Cervantes", "Ríos", "Lara", "Valdez",
    "Maldonado", "Padilla", "Acosta", "Bravo", "Orozco", "Villarreal", "Leal",
    "Garza", "Treviño", "Elizondo", "Benítez", "Cantú", "Solís", "Macías", "Pacheco",
]
MESES_ABR = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct",
             "nov", "dic"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))


def miles(n):
    return "{:,}".format(n)


def fecha_larga(d, con_anio=True):
    s = "%s %d de %s" % (DIAS[d.weekday()], d.day, MESES[d.month - 1])
    return s + (" de %d" % d.year if con_anio else "")


def corte_anterior(d):
    """El corte anterior de Cecilia: el lunes o miercoles mas reciente antes de d."""
    a = d - timedelta(days=1)
    while a.weekday() not in (0, 2):
        a -= timedelta(days=1)
    return a


def lineas_teams(corte, hora, periodo, n_pet, n_alum, por_estatus, ant, nuevas,
                 campus_top, n_top, nombre_tabla):
    l1 = "Reporte de inscritos, corte del %s a las %s (%s): %s peticiones de %s alumnos." % (
        fecha_larga(corte), hora, periodo, miles(n_pet), miles(n_alum))
    l2 = "Por estatus: %s nuevas, %s en proceso, %s resueltas y %s canceladas." % (
        miles(por_estatus["Nueva"]), miles(por_estatus["En proceso"]),
        miles(por_estatus["Resuelta"]), miles(por_estatus["Cancelada"]))
    l3 = ("Desde el corte del %s llegaron %s peticiones; el campus con más peticiones "
          "es %s (%s). Tabla completa: %s.") % (
        fecha_larga(ant, con_anio=False), miles(nuevas), campus_top, miles(n_top),
        nombre_tabla)
    return [l1, l2, l3]


# ------------------------------------------------------------------ xlsx (stdlib)
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def col_letra(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def escribir_xlsx(ruta, nombre_hoja, filas, negritas=()):
    """Escribe un xlsx de una hoja. Cadenas en linea (conservan espacios), numeros
    como numeros, renglones vacios se omiten. negritas: indices (base 0) de renglones."""
    partes = []
    for i, fila in enumerate(filas):
        if not fila:
            continue
        r = i + 1
        estilo = ' s="1"' if i in negritas else ""
        celdas = []
        for j, v in enumerate(fila):
            if v is None or v == "":
                continue
            ref = "%s%d" % (col_letra(j), r)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                celdas.append('<c r="%s"%s><v>%s</v></c>' % (ref, estilo, repr(v)))
            else:
                celdas.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                              % (ref, estilo, escape(str(v))))
        partes.append('<row r="%d">%s</row>' % (r, "".join(celdas)))
    hoja = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s" xmlns:r="%s"><sheetData>%s</sheetData></worksheet>'
            % (NS_MAIN, NS_REL, "".join(partes)))
    archivos = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="%s" xmlns:r="%s"><sheets>'
            '<sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>'
            % (NS_MAIN, NS_REL, escape(nombre_hoja)),
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            '</Relationships>',
        "xl/styles.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<styleSheet xmlns="%s">'
            '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
            '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="2"><fill><patternFill patternType="none"/></fill>'
            '<fill><patternFill patternType="gray125"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
            '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>' % NS_MAIN,
        "xl/worksheets/sheet1.xml": hoja,
    }
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre in ["[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                       "xl/_rels/workbook.xml.rels", "xl/styles.xml",
                       "xl/worksheets/sheet1.xml"]:
            info = zipfile.ZipInfo(nombre, date_time=(2026, 10, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, archivos[nombre].encode("utf-8"))


def escribir_csv(ruta, filas):
    with open(ruta, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        for fila in filas:
            w.writerow(["" if v is None else v for v in fila])


# ------------------------------------------------------------------ catalogos
def construir_materias(rng):
    materias, usadas = [], set()
    for ei, (_, letra, _, _, temas) in enumerate(ESCUELAS):
        combos = [(t, n) for t in temas for n in range(1, 5)]
        for tema, nivel in rng.sample(combos, 14):
            while True:
                clave = "%s%d%03d" % (letra, nivel, rng.randrange(1000))
                if clave not in usadas:
                    usadas.add(clave)
                    break
            cred = rng.choices([3, 6, 8], [0.5, 0.35, 0.15])[0]
            materias.append(dict(clave=clave, nombre="%s %s" % (tema, NIVELES[nivel - 1]),
                                 esc=ei, cred=cred))
    return materias


def construir_profesores(rng):
    combos = rng.sample(range(len(NOMBRES) * len(APELLIDOS)), N_PROFESORES)
    nominas = rng.sample(range(10_000_000, 100_000_000), N_PROFESORES)
    profs = []
    for i in range(N_PROFESORES):
        nom = NOMBRES[combos[i] // len(APELLIDOS)]
        ap1 = APELLIDOS[combos[i] % len(APELLIDOS)]
        profs.append(dict(nomina="L%08d" % nominas[i],
                          nombre="%s %s %s" % (nom, ap1, rng.choice(APELLIDOS)),
                          esc=rng.choices(range(len(ESCUELAS)), ESC_PESOS)[0]))
    return profs


def construir_grupos(rng, materias, profs):
    por_esc_mat = [[m for m in range(len(materias)) if materias[m]["esc"] == e]
                   for e in range(len(ESCUELAS))]
    por_esc_prof = [[p for p in range(len(profs)) if profs[p]["esc"] == e]
                    for e in range(len(ESCUELAS))]
    crns = rng.sample(range(10000, 100000), N_GRUPOS)
    contador, grupos = {}, []
    for g in range(N_GRUPOS):
        if g < 36:   # cada campus con al menos un grupo de cada escuela
            c, e = g % 6, g // 6
        else:
            c = rng.choices(range(6), CAMPUS_PESOS)[0]
            e = rng.choices(range(len(ESCUELAS)), ESC_PESOS)[0]
        m = rng.choice(por_esc_mat[e])
        contador[(c, m)] = contador.get((c, m), 0) + 1
        tit = rng.choice(por_esc_prof[e])
        co = -1
        if rng.random() < P_COTITULAR:
            co = rng.choice([p for p in por_esc_prof[e] if p != tit])
        grupos.append(dict(campus=c, esc=e, mat=m, num=contador[(c, m)], crn=crns[g],
                           mod=rng.choices(["Presencial", "Híbrida", "Remota"],
                                           [0.60, 0.25, 0.15])[0],
                           tit=tit, co=co))
    return grupos


def construir_alumnos(rng):
    mats = rng.sample(range(1_000_000, 10_000_000), N_ALUMNOS)
    alumnos = []
    for i in range(N_ALUMNOS):
        e = rng.choices(range(len(ESCUELAS)), ESC_PESOS)[0]
        mat = "A0%07d" % mats[i]
        alumnos.append(dict(
            mat=mat, correo=mat.lower() + DOMINIO_ALUMNO,
            nombre="%s %s %s" % (rng.choice(NOMBRES), rng.choice(APELLIDOS),
                                 rng.choice(APELLIDOS)),
            prog=rng.choice(ESCUELAS[e][3]), esc=e,
            campus=rng.choices(range(6), CAMPUS_PESOS)[0],
            sem=rng.randint(1, 9),
            prom=round(min(99.0, max(65.0, rng.gauss(84, 7))), 1)))
    return alumnos


def construir_peticiones(rng, alumnos, grupos, materias):
    dias = []
    d = INICIO_SOLICITUDES
    while d <= ULTIMO_DIA_SOLICITUDES:
        dias.append(d)
        d += timedelta(days=1)
    pesos_dia = [0.3 if x.weekday() == 5 else (0.2 if x.weekday() == 6 else 1.0)
                 for x in dias]
    pool_campus = [[g for g in range(len(grupos)) if grupos[g]["campus"] == c]
                   for c in range(6)]
    pool_ce = {}
    for g, gr in enumerate(grupos):
        pool_ce.setdefault((gr["campus"], gr["esc"]), []).append(g)
    peticiones = []
    for a, al in enumerate(alumnos):
        k = rng.choices([1, 2, 3, 4, 5], [0.08, 0.25, 0.35, 0.22, 0.10])[0]
        d0 = rng.choices(dias, pesos_dia)[0]
        propios = pool_ce.get((al["campus"], al["esc"]), [])
        todos = pool_campus[al["campus"]]
        elegidos, mats_el, intentos = [], set(), 0
        while len(elegidos) < k and intentos < 60:
            intentos += 1
            pool = propios if (propios and rng.random() < 0.8) else todos
            g = rng.choice(pool)
            if g in elegidos or grupos[g]["mat"] in mats_el:
                continue
            elegidos.append(g)
            mats_el.add(grupos[g]["mat"])
        for g in elegidos:
            f = d0 if rng.random() < 0.7 else d0 + timedelta(days=rng.randint(1, 4))
            d1 = rng.choices([1, 2, 3, 4], [0.35, 0.35, 0.20, 0.10])[0]
            if rng.random() < 0.05:
                d2, final, directa = 0, "Cancelada", True
            else:
                d2 = rng.choices([1, 2, 3, 4, 5, 6], [0.15, 0.25, 0.25, 0.15, 0.12, 0.08])[0]
                final = "Resuelta" if rng.random() < 0.85 else "Cancelada"
                directa = False
            minuto = (rng.randint(8, 19), rng.randint(0, 59))
            if f > ULTIMO_DIA_SOLICITUDES:
                continue
            peticiones.append(dict(
                alumno=a, grupo=g, fecha=f, minuto=minuto, d1=d1, d2=d2, final=final,
                directa=directa,
                canal=rng.choices(CANALES, CANAL_PESOS)[0],
                tipo="Ordinaria" if rng.random() < 0.90 else "Recursamiento",
                obs=rng.choice(OBSERVACIONES) if rng.random() < 0.04 else ""))
    return peticiones


def estatus_al_corte(p, corte):
    edad = (corte - p["fecha"]).days
    if edad < p["d1"]:
        return "Nueva"
    if p["directa"]:
        return "Cancelada"
    if edad < p["d1"] + p["d2"]:
        return "En proceso"
    return p["final"]


# ------------------------------------------------------------------ suciedad
def campus_texto(rng, nombre):
    if rng.random() < P_CAMPUS_VARIANTE:
        return nombre.upper() if rng.random() < 0.5 else nombre + " "
    return nombre


def estatus_texto(rng, e):
    if rng.random() >= P_ESTATUS_VARIANTE:
        return e
    v = rng.choice(["mayus", "minus", "titulo", "fin", "inicio", "doble"])
    if v == "mayus":
        return e.upper()
    if v == "minus":
        return e.lower()
    if v == "titulo":
        return e.title()
    if v == "fin":
        return e + " "
    if v == "inicio":
        return " " + e
    return e.replace(" ", "  ")


def fecha_texto(rng, p):
    d = p["fecha"]
    r = rng.random()
    if r < 0.50:
        return d.strftime("%d/%m/%Y"), "dd/mm/aaaa"
    if r < 0.70:
        return d.isoformat(), "aaaa-mm-dd"
    if r < 0.85:
        return "%d-%s-%d" % (d.day, MESES_ABR[d.month - 1], d.year), "d-mmm-aaaa"
    return "%s %02d:%02d" % (d.strftime("%d/%m/%Y"), p["minuto"][0], p["minuto"][1]), \
        "dd/mm/aaaa hh:mm"


ESTATUS_INSCRIPCION = {"Nueva": "Pendiente", "En proceso": "Pendiente",
                       "Resuelta": "Inscrito", "Cancelada": "Baja"}


# ------------------------------------------------------------------ una entrega
def armar_entrega(rng, corte, peticiones, alumnos, grupos, materias, profs):
    incluidas = [p for p in peticiones if p["fecha"] < corte]
    st = dict(duplicados=0, cotitular=0, campus_var=0, estatus_var=0, datos=0,
              formatos={})
    filas = [["SISTEMA DE INSCRIPCIONES · REPORTE DE INSCRITOS"],
             ["Periodo: %s" % PERIODO],
             ["Fecha de corte: %s %s" % (corte.strftime("%d/%m/%Y"), HORA_CORTE)],
             [],
             [ENCABEZADO_SUCIO.get(c, c) for c in COLUMNAS]]
    tabla = {c: {e: 0 for e in ESTATUS} for c in CAMPUS}
    estatus_de = {}
    orden_secciones = sorted(range(6), key=lambda c: sin_acentos(CAMPUS[c]))
    for c in orden_secciones:
        sec = [p for p in incluidas if grupos[p["grupo"]]["campus"] == c]
        sec.sort(key=lambda p: (p["fecha"], alumnos[p["alumno"]]["mat"],
                                grupos[p["grupo"]]["crn"]))
        filas.append(["CAMPUS: %s" % CAMPUS[c].upper()])
        n_sec = 0
        for p in sec:
            al, gr = alumnos[p["alumno"]], grupos[p["grupo"]]
            m, pt = materias[gr["mat"]], profs[gr["tit"]]
            est = estatus_al_corte(p, corte)
            estatus_de[(al["mat"], gr["crn"])] = est
            tabla[CAMPUS[c]][est] += 1
            ftxt, formato = fecha_texto(rng, p)
            ctxt = campus_texto(rng, CAMPUS[c])
            etxt = estatus_texto(rng, est)
            fila = [PERIODO, ctxt, m["clave"], m["nombre"], gr["num"], gr["crn"],
                    gr["mod"], pt["nomina"], pt["nombre"], al["mat"], al["nombre"],
                    al["correo"], al["prog"], al["sem"], ESTATUS_INSCRIPCION[est],
                    p["tipo"], m["cred"], al["prom"], "", "", p["obs"], ftxt, etxt,
                    p["canal"]]
            emitir = [(fila, False)]
            if gr["co"] >= 0:
                fila2 = list(fila)
                fila2[7], fila2[8] = profs[gr["co"]]["nomina"], profs[gr["co"]]["nombre"]
                emitir.append((fila2, True))
            for f, es_co in emitir:
                veces = 2 if rng.random() < P_DUPLICADO else 1
                for _ in range(veces):
                    filas.append(f)
                    n_sec += 1
                    st["campus_var"] += ctxt != CAMPUS[c]
                    st["estatus_var"] += etxt != est
                    st["formatos"][formato] = st["formatos"].get(formato, 0) + 1
                st["duplicados"] += veces - 1
                st["cotitular"] += es_co
        filas.append(["Subtotal %s: %s registros" % (CAMPUS[c], miles(n_sec))])
        st["datos"] += n_sec
    filas.append(["Total general: %s registros" % miles(st["datos"])])

    ant = corte_anterior(corte)
    st.update(
        corte=corte, anterior=ant, filas_hoja=len(filas), peticiones=len(incluidas),
        alumnos=len({p["alumno"] for p in incluidas}),
        nuevas=sum(1 for p in incluidas if p["fecha"] >= ant),
        tabla=tabla, estatus_de=estatus_de)
    assert st["datos"] == st["peticiones"] + st["cotitular"] + st["duplicados"]
    return filas, st


# ------------------------------------------------------------------ notas
def tabla_md(tabla):
    lin = ["| Campus | " + " | ".join(ESTATUS) + " | Total |",
           "|---|" + "---:|" * (len(ESTATUS) + 1)]
    tot = {e: 0 for e in ESTATUS}
    for c in CAMPUS:
        fila = tabla[c]
        for e in ESTATUS:
            tot[e] += fila[e]
        lin.append("| %s | %s | %s |" % (c, " | ".join(miles(fila[e]) for e in ESTATUS),
                                         miles(sum(fila.values()))))
    lin.append("| **Total** | %s | **%s** |" % (" | ".join("**%s**" % miles(tot[e])
                                                            for e in ESTATUS),
                                                 miles(sum(tot.values()))))
    return "\n".join(lin), tot


def bloque_entrega(nombre, st):
    t_md, tot = tabla_md(st["tabla"])
    totales = {c: sum(st["tabla"][c].values()) for c in CAMPUS}
    top = max(CAMPUS, key=lambda c: (totales[c], -CAMPUS.index(c)))
    nombre_tabla = "tabla-inscritos-%s.xlsx" % st["corte"].isoformat()
    lineas = lineas_teams(st["corte"], HORA_CORTE, PERIODO, st["peticiones"],
                          st["alumnos"], tot, st["anterior"], st["nuevas"], top,
                          totales[top], nombre_tabla)
    no_datos = st["filas_hoja"] - st["datos"]
    formatos = ", ".join("%s: %s" % (k, miles(st["formatos"].get(k, 0)))
                         for k in ["dd/mm/aaaa", "aaaa-mm-dd", "d-mmm-aaaa",
                                   "dd/mm/aaaa hh:mm"])
    s = []
    s.append("### %s · `%s`\n" % (fecha_larga(st["corte"]).capitalize(), nombre))
    s.append("Control de la limpieza (lo mismo que imprime `limpiar_reporte.py`):\n")
    s.append("| Concepto | Valor |\n|---|---:|")
    s.append("| Renglones en la hoja (título, encabezado, secciones, subtotales y datos) | %s |"
             % miles(st["filas_hoja"]))
    s.append("| Renglón del encabezado | 5 |")
    s.append("| Renglones que no son datos | %s |" % miles(no_datos))
    s.append("| Renglones de datos (lo que el sistema llama \"registros\") | %s |"
             % miles(st["datos"]))
    s.append("| Duplicados exactos quitados | %s |" % miles(st["duplicados"]))
    s.append("| Renglones repetidos por co-titular unidos | %s |" % miles(st["cotitular"]))
    s.append("| **Peticiones únicas (alumno + CRN)** | **%s** |" % miles(st["peticiones"]))
    s.append("| Alumnos distintos | %s |" % miles(st["alumnos"]))
    s.append("| Campus escritos con variante (corregidos) | %s |" % miles(st["campus_var"]))
    s.append("| Estatus escritos con variante (corregidos) | %s |" % miles(st["estatus_var"]))
    s.append("| Fechas de solicitud sin reconocer | 0 |")
    s.append("| Corte anterior | %s |" % fecha_larga(st["anterior"]))
    s.append("| Peticiones desde el corte anterior | %s |" % miles(st["nuevas"]))
    s.append("")
    s.append("Fechas de solicitud por formato (renglones de datos): %s.\n" % formatos)
    s.append("Tabla limpia (peticiones únicas por campus y estatus de la petición):\n")
    s.append(t_md)
    s.append("")
    s.append("Resumen de tres líneas para Teams (texto exacto que produce el skill):\n")
    s.append("```\n%s\n```\n" % "\n".join(lineas))
    return "\n".join(s), tot


def escribir_notas(ruta, semilla, formato, entregas):
    s = []
    s.append("# Notas de generación · reporte de inscritos ficticio (flujo de Cecilia, S6)\n")
    s.append("**Todos los datos de estos archivos son ficticios.** Ningún nombre, matrícula, "
             "nómina ni correo corresponde a una persona real. Este archivo lo escribe "
             "`generar_reporte_inscritos.py`; las cifras salen del modelo con el que se "
             "generaron los datos, no de leer el archivo, y `limpiar_reporte.py` las "
             "reproduce por su lado leyendo el reporte.\n")
    s.append("- Semilla: %d\n- Formato: %s\n- Periodo: %s\n- Comando: "
             "`python3 generar_reporte_inscritos.py --semilla %d --salida reportes`\n"
             % (semilla, formato, PERIODO, semilla))
    s.append("## Las dos entregas\n")
    s.append("| Archivo | Día | Corte | Renglones en la hoja | Renglones de datos | "
             "Peticiones únicas |\n|---|---|---|---:|---:|---:|")
    for nombre, st in entregas:
        s.append("| `%s` | %s | %s %s | %s | %s | %s |" % (
            nombre, DIAS[st["corte"].weekday()], st["corte"].strftime("%d/%m/%Y"),
            HORA_CORTE, miles(st["filas_hoja"]), miles(st["datos"]),
            miles(st["peticiones"])))
    s.append("")
    s.append("La entrega del miércoles es acumulada: trae todas las peticiones del lunes "
             "con su estatus al día y suma las que llegaron el 19 y el 20 de octubre. "
             "Es el mismo reporte dos veces con fecha distinta, para que el flujo se repita.\n")
    s.append("## Cómo viene sucio\n")
    s.append("1. El encabezado está en el renglón 5. Arriba hay tres renglones de título "
             "(sistema, periodo, fecha y hora de corte) y uno vacío.\n"
             "2. Filas de sección: el sistema agrupa por campus (en orden alfabético) con un "
             "renglón `CAMPUS: NOMBRE` antes de cada grupo, un renglón `Subtotal Campus: N "
             "registros` después y un `Total general` al final. Los subtotales cuentan "
             "renglones, no peticiones.\n"
             "3. Columnas con espacios de más: ` Campus`, `Clave materia `, `Nombre  materia`, "
             "`Matrícula alumno `, `Fecha de solicitud `, `Estatus de la petición  `, ` Canal`.\n"
             "4. `Estatus de la petición` con variantes de mayúsculas y espacios (`NUEVA`, "
             "`en proceso`, `En Proceso`, `En  proceso`, `Resuelta ` y así).\n"
             "5. `Campus` escrito de dos formas en algunos renglones: en MAYÚSCULAS o con un "
             "espacio al final (como en los CSV de la S5).\n"
             "6. Duplicados exactos: renglones repetidos tal cual, uno debajo del otro.\n"
             "7. Co-titulares: en los grupos con dos profesores, cada petición aparece dos "
             "veces, una por profesor; cambia la nómina y el nombre del profesor, lo demás "
             "es igual.\n"
             "8. `Fecha de solicitud` en texto con cuatro formatos: `14/10/2026`, "
             "`2026-10-14`, `14-oct-2026` y `14/10/2026 09:32`. Siempre día antes que mes.\n"
             "9. `Calificación parcial 1` y `Asistencia %` vienen vacías (el periodo no ha "
             "empezado). `Estatus inscripción` se deriva del estatus de la petición: "
             "Resuelta es Inscrito, Cancelada es Baja, Nueva y En proceso son Pendiente.\n")
    s.append("## Definición de la tabla limpia\n")
    s.append("- La unidad es la **petición**: un alumno pidiendo un grupo. Se identifica por "
             "`Matrícula alumno` + `CRN` (el reporte no trae folio).\n"
             "- Pasos: encontrar el encabezado; quedarse con los renglones que tienen "
             "matrícula y CRN; quitar duplicados exactos; unir los renglones de co-titular "
             "por matrícula + CRN; normalizar `Campus` y `Estatus de la petición` (sin "
             "espacios de más, sin distinguir mayúsculas); convertir las fechas.\n"
             "- La tabla cuenta peticiones únicas por campus y por estatus de la petición "
             "(Nueva, En proceso, Resuelta, Cancelada).\n"
             "- Corte anterior: el lunes o miércoles más reciente antes de la fecha de corte. "
             "\"Peticiones desde el corte anterior\" son las que tienen fecha de solicitud "
             "igual o posterior a ese día (el corte es a las 07:00 y el reporte llega hasta "
             "el día anterior).\n")
    s.append("## Respuesta correcta por entrega\n")
    totales = []
    for nombre, st in entregas:
        texto, tot = bloque_entrega(nombre, st)
        s.append(texto)
        totales.append((st, tot))
    (lun, _), (mie, _) = totales[0], totales[1]
    cambios = {}
    for k, e in lun["estatus_de"].items():
        e2 = mie["estatus_de"][k]
        if e2 != e:
            cambios[(e, e2)] = cambios.get((e, e2), 0) + 1
    s.append("## Del lunes al miércoles\n")
    s.append("- Peticiones nuevas en la entrega del miércoles: %s (fecha de solicitud 19 y "
             "20 de octubre)." % miles(mie["peticiones"] - lun["peticiones"]))
    s.append("- Peticiones del lunes que cambiaron de estatus: %s." % miles(sum(cambios.values())))
    for (a, b) in sorted(cambios, key=lambda x: (ESTATUS.index(x[0]), ESTATUS.index(x[1]))):
        s.append("  - %s a %s: %s" % (a, b, miles(cambios[(a, b)])))
    s.append("")
    s.append("## Cómo leer una respuesta equivocada\n")
    for nombre, st in entregas:
        s.append("- %s: si el total es %s, contó renglones de datos o tomó el `Total "
                 "general` del sistema (no quitó duplicados ni unió co-titulares); si es %s, "
                 "quitó duplicados exactos pero no unió los renglones de co-titular; la "
                 "respuesta correcta es %s." % (
                     DIAS[st["corte"].weekday()].capitalize(), miles(st["datos"]),
                     miles(st["datos"] - st["duplicados"]), miles(st["peticiones"])))
    s.append("- Si la tabla trae más de seis campus o más de cuatro estatus, no normalizó "
             "mayúsculas ni espacios.\n"
             "- Si \"desde el corte anterior\" no coincide, revisar cómo leyó las fechas: "
             "leer `05/10/2026` como mes/día (5 de mayo) o descartar los formatos con texto "
             "o con hora cambia la cuenta.\n"
             "- Si contó por la columna `Estatus inscripción` y no por `Estatus de la "
             "petición`, verá tres categorías (Pendiente, Inscrito, Baja) en vez de cuatro.\n")
    Path(ruta).write_text("\n".join(s), encoding="utf-8")


# ------------------------------------------------------------------ principal
def main():
    ap = argparse.ArgumentParser(description="Genera el reporte de inscritos ficticio (S6).")
    ap.add_argument("--semilla", type=int, default=42)
    ap.add_argument("--salida", default="reportes")
    ap.add_argument("--formato", choices=["xlsx", "csv", "ambos"], default="xlsx")
    ap.add_argument("--notas", default=None)
    args = ap.parse_args()

    rng = random.Random(args.semilla)
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)
    materias = construir_materias(rng)
    profs = construir_profesores(rng)
    grupos = construir_grupos(rng, materias, profs)
    alumnos = construir_alumnos(rng)
    peticiones = construir_peticiones(rng, alumnos, grupos, materias)

    entregas = []
    for corte in CORTES:
        filas, st = armar_entrega(rng, corte, peticiones, alumnos, grupos, materias, profs)
        base = "reporte-inscritos-%s" % corte.isoformat()
        nombre = base + (".csv" if args.formato == "csv" else ".xlsx")
        if args.formato in ("xlsx", "ambos"):
            escribir_xlsx(out / (base + ".xlsx"), "Reporte", filas, negritas=(0, 4))
        if args.formato in ("csv", "ambos"):
            escribir_csv(out / (base + ".csv"), filas)
        entregas.append((nombre, st))
        print("%s: %s renglones en la hoja, %s de datos, %s peticiones únicas" % (
            nombre, miles(st["filas_hoja"]), miles(st["datos"]), miles(st["peticiones"])))

    notas = Path(args.notas) if args.notas else out / "NOTAS-generacion.md"
    escribir_notas(notas, args.semilla, args.formato, entregas)
    print("Notas: %s" % notas.resolve())
    print("Salida: %s" % out.resolve())


if __name__ == "__main__":
    sys.exit(main())
