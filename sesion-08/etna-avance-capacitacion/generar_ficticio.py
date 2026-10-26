#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera la copia FICTICIA de los dos archivos del flujo de Etna (curso Ed
Digital, propuesta de Jorge, ronda 2 despues de la Sesion 7): la programacion diaria y el
historico de capacitacion. Siembra casos conocidos y los anota en NOTAS-ficticio.txt, con la
respuesta correcta.

Como correrlo (solo biblioteca estandar), desde la carpeta etna-avance-capacitacion:
    python3 -B generar_ficticio.py
    python3 -B generar_ficticio.py --semilla 22 --salida ficticio

Salen, en ficticio/:
    programacion-diaria-ficticia.xlsx     ENCABEZADOS SUPUESTOS (Etna no los ha compartido):
                                          Periodo, CRN, Clave UDF, UDF, Grupo, Modalidad,
                                          Idioma, Plan, Nómina, Nombre del profesor, Días,
                                          Horario y Zona horaria. Cada grupo viene repetido,
                                          uno por horario (TCM, TCH, TCP), como contó Etna
    historico-capacitacion-ficticio.xlsx  los encabezados REALES que Etna escribio en el muro,
                                          con la columna sin nombre que trae su lista
    NOTAS-ficticio.txt                    lo sembrado y las tres lineas que debe dar
                                          avance_capacitacion.py
Todo es inventado: ningun nombre ni nomina corresponde a una persona real. La columna Correo
del historico va vacia a proposito.
"""
import argparse
import random
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

# ------------------------------------------------------------------ escritura xlsx
# Las funciones de escritura de limpiar_reporte.py (Sesion 6).
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
            z.writestr(nombre, contenido.encode("utf-8"))

NOMBRES = ["Ana", "Luis", "María", "José", "Carmen", "Jorge", "Laura", "Ricardo", "Patricia",
           "Fernando", "Gabriela", "Héctor", "Sofía", "Manuel", "Elena", "Raúl", "Diana",
           "Arturo", "Mónica", "Javier", "Verónica", "Óscar", "Claudia", "Andrés", "Lucía",
           "Roberto", "Adriana", "Sergio", "Paola", "Emilio"]
APELLIDOS = ["García", "Hernández", "Martínez", "López", "González", "Pérez", "Rodríguez",
             "Sánchez", "Ramírez", "Cruz", "Flores", "Gómez", "Morales", "Vázquez", "Reyes",
             "Jiménez", "Torres", "Díaz", "Gutiérrez", "Ruiz", "Mendoza", "Aguilar", "Ortiz",
             "Castillo", "Romero", "Herrera", "Medina", "Navarro", "Salinas", "Treviño"]
UDFS = [  # clave, nombre, modalidad, idioma, plan (todo inventado)
    ("TE2001", "Pensamiento crítico aplicado", "Digital", "Español", "2019"),
    ("TE2002", "Comunicación profesional", "Digital", "Español", "2019"),
    ("TE2003", "Ética y ciudadanía", "Digital", "Español", "2019"),
    ("TE2004", "Business English", "Digital", "Inglés", "2019"),
    ("TE2005", "Estadística para la toma de decisiones", "Digital", "Español", "2019"),
    ("TE2006", "Taller de liderazgo", "Presencial", "Español", "2019"),
    ("TE2007", "Fundamentos de programación", "Digital", "Español", "2016"),
    ("TE2008", "Academic Writing", "Digital", "Inglés", "2016"),
    ("TE2009", "Historia del arte contemporáneo", "Presencial", "Español", "2016"),
    ("TE2010", "Innovación y emprendimiento", "Digital", "Español", "2019"),
    ("TE2011", "Cálculo aplicado", "Digital", "Español", "2016"),
    ("TE2012", "Diseño de experiencias", "Presencial", "Español", "2019"),
]
ZONAS = [("TCM", 0), ("TCH", -1), ("TCP", -2)]
HORAS = ["07:00", "08:30", "10:00", "11:30", "13:00", "14:30", "16:00", "17:30", "19:00"]
DIAS = ["Lu Ju", "Ma Vi", "Mi", "Lu Mi", "Ma Ju"]

ENC_PROG = ["Periodo", "CRN", "Clave UDF", "UDF", "Grupo", "Modalidad", "Idioma", "Plan",
            "Nómina", "Nombre del profesor", "Días", "Horario", "Zona horaria"]
# Los encabezados como los escribio Etna en el muro (con la columna sin nombre y el doble
# espacio de Reglamento).
ENC_HIST = ["Nómina", "Nombre", "Correo", "Tipo Prof", "Período del curso", "Plan que imparte",
            "Escuela", "Periodo Impartición", "Sensibilización a la microenseñanza",
            "Canvas: Conceptos básicos", "", "Fundamentos del Modelo Tec21",
            "Evaluación Competencias TEC21", "Estatus Tec21",
            "Nuevo plan de evaluación en Canvas y visor de calificaciones",
            "Modelo educativo Prepa Tec21",
            "Microenseñanza y estímulos de aprendizaje (8h) IT_VEM_EI02AD_PD_01_1282",
            "Reglamento General de Estudiantes  (1h) II.8.2.0.8", "Estatus Prepa",
            "Aula Virtual", "AA mediante HD", "Turnitin", "Respondus Monitor", "Estatus ED",
            "Inglés Acreditado", "Director/Coordinador", "Programa", "Fecha última Actualización"]
# Lo mismo que la tabla CURSOS del script (SUPUESTO): curso y cuando se pide
CURSOS = [("Sensibilización a la microenseñanza", "plan 2019"),
          ("Canvas: Conceptos básicos", "todos"),
          ("Fundamentos del Modelo Tec21", "plan 2019"),
          ("Evaluación Competencias TEC21", "plan 2019"),
          ("Aula Virtual", "digital"), ("AA mediante HD", "digital"),
          ("Turnitin", "digital"), ("Respondus Monitor", "digital"),
          ("Inglés Acreditado", "ingles")]
PREPA = ["Nuevo plan de evaluación en Canvas y visor de calificaciones",
         "Modelo educativo Prepa Tec21",
         "Microenseñanza y estímulos de aprendizaje (8h) IT_VEM_EI02AD_PD_01_1282",
         "Reglamento General de Estudiantes  (1h) II.8.2.0.8"]

N_TCM, N_SOLO_FUERA, N_SIN_HIST, N_NO_PROG = 40, 2, 3, 20


def miles(n):
    return "{:,}".format(n)


def y_lista(partes):
    partes = list(partes)
    if len(partes) <= 1:
        return "".join(partes)
    return ", ".join(partes[:-1]) + " y " + partes[-1]


def construir(rng):
    nominas, profes = set(), []
    total = N_TCM + N_SOLO_FUERA + N_NO_PROG
    while len(profes) < total:
        n = "L0%07d" % rng.randrange(1000000, 9999999)
        if n in nominas:
            continue
        nominas.add(n)
        nombre = "%s %s %s" % (rng.choice(NOMBRES), rng.choice(APELLIDOS), rng.choice(APELLIDOS))
        profes.append(dict(nomina=n, nombre=nombre))
    tcm, solo_fuera, no_prog = (profes[:N_TCM], profes[N_TCM:N_TCM + N_SOLO_FUERA],
                                profes[N_TCM + N_SOLO_FUERA:])

    # Programacion: cada grupo en TCM y casi siempre repetido en TCH y TCP
    renglones, crn = [], 41000
    grupos_de = {}
    for i, p in enumerate(tcm + solo_fuera):
        k = 1 + (i % 3)
        p["udfs"] = []
        for _ in range(k):
            crn += rng.randrange(3, 40)
            u = rng.choice(UDFS)
            hi = rng.randrange(len(HORAS) - 1)
            p["udfs"].append(u)
            zonas = ZONAS if p in tcm else ZONAS[1:]
            if p in tcm and rng.random() < 0.25:
                zonas = [ZONAS[0], ZONAS[1]]
            for z, delta in zonas:
                h1 = int(HORAS[hi][:2]) + delta
                h2 = int(HORAS[hi + 1][:2]) + delta
                horario = "%02d:%s-%02d:%s" % (h1, HORAS[hi][3:], h2, HORAS[hi + 1][3:])
                renglones.append(["Ene-Jun 2027", crn, u[0], u[1], 1 + rng.randrange(4), u[2],
                                  u[3], u[4], p["nomina"], p["nombre"], DIAS[crn % len(DIAS)], horario, z])
        grupos_de[p["nomina"]] = k
    # Formato que no debe contar: un TCM escrito en minusculas y con espacio
    tcm_rows = [r for r in renglones if r[12] == "TCM"]
    tcm_rows[5][12] = " tcm "
    total_tcm_rows = len(tcm_rows)

    # Historico: 37 de los 40 TCM (3 sin historico), los 2 de fuera y 20 no programados
    sin_hist = tcm[-N_SIN_HIST:]
    en_hist = tcm[:-N_SIN_HIST] + solo_fuera + no_prog
    rng.shuffle(en_hist)

    def requisitos(p):
        req = {"todos"}
        for u in p.get("udfs", []):
            if u[4] == "2019":
                req.add("plan 2019")
            if u[2] != "Presencial":
                req.add("digital")
            if u[3] == "Inglés":
                req.add("ingles")
        return req

    def valor_cursado():
        x = rng.random()
        if x < 0.7:
            return "Acreditado"
        if x < 0.85:
            return "%02d/%02d/2026" % (rng.randrange(1, 29), rng.randrange(1, 13))
        return 46000 + rng.randrange(300)        # fecha guardada como numero de Excel

    def valor_falta():
        return rng.choice(["", "", "", "Pendiente", "En curso"])

    # Casos forzados
    forzado_100, forzado_0 = tcm[0], tcm[1]
    nombre_distinto = [tcm[2], tcm[3]]
    no_reconocido = tcm[4]
    duplicado = tcm[6]
    nomina_formato = [tcm[7], tcm[8], tcm[9]]

    hist_rows = []
    for p in en_hist:
        req = requisitos(p)
        fila = {c: "" for c in ENC_HIST if c}
        fila["Nómina"] = p["nomina"]
        fila["Nombre"] = p["nombre"]
        fila["Tipo Prof"] = rng.choice(["Planta", "Cátedra"])
        fila["Período del curso"] = rng.choice(["Ago-Dic 2025", "Ene-Jun 2026", "Ago-Dic 2026"])
        fila["Plan que imparte"] = rng.choice(["2019", "2016"])
        fila["Escuela"] = rng.choice(["Escuela A", "Escuela B", "Escuela C", "Escuela D"])
        fila["Periodo Impartición"] = "Ene-Jun 2027"
        fila["Director/Coordinador"] = "Coordinación %s" % rng.choice("ABCDE")
        fila["Programa"] = rng.choice(["Programa 1", "Programa 2", "Programa 3"])
        fila["Fecha última Actualización"] = "21/10/2026"
        pide = [c for c, cond in CURSOS if cond in req]
        for c, cond in CURSOS:
            if p is forzado_100:
                ok = True
            elif p is forzado_0:
                ok = False
            else:
                ok = rng.random() < 0.6
            fila[c] = valor_cursado() if ok else valor_falta()
        if p is no_reconocido:
            fila[pide[0]] = "Revisar"            # el script no lo reconoce: cuenta como falta
        for c in PREPA:
            fila[c] = rng.choice(["Acreditado", "", "Pendiente"])
        for c in ("Estatus Tec21", "Estatus Prepa", "Estatus ED"):
            fila[c] = rng.choice(["Completo", "Incompleto"])
        if p in nombre_distinto:
            partes = p["nombre"].split()
            fila["Nombre"] = " ".join(partes[:2])          # un solo apellido
        if p in nomina_formato:
            fila["Nómina"] = " " + p["nomina"].lower() + " "
        hist_rows.append(fila)
        if p is duplicado:
            hist_rows.append(dict(fila))
    return dict(renglones=renglones, hist=hist_rows, tcm=tcm, solo_fuera=solo_fuera,
                sin_hist=sin_hist, no_prog=no_prog, requisitos=requisitos,
                total_tcm_rows=total_tcm_rows, casos=dict(
                    forzado_100=forzado_100, forzado_0=forzado_0,
                    nombre_distinto=nombre_distinto, no_reconocido=no_reconocido,
                    duplicado=duplicado, nomina_formato=nomina_formato))


def principal():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia del flujo de Etna.")
    ap.add_argument("--semilla", type=int, default=22, help="fija el azar (22)")
    ap.add_argument("--salida", default="ficticio", help="carpeta de salida (ficticio)")
    args = ap.parse_args()
    rng = random.Random(args.semilla)
    d = construir(rng)
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)
    np_, nh = "programacion-diaria-ficticia.xlsx", "historico-capacitacion-ficticio.xlsx"

    filas_p = [list(ENC_PROG)] + d["renglones"]
    escribir_xlsx(out / np_, [("Programación", filas_p, [12, 8, 10, 36, 6, 11, 9, 6, 12, 30, 8,
                                                         13, 12], True)])
    filas_h = [["Histórico de capacitación (copia ficticia)"], [], list(ENC_HIST)]
    filas_h += [[f.get(c, "") if c else "" for c in ENC_HIST] for f in d["hist"]]
    escribir_xlsx(out / nh, [("Histórico", filas_h, [12, 30] + [16] * (len(ENC_HIST) - 2),
                              False)])

    # Respuesta correcta, desde lo sembrado
    pcts = []
    faltan = {c: 0 for c, _ in CURSOS}
    completos = medios = bajos = 0
    for p in d["tcm"]:
        req = d["requisitos"](p)
        pide = [c for c, cond in CURSOS if cond in req]
        if p in d["sin_hist"]:
            for c in pide:
                faltan[c] += 1
            continue
        fila = next(f for f in d["hist"] if f["Nómina"].strip().upper() == p["nomina"])
        cursados = 0
        for c in pide:
            v = fila[c]
            ok = (v == "Acreditado" or isinstance(v, int) or (isinstance(v, str) and "/" in v))
            if ok:
                cursados += 1
            else:
                faltan[c] += 1
        pct = round(100.0 * cursados / len(pide), 1)
        pcts.append(pct)
        if pct >= 100:
            completos += 1
        elif pct >= 50:
            medios += 1
        else:
            bajos += 1
    promedio = round(sum(pcts) / len(pcts), 1)
    orden = [c for c, _ in CURSOS]
    top = sorted([c for c in faltan if faltan[c]], key=lambda c: (-faltan[c], orden.index(c)))

    def pct_txt(x):
        return ("%.1f" % x).rstrip("0").rstrip(".") + " %"

    lineas = [
        "Avance de capacitación de %s profesores programados en horario TCM (%s de %s renglones "
        "de %s), cruzado con %s por nómina." % (miles(len(d["tcm"])), miles(d["total_tcm_rows"]),
                                               miles(len(d["renglones"])), np_, nh),
        "Promedio de avance: %s. Con 100 %%: %d; de 50 a 99 %%: %d; menos de 50 %%: %d; sin "
        "histórico: %d." % (pct_txt(promedio), completos, medios, bajos, len(d["sin_hist"])),
        "Cursos que más faltan: %s. Detalle: avance-capacitacion-%s.xlsx."
        % (y_lista("%s (%d)" % (c, faltan[c]) for c in top[:3]), np_[:-5]),
    ]
    c = d["casos"]
    notas = [
        "NOTAS DE LA COPIA FICTICIA · avance de capacitación (propuesta para el flujo de Etna)",
        "Todos los datos son inventados por generar_ficticio.py. Ningún nombre ni nómina",
        "corresponde a una persona real. La columna Correo del histórico va vacía a propósito.",
        "Las cifras salen de lo que se sembró, no de leer los archivos; avance_capacitacion.py",
        "las debe reproducir por su lado.",
        "",
        "Semilla: %d" % args.semilla,
        "Comando: python3 -B generar_ficticio.py --semilla %d --salida %s" % (args.semilla,
                                                                            args.salida),
        "",
        "ARCHIVOS",
        "  %s: %s renglones; ENCABEZADOS SUPUESTOS (%s)." % (np_, miles(len(d["renglones"])),
                                                           ", ".join(ENC_PROG)),
        "    Cada grupo viene repetido por horario: TCM siempre; TCH casi siempre; TCP casi",
        "    siempre. Solo cambian Horario y Zona horaria.",
        "  %s: %d renglones; encabezados REALES del muro, con la columna sin nombre;" % (
            nh, len(d["hist"])),
        "    dos renglones de título arriba del encabezado (encabezado en la fila 3).",
        "",
        "LO QUE SE SEMBRÓ",
        "  Profesores con grupo en TCM: %d (renglones TCM: %d)." % (len(d["tcm"]),
                                                                  d["total_tcm_rows"]),
        "  Profesores que solo aparecen en TCH o TCP (no entran a la lista): %d: %s." % (
            len(d["solo_fuera"]), y_lista(p["nomina"] for p in d["solo_fuera"])),
        "  Programados en TCM sin histórico: %d: %s." % (len(d["sin_hist"]), y_lista(
            p["nomina"] for p in d["sin_hist"])),
        "  En el histórico sin estar programados (no salen): %d." % len(d["no_prog"]),
        "  Profesor con todos sus cursos: %s (100 %%)." % c["forzado_100"]["nomina"],
        "  Profesor sin ningún curso: %s (0 %%)." % c["forzado_0"]["nomina"],
        "  Nombre con un solo apellido en el histórico: %s." % y_lista(
            p["nomina"] for p in c["nombre_distinto"]),
        "  Valor 'Revisar' que el script no reconoce (cuenta como falta): %s." %
        c["no_reconocido"]["nomina"],
        "  Nómina repetida en el histórico (renglón idéntico): %s." % c["duplicado"]["nomina"],
        "",
        "LO QUE NO DEBE CONTAR (solo formato)",
        "  Un renglón TCM escrito ' tcm ' (minúsculas y espacios): sí es TCM.",
        "  Nóminas del histórico en minúsculas y con espacios: %s." % y_lista(
            p["nomina"] for p in c["nomina_formato"]),
        "  Cursado escrito de tres formas: Acreditado, una fecha 15/08/2026 o una fecha de Excel.",
        "  Las columnas de Prepa y de Estatus traen valores, pero no cuentan para el %.",
        "",
        "LAS TRES LÍNEAS QUE DEBE DAR avance_capacitacion.py",
    ] + ["  " + l for l in lineas]
    (out / "NOTAS-ficticio.txt").write_text("\n".join(notas) + "\n", encoding="utf-8")
    print("Listo. En %s/:" % out)
    print("  %s, %s y NOTAS-ficticio.txt" % (np_, nh))
    print("Lo que debe salir:")
    for l in lineas:
        print("  " + l)
    return 0


if __name__ == "__main__":
    sys.exit(principal())
