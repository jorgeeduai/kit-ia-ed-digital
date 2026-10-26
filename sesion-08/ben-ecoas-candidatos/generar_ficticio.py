#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera la copia FICTICIA de los dos archivos del flujo de Ben (curso Ed
Digital, propuesta de Jorge, ronda 2 despues de la Sesion 7): el archivo de ECOAs de tutores
y el archivo que se mando antes a escolar. Siembra casos conocidos y los anota en
NOTAS-ficticio.txt, con la respuesta correcta.

Como correrlo (solo biblioteca estandar), desde la carpeta ben-ecoas-candidatos:
    python3 -B generar_ficticio.py
    python3 -B generar_ficticio.py --semilla 22 --salida ficticio

Salen, en ficticio/:
    ecoas-tutores-ficticio.xlsx       los encabezados REALES que Ben escribio en el muro: CRN,
                                      Clave, Grupo, Alumnos Candidatos, Nómina del Profesor,
                                      Nombre del profesor. Un renglon por tutor de cada CRN; si
                                      el CRN tiene varios tutores, cada uno trae la suma del CRN
    enviado-a-escolar-ficticio.xlsx   ENCABEZADOS SUPUESTOS (Ben no los ha compartido): CRN,
                                      Matrícula, Nómina del Profesor, Nombre del profesor. Un
                                      renglon por alumno, con el tutor que le toca
    NOTAS-ficticio.txt                lo sembrado y las tres lineas que debe dar
                                      corregir_candidatos.py
Todo es inventado: ningun nombre, nomina ni matricula corresponde a una persona real. Las
matriculas empiezan con FIC para que no se confundan con una real.
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
           "Arturo", "Mónica", "Javier", "Verónica", "Óscar", "Claudia", "Andrés", "Lucía"]
APELLIDOS = ["García", "Hernández", "Martínez", "López", "González", "Pérez", "Rodríguez",
             "Sánchez", "Ramírez", "Cruz", "Flores", "Gómez", "Morales", "Vázquez", "Reyes",
             "Jiménez", "Torres", "Díaz", "Gutiérrez", "Ruiz", "Mendoza", "Aguilar"]
MATERIAS = ["TU1001", "TU1002", "TU1003", "TU2001", "TU2002", "TU3001"]
TUTORES_POR_CRN = [1, 1, 1, 1, 2, 2, 3, 3, 4, 5, 7]
N_CRN = 40

ENC_ECOAS = ["CRN", "Clave", "Grupo", "Alumnos Candidatos", "Nómina del Profesor",
             "Nombre del profesor"]
ENC_ESCOLAR = ["CRN", "Matrícula", "Nómina del Profesor", "Nombre del profesor"]


def miles(n):
    return "{:,}".format(n)


def construir(rng):
    usadas = set()

    def tutor():
        while True:
            n = "L0%07d" % rng.randrange(1000000, 9999999)
            if n not in usadas:
                usadas.add(n)
                return dict(nomina=n, nombre="%s %s %s" % (
                    rng.choice(NOMBRES), rng.choice(APELLIDOS), rng.choice(APELLIDOS)))

    crns, crn = [], 60000
    for i in range(N_CRN):
        crn += rng.randrange(5, 60)
        k = rng.choice(TUTORES_POR_CRN)
        ts = [dict(tutor(), alumnos=rng.randrange(4, 26)) for _ in range(k)]
        crns.append(dict(crn=crn, clave=rng.choice(MATERIAS), grupo=1 + rng.randrange(6),
                         tutores=ts))
    multi = [c for c in crns if len(c["tutores"]) > 1]
    uno = [c for c in crns if len(c["tutores"]) == 1]

    # Casos sembrados
    ya_bien = multi[:2]                 # CRN con varios tutores que ya vienen bien
    otro_numero = uno[0]                # un tutor solo con otro numero (+3)
    sin_escolar = uno[1]                # el tutor no esta en el archivo de escolar
    extra = uno[2]                      # en escolar hay un tutor mas que no esta en ECOAs
    extra_tutor = dict(tutor(), alumnos=6)
    repetidos = [uno[3], multi[2]]      # un alumno repetido en escolar en estos CRN
    formato = [multi[3]]                # nominas en minusculas y con espacios en escolar

    ecoas = []
    esperado = dict(suma=0, otro=0, bien=0, revisar=0, solo_escolar=1)
    antes = despues = 0
    for c in crns:
        total = sum(t["alumnos"] for t in c["tutores"])
        for t in c["tutores"]:
            if c in ya_bien or len(c["tutores"]) == 1:
                valor = t["alumnos"]
            else:
                valor = total                         # el error que describio Ben
            if c is otro_numero:
                valor = t["alumnos"] + 3
            ecoas.append([c["crn"], c["clave"], c["grupo"], valor, t["nomina"], t["nombre"]])
            antes += valor
            if c is sin_escolar:
                esperado["revisar"] += 1
                despues += valor
            elif valor == t["alumnos"]:
                esperado["bien"] += 1
                despues += valor
            elif c is otro_numero:
                esperado["otro"] += 1
                despues += t["alumnos"]
            else:
                esperado["suma"] += 1
                despues += t["alumnos"]
    rng.shuffle(ecoas)
    ecoas.sort(key=lambda f: f[0])

    escolar, folio = [], 10000
    en_escolar = 0
    for c in crns:
        ts = list(c["tutores"]) + ([extra_tutor] if c is extra else [])
        for t in ts:
            if c is sin_escolar:
                continue
            for _ in range(t["alumnos"]):
                folio += rng.randrange(1, 9)
                nom = t["nomina"]
                if c in formato:
                    nom = " " + nom.lower()
                escolar.append([str(c["crn"]), "FIC%06d" % folio, nom, t["nombre"]])
                en_escolar += 1
            if c in repetidos and t is ts[0]:
                escolar.append(list(escolar[-1]))     # el mismo alumno dos veces
    rng.shuffle(escolar)
    escolar.sort(key=lambda f: (f[0], f[2].strip().upper()))
    return dict(crns=crns, ecoas=ecoas, escolar=escolar, esperado=esperado, antes=antes,
                despues=despues, en_escolar=en_escolar, casos=dict(
                    ya_bien=ya_bien, otro_numero=otro_numero, sin_escolar=sin_escolar,
                    extra=extra, extra_tutor=extra_tutor, repetidos=repetidos,
                    formato=formato))


def principal():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia del flujo de Ben.")
    ap.add_argument("--semilla", type=int, default=22, help="fija el azar (22)")
    ap.add_argument("--salida", default="ficticio", help="carpeta de salida (ficticio)")
    args = ap.parse_args()
    d = construir(random.Random(args.semilla))
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)
    ne, ns = "ecoas-tutores-ficticio.xlsx", "enviado-a-escolar-ficticio.xlsx"
    escribir_xlsx(out / ne, [("ECOAs tutores", [ENC_ECOAS] + d["ecoas"],
                              [8, 9, 7, 18, 20, 30], True)])
    escribir_xlsx(out / ns, [("Enviado a escolar", [["Alumnos por tutor enviados a escolar "
                                                     "(copia ficticia)"], [], ENC_ESCOLAR]
                              + d["escolar"], [8, 12, 20, 30], False)])

    e, c = d["esperado"], d["casos"]
    corregidos = e["suma"] + e["otro"]
    por_revisar = e["revisar"] + e["solo_escolar"]
    lineas = [
        "Corrección de Alumnos Candidatos de %s con %s, por CRN y Nómina del Profesor." % (ne, ns),
        "De %s tutores en %s CRN: %s corregidos (%s traían la suma del CRN y %s otro número), %s "
        "ya estaban bien y %s por revisar." % (miles(len(d["ecoas"])), miles(len(d["crns"])),
                                               miles(corregidos), miles(e["suma"]),
                                               miles(e["otro"]), miles(e["bien"]),
                                               miles(por_revisar)),
        "Alumnos candidatos en total: %s antes y %s después; en el archivo de escolar, %s. "
        "Detalle: candidatos-corregidos-%s.xlsx." % (miles(d["antes"]), miles(d["despues"]),
                                                     miles(d["en_escolar"]), ne[:-5]),
    ]
    multi = sum(1 for x in d["crns"] if len(x["tutores"]) > 1)
    notas = [
        "NOTAS DE LA COPIA FICTICIA · alumnos candidatos de las ECOAs de tutores (flujo de Ben)",
        "Todos los datos son inventados por generar_ficticio.py. Ningún nombre, nómina ni",
        "matrícula corresponde a una persona real. Las cifras salen de lo que se sembró, no de",
        "leer los archivos; corregir_candidatos.py las debe reproducir por su lado.",
        "",
        "Semilla: %d" % args.semilla,
        "Comando: python3 -B generar_ficticio.py --semilla %d --salida %s" % (args.semilla,
                                                                            args.salida),
        "",
        "ARCHIVOS",
        "  %s: %d tutores en %d CRN (%d CRN con más de un tutor). Encabezados REALES" % (
            ne, len(d["ecoas"]), len(d["crns"]), multi),
        "    del muro; encabezado en la fila 1.",
        "  %s: %s renglones, uno por alumno. ENCABEZADOS SUPUESTOS (%s);" % (
            ns, miles(len(d["escolar"])), ", ".join(ENC_ESCOLAR)),
        "    dos renglones de título arriba del encabezado (encabezado en la fila 3).",
        "",
        "LO QUE SE SEMBRÓ",
        "  En cada CRN con varios tutores, cada tutor trae la suma del CRN (el error que contó",
        "  Ben), salvo en dos CRN que ya vienen bien: %s." % " y ".join(
            str(x["crn"]) for x in c["ya_bien"]),
        "  CRN %d: un solo tutor con 3 alumnos de más en ECOAs (otro número)." %
        c["otro_numero"]["crn"],
        "  CRN %d: el tutor no está en el archivo de escolar (por revisar, no se toca)." %
        c["sin_escolar"]["crn"],
        "  CRN %d: en escolar hay un tutor más, %s, que no está en ECOAs (por revisar)." % (
            c["extra"]["crn"], c["extra_tutor"]["nomina"]),
        "  CRN %s: un alumno repetido en escolar (se cuenta una vez)." % " y ".join(
            str(x["crn"]) for x in c["repetidos"]),
        "",
        "LO QUE NO DEBE CONTAR (solo formato)",
        "  CRN como texto en escolar y como número en ECOAs.",
        "  CRN %s: nóminas en minúsculas y con un espacio adelante en escolar." % " y ".join(
            str(x["crn"]) for x in c["formato"]),
        "",
        "LO ESPERADO",
        "  Corregidos porque traían la suma del CRN: %d" % e["suma"],
        "  Corregidos porque traían otro número: %d" % e["otro"],
        "  Ya estaban bien: %d" % e["bien"],
        "  Por revisar: %d (1 tutor sin alumnos en escolar y 1 tutor solo en escolar)" %
        por_revisar,
        "",
        "LAS TRES LÍNEAS QUE DEBE DAR corregir_candidatos.py",
    ] + ["  " + l for l in lineas]
    (out / "NOTAS-ficticio.txt").write_text("\n".join(notas) + "\n", encoding="utf-8")
    print("Listo. En %s/:" % out)
    print("  %s, %s y NOTAS-ficticio.txt" % (ne, ns))
    print("Lo que debe salir:")
    for l in lineas:
        print("  " + l)
    return 0


if __name__ == "__main__":
    sys.exit(principal())
