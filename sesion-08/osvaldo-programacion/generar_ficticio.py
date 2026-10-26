#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera la copia FICTICIA de los dos archivos de la propuesta para
Osvaldo (curso Ed Digital, ronda 2 despues de la Sesion 7): la programacion de un periodo y
los resultados de ECOAs del mismo periodo. Siembra inconsistencias conocidas y las anota en
NOTAS-ficticio.txt, con la respuesta correcta.

TODO ES SUPUESTO: Osvaldo no ha escrito su flujo ni sus encabezados. Se parte de lo que dijo
en la Sesion 5 (ECOAs e inconsistencias de programacion).

Como correrlo (solo biblioteca estandar), desde la carpeta osvaldo-programacion:
    python3 -B generar_ficticio.py
    python3 -B generar_ficticio.py --semilla 22 --salida ficticio

Salen, en ficticio/:
    programacion-periodo-ficticia.xlsx   un renglon por profesor de cada grupo (CRN)
    resultados-ecoas-ficticio.xlsx       un renglon por profesor de cada grupo con ECOA
    NOTAS-ficticio.txt                   lo sembrado y las tres lineas que debe dar
                                         revisar_programacion.py
Todo es inventado: ningun nombre ni nomina corresponde a una persona real.
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
MATERIAS = [("HU1001", "Pensamiento crítico"), ("HU1002", "Ética y ciudadanía"),
            ("HU2003", "Comunicación oral y escrita"), ("HU2004", "Historia de México"),
            ("NE1005", "Fundamentos de negocios"), ("NE2006", "Finanzas personales"),
            ("IN1007", "Programación básica"), ("IN2008", "Bases de datos"),
            ("CS1009", "Estadística"), ("CS2010", "Cálculo diferencial")]
MODALIDADES = ["Presencial", "Digital", "Híbrida", "Flexible digital"]
PATRONES = ["Lu Ju", "Ma Vi", "Mi"]
INICIOS = ["07:00", "08:30", "10:00", "11:30", "13:00", "14:30", "16:00", "17:30", "19:00"]
ENC_PROG = ["Periodo", "CRN", "Clave", "Nombre de la materia", "Grupo", "Modalidad",
            "Nómina del profesor", "Nombre del profesor", "% de responsabilidad", "Días",
            "Horario inicio", "Horario fin"]
ENC_ECOAS = ["Periodo", "CRN", "Clave", "Nómina del profesor", "Nombre del profesor",
             "Alumnos candidatos", "Respuestas", "Promedio"]
N_CRN, N_PROF = 70, 32
PERIODO = "Ago-Dic 2026"


def miles(n):
    return "{:,}".format(n)


def y_lista(partes):
    partes = list(partes)
    if len(partes) <= 1:
        return "".join(partes)
    return ", ".join(partes[:-1]) + " y " + partes[-1]


def fin(inicio):
    h, m = int(inicio[:2]), int(inicio[3:])
    t = h * 60 + m + 90
    return "%02d:%02d" % divmod(t, 60)


def construir(rng):
    usadas, profes = set(), []
    while len(profes) < N_PROF:
        n = "L0%07d" % rng.randrange(1000000, 9999999)
        if n in usadas:
            continue
        usadas.add(n)
        profes.append(dict(nomina=n, nombre="%s %s %s" % (rng.choice(NOMBRES),
                                                         rng.choice(APELLIDOS),
                                                         rng.choice(APELLIDOS)),
                           ocupados=set(), crns=[]))
    slots = [(p, h) for p in PATRONES for h in INICIOS]
    grupos, crn = [], 30000
    while len(grupos) < N_CRN:
        k = 2 if rng.random() < 0.2 else 1
        ps = rng.sample(profes, k)
        if any(len(p["crns"]) >= 4 for p in ps):
            continue
        libres = [s for s in slots if all(s not in p["ocupados"] for p in ps)]
        s = rng.choice(libres)
        crn += rng.randrange(3, 30)
        clave, materia = rng.choice(MATERIAS)
        pct = [100] if k == 1 else rng.choice([[50, 50], [60, 40]])
        g = dict(crn=crn, clave=clave, materia=materia, grupo=1 + rng.randrange(5),
                 modalidad=rng.choice(MODALIDADES), dias=s[0], inicio=s[1], fin=fin(s[1]),
                 profes=[dict(nomina=p["nomina"], nombre=p["nombre"], pct=pc, prof=p)
                         for p, pc in zip(ps, pct)])
        for p in ps:
            p["ocupados"].add(s)
            p["crns"].append(g)
        grupos.append(g)
    return profes, grupos


def renglones(grupos, rng):
    filas = []
    for g in grupos:
        for i, p in enumerate(g["profes"]):
            pct = p["pct"]
            pct_v = "%d%%" % pct if rng.random() < 0.5 else pct          # 50% o 50
            ini = g["inicio"]
            if rng.random() < 0.15:
                ini = ini.replace(":", "")                              # 0700: solo formato
            filas.append([PERIODO, g["crn"], g["clave"], g.get("materia_" + str(i), g["materia"]),
                          g.get("grupo_" + str(i), g["grupo"]), g["modalidad"],
                          p.get("nomina_txt", p["nomina"]), p.get("nombre_txt", p["nombre"]),
                          p.get("pct_txt", pct_v), g["dias"], ini, g["fin"]])
    return filas


def principal():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia de la propuesta para "
                                             "Osvaldo.")
    ap.add_argument("--semilla", type=int, default=22, help="fija el azar (22)")
    ap.add_argument("--salida", default="ficticio", help="carpeta de salida (ficticio)")
    args = ap.parse_args()
    rng = random.Random(args.semilla)
    profes, grupos = construir(rng)
    uno = [g for g in grupos if len(g["profes"]) == 1]
    dos = [g for g in grupos if len(g["profes"]) == 2]
    usados = set()

    def tomar(lista):
        g = next(x for x in lista if x["crn"] not in usados)
        usados.add(g["crn"])
        return g

    s = {}
    # Profesores para nombres y choques: todos sus grupos con un solo profesor, 2 o mas grupos
    candidatos = [p for p in profes if len(p["crns"]) >= 2
                  and all(len(g["profes"]) == 1 for g in p["crns"])]
    s["nombres"] = candidatos[0]
    s["choques"] = candidatos[1:3]
    for p in candidatos[:3]:
        usados.update(g["crn"] for g in p["crns"])
    s["repetido"] = tomar(uno)
    s["grupo_materia"] = tomar(dos)
    s["grupo_grupo"] = tomar(dos)
    s["pct_dos"] = tomar(dos)
    s["pct_uno"] = tomar(uno)
    s["nomina_vacia"] = tomar(uno)
    s["nomina_formato"] = tomar(uno)
    s["mod_fuera"] = [tomar(uno), tomar(uno)]
    s["mod_formato"] = [tomar(uno), tomar(uno)]
    s["sin_ecoa"] = [tomar(uno), tomar(uno), tomar(uno)]

    # Aplicar lo sembrado
    g = s["grupo_materia"]
    g["materia_1"] = next(m for _, m in MATERIAS if m != g["materia"])
    g = s["grupo_grupo"]
    g["grupo_1"] = g["grupo"] % 5 + 1
    g = s["pct_dos"]
    g["profes"][0]["pct_txt"], g["profes"][1]["pct_txt"] = "60%", 50
    s["pct_uno"]["profes"][0]["pct_txt"] = 80
    s["nomina_vacia"]["profes"][0]["nomina_txt"] = ""
    s["nomina_formato"]["profes"][0]["nomina_txt"] = "12345678"
    p = s["nombres"]
    p["crns"][-1]["profes"][0]["nombre_txt"] = " ".join(p["nombre"].split()[:2])
    pares_choque = []
    for p in s["choques"]:
        a, b = p["crns"][0], p["crns"][1]
        b["dias"], b["inicio"], b["fin"] = a["dias"], a["inicio"], a["fin"]
        pares_choque.append((min(a["crn"], b["crn"]), max(a["crn"], b["crn"]), p["nomina"]))
    s["mod_fuera"][0]["modalidad"] = "Virtual"
    s["mod_fuera"][1]["modalidad"] = "Por definir"
    s["mod_formato"][0]["modalidad"] = s["mod_formato"][0]["modalidad"].lower()
    s["mod_formato"][1]["modalidad"] = s["mod_formato"][1]["modalidad"].upper()

    filas = renglones(grupos, rng)
    rep = next(i for i, f in enumerate(filas) if f[1] == s["repetido"]["crn"])
    filas.insert(rep + 1, list(filas[rep]))

    # ECOAs: un renglon por profesor de cada grupo, menos los tres sin ECOA, mas dos de
    # grupos que no estan en la programacion
    ecoas = []
    sin_ecoa = {(g["crn"], g["profes"][0]["nomina"]) for g in s["sin_ecoa"]}
    for g in grupos:
        for p in g["profes"]:
            if (g["crn"], p["nomina"]) in sin_ecoa:
                continue
            nom = p.get("nomina_txt") if p.get("nomina_txt") == "12345678" else p["nomina"]
            if rng.random() < 0.1:
                nom = nom.lower()
            cand = 10 + rng.randrange(30)
            ecoas.append([PERIODO, g["crn"], g["clave"], nom, p["nombre"], cand,
                          rng.randrange(3, cand), round(7 + 3 * rng.random(), 2)])
    extra = []
    for _ in range(2):
        crn = grupos[-1]["crn"] + 50 + len(extra) * 7
        pr = rng.choice(profes)
        extra.append((crn, pr["nomina"]))
        ecoas.append([PERIODO, crn, "HU9999", pr["nomina"], pr["nombre"], 20, 12, 9.1])
    ecoas.sort(key=lambda f: (f[1], f[3].upper()))

    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)
    np_, ne = "programacion-periodo-ficticia.xlsx", "resultados-ecoas-ficticio.xlsx"
    escribir_xlsx(out / np_, [("Programación", [ENC_PROG] + filas,
                               [12, 8, 9, 30, 6, 16, 18, 30, 12, 8, 10, 10], True)])
    escribir_xlsx(out / ne, [("ECOAs", [["Resultados de ECOAs %s (copia ficticia)" % PERIODO],
                                        [], ENC_ECOAS] + ecoas,
                              [12, 8, 9, 18, 30, 10, 10, 9], False)])

    # Respuesta correcta, desde lo sembrado
    casos = []   # (tipo, crn)
    casos.append(("Renglón repetido", s["repetido"]["crn"]))
    casos.append(("Datos del grupo distintos en el mismo CRN", s["grupo_materia"]["crn"]))
    casos.append(("Datos del grupo distintos en el mismo CRN", s["grupo_grupo"]["crn"]))
    casos.append(("% de responsabilidad que no suma 100", s["pct_dos"]["crn"]))
    casos.append(("% de responsabilidad que no suma 100", s["pct_uno"]["crn"]))
    casos.append(("Nómina vacía o con otro formato", s["nomina_vacia"]["crn"]))
    casos.append(("Nómina vacía o con otro formato", s["nomina_formato"]["crn"]))
    casos.append(("Una nómina con dos nombres", min(g["crn"] for g in s["nombres"]["crns"])))
    for a, b, n in pares_choque:
        casos.append(("Choque de horario del profesor", a))
    for g in s["mod_fuera"]:
        casos.append(("Modalidad fuera del catálogo", g["crn"]))
    for g in s["sin_ecoa"]:
        casos.append(("Programado sin ECOA", g["crn"]))
    casos.append(("ECOA sin programación", s["nomina_vacia"]["crn"]))
    for crn, _ in extra:
        casos.append(("ECOA sin programación", crn))
    tipos = ["Renglón repetido", "Datos del grupo distintos en el mismo CRN",
             "% de responsabilidad que no suma 100", "Nómina vacía o con otro formato",
             "Una nómina con dos nombres", "Choque de horario del profesor",
             "Modalidad fuera del catálogo", "Programado sin ECOA", "ECOA sin programación"]
    por_tipo = {t: sum(1 for c in casos if c[0] == t) for t in tipos}
    top = sorted([t for t in tipos if por_tipo[t]], key=lambda t: (-por_tipo[t], tipos.index(t)))
    en_ecoas = por_tipo["Programado sin ECOA"] + por_tipo["ECOA sin programación"]
    lineas = [
        "Revisión de %s (%s renglones, %s CRN) y su cruce con %s, por CRN y nómina del profesor."
        % (np_, miles(len(filas)), miles(len(grupos)), ne),
        "Hay %d inconsistencias en %d CRN: %d en la programación y %d en el cruce con ECOAs."
        % (len(casos), len({c[1] for c in casos}), len(casos) - en_ecoas, en_ecoas),
        "Tipos con más casos: %s. Detalle: inconsistencias-%s.xlsx."
        % (y_lista("%s (%d)" % (t, por_tipo[t]) for t in top[:3]), np_[:-5]),
    ]
    notas = [
        "NOTAS DE LA COPIA FICTICIA · inconsistencias de programación y ECOAs (propuesta para",
        "Osvaldo). TODO ES SUPUESTO: encabezados, tipos de inconsistencia y catálogo.",
        "Todos los datos son inventados por generar_ficticio.py. Ningún nombre ni nómina",
        "corresponde a una persona real. Las cifras salen de lo que se sembró, no de leer los",
        "archivos; revisar_programacion.py las debe reproducir por su lado.",
        "",
        "Semilla: %d" % args.semilla,
        "Comando: python3 -B generar_ficticio.py --semilla %d --salida %s" % (args.semilla,
                                                                            args.salida),
        "",
        "ARCHIVOS",
        "  %s: %d renglones, %d CRN (%d con dos profesores), %d profesores." % (
            np_, len(filas), len(grupos), len(dos), N_PROF),
        "    Encabezados supuestos: %s." % ", ".join(ENC_PROG),
        "  %s: %d renglones; dos renglones de título (encabezado en la fila 3)." % (
            ne, len(ecoas)),
        "    Encabezados supuestos: %s." % ", ".join(ENC_ECOAS),
        "",
        "LO QUE SE SEMBRÓ (CRN y qué)",
        "  %d · renglón repetido idéntico." % s["repetido"]["crn"],
        "  %d · el segundo profesor trae otro Nombre de la materia." % s["grupo_materia"]["crn"],
        "  %d · el segundo profesor trae otro Grupo." % s["grupo_grupo"]["crn"],
        "  %d · %% de responsabilidad 60%% y 50 (suma 110)." % s["pct_dos"]["crn"],
        "  %d · un solo profesor con 80 %% (suma 80)." % s["pct_uno"]["crn"],
        "  %d · nómina vacía; su ECOA sí trae la nómina, así que también sale como ECOA sin "
        "programación." % s["nomina_vacia"]["crn"],
        "  %d · nómina 12345678 (sin la L); su ECOA trae la misma, así que el cruce sí la "
        "encuentra." % s["nomina_formato"]["crn"],
        "  %s · la nómina %s aparece con un solo apellido en uno de sus grupos." % (
            min(g["crn"] for g in s["nombres"]["crns"]), s["nombres"]["nomina"]),
    ] + ["  %d y %d · choque de horario de la nómina %s (mismos días y horas)." % x
         for x in pares_choque] + [
        "  %d y %d · modalidad fuera del catálogo: Virtual y Por definir." % (
            s["mod_fuera"][0]["crn"], s["mod_fuera"][1]["crn"]),
        "  %s · programados sin ECOA." % y_lista(str(g["crn"]) for g in s["sin_ecoa"]),
        "  %s · ECOAs de grupos que no están en la programación." % y_lista(
            str(c) for c, _ in extra),
        "",
        "LO QUE NO DEBE CONTAR (solo formato)",
        "  % de responsabilidad escrito 50% o 50; hora 07:00 o 0700.",
        "  %d y %d · modalidad en minúsculas o mayúsculas (sí está en el catálogo)." % (
            s["mod_formato"][0]["crn"], s["mod_formato"][1]["crn"]),
        "  Nóminas en minúsculas en unos renglones de ECOAs.",
        "",
        "LO ESPERADO POR TIPO",
    ] + ["  %s: %d" % (t, por_tipo[t]) for t in tipos] + [
        "",
        "LAS TRES LÍNEAS QUE DEBE DAR revisar_programacion.py",
    ] + ["  " + l for l in lineas]
    (out / "NOTAS-ficticio.txt").write_text("\n".join(notas) + "\n", encoding="utf-8")
    print("Listo. En %s/:" % out)
    print("  %s, %s y NOTAS-ficticio.txt" % (np_, ne))
    print("Lo que debe salir:")
    for l in lineas:
        print("  " + l)
    return 0


if __name__ == "__main__":
    sys.exit(principal())
