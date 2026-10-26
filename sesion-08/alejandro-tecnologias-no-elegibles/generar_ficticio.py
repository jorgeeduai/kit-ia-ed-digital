#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera la copia FICTICIA de los dos archivos del flujo de tecnologias no
elegibles (curso Ed Digital, borrador de skill del grupo, Sesion 7):

  1. El archivo matriz TeachingHub-AppOTecnologias, con varias pestanas (una por coleccion).
     Cada pestana trae las columnas Ficha y AppOTecnologia.
  2. El listado Tecnologias no elegibles, con las columnas Tecnologia y Motivo, en dos cortes:
     el anterior y el de hoy (el de hoy agrega dos tecnologias).

Como correrlo (solo biblioteca estandar, Python 3.8 o mas reciente), desde esta carpeta:
    python3 -B generar_ficticio.py
    python3 -B generar_ficticio.py --semilla 22 --salida ficticio

Salen, dentro de --salida (ficticio/ por omision):
    TeachingHub-AppOTecnologias/            un CSV por pestana (para la laptop, sin openpyxl)
    TeachingHub-AppOTecnologias.xlsx        las mismas pestanas en un solo Excel
    Tecnologias-no-elegibles.csv            el listado de hoy (8 tecnologias)
    Tecnologias-no-elegibles-corte-anterior.csv   el listado del corte anterior (6)
    NOTAS-ficticio.txt                      lo que se sembro y la respuesta correcta

Todas las tecnologias, fichas y motivos son inventados: ningun nombre corresponde a una
herramienta real ni a una ficha real del portal.
"""
import argparse
import csv
import random
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

# Encabezados tal como los escribio la persona en el muro
COL_FICHA, COL_APP = "Ficha", "AppOTecnologia"
COL_TEC, COL_MOTIVO = "Tecnología", "Motivo"

# Pestanas del archivo matriz y cuantas fichas trae cada una
PESTANAS = [("Actividades", "ACT", 60), ("Estrategias", "EST", 45),
            ("Evaluación", "EVA", 35), ("Trabajo colaborativo", "COL", 30)]

# Tecnologias permitidas (inventadas). Ninguna contiene, como palabra suelta, el nombre de una
# no elegible; LienzoEdu es la trampa a proposito: contiene "Lienzo" pegado a otra palabra.
PERMITIDAS = ["AulaNube", "PizarraViva", "EncuestaYa", "VideoAula", "TableroÁgil",
              "MapaIdeas", "RedactaMás", "LabSim", "PodClase", "GaleríaVisual",
              "AgendaGrupo", "BitácoraDigital", "RúbricaFácil", "SimulaQuímica",
              "LectorPDF Aula", "Presenta360"]

# Listado de no elegibles: (tecnologia, motivo). Las dos ultimas solo estan en el corte de hoy.
NO_ELEGIBLES = [
    ("QuizRayo", "Guarda datos de estudiantes fuera del convenio institucional"),
    ("NotaVoz", "Sin aviso de privacidad vigente"),
    ("Lienzo", "Versión gratuita descontinuada"),
    ("MuroLibre", "Publica contenido sin moderación"),
    ("ChatAula", "Pide cuenta personal al estudiante"),
    ("FotoCollage", "Sin soporte institucional"),
    ("TriviaMax", "Cambió sus términos de uso"),
    ("DibujaYa", "Guarda datos de estudiantes fuera del convenio institucional"),
]
N_ANTERIOR = 6

TEMAS = ["Debate guiado", "Lluvia de ideas", "Estudio de caso", "Aprendizaje basado en retos",
         "Mapa conceptual", "Diario reflexivo", "Simulación de laboratorio", "Coevaluación",
         "Rúbrica analítica", "Proyecto integrador", "Aula invertida", "Exposición breve",
         "Lectura comentada", "Resolución de problemas", "Portafolio de evidencias",
         "Juego de roles", "Galería de aprendizaje", "Pregunta detonadora", "Cuestionario rápido",
         "Video reflexivo", "Infografía del tema", "Línea del tiempo", "Glosario colaborativo",
         "Entrevista simulada", "Bitácora de avance"]

# Lo sembrado: (pestana, contenido de AppOTecnologia con la no elegible, que se debe detectar)
# Cada renglon sembrado lleva ademas una tecnologia permitida para que sea creible.
SEMBRADOS = [
    ("Actividades", "{p}, QuizRayo", ["QuizRayo"], "escrita normal"),
    ("Actividades", "QUIZRAYO", ["QuizRayo"], "en mayúsculas"),
    ("Actividades", "  notavoz , {p}", ["NotaVoz"], "en minúsculas y con espacios de más"),
    ("Actividades", "Lienzo; {p}", ["Lienzo"], "separada con punto y coma"),
    ("Estrategias", "MuroLibre; ChatAula", ["MuroLibre", "ChatAula"],
     "dos no elegibles en la misma celda"),
    ("Estrategias", "{p} / ChatAula", ["ChatAula"], "separada con diagonal"),
    ("Estrategias", "NotaVoz (versión web)", ["NotaVoz"], "con un comentario entre paréntesis"),
    ("Evaluación", "TriviaMax, {p}", ["TriviaMax"], "nueva en el corte de hoy"),
    ("Evaluación", "{p}, DibujaYa", ["DibujaYa"], "nueva en el corte de hoy"),
]
# Trampas que NO deben salir: LienzoEdu en Trabajo colaborativo (contiene "Lienzo" pegado)
TRAMPAS = [("Trabajo colaborativo", "LienzoEdu, {p}"), ("Trabajo colaborativo", "LienzoEdu"),
           ("Trabajo colaborativo", "{p}; LienzoEdu")]


# ------------------------------------------------------------------ escritura
def escribir_csv(ruta, filas):
    with open(ruta, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        for fila in filas:
            w.writerow(fila)


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
            estilo = ' s="1"' if i == 0 else ""
            celdas.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                          % (ref, estilo, escape(str(v))))
        if celdas:
            partes.append('<row r="%d">%s</row>' % (i + 1, "".join(celdas)))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<cols><col min="1" max="1" width="40" customWidth="1"/>'
            '<col min="2" max="2" width="40" customWidth="1"/></cols>'
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
         '%s<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
         '</Types>' % "".join(
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
         '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
         '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
         '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
         '</styleSheet>' % ns),
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
def generar(semilla):
    rnd = random.Random(semilla)
    pestanas = {}
    sembrados_por = {}
    for p, *_ in PESTANAS:
        sembrados_por[p] = [s for s in SEMBRADOS if s[0] == p] + \
            [(t[0], t[1], [], "trampa") for t in TRAMPAS if t[0] == p]
    notas_renglones = []
    for nombre, prefijo, n in PESTANAS:
        especiales = sembrados_por[nombre]
        # Renglones de datos donde caen los sembrados (sin repetir), el resto es normal
        posiciones = sorted(rnd.sample(range(n), len(especiales)))
        asignados = dict(zip(posiciones, especiales))
        filas = [[COL_FICHA, COL_APP]]
        for i in range(n):
            ficha = "%s-%03d %s" % (prefijo, i + 1, rnd.choice(TEMAS))
            if i in asignados:
                _, plantilla, esperadas, como = asignados[i]
                app = plantilla.replace("{p}", rnd.choice(PERMITIDAS))
                notas_renglones.append((nombre, i + 2, ficha, app, esperadas, como))
            else:
                k = rnd.choice([1, 1, 2, 2, 3])
                app = ", ".join(rnd.sample(PERMITIDAS, k))
            filas.append([ficha, app])
        # Un renglon vacio a la mitad y una ficha sin tecnologia, como en un archivo real
        mitad = len(filas) // 2
        filas.insert(mitad, ["", ""])
        notas_renglones = [(p, r + 1 if (p == nombre and r > mitad) else r, f, a, e, c)
                           for (p, r, f, a, e, c) in notas_renglones]
        filas.append(["%s-%03d %s" % (prefijo, n + 1, rnd.choice(TEMAS)), ""])
        pestanas[nombre] = filas
    return pestanas, notas_renglones


def listado(n, titulo):
    filas = [[titulo, ""], ["", ""], [COL_TEC, COL_MOTIVO]]
    filas += [[t, m] for t, m in NO_ELEGIBLES[:n]]
    return filas


def principal():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia del flujo de tecnologías "
                                             "no elegibles.")
    ap.add_argument("--semilla", type=int, default=22, help="fija el azar (22)")
    ap.add_argument("--salida", default="ficticio", help="carpeta de salida (ficticio)")
    args = ap.parse_args()

    out = Path(args.salida)
    carpeta = out / "TeachingHub-AppOTecnologias"
    carpeta.mkdir(parents=True, exist_ok=True)
    pestanas, notas = generar(args.semilla)

    for i, (nombre, filas) in enumerate(pestanas.items(), start=1):
        escribir_csv(carpeta / ("%d-%s.csv" % (i, nombre)), filas)
    escribir_xlsx(out / "TeachingHub-AppOTecnologias.xlsx", list(pestanas.items()))
    escribir_csv(out / "Tecnologias-no-elegibles.csv",
                 listado(len(NO_ELEGIBLES), "Tecnologías no elegibles, corte de hoy"))
    escribir_csv(out / "Tecnologias-no-elegibles-corte-anterior.csv",
                 listado(N_ANTERIOR, "Tecnologías no elegibles, corte anterior"))

    # Respuesta correcta
    fichas = [x for x in notas if x[4]]
    menciones = sum(len(x[4]) for x in fichas)
    por_pestana = {p: sum(1 for x in fichas if x[0] == p) for p, *_ in PESTANAS}
    encontradas = []
    for x in fichas:
        for t in x[4]:
            if t not in encontradas:
                encontradas.append(t)
    nuevas = [t for t, _ in NO_ELEGIBLES[N_ANTERIOR:]]
    total_fichas = sum(n + 1 for _, _, n in PESTANAS)

    lineas = [
        "NOTAS-ficticio.txt: lo que sembró generar_ficticio.py (semilla %d)" % args.semilla,
        "Todo es inventado: tecnologías, fichas y motivos.",
        "",
        "Archivo matriz: TeachingHub-AppOTecnologias (.xlsx y una carpeta con un CSV por pestaña)",
        "  Pestañas: %s" % ", ".join("%s (%d fichas)" % (p, n + 1) for p, _, n in PESTANAS),
        "  Total: %d fichas. Cada pestaña trae además un renglón vacío a la mitad y, al final,"
        % total_fichas,
        "  una ficha sin tecnología.",
        "Listado de hoy: Tecnologias-no-elegibles.csv, %d tecnologías (título arriba del encabezado)"
        % len(NO_ELEGIBLES),
        "Listado del corte anterior: Tecnologias-no-elegibles-corte-anterior.csv, %d tecnologías"
        % N_ANTERIOR,
        "  Nuevas en el corte de hoy: %s" % ", ".join(nuevas),
        "",
        "Sembrado (pestaña, renglón de Excel, ficha, AppOTecnologia, qué se debe encontrar, cómo):",
    ]
    for p, r, f, a, e, c in notas:
        lineas.append("  %s | renglón %d | %s | %r | %s | %s"
                      % (p, r, f, a, ", ".join(e) if e else "NADA (trampa)", c))
    lineas += [
        "",
        "Respuesta correcta:",
        "  Fichas con alguna tecnología no elegible: %d (menciones: %d)" % (len(fichas), menciones),
        "  Por pestaña: %s" % ", ".join("%s %d" % (p, n) for p, n in por_pestana.items()),
        "  Tecnologías encontradas: %d de %d (%s)" % (len(encontradas), len(NO_ELEGIBLES),
                                                     ", ".join(encontradas)),
        "  No aparece en ninguna pestaña: %s" % ", ".join(
            t for t, _ in NO_ELEGIBLES if t not in encontradas),
        "  Nuevas en este corte que sí aparecen: %s (%d fichas)" % (
            ", ".join(nuevas), sum(1 for x in fichas if set(x[4]) & set(nuevas))),
        "  Trampas que NO deben salir: %d renglones con LienzoEdu en Trabajo colaborativo"
        % len(TRAMPAS),
    ]
    (out / "NOTAS-ficticio.txt").write_text("\n".join(lineas) + "\n", encoding="utf-8")

    print("Listo. Archivos en %s/:" % out)
    print("  TeachingHub-AppOTecnologias/ (%d CSV, uno por pestaña)" % len(pestanas))
    print("  TeachingHub-AppOTecnologias.xlsx (%d pestañas)" % len(pestanas))
    print("  Tecnologias-no-elegibles.csv y Tecnologias-no-elegibles-corte-anterior.csv")
    print("  NOTAS-ficticio.txt")
    print()
    print("Sembrado: %d fichas con %d menciones de tecnologías no elegibles; %d trampas con LienzoEdu."
          % (len(fichas), menciones, len(TRAMPAS)))
    return 0


if __name__ == "__main__":
    sys.exit(principal())
