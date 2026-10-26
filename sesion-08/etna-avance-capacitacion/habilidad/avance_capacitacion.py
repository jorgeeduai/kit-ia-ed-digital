#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
avance_capacitacion.py: del archivo de programacion diaria deja solo los horarios TCM, arma
la lista de profesores programados (uno por nomina) y la cruza con el historico de
capacitacion para dar, por profesor, el detalle de cursos y el % de avance.
Es parte del borrador de habilidad avance-capacitacion (curso Ed Digital, propuesta de Jorge
para el flujo de Etna, ronda 2 despues de la Sesion 7). Solo biblioteca estandar: corre con
cualquier python3, sin instalar nada.

Uso:
    python3 avance_capacitacion.py PROGRAMACION HISTORICO
    python3 avance_capacitacion.py PROGRAMACION HISTORICO --horario "Zona horaria" --valor TCM

PROGRAMACION y HISTORICO son .xlsx o .csv (del Excel se lee la primera hoja).

SUPUESTOS (estan en la seccion Lo que supuse del LEEME.md y se confirman con Etna):
  - Los encabezados del archivo de programacion son SUPUESTOS: Nómina, Nombre del profesor,
    UDF, Modalidad, Idioma, Plan y una columna que dice el horario (TCM, TCH, TCP). Los del
    historico son los que Etna escribio en el muro.
  - Un renglon es TCM si la columna del horario contiene la palabra TCM.
  - Que cursos pide cada profesor y que valor cuenta como cursado: tablas CURSOS y
    CURSADO de abajo. Se cambian ahi, en un solo lugar.

Deja en la carpeta salida/ (junto a la carpeta ficticio/ o entrada/ donde esta el archivo de
programacion; si no, dentro de su carpeta; o en --salida):
    avance-capacitacion-NOMBRE.xlsx           hojas Resumen, Avance por profesor, Detalle,
                                              Sin histórico y Control TCM
    resumen-avance-capacitacion-NOMBRE.txt    las tres lineas del resumen
donde NOMBRE es el nombre del archivo de programacion sin extension. No modifica las entradas.
"""
import argparse
import csv
import io
import math
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

FILAS_ENCABEZADO = 10      # el encabezado se busca en los primeros 10 renglones
TOP_CURSOS = 3             # cursos que mas faltan en la tercera linea

# ------------------------------------------------------------------ SUPUESTOS DE LA REGLA
# Columnas del archivo de programacion (SUPUESTAS: Etna no las ha compartido).
PROG_NOMINA = "Nómina"
PROG_NOMBRE = "Nombre del profesor"
PROG_UDF = "UDF"
PROG_MODALIDAD = "Modalidad"
PROG_IDIOMA = "Idioma"
PROG_PLAN = "Plan"
PROG_HORARIO = "Zona horaria"
VALOR_TCM = "TCM"

# Columnas del historico de capacitacion (las que Etna escribio en el muro).
HIST_NOMINA = "Nómina"
HIST_NOMBRE = "Nombre"
HIST_EXTRA = ["Tipo Prof", "Escuela", "Director/Coordinador"]

# Que cursos cuentan para el % de avance y cuando los pide un profesor (SUPUESTO).
#   "todos"        lo pide todo profesor programado
#   "plan 2019"    lo pide quien imparte una UDF del plan 2019 (Tec21)
#   "digital"      lo pide quien imparte una UDF con modalidad distinta de presencial
#   "ingles"       lo pide quien imparte una UDF en inglés
# Las columnas Estatus Tec21, Estatus Prepa y Estatus ED no cuentan: son resúmenes.
# El bloque Prepa no cuenta: la programación es de profesional.
CURSOS = [
    ("Sensibilización a la microenseñanza", "plan 2019"),
    ("Canvas: Conceptos básicos", "todos"),
    ("Fundamentos del Modelo Tec21", "plan 2019"),
    ("Evaluación Competencias TEC21", "plan 2019"),
    ("Aula Virtual", "digital"),
    ("AA mediante HD", "digital"),
    ("Turnitin", "digital"),
    ("Respondus Monitor", "digital"),
    ("Inglés Acreditado", "ingles"),
]
REQUISITO_TEXTO = {"todos": "todos", "plan 2019": "plan 2019", "digital": "UDF digital",
                   "ingles": "UDF en inglés"}

# Qué valor de una celda del histórico cuenta como curso cursado (SUPUESTO). Además de estos
# textos, cuenta como cursado una fecha (15/08/2026, 2026-08-15 o fecha de Excel).
CURSADO = ["acreditado", "acreditada", "si", "sí", "x", "completo", "completado", "aprobado",
           "cursado", "concluido"]
NO_CURSADO = ["", "no", "pendiente", "no acreditado", "en curso", "inscrito", "no aplica",
              "-", "n/a", "na"]


class Falla(Exception):
    """Un problema con los archivos que la persona tiene que resolver; se explica y se sale."""

# ------------------------------------------------------------------ utilidades
def fold(s):
    """Para comparar nombres de columna: sin acentos, sin mayusculas y sin espacios de mas."""
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s))
                if not unicodedata.combining(c))
    return " ".join(s.casefold().split())


def limpio(s):
    return " ".join(str(s).split())


def miles(n):
    return "{:,}".format(n)


def y_lista(partes):
    partes = list(partes)
    if len(partes) <= 1:
        return "".join(partes)
    return ", ".join(partes[:-1]) + " y " + partes[-1]


def canonico(v):
    """La forma con la que se compara un valor: sin espacios de mas y, si es un numero,
    escrito siempre igual (35, 35.0 y ' 35 ' son lo mismo). Los textos que empiezan con cero,
    como 007, se quedan como texto."""
    s = limpio(v)
    if not s or (len(s) > 1 and s[0] == "0" and s[1].isdigit()) or "_" in s:
        return s
    try:
        f = float(s)
    except ValueError:
        return s
    if not math.isfinite(f):
        return s
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return repr(f)


def orden_natural(clave):
    """Ordena claves con numeros como numeros (41003 antes de 41027) y textos despues."""
    partes = []
    for v in clave:
        try:
            partes.append((0, float(v), ""))
        except ValueError:
            partes.append((1, 0.0, fold(v)))
    return partes


def a_numero(s):
    s = limpio(s)
    if s.isdigit() and len(s) < 16 and (s == "0" or not s.startswith("0")):
        return int(s)
    if s.count(".") == 1 and s.replace(".", "").isdigit() and not s.startswith("."):
        return float(s)
    return s


# ------------------------------------------------------------------ lectura
# Las funciones de lectura de limpiar_reporte.py (Sesion 6).
def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _texto_rico(el):
    partes = []
    for h in el:
        n = _local(h.tag)
        if n == "t":
            partes.append(h.text or "")
        elif n == "r":
            partes.extend(t.text or "" for t in h if _local(t.tag) == "t")
    return "".join(partes)


def _indice_columna(ref):
    j = 0
    for ch in ref:
        if ch.isalpha():
            j = j * 26 + (ord(ch.upper()) - 64)
        else:
            break
    return j - 1


def _numero_texto(raw):
    if raw is None or raw == "":
        return ""
    try:
        f = float(raw)
    except ValueError:
        return raw
    if f == int(f) and abs(f) < 1e15:
        return str(int(f))
    return "%.10g" % f


def leer_xlsx(ruta):
    """Primera hoja de un .xlsx como lista de renglones (listas de texto)."""
    with zipfile.ZipFile(ruta) as z:
        nombres = set(z.namelist())
        compartidas = []
        if "xl/sharedStrings.xml" in nombres:
            raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
            compartidas = [_texto_rico(si) for si in raiz if _local(si.tag) == "si"]
        hoja = None
        try:
            wb = ET.fromstring(z.read("xl/workbook.xml"))
            primera = next(e for e in wb.iter() if _local(e.tag) == "sheet")
            rid = next(v for k, v in primera.attrib.items() if _local(k) == "id")
            rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            destino = next(e.get("Target") for e in rels if e.get("Id") == rid)
            hoja = destino.lstrip("/") if destino.startswith("/") else "xl/" + destino
        except (KeyError, StopIteration):
            hoja = None
        if hoja not in nombres:
            hoja = sorted(n for n in nombres if n.startswith("xl/worksheets/sheet"))[0]
        raiz = ET.fromstring(z.read(hoja))
    filas, ultimo = {}, 0
    for row in raiz.iter():
        if _local(row.tag) != "row":
            continue
        r = int(row.get("r")) if row.get("r") else ultimo + 1
        ultimo = r
        celdas, siguiente = {}, 0
        for c in row:
            if _local(c.tag) != "c":
                continue
            j = _indice_columna(c.get("r")) if c.get("r") else siguiente
            siguiente = j + 1
            t = c.get("t", "n")
            if t == "inlineStr":
                val = "".join(_texto_rico(h) for h in c if _local(h.tag) == "is")
            else:
                v = next((h for h in c if _local(h.tag) == "v"), None)
                raw = v.text if v is not None and v.text is not None else ""
                if t == "s":
                    val = compartidas[int(raw)] if raw != "" else ""
                elif t in ("str", "e"):
                    val = raw
                elif t == "b":
                    val = "VERDADERO" if raw == "1" else "FALSO"
                else:
                    val = _numero_texto(raw)
            celdas[j] = val
        fila = [""] * (max(celdas) + 1) if celdas else []
        for j, val in celdas.items():
            fila[j] = val
        filas[r] = fila
    return [filas.get(i, []) for i in range(1, (max(filas) if filas else 0) + 1)]


def leer_csv(ruta):
    datos = Path(ruta).read_bytes()
    texto = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            texto = datos.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    muestra = texto[:20000]
    delim = max([",", ";", "\t"], key=muestra.count)
    return [list(f) for f in csv.reader(io.StringIO(texto, newline=""), delimiter=delim)]


def leer_filas(ruta):
    if not ruta.is_file():
        raise Falla("No encontré el archivo %s. Revisa la ruta y el nombre." % ruta)
    if ruta.stat().st_size == 0:
        raise Falla("El archivo %s llegó vacío: pesa 0 bytes. Revisa que sea la versión "
                    "correcta." % ruta.name)
    ext = ruta.suffix.lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            return leer_xlsx(ruta)
        if ext in (".csv", ".txt"):
            return leer_csv(ruta)
    except (zipfile.BadZipFile, ET.ParseError, KeyError, IndexError, ValueError):
        raise Falla("No pude leer %s: no parece un archivo %s válido o está dañado. Ábrelo en "
                    "Excel y guárdalo otra vez como .xlsx o como .csv." % (ruta.name, ext))
    raise Falla("No sé leer archivos %s (%s). Guarda el archivo como .xlsx o como .csv."
                % (ext or "sin extensión", ruta.name))


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


# ------------------------------------------------------------------ tabla con encabezado
def leer_tabla(ruta, requeridas, que_es, opcion):
    """Lee un archivo y encuentra el encabezado: la fila, en los primeros 10 renglones, que
    trae todas las columnas requeridas. Devuelve el encabezado, un indice de columnas (por
    nombre sin acentos ni mayusculas) y los renglones con datos, con su numero de renglon."""
    filas = leer_filas(ruta)
    if not any(limpio(c) for f in filas for c in f):
        raise Falla("El archivo %s (%s) llegó vacío: no trae ningún renglón. Revisa que sea "
                    "el archivo correcto." % (ruta.name, que_es))
    req_f = [fold(k) for k in requeridas]
    enc = None
    for i, f in enumerate(filas[:FILAS_ENCABEZADO]):
        nombres = [fold(x) for x in f]
        if all(k in nombres for k in req_f):
            enc = i
            break
    if enc is None:
        candidatas = [f for f in filas[:FILAS_ENCABEZADO] if any(limpio(x) for x in f)]
        if not candidatas:
            raise Falla("Los primeros %d renglones de %s están vacíos; el encabezado (la fila "
                        "con los nombres de las columnas) tiene que estar ahí."
                        % (FILAS_ENCABEZADO, ruta.name))
        mejor = max(candidatas, key=lambda f: sum(1 for x in f if limpio(x)))
        nombres = [fold(x) for x in mejor]
        faltan = [k for k, kf in zip(requeridas, req_f) if kf not in nombres]
        raise Falla("%s %s en %s (%s): no %s encontré en los primeros %d renglones. "
                    "Columnas que sí hay: %s. Si en tu archivo se llama de otra forma, "
                    "indícalo con %s."
                    % ("Falta la columna" if len(faltan) == 1 else "Faltan las columnas",
                       y_lista('"%s"' % k for k in faltan), ruta.name, que_es,
                       "la" if len(faltan) == 1 else "las", FILAS_ENCABEZADO,
                       " | ".join(limpio(x) for x in mejor if limpio(x)), opcion))
    encabezado = [limpio(x) for x in filas[enc]]
    col, avisos = {}, []
    for j, nombre in enumerate(encabezado):
        if not nombre:
            continue
        if fold(nombre) in col:
            avisos.append("La columna %s aparece dos veces en %s; uso solo la primera."
                          % (nombre, ruta.name))
            continue
        col[fold(nombre)] = j
    datos = []
    for r, f in enumerate(filas[enc + 1:], start=enc + 2):
        if not any(limpio(c) for c in f):
            continue
        datos.append((r, list(f) + [""] * (len(encabezado) - len(f))))
    if not datos:
        raise Falla("El archivo %s trae el encabezado (renglón %d) pero ningún renglón de "
                    "datos. Revisa que sea el archivo correcto." % (ruta.name, enc + 1))
    return dict(ruta=ruta, renglon_enc=enc + 1, encabezado=encabezado, col=col,
                datos=datos, avisos=avisos)


def celda(tabla, fila, nombre):
    """Valor limpio de la columna NOMBRE en un renglon; vacio si la columna no esta."""
    j = tabla["col"].get(fold(nombre))
    return limpio(fila[j]) if j is not None and j < len(fila) else ""


def tiene(tabla, nombre):
    return fold(nombre) in tabla["col"]


def nomina(v):
    """La nomina con la que se cruza: sin espacios y en mayusculas (l0123 y L0123 son la
    misma)."""
    return "".join(str(v).split()).upper()


def carpeta_salida(ruta, salida):
    if salida:
        return Path(salida).expanduser()
    if fold(ruta.parent.name) in ("ficticio", "entrada"):
        return ruta.parent.parent / "salida"
    return ruta.parent / "salida"


# ------------------------------------------------------------------ reglas
FECHA_TEXTO = re.compile(r"^(\d{1,2}/\d{1,2}/\d{4}|\d{4}-\d{2}-\d{2})$")


def estado_curso(valor):
    """('cursado' | 'falta' | 'no reconocido') para el valor de una celda del historico."""
    v = fold(valor)
    if v in [fold(x) for x in CURSADO]:
        return "cursado"
    if v in [fold(x) for x in NO_CURSADO]:
        return "falta"
    if FECHA_TEXTO.match(v):
        return "cursado"
    try:
        f = float(v)
        if 20000 <= f <= 80000:          # fecha guardada como numero de Excel
            return "cursado"
    except ValueError:
        pass
    return "no reconocido"


def es_tcm(valor, buscado):
    """El renglon es TCM si la palabra TCM aparece en la celda (sola o dentro de un texto
    como 'LuJu 07:00-08:30 TCM')."""
    palabras = re.split(r"[^0-9a-z]+", fold(valor))
    return fold(buscado) in palabras


def requisitos_de(udfs):
    """Que condiciones aplican a un profesor, por las UDF que imparte en TCM."""
    req = {"todos"}
    for u in udfs:
        if "2019" in fold(u["plan"]):
            req.add("plan 2019")
        if u["modalidad"] and "presencial" not in fold(u["modalidad"]):
            req.add("digital")
        if fold(u["idioma"]) in ("ingles", "en", "english"):
            req.add("ingles")
    return req


# ------------------------------------------------------------------ principal
def principal(args):
    ruta_p = Path(args.programacion).expanduser()
    ruta_h = Path(args.historico).expanduser()
    col_nom_p = args.nomina_programacion or PROG_NOMINA
    col_nom_h = args.nomina_historico or HIST_NOMINA
    col_hor = args.horario or PROG_HORARIO
    buscado = args.valor or VALOR_TCM

    prog = leer_tabla(ruta_p, [col_nom_p, col_hor], "programación",
                      "--nomina-programacion o --horario")
    hist = leer_tabla(ruta_h, [col_nom_h], "histórico de capacitación", "--nomina-historico")
    avisos = prog["avisos"] + hist["avisos"]
    for nombre in (PROG_NOMBRE, PROG_UDF, PROG_MODALIDAD, PROG_IDIOMA, PROG_PLAN):
        if not tiene(prog, nombre):
            avisos.append("La columna %s no está en %s: queda vacía en la lista de profesores "
                          "y no decide qué cursos se piden." % (nombre, ruta_p.name))
    cursos = [(c, cond) for c, cond in CURSOS if tiene(hist, c)]
    faltan_cursos = [c for c, _ in CURSOS if not tiene(hist, c)]
    if faltan_cursos:
        avisos.append("%s no %s en %s: no %s para el %% de avance." % (
            y_lista(faltan_cursos), "está" if len(faltan_cursos) == 1 else "están",
            ruta_h.name, "cuenta" if len(faltan_cursos) == 1 else "cuentan"))
    if not cursos:
        raise Falla("Ninguna columna de cursos de la tabla CURSOS está en %s. Revisa que sea el "
                    "histórico de capacitación." % ruta_h.name)

    # 1. Quitar los horarios distintos de TCM
    por_horario, tcm, sin_nomina = {}, [], 0
    for r, f in prog["datos"]:
        h = celda(prog, f, col_hor) or "(vacío)"
        por_horario[h] = por_horario.get(h, 0) + 1
        if es_tcm(h, buscado):
            if not nomina(celda(prog, f, col_nom_p)):
                sin_nomina += 1
                continue
            tcm.append((r, f))
    if not tcm:
        raise Falla("Ningún renglón de %s trae %s en la columna %s. Valores que sí hay: %s. "
                    "Si el horario se marca de otra forma, indícalo con --valor."
                    % (ruta_p.name, buscado, col_hor, y_lista(sorted(por_horario))))
    nominas_todas = {nomina(celda(prog, f, col_nom_p)) for _, f in prog["datos"]} - {""}

    # 2. Lista nueva: un profesor por nómina, con sus UDF en TCM
    profes, orden = {}, []
    for r, f in tcm:
        n = nomina(celda(prog, f, col_nom_p))
        if n not in profes:
            profes[n] = dict(nomina=n, nombre=celda(prog, f, PROG_NOMBRE), udfs=[], vistas=set(),
                             nombres={celda(prog, f, PROG_NOMBRE)})
            orden.append(n)
        p = profes[n]
        p["nombres"].add(celda(prog, f, PROG_NOMBRE))
        u = dict(udf=celda(prog, f, PROG_UDF), modalidad=celda(prog, f, PROG_MODALIDAD),
                 idioma=celda(prog, f, PROG_IDIOMA), plan=celda(prog, f, PROG_PLAN))
        llave = tuple(fold(x) for x in u.values())
        if llave not in p["vistas"]:
            p["vistas"].add(llave)
            p["udfs"].append(u)
    solo_fuera = sorted(nominas_todas - set(profes))
    for n in orden:
        if len({fold(x) for x in profes[n]["nombres"] if x}) > 1:
            avisos.append("La nómina %s aparece con %d nombres en la programación: %s."
                          % (n, len(profes[n]["nombres"]), y_lista(sorted(profes[n]["nombres"]))))

    # 3. Histórico por nómina
    historico, repetidas = {}, []
    for r, f in hist["datos"]:
        n = nomina(celda(hist, f, col_nom_h))
        if not n:
            continue
        if n in historico:
            repetidas.append(n)
            continue
        historico[n] = f
    if repetidas:
        avisos.append("%d %s más de una vez en %s (%s); uso el primer renglón."
                      % (len(repetidas), "nómina aparece" if len(repetidas) == 1
                         else "nóminas aparecen", ruta_h.name,
                         y_lista(sorted(set(repetidas))[:5])))

    # 4. Avance por profesor
    filas_avance, filas_detalle, sin_hist, no_reconocidos = [], [], [], {}
    faltan_por_curso = {c: 0 for c, _ in cursos}
    nombres_distintos = 0
    for n in sorted(orden, key=lambda x: fold(profes[x]["nombre"])):
        p = profes[n]
        udfs = p["udfs"]
        req = requisitos_de(udfs)
        pide = [(c, cond) for c, cond in cursos if cond in req]
        base = [n, p["nombre"], "; ".join(sorted({u["udf"] for u in udfs if u["udf"]})),
                "; ".join(sorted({u["modalidad"] for u in udfs if u["modalidad"]})),
                "; ".join(sorted({u["idioma"] for u in udfs if u["idioma"]})),
                "; ".join(sorted({u["plan"] for u in udfs if u["plan"]}))]
        if n not in historico:
            sin_hist.append(base + [len(pide), y_lista(c for c, _ in pide)])
            filas_avance.append(base + ["No", len(pide), 0, 0.0, y_lista(c for c, _ in pide),
                                        "", "", "", ""])
            for c in [c for c, _ in pide]:
                faltan_por_curso[c] += 1
            continue
        h = historico[n]
        cursados, falta = 0, []
        for c, cond in pide:
            valor = celda(hist, h, c)
            e = estado_curso(valor)
            if e == "no reconocido":
                no_reconocidos[valor] = no_reconocidos.get(valor, 0) + 1
            if e == "cursado":
                cursados += 1
            else:
                falta.append(c)
                faltan_por_curso[c] += 1
            filas_detalle.append([n, p["nombre"], c, REQUISITO_TEXTO[cond], valor or "(vacío)",
                                  {"cursado": "Cursado", "falta": "Falta",
                                   "no reconocido": "No reconocido (cuenta como falta)"}[e]])
        pct = round(100.0 * cursados / len(pide), 1) if pide else 100.0
        nombre_h = celda(hist, h, HIST_NOMBRE)
        distinto = nombre_h if nombre_h and fold(nombre_h) != fold(p["nombre"]) else ""
        if distinto:
            nombres_distintos += 1
        filas_avance.append(base + ["Sí", len(pide), cursados, pct, y_lista(falta) or "nada",
                                    distinto] + [celda(hist, h, x) for x in HIST_EXTRA])

    con_hist = [f for f in filas_avance if f[6] == "Sí"]
    completos = sum(1 for f in con_hist if f[9] >= 100)
    medios = sum(1 for f in con_hist if 50 <= f[9] < 100)
    bajos = sum(1 for f in con_hist if f[9] < 50)
    promedio = round(sum(f[9] for f in con_hist) / len(con_hist), 1) if con_hist else 0.0
    top = sorted([c for c in faltan_por_curso if faltan_por_curso[c]],
                 key=lambda c: (-faltan_por_curso[c], [x for x, _ in cursos].index(c)))

    out = carpeta_salida(ruta_p, args.salida)
    out.mkdir(parents=True, exist_ok=True)
    base_nombre = ruta_p.stem
    nombre_detalle = "avance-capacitacion-%s.xlsx" % base_nombre
    ruta_detalle = out / nombre_detalle
    ruta_resumen = out / ("resumen-avance-capacitacion-%s.txt" % base_nombre)

    def pct_txt(x):
        return ("%.1f" % x).rstrip("0").rstrip(".") + " %"

    l1 = ("Avance de capacitación de %s profesores programados en horario %s (%s de %s "
          "renglones de %s), cruzado con %s por nómina."
          % (miles(len(orden)), buscado, miles(len(tcm)), miles(len(prog["datos"])),
             ruta_p.name, ruta_h.name))
    l2 = ("Promedio de avance: %s. Con 100 %%: %d; de 50 a 99 %%: %d; menos de 50 %%: %d; "
          "sin histórico: %d." % (pct_txt(promedio), completos, medios, bajos, len(sin_hist)))
    mas = y_lista("%s (%d)" % (c, faltan_por_curso[c]) for c in top[:TOP_CURSOS]) or "ninguno"
    l3 = "Cursos que más faltan: %s. Detalle: %s." % (mas, nombre_detalle)
    lineas = [l1, l2, l3]

    control = [
        ("Archivo de programación", ruta_p.name),
        ("Histórico de capacitación", ruta_h.name),
        ("Renglón del encabezado (programación, histórico)", "%d, %d" % (prog["renglon_enc"],
                                                                         hist["renglon_enc"])),
        ("Renglones en programación", len(prog["datos"])),
        ("Renglones por horario", ", ".join("%s %d" % (k, por_horario[k])
                                            for k in sorted(por_horario))),
        ("Renglones %s que se quedan" % buscado, len(tcm)),
        ("Profesores en programación (cualquier horario)", len(nominas_todas)),
        ("Profesores en %s (lista nueva)" % buscado, len(orden)),
        ("Profesores que solo aparecen fuera de %s" % buscado, len(solo_fuera)),
        ("Profesores en el histórico", len(historico)),
        ("Programados con histórico", len(con_hist)),
        ("Programados sin histórico", len(sin_hist)),
        ("Nombre distinto en el histórico", nombres_distintos),
        ("Cursos que cuentan", len(cursos)),
        ("Valores no reconocidos (cuentan como falta)", sum(no_reconocidos.values())),
    ]
    if sin_nomina:
        control.append(("Renglones %s sin nómina (no entran)" % buscado, sin_nomina))

    res = [[(lineas[0], NEGRITA)], [lineas[1]], [lineas[2]], [],
           [("Concepto", NEGRITA), ("Valor", NEGRITA)]]
    res += [[k, v] for k, v in control]
    res += [[], [("Cursos que faltan", NEGRITA), ("Profesores", NEGRITA)]]
    res += [[c, faltan_por_curso[c]] for c in top] or [["ninguno", 0]]
    if no_reconocidos:
        res += [[], [("Valores no reconocidos", NEGRITA), ("Veces", NEGRITA)]]
        res += [[k, v] for k, v in sorted(no_reconocidos.items())]
    res += [[], [("Supuestos de esta propuesta", NEGRITA)],
            ["Encabezados del archivo de programación supuestos: %s, %s, %s, %s, %s, %s y %s."
             % (PROG_NOMINA, PROG_NOMBRE, PROG_UDF, PROG_MODALIDAD, PROG_IDIOMA, PROG_PLAN,
                col_hor)],
            ["Qué cursos pide cada profesor y qué cuenta como cursado: ver PREGUNTAS."]]
    if avisos:
        res += [[], [("Avisos", NEGRITA)]] + [[a] for a in avisos]

    cab_av = ["Nómina", "Nombre del profesor", "UDF que imparte", "Modalidad", "Idioma", "Plan",
              "En histórico", "Cursos que se piden", "Cursados", "% de avance", "Le falta",
              "Nombre en el histórico (si es distinto)"] + HIST_EXTRA
    h_av = [[(c, NEGRITA) for c in cab_av]] + filas_avance
    h_det = [[(c, NEGRITA) for c in ["Nómina", "Nombre del profesor", "Curso", "Lo pide por",
                                     "Valor en el histórico", "Cuenta como"]]] + filas_detalle
    h_sin = [[(c, NEGRITA) for c in ["Nómina", "Nombre del profesor", "UDF que imparte",
                                     "Modalidad", "Idioma", "Plan", "Cursos que se piden",
                                     "Cursos"]]] + sin_hist
    h_tcm = [[("Horario", NEGRITA), ("Renglones", NEGRITA), ("Se queda", NEGRITA)]]
    h_tcm += [[k, por_horario[k], "Sí" if es_tcm(k, buscado) else "No"]
              for k in sorted(por_horario)]
    h_tcm += [[], [("Nómina que solo aparece fuera de %s" % buscado, NEGRITA)]]
    h_tcm += [[n] for n in solo_fuera]

    escribir_xlsx(ruta_detalle, [
        ("Resumen", res, [80, 30], False),
        ("Avance por profesor", h_av, [12, 30, 40, 14, 10, 8, 10, 10, 9, 10, 60, 30, 12, 20, 24],
         True),
        ("Detalle", h_det, [12, 30, 44, 14, 22, 30], True),
        ("Sin histórico", h_sin, [12, 30, 40, 14, 10, 8, 10, 60], True),
        ("Control TCM", h_tcm, [36, 12, 10], True),
    ])
    ruta_resumen.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    for a in avisos:
        print("Aviso: %s" % a)
    if avisos:
        print()
    print("Resumen:")
    for l in lineas:
        print(l)
    print()
    print("Control:")
    for k, v in control:
        print("  %s: %s" % (k, miles(v) if isinstance(v, int) else v))
    if no_reconocidos:
        print("  Valores no reconocidos: %s" % ", ".join("'%s' %d" % (k, v) for k, v in
                                                       sorted(no_reconocidos.items())))
    print()
    print("Archivos:")
    print("  %s" % ruta_detalle)
    print("  %s" % ruta_resumen)
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Avance de capacitación de los profesores programados en TCM "
                    "(habilidad avance-capacitacion).")
    ap.add_argument("programacion", help="archivo de programación diaria (.xlsx o .csv)")
    ap.add_argument("historico", help="histórico de capacitación (.xlsx o .csv)")
    ap.add_argument("--horario", default=None,
                    help="columna de programación que dice el horario (por omisión: %s)"
                         % PROG_HORARIO)
    ap.add_argument("--valor", default=None,
                    help="valor que se queda (por omisión: %s)" % VALOR_TCM)
    ap.add_argument("--nomina-programacion", default=None,
                    help="columna de nómina en programación (por omisión: %s)" % PROG_NOMINA)
    ap.add_argument("--nomina-historico", default=None,
                    help="columna de nómina en el histórico (por omisión: %s)" % HIST_NOMINA)
    ap.add_argument("--salida", default=None, help="carpeta de salida")
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude calcular el avance: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
