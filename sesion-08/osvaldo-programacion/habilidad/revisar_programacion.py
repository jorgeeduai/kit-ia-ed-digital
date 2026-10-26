#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
revisar_programacion.py: busca inconsistencias en el archivo de programacion de un periodo y
lo cruza con los resultados de ECOAs del mismo periodo, por CRN y nomina del profesor.
Es parte del borrador de habilidad revisar-programacion-ecoas (curso Ed Digital, propuesta de
Jorge para Osvaldo, ronda 2 despues de la Sesion 7). Solo biblioteca estandar: corre con
cualquier python3, sin instalar nada.

Uso:
    python3 revisar_programacion.py PROGRAMACION ECOAS
    python3 revisar_programacion.py PROGRAMACION            (solo la revision de programacion)

PROGRAMACION y ECOAS son .xlsx o .csv (del Excel se lee la primera hoja).

TODO ES SUPUESTO (Osvaldo no ha escrito su flujo; se parte de lo que dijo en la Sesion 5:
"el tema de ECOAs" y "la programacion, identificar inconsistencias o las relaciones"):
  - Los encabezados de los dos archivos (tablas COLUMNAS de abajo).
  - Que es una inconsistencia: la tabla TIPOS de abajo.
  - El catalogo de modalidades y el formato de la nomina (MODALIDADES y NOMINA_VALIDA).

Deja en la carpeta salida/ (junto a la carpeta ficticio/ o entrada/ donde esta el archivo de
programacion; si no, dentro de su carpeta; o en --salida):
    inconsistencias-NOMBRE.xlsx           hojas Resumen, Todas y una por tipo
    resumen-inconsistencias-NOMBRE.txt    las tres lineas del resumen
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

FILAS_ENCABEZADO = 10
TOP_TIPOS = 3

# ------------------------------------------------------------------ SUPUESTOS
CRN = "CRN"
NOMINA = "Nómina del profesor"
NOMBRE = "Nombre del profesor"
PORCENTAJE = "% de responsabilidad"
DIAS = "Días"
H_INICIO = "Horario inicio"
H_FIN = "Horario fin"
MODALIDAD = "Modalidad"
# Datos del grupo: deben ser iguales en todos los renglones de un mismo CRN
DATOS_GRUPO = ["Clave", "Nombre de la materia", "Grupo", "Modalidad", "Días", "Horario inicio",
               "Horario fin"]
MODALIDADES = ["Presencial", "Digital", "Híbrida", "Flexible digital"]
NOMINA_VALIDA = re.compile(r"^L[0-9]{8}$")       # L y ocho digitos (supuesto)

TIPOS = [
    ("repetido", "Renglón repetido", "programación"),
    ("grupo", "Datos del grupo distintos en el mismo CRN", "programación"),
    ("porcentaje", "% de responsabilidad que no suma 100", "programación"),
    ("nomina", "Nómina vacía o con otro formato", "programación"),
    ("nombres", "Una nómina con dos nombres", "programación"),
    ("choque", "Choque de horario del profesor", "programación"),
    ("modalidad", "Modalidad fuera del catálogo", "programación"),
    ("sin_ecoa", "Programado sin ECOA", "ECOAs"),
    ("sin_prog", "ECOA sin programación", "ECOAs"),
]
NOMBRE_TIPO = {k: n for k, n, _ in TIPOS}
HOJA_TIPO = {"repetido": "Repetidos", "grupo": "Datos del grupo", "porcentaje": "Porcentaje",
             "nomina": "Nómina", "nombres": "Dos nombres", "choque": "Choques",
             "modalidad": "Modalidad", "sin_ecoa": "Sin ECOA", "sin_prog": "ECOA sin programar"}


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


# ------------------------------------------------------------------ formatos
def porcentaje(v):
    s = limpio(v).replace("%", "").strip()
    try:
        f = float(s)
    except ValueError:
        return None
    return f * 100 if 0 < f <= 1 and "%" not in limpio(v) and "." in s else f


def minutos(v):
    """07:30, 7:30, 0730 o fraccion de dia de Excel (0.3125) -> minutos desde medianoche."""
    s = limpio(v)
    if not s:
        return None
    if ":" in s:
        h, _, m = s.partition(":")
        try:
            return int(h) * 60 + int(m[:2])
        except ValueError:
            return None
    try:
        f = float(s)
    except ValueError:
        return None
    if 0 <= f < 1:
        return int(round(f * 24 * 60))
    if s.isdigit() and len(s) in (3, 4):
        return int(s[:-2]) * 60 + int(s[-2:])
    return None


DIA = {"lu": "Lu", "ma": "Ma", "mi": "Mi", "ju": "Ju", "vi": "Vi", "sa": "Sa", "do": "Do"}


def dias(v):
    partes = [p for p in re.split(r"[^a-z]+", fold(v)) if p]
    return {DIA[p[:2]] for p in partes if p[:2] in DIA}


def semana(conjunto):
    return " ".join(d for d in ["Lu", "Ma", "Mi", "Ju", "Vi", "Sa", "Do"] if d in conjunto)


def hhmm(m):
    return "%02d:%02d" % divmod(m, 60) if m is not None else ""


# ------------------------------------------------------------------ revision
def revisar(prog, ecoas):
    casos = []

    def caso(tipo, crn, nom, detalle):
        casos.append(dict(tipo=tipo, crn=crn, nomina=nom, detalle=detalle))

    # Un registro por CRN y nomina; los repetidos se avisan
    registros, orden = {}, []
    for r, f in prog["datos"]:
        crn = canonico(celda(prog, f, CRN))
        if not crn:
            continue
        nom = nomina(celda(prog, f, NOMINA))
        k = (crn, nom)
        if k in registros and nom:
            caso("repetido", crn, nom, "Renglones %d y %d" % (registros[k]["renglon"], r))
            continue
        if k in registros:
            k = (crn, "#%d" % r)
        registros[k] = dict(renglon=r, fila=f, crn=crn, nomina=nom)
        orden.append(k)
    por_crn = {}
    for k in orden:
        por_crn.setdefault(k[0], []).append(registros[k])

    for crn, regs in por_crn.items():
        # Datos del grupo iguales en todos los renglones del CRN
        for col in DATOS_GRUPO:
            if not tiene(prog, col):
                continue
            valores = []
            for g in regs:
                v = celda(prog, g["fila"], col)
                if col in (H_INICIO, H_FIN):
                    v = hhmm(minutos(v)) or v
                elif col == DIAS:
                    v = semana(dias(v)) or v
                elif col == MODALIDAD:
                    v = fold(v)
                else:
                    v = canonico(v)
                valores.append(v)
            if len(set(valores)) > 1:
                caso("grupo", crn, "", "%s: %s" % (col, y_lista(sorted(
                    {celda(prog, g["fila"], col) for g in regs}))))
        # % de responsabilidad
        if tiene(prog, PORCENTAJE):
            ps = [porcentaje(celda(prog, g["fila"], PORCENTAJE)) for g in regs]
            if None in ps:
                caso("porcentaje", crn, "", "Hay un %% que no es número: %s" % y_lista(
                    celda(prog, g["fila"], PORCENTAJE) or "(vacío)" for g in regs))
            elif abs(sum(ps) - 100) > 0.5:
                caso("porcentaje", crn, "", "Suma %s %% (%s)" % (
                    ("%.1f" % sum(ps)).rstrip("0").rstrip("."),
                    y_lista(celda(prog, g["fila"], PORCENTAJE) for g in regs)))
        # Modalidad
        if tiene(prog, MODALIDAD):
            m = celda(prog, regs[0]["fila"], MODALIDAD)
            if fold(m) not in [fold(x) for x in MODALIDADES]:
                caso("modalidad", crn, "", "'%s' no está en el catálogo (%s)"
                     % (m or "(vacío)", y_lista(MODALIDADES)))

    # Nominas
    validos = []
    for k in orden:
        g = registros[k]
        if not g["nomina"]:
            caso("nomina", g["crn"], "", "Renglón %d sin nómina" % g["renglon"])
        elif not NOMINA_VALIDA.match(g["nomina"]):
            caso("nomina", g["crn"], g["nomina"], "Renglón %d: '%s' no tiene el formato "
                 "L y ocho dígitos" % (g["renglon"], g["nomina"]))
            validos.append(g)
        else:
            validos.append(g)
    por_nomina = {}
    for g in validos:
        por_nomina.setdefault(g["nomina"], []).append(g)
    for nom, regs in sorted(por_nomina.items()):
        nombres = sorted({celda(prog, g["fila"], NOMBRE) for g in regs
                          if celda(prog, g["fila"], NOMBRE)}, key=fold)
        if len({fold(x) for x in nombres}) > 1:
            caso("nombres", regs[0]["crn"], nom, y_lista("'%s'" % x for x in nombres))
        # Choques: mismo profesor, un dia en comun y horas que se enciman
        if not (tiene(prog, DIAS) and tiene(prog, H_INICIO) and tiene(prog, H_FIN)):
            continue
        bloques = []
        for g in regs:
            d = dias(celda(prog, g["fila"], DIAS))
            a = minutos(celda(prog, g["fila"], H_INICIO))
            b = minutos(celda(prog, g["fila"], H_FIN))
            if d and a is not None and b is not None:
                bloques.append((g["crn"], d, a, b))
        for i in range(len(bloques)):
            for j in range(i + 1, len(bloques)):
                c1, d1, a1, b1 = bloques[i]
                c2, d2, a2, b2 = bloques[j]
                if c1 != c2 and d1 & d2 and a1 < b2 and a2 < b1:
                    caso("choque", c1, nom, "CRN %s y CRN %s: %s, %s-%s y %s-%s" % (
                        c1, c2, semana(d1 & d2), hhmm(a1), hhmm(b1), hhmm(a2),
                        hhmm(b2)))

    # Cruce con ECOAs
    con_ecoa = None
    if ecoas is not None:
        llaves_e = {}
        for r, f in ecoas["datos"]:
            crn = canonico(celda(ecoas, f, CRN))
            nom = nomina(celda(ecoas, f, NOMINA))
            if crn and nom:
                llaves_e.setdefault((crn, nom), r)
        llaves_p = {(g["crn"], g["nomina"]) for g in validos}
        con_ecoa = len(llaves_p & set(llaves_e))
        for k in sorted(llaves_p - set(llaves_e), key=lambda x: orden_natural(x)):
            caso("sin_ecoa", k[0], k[1], "Está en la programación y no tiene ECOA")
        for k in sorted(set(llaves_e) - llaves_p, key=lambda x: orden_natural(x)):
            caso("sin_prog", k[0], k[1], "Renglón %d de ECOAs; no está en la programación"
                 % llaves_e[k])
    return casos, registros, por_crn, con_ecoa


# ------------------------------------------------------------------ principal
def principal(args):
    ruta_p = Path(args.programacion).expanduser()
    prog = leer_tabla(ruta_p, [CRN, NOMINA], "programación", "--crn o --nomina")
    avisos = list(prog["avisos"])
    for col in DATOS_GRUPO + [NOMBRE, PORCENTAJE]:
        if not tiene(prog, col):
            avisos.append("La columna %s no está en %s: lo que depende de ella no se revisa."
                          % (col, ruta_p.name))
    ecoas, ruta_e = None, None
    if args.ecoas:
        ruta_e = Path(args.ecoas).expanduser()
        ecoas = leer_tabla(ruta_e, [CRN, NOMINA], "resultados de ECOAs", "--crn o --nomina")
        avisos += ecoas["avisos"]
    casos, registros, por_crn, con_ecoa = revisar(prog, ecoas)

    orden_tipos = [k for k, _, _ in TIPOS]
    casos.sort(key=lambda c: (orden_natural((c["crn"],)), orden_tipos.index(c["tipo"])))
    por_tipo = {}
    for c in casos:
        por_tipo[c["tipo"]] = por_tipo.get(c["tipo"], 0) + 1
    en_prog = sum(n for t, n in por_tipo.items() if t not in ("sin_ecoa", "sin_prog"))
    en_ecoas = por_tipo.get("sin_ecoa", 0) + por_tipo.get("sin_prog", 0)
    crn_con = len({c["crn"] for c in casos})
    top = sorted(por_tipo, key=lambda t: (-por_tipo[t], orden_tipos.index(t)))

    out = carpeta_salida(ruta_p, args.salida)
    out.mkdir(parents=True, exist_ok=True)
    base = ruta_p.stem
    nombre_detalle = "inconsistencias-%s.xlsx" % base
    ruta_detalle = out / nombre_detalle
    ruta_resumen = out / ("resumen-inconsistencias-%s.txt" % base)

    if ecoas is not None:
        l1 = ("Revisión de %s (%s renglones, %s CRN) y su cruce con %s, por CRN y nómina del "
              "profesor." % (ruta_p.name, miles(len(prog["datos"])), miles(len(por_crn)),
                             ruta_e.name))
        l2 = ("Hay %s inconsistencias en %s CRN: %s en la programación y %s en el cruce con "
              "ECOAs." % (miles(len(casos)), miles(crn_con), miles(en_prog), miles(en_ecoas)))
    else:
        l1 = ("Revisión de %s (%s renglones, %s CRN), sin cruce con ECOAs."
              % (ruta_p.name, miles(len(prog["datos"])), miles(len(por_crn))))
        l2 = "Hay %s inconsistencias en %s CRN." % (miles(len(casos)), miles(crn_con))
    mas = y_lista("%s (%d)" % (NOMBRE_TIPO[t], por_tipo[t]) for t in top[:TOP_TIPOS]) \
        or "ninguno"
    l3 = "Tipos con más casos: %s. Detalle: %s." % (mas, nombre_detalle)
    lineas = [l1, l2, l3]

    control = [
        ("Archivo de programación", ruta_p.name),
        ("Resultados de ECOAs", ruta_e.name if ruta_e else "(sin cruce)"),
        ("Renglón del encabezado (programación%s)" % (", ECOAs" if ecoas else ""),
         "%d%s" % (prog["renglon_enc"], ", %d" % ecoas["renglon_enc"] if ecoas else "")),
        ("Renglones en programación", len(prog["datos"])),
        ("CRN en programación", len(por_crn)),
        ("Profesores (nóminas) en programación",
         len({k[1] for k in registros if k[1] and not k[1].startswith("#")})),
    ]
    if ecoas is not None:
        control += [("Renglones en ECOAs", len(ecoas["datos"])),
                    ("CRN y profesor con ECOA", con_ecoa)]
    control += [("Inconsistencias", len(casos)), ("CRN con alguna inconsistencia", crn_con)]
    control.append(("Por tipo", ", ".join("%s %d" % (NOMBRE_TIPO[t], por_tipo.get(t, 0))
                                          for t in orden_tipos
                                          if ecoas is not None or t not in ("sin_ecoa",
                                                                            "sin_prog"))))

    res = [[(lineas[0], NEGRITA)], [lineas[1]], [lineas[2]], [],
           [("Concepto", NEGRITA), ("Valor", NEGRITA)]]
    res += [[k, v] for k, v in control]
    res += [[], [("Supuestos de esta propuesta", NEGRITA)],
            ["Todos los encabezados y tipos de inconsistencia son supuestos: ver "
             "la sección Lo que supuse del LEEME.md."]]
    if avisos:
        res += [[], [("Avisos", NEGRITA)]] + [[a] for a in avisos]
    cab = [("CRN", NEGRITA), ("Tipo", NEGRITA), ("Nómina del profesor", NEGRITA),
           ("Detalle", NEGRITA)]
    todas = [cab] + [[a_numero(c["crn"]), NOMBRE_TIPO[c["tipo"]], c["nomina"], c["detalle"]]
                     for c in casos]
    hojas = [("Resumen", res, [80, 60], False), ("Todas", todas, [10, 40, 16, 80], True)]
    for t in orden_tipos:
        sel = [c for c in casos if c["tipo"] == t]
        if sel:
            hojas.append((HOJA_TIPO[t], [cab] + [[a_numero(c["crn"]), NOMBRE_TIPO[t],
                                                  c["nomina"], c["detalle"]] for c in sel],
                          [10, 40, 16, 80], True))
    escribir_xlsx(ruta_detalle, hojas)
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
    print()
    print("Archivos:")
    print("  %s" % ruta_detalle)
    print("  %s" % ruta_resumen)
    return 0


def main():
    global CRN, NOMINA
    ap = argparse.ArgumentParser(
        description="Inconsistencias de programación y cruce con ECOAs "
                    "(habilidad revisar-programacion-ecoas).")
    ap.add_argument("programacion", help="archivo de programación del periodo (.xlsx o .csv)")
    ap.add_argument("ecoas", nargs="?", default=None,
                    help="resultados de ECOAs del mismo periodo (.xlsx o .csv); opcional")
    ap.add_argument("--crn", default=None, help="columna del CRN (por omisión: CRN)")
    ap.add_argument("--nomina", default=None,
                    help="columna de nómina en los dos archivos (por omisión: %s)" % NOMINA)
    ap.add_argument("--salida", default=None, help="carpeta de salida")
    args = ap.parse_args()
    CRN = args.crn or CRN
    NOMINA = args.nomina or NOMINA
    try:
        return principal(args)
    except Falla as e:
        print("No pude revisar: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
