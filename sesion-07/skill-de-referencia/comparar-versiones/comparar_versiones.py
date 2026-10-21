#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
comparar_versiones.py: compara dos versiones del mismo archivo (Excel o CSV) por una clave y
dice que renglones son nuevos, cuales ya no aparecen y que celdas cambiaron.
Es parte del skill comparar-versiones (curso Ed Digital, Sesion 7). Solo biblioteca estandar:
corre con cualquier python3, sin instalar nada.

Uso:
    python3 comparar_versiones.py ANTERIOR NUEVA --clave CRN
    python3 comparar_versiones.py ANTERIOR NUEVA --clave "Clave de materia,Grupo" --salida CARPETA

ANTERIOR y NUEVA son .xlsx o .csv. --clave es la columna (o las columnas, separadas por coma)
que identifica cada renglon. Compara por clave, nunca por posicion: el orden de los renglones
no importa. Valores con espacios de mas y numeros como 35 y 35.0 cuentan como iguales.

Deja en la carpeta salida/ (junto a la carpeta versiones/ o entrada/ donde esta la version
nueva; si no, dentro de la carpeta de la version nueva; o en --salida):
    cambios-NOMBRE.xlsx           hojas Resumen, Nuevos, Ya no aparecen y Cambios
    resumen-cambios-NOMBRE.txt    las tres lineas del resumen
donde NOMBRE es el nombre de la version nueva sin extension. Imprime el resumen y el control.
No modifica ninguna de las dos versiones.
"""
import argparse
import csv
import io
import math
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

FILAS_ENCABEZADO = 10      # el encabezado se busca en los primeros 10 renglones
TOP_COLUMNAS = 3           # columnas con mas cambios que van en la tercera linea
# Columna que acompana a la clave en la hoja Cambios, para reconocer el renglon
ETIQUETAS = ["materia", "nombre de la materia", "nombre materia", "curso", "nombre del curso",
             "nombre", "descripcion", "indicador", "concepto"]


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


# ------------------------------------------------------------------ una version
def explicar_repetidas(ruta, encabezado, columnas, idx_clave, repetidas, todas):
    """Mensaje para una clave que se repite: cuales se repiten, en que renglones, y una
    combinacion de columnas que si identifica cada renglon, si la hay."""
    nombre_clave = " + ".join(encabezado[j] for j in idx_clave)
    ejemplos = ["%s (renglones %s)" % (" + ".join(k), y_lista(str(x) for x in rs))
                for k, rs in sorted(repetidas.items(), key=lambda x: orden_natural(x[0]))[:5]]
    msg = ("La clave %s se repite en %s: %d %s más de una vez, por ejemplo %s. Para comparar, "
           "cada renglón necesita una clave que no se repita. "
           % (nombre_clave, ruta.name, len(repetidas),
              "valor aparece" if len(repetidas) == 1 else "valores aparecen", "; ".join(ejemplos)))
    sugerida = None
    for nombre, j in columnas:
        if j in idx_clave:
            continue
        combos = [tuple(canonico(f[i]) for i in idx_clave + [j]) for f in todas]
        if len(set(combos)) == len(combos):
            sugerida = nombre
            break
    actual = ",".join(encabezado[j] for j in idx_clave)
    if sugerida:
        msg += ("Dime qué combinación de columnas identifica un renglón y pásala separada por "
                "coma; en este archivo, por ejemplo, --clave \"%s,%s\" no se repite."
                % (actual, sugerida))
    else:
        msg += ("Dime qué combinación de columnas identifica un renglón y pásala separada por "
                "coma, por ejemplo --clave \"%s,OTRA COLUMNA\". Si los renglones repetidos son "
                "idénticos, son duplicados: quítalos antes de comparar." % actual)
    return msg


def leer_version(ruta, claves):
    """Lee una version: encuentra el encabezado (la fila que trae la clave), arma un renglon
    por clave y avisa si la clave falta o se repite."""
    filas = leer_filas(ruta)
    if not any(limpio(c) for f in filas for c in f):
        raise Falla("El archivo %s llegó vacío: no trae ningún renglón. Revisa que sea la "
                    "versión correcta." % ruta.name)
    claves_f = [fold(k) for k in claves]
    enc = None
    for i, f in enumerate(filas[:FILAS_ENCABEZADO]):
        nombres = [fold(x) for x in f]
        if all(k in nombres for k in claves_f):
            enc = i
            break
    if enc is None:
        candidatas = [f for f in filas[:FILAS_ENCABEZADO] if any(limpio(x) for x in f)]
        if not candidatas:
            raise Falla("Los primeros %d renglones de %s están vacíos; el encabezado (la fila "
                        "con los nombres de las columnas) tiene que estar ahí."
                        % (FILAS_ENCABEZADO, ruta.name))
        # La fila que mas se parece a un encabezado: la que trae mas celdas con texto
        mejor = max(candidatas, key=lambda f: sum(1 for x in f if limpio(x)))
        nombres = [fold(x) for x in mejor]
        faltan = [k for k, kf in zip(claves, claves_f) if kf not in nombres]
        raise Falla("%s %s en %s: no %s encontré en los primeros %d renglones. "
                    "Columnas que sí hay: %s. Revisa cómo se llama la columna que identifica "
                    "cada renglón y pásala con --clave."
                    % ("Falta la columna" if len(faltan) == 1 else "Faltan las columnas",
                       y_lista('"%s"' % k for k in faltan), ruta.name,
                       "la" if len(faltan) == 1 else "las", FILAS_ENCABEZADO,
                       " | ".join(limpio(x) for x in mejor if limpio(x))))
    encabezado = [limpio(x) for x in filas[enc]]
    columnas, avisos, vistas = [], [], {}
    for j, nombre in enumerate(encabezado):
        if not nombre:
            continue
        if fold(nombre) in vistas:
            avisos.append("La columna %s aparece dos veces en %s; comparo solo la primera."
                          % (nombre, ruta.name))
            continue
        vistas[fold(nombre)] = j
        columnas.append((nombre, j))
    idx_clave = [vistas[k] for k in claves_f]
    registros, repetidas, sin_clave, todas = {}, {}, 0, []
    for r, f in enumerate(filas[enc + 1:], start=enc + 2):
        if not any(limpio(c) for c in f):
            continue
        f = list(f) + [""] * (len(encabezado) - len(f))
        clave = tuple(canonico(f[j]) for j in idx_clave)
        if all(c == "" for c in clave):
            sin_clave += 1
            continue
        todas.append(f)
        if clave in registros:
            repetidas.setdefault(clave, [registros[clave]["renglon"]]).append(r)
            continue
        registros[clave] = dict(renglon=r, fila=f)
    if repetidas:
        raise Falla(explicar_repetidas(ruta, encabezado, columnas, idx_clave, repetidas, todas))
    if not registros:
        raise Falla("El archivo %s trae el encabezado (renglón %d) pero ningún renglón de datos. "
                    "Revisa que sea la versión correcta." % (ruta.name, enc + 1))
    return dict(ruta=ruta, encabezado_renglon=enc + 1, columnas=columnas, idx_clave=idx_clave,
                registros=registros, sin_clave=sin_clave, avisos=avisos)


# ------------------------------------------------------------------ comparacion
def comparar(ant, nue, claves):
    col_ant = {fold(n): (n, j) for n, j in ant["columnas"]}
    col_nue = {fold(n): (n, j) for n, j in nue["columnas"]}
    claves_f = [fold(k) for k in claves]
    avisos = ant["avisos"] + nue["avisos"]
    solo_ant = [n for n, _ in ant["columnas"] if fold(n) not in col_nue]
    solo_nue = [n for n, _ in nue["columnas"] if fold(n) not in col_ant]
    for solo, en, no_en in ((solo_ant, ant, nue), (solo_nue, nue, ant)):
        if solo:
            uno = len(solo) == 1
            avisos.append("%s %s %s en %s y no en %s: %s." % (
                "La columna" if uno else "Las columnas", y_lista(solo),
                "está" if uno else "están", en["ruta"].name, no_en["ruta"].name,
                "no se compara" if uno else "no se comparan"))
    # Columnas que se comparan: las de las dos versiones, sin la clave, en el orden del archivo
    comunes = [(n, j, col_nue[fold(n)][1]) for n, j in ant["columnas"]
               if fold(n) in col_nue and fold(n) not in claves_f]
    presentes = [fold(n) for n, _, _ in comunes]
    etiqueta = next((e for e in ETIQUETAS if e in presentes), None)

    def valor_etiqueta(fila_nue):
        return limpio(fila_nue[col_nue[etiqueta][1]]) if etiqueta else ""

    ra, rn = ant["registros"], nue["registros"]
    nuevos = sorted([k for k in rn if k not in ra], key=orden_natural)
    ya_no = sorted([k for k in ra if k not in rn], key=orden_natural)
    cambios = []
    for k in sorted([k for k in ra if k in rn], key=orden_natural):
        fa, fn = ra[k]["fila"], rn[k]["fila"]
        for nombre, ja, jn in comunes:
            if canonico(fa[ja]) != canonico(fn[jn]):
                cambios.append(dict(clave=k, etiqueta=valor_etiqueta(fn), columna=nombre,
                                    antes=limpio(fa[ja]), despues=limpio(fn[jn])))
    por_col = {}
    for c in cambios:
        por_col[c["columna"]] = por_col.get(c["columna"], 0) + 1
    orden = [n for n, _, _ in comunes]
    top = sorted(por_col, key=lambda n: (-por_col[n], orden.index(n)))
    con_cambios = len({c["clave"] for c in cambios})
    comunes_n = sum(1 for k in ra if k in rn)
    nombre_etiqueta = col_nue[etiqueta][0] if etiqueta else None
    return dict(nuevos=nuevos, ya_no=ya_no, cambios=cambios, por_col=por_col, top=top,
                con_cambios=con_cambios, sin_cambios=comunes_n - con_cambios,
                avisos=avisos, solo_ant=solo_ant, solo_nue=solo_nue,
                etiqueta=nombre_etiqueta)


def unidad(claves):
    """Como se llaman los renglones en el resumen: grupos si la clave es el CRN."""
    if "crn" in [fold(k) for k in claves]:
        return "grupos", "grupo"
    return "renglones", "renglón"


def lineas_resumen(ant, nue, nombres_clave, r, nombre_detalle, claves):
    plural, singular = unidad(claves)

    def cuantos(n, texto_plural, texto_singular):
        return "%s %s" % (miles(n), texto_plural if n != 1 else texto_singular)

    l1 = "Comparación de %s contra %s, por %s." % (ant["ruta"].name, nue["ruta"].name,
                                                   y_lista(nombres_clave))
    celdas = len(r["cambios"])
    l2 = "Hay %s, %s y %s (%s)." % (
        cuantos(len(r["nuevos"]), plural + " nuevos", singular + " nuevo"),
        cuantos(len(r["ya_no"]), plural + " que ya no aparecen", singular + " que ya no aparece"),
        cuantos(r["con_cambios"], plural + " con cambios", singular + " con cambios"),
        cuantos(celdas, "celdas", "celda"))
    if r["top"]:
        mas = y_lista("%s (%s)" % (n, miles(r["por_col"][n])) for n in r["top"][:TOP_COLUMNAS])
    else:
        mas = "ninguna"
    l3 = "Columnas con más cambios: %s. Detalle: %s." % (mas, nombre_detalle)
    return [l1, l2, l3]


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


# ------------------------------------------------------------------ principal
def carpeta_salida(nueva, salida):
    if salida:
        return Path(salida).expanduser()
    if fold(nueva.parent.name) in ("versiones", "entrada"):
        return nueva.parent.parent / "salida"
    return nueva.parent / "salida"


def principal(args):
    if not args.clave:
        raise Falla("Falta --clave: dime qué columna identifica cada renglón, por ejemplo "
                    "--clave CRN (o varias separadas por coma: --clave \"Clave de materia,Grupo\").")
    claves = [limpio(k) for k in args.clave.split(",") if limpio(k)]
    ruta_ant = Path(args.anterior).expanduser()
    ruta_nue = Path(args.nueva).expanduser()
    ant = leer_version(ruta_ant, claves)
    nue = leer_version(ruta_nue, claves)
    r = comparar(ant, nue, claves)

    out = carpeta_salida(ruta_nue, args.salida)
    out.mkdir(parents=True, exist_ok=True)
    base = ruta_nue.stem
    nombre_detalle = "cambios-%s.xlsx" % base
    ruta_detalle, ruta_resumen = out / nombre_detalle, out / ("resumen-cambios-%s.txt" % base)
    nombre_de = {j: n for n, j in nue["columnas"]}
    nombres_clave = [nombre_de[j] for j in nue["idx_clave"]]
    lineas = lineas_resumen(ant, nue, nombres_clave, r, nombre_detalle, claves)
    plural, _ = unidad(claves)

    control = [
        ("Versión anterior", ruta_ant.name),
        ("Versión nueva", ruta_nue.name),
        ("Clave", y_lista(nombres_clave)),
        ("Renglón del encabezado (anterior, nueva)", "%d, %d" % (ant["encabezado_renglon"],
                                                                 nue["encabezado_renglon"])),
        ("Renglones con clave en la anterior", len(ant["registros"])),
        ("Renglones con clave en la nueva", len(nue["registros"])),
        ("%s nuevos" % plural.capitalize(), len(r["nuevos"])),
        ("%s que ya no aparecen" % plural.capitalize(), len(r["ya_no"])),
        ("%s con cambios" % plural.capitalize(), r["con_cambios"]),
        ("Celdas que cambiaron", len(r["cambios"])),
        ("%s sin cambios" % plural.capitalize(), r["sin_cambios"]),
    ]
    if ant["sin_clave"] or nue["sin_clave"]:
        control.append(("Renglones sin clave, que no se comparan (anterior, nueva)",
                        "%d, %d" % (ant["sin_clave"], nue["sin_clave"])))

    # Hoja Resumen
    res = [[(lineas[0], NEGRITA)], [lineas[1]], [lineas[2]], [],
           [("Concepto", NEGRITA), ("Valor", NEGRITA)]]
    res += [[k, v] for k, v in control]
    res += [[], [("Cambios por columna", NEGRITA), ("Celdas", NEGRITA)]]
    res += [[n, r["por_col"][n]] for n in r["top"]] or [["ninguna", 0]]
    if r["avisos"]:
        res += [[], [("Avisos", NEGRITA)]] + [[a] for a in r["avisos"]]

    # Hojas Nuevos y Ya no aparecen: los renglones completos
    def hoja_renglones(version, claves_lista):
        cols = version["columnas"]
        filas = [[(n, NEGRITA) for n, _ in cols]]
        for k in claves_lista:
            f = version["registros"][k]["fila"]
            filas.append([a_numero(f[j]) for _, j in cols])
        return filas, [max(8, min(40, len(n) + 4)) for n, _ in cols]

    h_nuevos, a_nuevos = hoja_renglones(nue, r["nuevos"])
    h_ya_no, a_ya_no = hoja_renglones(ant, r["ya_no"])

    # Hoja Cambios: una fila por celda que cambio
    cab = list(nombres_clave) + ([r["etiqueta"]] if r["etiqueta"] else []) + \
        ["Columna", "Antes", "Después"]
    h_cambios = [[(c, NEGRITA) for c in cab]]
    for c in r["cambios"]:
        fila = [a_numero(v) for v in c["clave"]]
        if r["etiqueta"]:
            fila.append(c["etiqueta"])
        fila += [c["columna"], a_numero(c["antes"]), a_numero(c["despues"])]
        h_cambios.append(fila)
    a_cambios = [10] * len(nombres_clave) + ([36] if r["etiqueta"] else []) + [20, 28, 28]

    escribir_xlsx(ruta_detalle, [
        ("Resumen", res, [70, 40], False),
        ("Nuevos", h_nuevos, a_nuevos, True),
        ("Ya no aparecen", h_ya_no, a_ya_no, True),
        ("Cambios", h_cambios, a_cambios, True),
    ])
    ruta_resumen.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    for a in r["avisos"]:
        print("Aviso: %s" % a)
    if r["avisos"]:
        print()
    print("Resumen:")
    for l in lineas:
        print(l)
    print()
    print("Control de la comparación:")
    for k, v in control:
        print("  %s: %s" % (k, miles(v) if isinstance(v, int) else v))
    if r["top"]:
        print("  Cambios por columna: %s" % ", ".join("%s %s" % (n, miles(r["por_col"][n]))
                                                     for n in r["top"]))
    print()
    print("Archivos:")
    print("  %s" % ruta_detalle)
    print("  %s" % ruta_resumen)
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Compara dos versiones de un archivo por una clave (skill comparar-versiones).")
    ap.add_argument("anterior", help="versión anterior (.xlsx o .csv)")
    ap.add_argument("nueva", help="versión nueva (.xlsx o .csv)")
    ap.add_argument("--clave", default=None,
                    help="columna que identifica cada renglón; varias, separadas por coma")
    ap.add_argument("--salida", default=None,
                    help="carpeta de salida (por omisión, salida/ junto a la carpeta de la versión nueva)")
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude comparar: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
