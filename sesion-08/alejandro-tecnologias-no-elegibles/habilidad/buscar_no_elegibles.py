#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
buscar_no_elegibles.py: busca, pestana por pestana del archivo matriz TeachingHub, las fichas
(actividades y estrategias) que mencionan alguna tecnologia del listado Tecnologias no
elegibles. Es parte del skill tecnologias-no-elegibles (curso Ed Digital, borrador del grupo).
Solo biblioteca estandar: corre con cualquier python3, sin instalar nada. Si openpyxl esta
instalado (TECgpt lo trae), lo usa para leer el .xlsx; si no, lo lee con la biblioteca estandar.

Uso:
    python3 buscar_no_elegibles.py MATRIZ LISTADO
    python3 buscar_no_elegibles.py MATRIZ LISTADO --anterior LISTADO_DEL_CORTE_ANTERIOR
    python3 buscar_no_elegibles.py MATRIZ LISTADO --en-toda-la-fila --salida CARPETA

MATRIZ es el .xlsx con todas las pestanas, o una carpeta con un .csv por pestana (el nombre del
archivo es el nombre de la pestana; un numero al inicio, como "1-", solo fija el orden).
LISTADO es el .xlsx (primera hoja) o .csv con la columna Tecnologia (y Motivo, si la trae).
--anterior es el listado del corte anterior: con el, el resumen dice que tecnologias son nuevas.

Por omision busca en la columna AppOTecnologia de cada pestana. Si una pestana no la trae, o
con --en-toda-la-fila, busca en todas las celdas del renglon y lo avisa.

Deja en salida/ (junto a la matriz; o en --salida):
    no-elegibles-NOMBRE.xlsx          hojas Resumen, Coincidencias, Por pestaña y Por tecnología
    resumen-no-elegibles-NOMBRE.txt   las tres lineas del resumen
donde NOMBRE es el nombre de la matriz sin extension. No modifica ningun archivo de entrada.
"""
import argparse
import csv
import io
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

FILAS_ENCABEZADO = 10           # el encabezado se busca en los primeros 10 renglones
COL_FICHA, COL_APP = "Ficha", "AppOTecnologia"
COL_TEC, COL_MOTIVO = "Tecnología", "Motivo"


class Falla(Exception):
    """Un problema con los archivos que la persona tiene que resolver; se explica y se sale."""


# ------------------------------------------------------------------ utilidades
def fold(s):
    """Sin acentos, sin mayusculas y sin espacios de mas."""
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


def menciona(texto_fold, tec_fold):
    """True si la tecnologia aparece en el texto como palabra suelta: antes y despues hay un
    borde (inicio, fin, espacio, coma, punto y coma, diagonal, parentesis...). Asi "Lienzo"
    no se confunde con "LienzoEdu"."""
    if not tec_fold:
        return False
    desde = 0
    while True:
        i = texto_fold.find(tec_fold, desde)
        if i < 0:
            return False
        antes = texto_fold[i - 1] if i > 0 else " "
        fin = i + len(tec_fold)
        despues = texto_fold[fin] if fin < len(texto_fold) else " "
        if not antes.isalnum() and not despues.isalnum():
            return True
        desde = i + 1


# ------------------------------------------------------------------ lectura
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


def _filas_de_hoja(raiz, compartidas):
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


def _xlsx_biblioteca_estandar(ruta):
    """Todas las hojas de un .xlsx, en orden: lista de (nombre, filas)."""
    with zipfile.ZipFile(ruta) as z:
        nombres = set(z.namelist())
        compartidas = []
        if "xl/sharedStrings.xml" in nombres:
            raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
            compartidas = [_texto_rico(si) for si in raiz if _local(si.tag) == "si"]
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        destinos = {e.get("Id"): e.get("Target") for e in rels}
        hojas = []
        for e in wb.iter():
            if _local(e.tag) != "sheet":
                continue
            rid = next(v for k, v in e.attrib.items() if _local(k) == "id")
            destino = destinos[rid]
            parte = destino.lstrip("/") if destino.startswith("/") else "xl/" + destino
            hojas.append((e.get("name"), _filas_de_hoja(ET.fromstring(z.read(parte)),
                                                        compartidas)))
    return hojas


def _xlsx_openpyxl(ruta, openpyxl):
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hojas = []
    for ws in wb.worksheets:
        filas = []
        for fila in ws.iter_rows(values_only=True):
            filas.append(["" if v is None else _numero_texto(v) if isinstance(v, float)
                          else str(v) for v in fila])
        hojas.append((ws.title, filas))
    wb.close()
    return hojas


def leer_xlsx(ruta):
    """Todas las hojas: con openpyxl si esta instalado, si no con la biblioteca estandar."""
    try:
        import openpyxl  # noqa: F401  (TECgpt lo trae; la laptop no)
    except ImportError:
        return _xlsx_biblioteca_estandar(ruta), "biblioteca estándar"
    return _xlsx_openpyxl(ruta, openpyxl), "openpyxl"


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


def nombre_pestana_csv(ruta):
    """'1-Actividades.csv' -> 'Actividades'. El numero del inicio solo fija el orden."""
    stem = ruta.stem
    i = 0
    while i < len(stem) and stem[i].isdigit():
        i += 1
    if 0 < i < len(stem) and stem[i] in "-_ ":
        return stem[i + 1:]
    return stem


def orden_csv(ruta):
    stem, i = ruta.stem, 0
    while i < len(stem) and stem[i].isdigit():
        i += 1
    return (int(stem[:i]) if i else 10 ** 9, fold(stem))


def revisar_archivo(ruta, que):
    if not ruta.exists():
        raise Falla("No encontré %s: %s. Revisa la ruta y el nombre." % (que, ruta))
    if ruta.is_file() and ruta.stat().st_size == 0:
        raise Falla("%s (%s) llegó vacío: pesa 0 bytes. Revisa que sea el archivo correcto."
                    % (que.capitalize(), ruta.name))


def leer_libro(ruta, que):
    """Lista de (pestana, filas) y como se leyo. Acepta .xlsx, .csv o carpeta de .csv."""
    revisar_archivo(ruta, que)
    if ruta.is_dir():
        csvs = sorted([p for p in ruta.iterdir() if p.suffix.lower() == ".csv"
                       and not p.name.startswith(".")], key=orden_csv)
        if not csvs:
            raise Falla("La carpeta %s no trae ningún .csv. Pon un .csv por pestaña, o pasa el "
                        ".xlsx completo." % ruta.name)
        return [(nombre_pestana_csv(p), leer_csv(p)) for p in csvs], "CSV, uno por pestaña"
    ext = ruta.suffix.lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            return leer_xlsx(ruta)
        if ext in (".csv", ".txt"):
            return [(ruta.stem, leer_csv(ruta))], "CSV"
    except (zipfile.BadZipFile, ET.ParseError, KeyError, IndexError, ValueError,
            StopIteration):
        raise Falla("No pude leer %s: no parece un archivo %s válido o está dañado. Ábrelo en "
                    "Excel y guárdalo otra vez como .xlsx, o cada pestaña como .csv."
                    % (ruta.name, ext))
    raise Falla("No sé leer archivos %s (%s). Guarda el archivo como .xlsx o como .csv."
                % (ext or "sin extensión", ruta.name))


def buscar_encabezado(filas, buscadas):
    """Indice de la primera fila (de las primeras 10) que trae alguna de las columnas buscadas."""
    buscadas_f = [fold(b) for b in buscadas]
    for i, f in enumerate(filas[:FILAS_ENCABEZADO]):
        nombres = [fold(x) for x in f]
        if any(b in nombres for b in buscadas_f):
            return i
    return None


def columnas_que_hay(filas):
    candidatas = [f for f in filas[:FILAS_ENCABEZADO] if any(limpio(x) for x in f)]
    if not candidatas:
        return "ninguna"
    mejor = max(candidatas, key=lambda f: sum(1 for x in f if limpio(x)))
    return " | ".join(limpio(x) for x in mejor if limpio(x))


# ------------------------------------------------------------------ listado
def leer_listado(ruta, que):
    hojas, _ = leer_libro(ruta, que)
    nombre, filas = hojas[0]
    if not any(limpio(c) for f in filas for c in f):
        raise Falla("%s (%s) llegó vacío: no trae ningún renglón." % (que.capitalize(),
                                                                      ruta.name))
    enc = buscar_encabezado(filas, [COL_TEC])
    if enc is None:
        raise Falla("Falta la columna \"%s\" en %s (%s): no la encontré en los primeros %d "
                    "renglones. Columnas que sí hay: %s. Revisa cómo se llama la columna con el "
                    "nombre de la tecnología." % (COL_TEC, que, ruta.name, FILAS_ENCABEZADO,
                                                  columnas_que_hay(filas)))
    nombres = [fold(x) for x in filas[enc]]
    j_tec = nombres.index(fold(COL_TEC))
    j_mot = nombres.index(fold(COL_MOTIVO)) if fold(COL_MOTIVO) in nombres else None
    tecs, vistas, avisos = [], {}, []
    for f in filas[enc + 1:]:
        f = list(f) + [""] * (len(nombres) - len(f))
        t = limpio(f[j_tec])
        if not t:
            continue
        if fold(t) in vistas:
            avisos.append("La tecnología %s está dos veces en %s; la cuento una vez."
                          % (t, ruta.name))
            continue
        vistas[fold(t)] = True
        tecs.append(dict(nombre=t, fold=fold(t),
                         motivo=limpio(f[j_mot]) if j_mot is not None else ""))
    if not tecs:
        raise Falla("%s (%s) trae el encabezado (renglón %d) pero ninguna tecnología debajo."
                    % (que.capitalize(), ruta.name, enc + 1))
    if j_mot is None:
        avisos.append("%s no trae la columna %s; el reporte sale sin motivo."
                      % (ruta.name, COL_MOTIVO))
    return tecs, avisos


# ------------------------------------------------------------------ busqueda
def revisar_pestana(nombre, filas, tecs, en_toda_la_fila):
    """Coincidencias de una pestana: una por (ficha, tecnologia)."""
    avisos, coincidencias = [], []
    if not any(limpio(c) for f in filas for c in f):
        avisos.append("La pestaña %s está vacía; no la reviso." % nombre)
        return dict(nombre=nombre, fichas=0, con=0, coincidencias=[], avisos=avisos,
                    donde="vacía")
    enc = buscar_encabezado(filas, [COL_APP, COL_FICHA])
    if enc is None:
        enc = next(i for i, f in enumerate(filas) if any(limpio(x) for x in f))
        avisos.append("La pestaña %s no trae las columnas %s ni %s en los primeros %d renglones "
                      "(trae: %s). Busqué en todas las celdas, desde el renglón %d."
                      % (nombre, COL_FICHA, COL_APP, FILAS_ENCABEZADO,
                         columnas_que_hay(filas), enc + 2))
    encabezado = [limpio(x) for x in filas[enc]]
    nombres = [fold(x) for x in encabezado]
    j_ficha = nombres.index(fold(COL_FICHA)) if fold(COL_FICHA) in nombres else None
    j_app = nombres.index(fold(COL_APP)) if fold(COL_APP) in nombres else None
    if j_ficha is None and enc is not None and nombres and any(nombres):
        avisos.append("La pestaña %s no trae la columna %s; en el reporte va el renglón."
                      % (nombre, COL_FICHA))
    toda = en_toda_la_fila or j_app is None
    if j_app is None and any(nombres):
        avisos.append("La pestaña %s no trae la columna %s; busqué en todas las celdas del "
                      "renglón." % (nombre, COL_APP))
    fichas, con = 0, 0
    for r, f in enumerate(filas[enc + 1:], start=enc + 2):
        if not any(limpio(c) for c in f):
            continue
        fichas += 1
        f = list(f) + [""] * (len(encabezado) - len(f))
        if toda:
            texto = " | ".join(limpio(c) for j, c in enumerate(f) if j != j_ficha)
        else:
            texto = limpio(f[j_app])
        tf = fold(texto)
        halladas = [t for t in tecs if menciona(tf, t["fold"])]
        if halladas:
            con += 1
        for t in halladas:
            coincidencias.append(dict(
                pestana=nombre, renglon=r,
                ficha=limpio(f[j_ficha]) if j_ficha is not None else "",
                texto=texto, tec=t["nombre"], motivo=t["motivo"]))
    return dict(nombre=nombre, fichas=fichas, con=con, coincidencias=coincidencias,
                avisos=avisos, donde="todas las celdas" if toda else COL_APP)


# ------------------------------------------------------------------ escritura xlsx
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
def principal(args):
    ruta_matriz = Path(args.matriz).expanduser()
    ruta_listado = Path(args.listado).expanduser()
    tecs, avisos = leer_listado(ruta_listado, "el listado de tecnologías no elegibles")
    nuevas = None
    if args.anterior:
        ruta_ant = Path(args.anterior).expanduser()
        tecs_ant, _ = leer_listado(ruta_ant, "el listado del corte anterior")
        previas = {t["fold"] for t in tecs_ant}
        nuevas = [t["nombre"] for t in tecs if t["fold"] not in previas]
    hojas, como = leer_libro(ruta_matriz, "el archivo matriz")
    pestanas = [revisar_pestana(n, f, tecs, args.en_toda_la_fila) for n, f in hojas]
    for p in pestanas:
        avisos += p["avisos"]
    revisadas = [p for p in pestanas if p["donde"] != "vacía"]
    if not revisadas:
        raise Falla("El archivo matriz %s no trae ninguna pestaña con datos." % ruta_matriz.name)
    total_fichas = sum(p["fichas"] for p in revisadas)
    coinc = [c for p in pestanas for c in p["coincidencias"]]
    con = sum(p["con"] for p in pestanas)
    por_tec = {t["nombre"]: set() for t in tecs}
    for c in coinc:
        por_tec[c["tec"]].add((c["pestana"], c["renglon"]))
    encontradas = [t["nombre"] for t in tecs if por_tec[t["nombre"]]]

    base = ruta_matriz.stem if ruta_matriz.is_file() else ruta_matriz.name
    nombre_detalle = "no-elegibles-%s.xlsx" % base
    l1 = ("Revisión de %s contra %s (%s), en %s y %s." % (
        ruta_matriz.name, ruta_listado.name,
        "1 tecnología" if len(tecs) == 1 else "%s tecnologías" % miles(len(tecs)),
        "1 pestaña" if len(revisadas) == 1 else "%d pestañas" % len(revisadas),
        "1 ficha" if total_fichas == 1 else "%s fichas" % miles(total_fichas)))
    l2 = ("Hay %s que %s alguna tecnología no elegible (%s): %s." % (
        "1 ficha" if con == 1 else "%s fichas" % miles(con),
        "menciona" if con == 1 else "mencionan",
        "1 mención" if len(coinc) == 1 else "%s menciones" % miles(len(coinc)),
        y_lista("%s %s" % (p["nombre"], miles(p["con"])) for p in revisadas)))
    if nuevas is None:
        tail = "sin listado anterior, no digo cuáles son nuevas"
    elif not nuevas:
        tail = "ninguna es nueva en este corte"
    else:
        fichas_nuevas = len({k for t in nuevas for k in por_tec[t]})
        tail = "nuevas en este corte: %s (%s)" % (
            y_lista(nuevas), "1 ficha" if fichas_nuevas == 1 else "%d fichas" % fichas_nuevas)
    l3 = "Tecnologías encontradas: %d de %d; %s. Detalle: %s." % (
        len(encontradas), len(tecs), tail, nombre_detalle)
    lineas = [l1, l2, l3]

    if args.salida:
        out = Path(args.salida).expanduser()
    else:
        out = (ruta_matriz.parent if ruta_matriz.is_file() else ruta_matriz.parent) / "salida"
    out.mkdir(parents=True, exist_ok=True)
    ruta_detalle = out / nombre_detalle
    ruta_resumen = out / ("resumen-no-elegibles-%s.txt" % base)

    control = [
        ("Archivo matriz", ruta_matriz.name),
        ("Leído con", como),
        ("Listado de no elegibles", ruta_listado.name),
        ("Listado del corte anterior", Path(args.anterior).name if args.anterior else "no se dio"),
        ("Tecnologías en el listado", len(tecs)),
        ("Pestañas revisadas", len(revisadas)),
        ("Fichas revisadas", total_fichas),
        ("Fichas con alguna no elegible", con),
        ("Menciones (ficha y tecnología)", len(coinc)),
        ("Tecnologías encontradas", len(encontradas)),
    ]
    nueva_txt = (lambda t: "") if nuevas is None else (lambda t: "Sí" if t in nuevas else "No")

    res = [[(l1, NEGRITA)], [l2], [l3], [], [("Concepto", NEGRITA), ("Valor", NEGRITA)]]
    res += [[k, v] for k, v in control]
    if avisos:
        res += [[], [("Avisos", NEGRITA)]] + [[a] for a in avisos]
    h_coinc = [[(c, NEGRITA) for c in ["Pestaña", "Renglón", COL_FICHA, COL_APP,
                                       "Tecnología no elegible", COL_MOTIVO,
                                       "Nueva en este corte"]]]
    for c in coinc:
        h_coinc.append([c["pestana"], c["renglon"], c["ficha"], c["texto"], c["tec"],
                        c["motivo"], nueva_txt(c["tec"])])
    h_pest = [[(c, NEGRITA) for c in ["Pestaña", "Fichas revisadas",
                                      "Fichas con alguna no elegible", "Dónde se buscó"]]]
    h_pest += [[p["nombre"], p["fichas"], p["con"], p["donde"]] for p in pestanas]
    h_tec = [[(c, NEGRITA) for c in ["Tecnología", COL_MOTIVO, "Fichas que la mencionan",
                                     "Nueva en este corte"]]]
    h_tec += [[t["nombre"], t["motivo"], len(por_tec[t["nombre"]]), nueva_txt(t["nombre"])]
              for t in tecs]
    escribir_xlsx(ruta_detalle, [
        ("Resumen", res, [90, 30], False),
        ("Coincidencias", h_coinc, [22, 9, 36, 40, 24, 50, 12], True),
        ("Por pestaña", h_pest, [24, 16, 28, 20], True),
        ("Por tecnología", h_tec, [24, 50, 22, 12], True),
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
    print("Control de la revisión:")
    for k, v in control:
        print("  %s: %s" % (k, miles(v) if isinstance(v, int) else v))
    print()
    print("Fichas encontradas:")
    for c in coinc:
        print("  %s, renglón %d: %s -> %s" % (c["pestana"], c["renglon"], c["ficha"] or
                                              "(sin ficha)", c["tec"]))
    print()
    print("Archivos:")
    print("  %s" % ruta_detalle)
    print("  %s" % ruta_resumen)
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Busca en cada pestaña de TeachingHub las fichas que mencionan una "
                    "tecnología no elegible (skill tecnologias-no-elegibles).")
    ap.add_argument("matriz", help="archivo matriz: .xlsx con todas las pestañas, o carpeta "
                                   "con un .csv por pestaña")
    ap.add_argument("listado", help="listado de tecnologías no elegibles (.xlsx o .csv)")
    ap.add_argument("--anterior", default=None,
                    help="listado del corte anterior, para decir qué tecnologías son nuevas")
    ap.add_argument("--en-toda-la-fila", action="store_true",
                    help="buscar en todas las celdas del renglón, no solo en AppOTecnologia")
    ap.add_argument("--salida", default=None,
                    help="carpeta de salida (por omisión, salida/ junto al archivo matriz)")
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude revisar: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
