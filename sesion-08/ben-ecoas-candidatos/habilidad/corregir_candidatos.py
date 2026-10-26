#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
corregir_candidatos.py: corrige la columna Alumnos Candidatos del archivo de ECOAs de tutores
con lo que se mando antes a escolar. En el archivo de ECOAs, cuando un CRN tiene varios
tutores, cada tutor trae la suma de los alumnos de todo el CRN; este script le pone a cada
tutor solo sus alumnos, contados en el archivo de escolar por CRN y nomina.
Es parte del borrador de habilidad corregir-candidatos-ecoas (curso Ed Digital, propuesta de
Jorge para el flujo de Ben, ronda 2 despues de la Sesion 7). Solo biblioteca estandar: corre
con cualquier python3, sin instalar nada.

Uso:
    python3 corregir_candidatos.py ECOAS ESCOLAR
    python3 corregir_candidatos.py ECOAS ESCOLAR --nomina-escolar "Nómina del tutor"

ECOAS y ESCOLAR son .xlsx o .csv (del Excel se lee la primera hoja).

SUPUESTOS (estan en la seccion Lo que supuse del LEEME.md y se confirman con Ben):
  - Los encabezados del archivo de ECOAs son los que Ben escribio en el muro: CRN, Clave,
    Grupo, Alumnos Candidatos, Nómina del Profesor, Nombre del profesor.
  - El archivo que se mando a escolar es SUPUESTO: un renglon por alumno, con CRN, Matrícula y
    Nómina del Profesor (la nomina del tutor que le toca). Si en lugar de un renglon por alumno
    trae ya el numero por tutor (una columna como "Alumnos asignados"), el script suma esa
    columna en lugar de contar renglones.
  - La llave es CRN mas Nómina del Profesor.

Deja en la carpeta salida/ (junto a la carpeta ficticio/ o entrada/ donde esta el archivo de
ECOAs; si no, dentro de su carpeta; o en --salida):
    candidatos-corregidos-NOMBRE.xlsx          hojas Resumen, ECOAs corregidas, Cambios y
                                               Por revisar
    resumen-candidatos-corregidos-NOMBRE.txt   las tres lineas del resumen
donde NOMBRE es el nombre del archivo de ECOAs sin extension. No modifica las entradas.
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

FILAS_ENCABEZADO = 10

# ------------------------------------------------------------------ SUPUESTOS
COL_CRN = "CRN"
COL_NOMINA = "Nómina del Profesor"
COL_CANDIDATOS = "Alumnos Candidatos"
COL_NOMBRE = "Nombre del profesor"
COL_MATRICULA = "Matrícula"          # SUPUESTA, en el archivo de escolar
# Si el archivo de escolar trae el numero por tutor, se suma la primera de estas columnas
COLS_CONTEO = ["Alumnos asignados", "Alumnos Candidatos", "Número de alumnos", "Alumnos"]

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


# ------------------------------------------------------------------ principal
def entero(v):
    c = canonico(v)
    try:
        return int(float(c))
    except ValueError:
        return None


def principal(args):
    ruta_e = Path(args.ecoas).expanduser()
    ruta_s = Path(args.escolar).expanduser()
    nom_e = args.nomina_ecoas or COL_NOMINA
    nom_s = args.nomina_escolar or COL_NOMINA
    crn_c = args.crn or COL_CRN

    ecoas = leer_tabla(ruta_e, [crn_c, nom_e, COL_CANDIDATOS], "ECOAs de tutores",
                       "--crn o --nomina-ecoas")
    esc = leer_tabla(ruta_s, [crn_c, nom_s], "lo que se mandó a escolar",
                     "--crn o --nomina-escolar")
    avisos = ecoas["avisos"] + esc["avisos"]

    # 1. Contar los alumnos de cada tutor en el archivo de escolar
    col_conteo = next((c for c in COLS_CONTEO if tiene(esc, c)), None)
    usa_matricula = col_conteo is None and tiene(esc, COL_MATRICULA)
    conteo, vistos, repetidos, sin_llave = {}, set(), 0, 0
    for r, f in esc["datos"]:
        k = (canonico(celda(esc, f, crn_c)), nomina(celda(esc, f, nom_s)))
        if not k[0] or not k[1]:
            sin_llave += 1
            continue
        if col_conteo:
            n = entero(celda(esc, f, col_conteo))
            if n is None:
                avisos.append("Renglón %d de %s: %s no es un número; no cuenta."
                              % (r, ruta_s.name, col_conteo))
                continue
            conteo[k] = conteo.get(k, 0) + n
            continue
        if usa_matricula:
            m = fold(celda(esc, f, COL_MATRICULA))
            if m and (k, m) in vistos:
                repetidos += 1
                continue
            vistos.add((k, m))
        conteo[k] = conteo.get(k, 0) + 1
    if repetidos:
        avisos.append("%d %s en %s (mismo CRN, tutor y matrícula): se cuenta una vez."
                      % (repetidos, "alumno viene repetido" if repetidos == 1
                         else "alumnos vienen repetidos", ruta_s.name))
    por_crn = {}
    for (crn, _), n in conteo.items():
        por_crn[crn] = por_crn.get(crn, 0) + n

    # 2. Revisar cada tutor del archivo de ECOAs
    j_cand = ecoas["col"][fold(COL_CANDIDATOS)]
    filas_out, cambios, revisar, vistas = [], [], [], set()
    cuenta = dict(suma=0, otro=0, bien=0, revisar=0, repetido=0)
    antes_total = despues_total = 0
    crns = set()
    for r, f in ecoas["datos"]:
        crn = canonico(celda(ecoas, f, crn_c))
        n = nomina(celda(ecoas, f, nom_e))
        orig = entero(f[j_cand])
        nueva = list(f)
        if not crn or not n:
            qp = "Por revisar: sin CRN o sin nómina"
            cuenta["revisar"] += 1
            revisar.append([crn, n, celda(ecoas, f, COL_NOMBRE), f[j_cand], qp])
        elif (crn, n) in vistas:
            qp = "Por revisar: el tutor aparece dos veces en el mismo CRN"
            cuenta["repetido"] += 1
            revisar.append([crn, n, celda(ecoas, f, COL_NOMBRE), f[j_cand], qp])
        elif (crn, n) not in conteo:
            qp = "Por revisar: el tutor no está en el archivo de escolar"
            cuenta["revisar"] += 1
            revisar.append([crn, n, celda(ecoas, f, COL_NOMBRE), f[j_cand], qp])
        else:
            correcto = conteo[(crn, n)]
            if orig == correcto:
                qp = "Ya estaba bien"
                cuenta["bien"] += 1
            else:
                if orig is not None and orig == por_crn.get(crn) and orig != correcto:
                    qp = "Corregido: traía la suma del CRN"
                    cuenta["suma"] += 1
                else:
                    qp = "Corregido: traía otro número"
                    cuenta["otro"] += 1
                nueva[j_cand] = correcto
                cambios.append([crn, celda(ecoas, f, "Clave"), celda(ecoas, f, "Grupo"), n,
                                celda(ecoas, f, COL_NOMBRE), orig if orig is not None
                                else limpio(f[j_cand]), correcto, qp.split(": ", 1)[1]])
        vistas.add((crn, n))
        crns.add(crn)
        antes_total += orig or 0
        despues_total += entero(nueva[j_cand]) or 0
        filas_out.append((nueva, f[j_cand], qp))
    solo_escolar = sorted(k for k in conteo if k not in vistas)
    for crn, n in solo_escolar:
        revisar.append([crn, n, "", conteo[(crn, n)],
                        "Por revisar: está en escolar y no en el archivo de ECOAs"])
    if cuenta["repetido"]:
        avisos.append("%d tutores aparecen dos veces en el mismo CRN en %s: el segundo renglón "
                      "no se corrige." % (cuenta["repetido"], ruta_e.name))

    out = carpeta_salida(ruta_e, args.salida)
    out.mkdir(parents=True, exist_ok=True)
    base = ruta_e.stem
    nombre_detalle = "candidatos-corregidos-%s.xlsx" % base
    ruta_detalle = out / nombre_detalle
    ruta_resumen = out / ("resumen-candidatos-corregidos-%s.txt" % base)

    corregidos = cuenta["suma"] + cuenta["otro"]
    por_revisar = cuenta["revisar"] + cuenta["repetido"] + len(solo_escolar)
    l1 = ("Corrección de %s de %s con %s, por %s y %s."
          % (COL_CANDIDATOS, ruta_e.name, ruta_s.name, crn_c, nom_e))
    l2 = ("De %s tutores en %s CRN: %s corregidos (%s traían la suma del CRN y %s otro número), "
          "%s ya estaban bien y %s por revisar."
          % (miles(len(filas_out)), miles(len(crns)), miles(corregidos), miles(cuenta["suma"]),
             miles(cuenta["otro"]), miles(cuenta["bien"]), miles(por_revisar)))
    l3 = ("Alumnos candidatos en total: %s antes y %s después; en el archivo de escolar, %s. "
          "Detalle: %s." % (miles(antes_total), miles(despues_total),
                            miles(sum(conteo.values())), nombre_detalle))
    lineas = [l1, l2, l3]

    control = [
        ("Archivo de ECOAs", ruta_e.name),
        ("Archivo de escolar", ruta_s.name),
        ("Llave", "%s y %s" % (crn_c, nom_e)),
        ("Renglón del encabezado (ECOAs, escolar)", "%d, %d" % (ecoas["renglon_enc"],
                                                                esc["renglon_enc"])),
        ("Cómo se cuentan los alumnos de escolar",
         "suma de la columna %s" % col_conteo if col_conteo else
         ("un renglón por alumno, sin repetir matrícula" if usa_matricula
          else "un renglón por alumno")),
        ("Renglones de ECOAs (tutores)", len(filas_out)),
        ("CRN en ECOAs", len(crns)),
        ("Renglones en escolar", len(esc["datos"])),
        ("Tutores en escolar (CRN y nómina)", len(conteo)),
        ("Corregidos: traían la suma del CRN", cuenta["suma"]),
        ("Corregidos: traían otro número", cuenta["otro"]),
        ("Ya estaban bien", cuenta["bien"]),
        ("Por revisar: tutor sin alumnos en escolar", cuenta["revisar"]),
        ("Por revisar: tutor solo en escolar", len(solo_escolar)),
    ]
    if cuenta["repetido"]:
        control.append(("Por revisar: tutor repetido en el mismo CRN", cuenta["repetido"]))
    if repetidos:
        control.append(("Alumnos repetidos en escolar (se cuentan una vez)", repetidos))
    if sin_llave:
        control.append(("Renglones de escolar sin CRN o sin nómina (no cuentan)", sin_llave))

    res = [[(lineas[0], NEGRITA)], [lineas[1]], [lineas[2]], [],
           [("Concepto", NEGRITA), ("Valor", NEGRITA)]]
    res += [[k, v] for k, v in control]
    res += [[], [("Supuestos de esta propuesta", NEGRITA)],
            ["El archivo de escolar trae un renglón por alumno con CRN, Matrícula y Nómina del "
             "Profesor (encabezados supuestos)."],
            ["La llave que une a cada tutor con sus alumnos es CRN más nómina."]]
    if avisos:
        res += [[], [("Avisos", NEGRITA)]] + [[a] for a in avisos]

    enc = ecoas["encabezado"]
    h_ecoas = [[(c, NEGRITA) for c in enc] + [("%s (antes)" % COL_CANDIDATOS, NEGRITA),
                                              ("Qué pasó", NEGRITA)]]
    for nueva, antes, qp in filas_out:
        h_ecoas.append([a_numero(v) if not isinstance(v, int) else v for v in nueva[:len(enc)]]
                       + [a_numero(antes), qp])
    h_cambios = [[(c, NEGRITA) for c in ["CRN", "Clave", "Grupo", "Nómina del Profesor",
                                         "Nombre del profesor", "Antes", "Después", "Por qué"]]]
    h_cambios += [[a_numero(x) if isinstance(x, str) else x for x in c] for c in cambios]
    h_rev = [[(c, NEGRITA) for c in ["CRN", "Nómina del Profesor", "Nombre del profesor",
                                     "Alumnos Candidatos", "Qué revisar"]]]
    h_rev += [[a_numero(x) if isinstance(x, str) else x for x in c] for c in revisar]

    escribir_xlsx(ruta_detalle, [
        ("Resumen", res, [80, 40], False),
        ("ECOAs corregidas", h_ecoas, [max(10, min(30, len(c) + 4)) for c in enc] + [14, 34],
         True),
        ("Cambios", h_cambios, [10, 10, 8, 14, 30, 8, 9, 22], True),
        ("Por revisar", h_rev, [10, 14, 30, 12, 50], True),
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
    if revisar:
        print()
        print("Por revisar:")
        for c in revisar:
            print("  CRN %s, nómina %s: %s" % (c[0], c[1], c[4].split(": ", 1)[1]))
    print()
    print("Archivos:")
    print("  %s" % ruta_detalle)
    print("  %s" % ruta_resumen)
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Corrige Alumnos Candidatos de las ECOAs de tutores "
                    "(habilidad corregir-candidatos-ecoas).")
    ap.add_argument("ecoas", help="archivo de ECOAs de tutores (.xlsx o .csv)")
    ap.add_argument("escolar", help="archivo que se mandó a escolar (.xlsx o .csv)")
    ap.add_argument("--crn", default=None, help="columna del CRN (por omisión: CRN)")
    ap.add_argument("--nomina-ecoas", default=None,
                    help="columna de nómina en ECOAs (por omisión: %s)" % COL_NOMINA)
    ap.add_argument("--nomina-escolar", default=None,
                    help="columna de nómina en escolar (por omisión: %s)" % COL_NOMINA)
    ap.add_argument("--salida", default=None, help="carpeta de salida")
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude corregir: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
