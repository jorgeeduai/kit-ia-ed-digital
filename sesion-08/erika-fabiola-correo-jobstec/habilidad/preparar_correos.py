#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
preparar_correos.py: arma un borrador de correo por renglon del Excel que llega del Forms de
altas de profesores, con la plantilla de instrucciones para crear el perfil en JobsTec.
Es parte del skill correo-jobstec (curso Ed Digital, borrador post-S7 para Erika Fabiola).
Solo biblioteca estandar: corre con cualquier python3, sin instalar nada.

EL SCRIPT NO MANDA NADA. Solo prepara los borradores; mandarlos es la entrega y la hace la
persona o TECgpt, con su confirmacion.

Uso:
    python3 preparar_correos.py ARCHIVO
    python3 preparar_correos.py ARCHIVO --salida CARPETA --registro RUTA --plantilla RUTA
    python3 preparar_correos.py ARCHIVO --columna-nombre "Nombre" --columna-correo "Correo personal"

ARCHIVO es .xlsx o .csv (del Excel se lee la primera hoja). Por omision espera las columnas
"Nombre del profesor" y "Correo electrónico" (sin distinguir mayusculas, acentos ni espacios).

La plantilla es el bloque ```text de la seccion "Plantilla del correo" del SKILL.md que esta
junto a este script (o el archivo de --plantilla). Su primera linea es "Asunto: ...". Lo que va
entre llaves, como {Nombre del profesor}, se llena con la columna de ese nombre.

Registro de atendidos: un CSV (por omision salida/atendidos-jobstec.csv) con un renglon por
correo que ya tiene borrador. Un correo que ya esta en el registro no lleva borrador otra vez.
SUPUESTO: "atendido" quiere decir "ya se le preparo borrador", no "ya se le mando".

Deja en la carpeta salida/ (junto a la carpeta entrada/ donde esta el archivo; si no, dentro
de la carpeta del archivo; o en --salida):
    borradores-NOMBRE.xlsx    hojas Resumen, Borradores, Por revisar y Sin borrador
    borradores-NOMBRE.txt     cada borrador en texto: Para, Asunto y el cuerpo
    resumen-NOMBRE.txt        las tres lineas del resumen
    atendidos-jobstec.csv     el registro, con los correos de esta corrida agregados al final
Si un archivo de salida ya existe, no lo sobrescribe: agrega -2, -3. No modifica el archivo
del Forms.
"""
import argparse
import csv
import datetime
import io
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

FILAS_ENCABEZADO = 10      # el encabezado se busca en los primeros 10 renglones
COL_NOMBRE = "Nombre del profesor"
COL_CORREO = "Correo electrónico"
REGISTRO = "atendidos-jobstec.csv"
REG_COLUMNAS = ["Correo", "Nombre del profesor", "Fecha", "Archivo", "Renglón"]
TITULO_PLANTILLA = "plantilla del correo"

BORRADOR = "Borrador"
YA_ATENDIDO = "Ya atendido"
REPETIDO = "Repetido en este archivo"
REVISAR = "Por revisar"


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


def cuantos(n, plural, singular):
    return "%s %s" % (miles(n), plural if n != 1 else singular)


def correo_limpio(c):
    """El correo como se compara y se usa: sin espacios en los extremos y en minusculas."""
    return str(c).strip().lower()


def por_que_no_sirve(c):
    """Devuelve el motivo si el correo no sirve, o "" si sirve. Revision simple, sin regex."""
    if not c:
        return "sin correo"
    if any(ch.isspace() for ch in c):
        return "el correo trae espacios por dentro"
    if c.count("@") != 1:
        return "el correo no trae arroba" if "@" not in c else "el correo trae más de una arroba"
    usuario, dominio = c.split("@")
    if not usuario:
        return "el correo no trae nada antes de la arroba"
    if "." not in dominio or dominio.startswith(".") or dominio.endswith(".") or ".." in dominio:
        return "el dominio del correo está incompleto"
    return ""


def entre(texto, abre, cierra):
    """Los textos que van entre abre y cierra, en orden, sin repetir."""
    vistos, i = [], 0
    while True:
        a = texto.find(abre, i)
        if a < 0:
            return vistos
        b = texto.find(cierra, a + 1)
        if b < 0:
            return vistos
        dentro = texto[a + 1:b]
        if dentro and abre not in dentro and dentro not in vistos:
            vistos.append(dentro)
        i = b + 1


def ruta_libre(ruta):
    """La ruta si no existe; si existe, la misma con -2, -3... para no sobrescribir."""
    if not ruta.exists():
        return ruta
    n = 2
    while True:
        otra = ruta.with_name("%s-%d%s" % (ruta.stem, n, ruta.suffix))
        if not otra.exists():
            return otra
        n += 1


# ------------------------------------------------------------------ lectura
# Las funciones de lectura de comparar_versiones.py (Sesion 7) y limpiar_reporte.py (Sesion 6).
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
        raise Falla("El archivo %s llegó vacío: pesa 0 bytes. Revisa que sea el archivo "
                    "correcto." % ruta.name)
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


# ------------------------------------------------------------------ plantilla
def leer_plantilla(ruta_plantilla):
    """Devuelve (asunto, cuerpo, origen). Sin --plantilla, la lee del SKILL.md de junto."""
    if ruta_plantilla:
        ruta = Path(ruta_plantilla).expanduser()
        if not ruta.is_file():
            raise Falla("No encontré la plantilla %s. Revisa la ruta." % ruta)
        texto = ruta.read_text(encoding="utf-8-sig")
        origen = ruta.name
    else:
        ruta = Path(__file__).resolve().parent / "SKILL.md"
        if not ruta.is_file():
            raise Falla("No encontré el SKILL.md junto al script, que es donde vive la plantilla "
                        "del correo. Copia la carpeta completa de la habilidad, o pasa la "
                        "plantilla con --plantilla ARCHIVO.txt.")
        lineas = ruta.read_text(encoding="utf-8-sig").splitlines()
        inicio = None
        for i, l in enumerate(lineas):
            if l.startswith("#") and fold(l.lstrip("#")) == TITULO_PLANTILLA:
                inicio = i
                break
        if inicio is None:
            raise Falla("El SKILL.md no trae la sección \"## Plantilla del correo\". Vuelve a "
                        "ponerla con el texto del correo dentro de un bloque ```text.")
        a = next((i for i in range(inicio + 1, len(lineas))
                  if lineas[i].strip().startswith("```")), None)
        b = next((i for i in range(a + 1, len(lineas)) if lineas[i].strip() == "```"), None) \
            if a is not None else None
        if a is None or b is None:
            raise Falla("La sección \"Plantilla del correo\" del SKILL.md no trae el texto "
                        "dentro de un bloque que abre con ```text y cierra con ```.")
        texto = "\n".join(lineas[a + 1:b])
        origen = "SKILL.md, sección Plantilla del correo"
    lineas = texto.strip("\n").splitlines()
    if not lineas or not fold(lineas[0]).startswith("asunto:"):
        raise Falla("La primera línea de la plantilla tiene que ser el asunto, así: "
                    "\"Asunto: Instrucciones para crear tu perfil en JobsTec\".")
    asunto = lineas[0].split(":", 1)[1].strip()
    cuerpo = "\n".join(lineas[1:]).strip("\n")
    if not asunto or not cuerpo.strip():
        raise Falla("La plantilla está incompleta: necesita el asunto en la primera línea y "
                    "debajo el texto del correo.")
    return asunto, cuerpo, origen


def llenar(texto, valores):
    for nombre, v in valores.items():
        texto = texto.replace("{%s}" % nombre, v)
    return texto


# ------------------------------------------------------------------ el archivo del Forms
def leer_forms(ruta, col_nombre, col_correo, campos):
    """Encuentra el encabezado y devuelve los renglones con su numero de renglon de Excel."""
    filas = leer_filas(ruta)
    if not any(limpio(c) for f in filas for c in f):
        raise Falla("El archivo %s llegó vacío: no trae ningún renglón. Revisa que sea el "
                    "Excel correcto del Forms." % ruta.name)
    necesarias = [col_nombre, col_correo] + [c for c in campos
                                            if fold(c) not in (fold(col_nombre), fold(col_correo))]
    nec_f = [fold(c) for c in necesarias]
    enc = None
    for i, f in enumerate(filas[:FILAS_ENCABEZADO]):
        nombres = [fold(x) for x in f]
        if fold(col_nombre) in nombres and fold(col_correo) in nombres:
            enc = i
            break
    if enc is None:
        candidatas = [f for f in filas[:FILAS_ENCABEZADO] if any(limpio(x) for x in f)]
        mejor = max(candidatas, key=lambda f: sum(1 for x in f if limpio(x))) if candidatas else []
        nombres = [fold(x) for x in mejor]
        faltan = [c for c in (col_nombre, col_correo) if fold(c) not in nombres]
        raise Falla("%s %s en %s: no %s encontré en los primeros %d renglones. Columnas que sí "
                    "hay: %s. Si en tu archivo se llaman de otra forma, pásalas con "
                    "--columna-nombre y --columna-correo."
                    % ("Falta la columna" if len(faltan) == 1 else "Faltan las columnas",
                       y_lista('"%s"' % c for c in faltan), ruta.name,
                       "la" if len(faltan) == 1 else "las", FILAS_ENCABEZADO,
                       " | ".join(limpio(x) for x in mejor if limpio(x)) or "ninguna"))
    encabezado = [limpio(x) for x in filas[enc]]
    pos = {}
    for j, nombre in enumerate(encabezado):
        if not nombre:
            continue
        k = fold(nombre)
        if k in pos and k in nec_f:
            raise Falla("La columna \"%s\" aparece dos veces en %s (columnas %d y %d). No sé "
                        "cuál es la buena: en un Excel de Forms, una puede ser el correo de quien "
                        "llenó el formulario. En una copia del archivo, cambia el nombre de la "
                        "columna que no es (por ejemplo, a \"Correo de quien llena\") y vuelve "
                        "a correr con la copia."
                        % (nombre, ruta.name, pos[k] + 1, j + 1))
        pos.setdefault(k, j)
    faltan = [c for c in necesarias if fold(c) not in pos]
    if faltan:
        raise Falla("La plantilla usa %s, pero %s no %s esa columna. Columnas que sí hay: %s. "
                    "Corrige el nombre entre llaves en la plantilla o agrega la columna."
                    % (y_lista("{%s}" % c for c in faltan), ruta.name, "trae",
                       " | ".join(x for x in encabezado if x)))
    renglones = []
    for r, f in enumerate(filas[enc + 1:], start=enc + 2):
        if not any(limpio(c) for c in f):
            continue
        f = list(f) + [""] * (len(encabezado) - len(f))
        valores = {c: limpio(f[pos[fold(c)]]) for c in necesarias}
        renglones.append(dict(renglon=r, nombre=valores[col_nombre],
                              correo=correo_limpio(valores[col_correo]), valores=valores))
    if not renglones:
        raise Falla("El archivo %s trae el encabezado (renglón %d) pero ningún renglón de datos. "
                    "Todavía no hay respuestas del Forms, o no es el archivo correcto."
                    % (ruta.name, enc + 1))
    return dict(encabezado_renglon=enc + 1, renglones=renglones,
                nombre_col=encabezado[pos[fold(col_nombre)]],
                correo_col=encabezado[pos[fold(col_correo)]])


# ------------------------------------------------------------------ registro
def leer_registro(ruta):
    """Los correos que ya tienen borrador, del registro de atendidos (si existe)."""
    if not ruta.is_file() or ruta.stat().st_size == 0:
        return set()
    with ruta.open(encoding="utf-8-sig", newline="") as fh:
        lector = csv.reader(fh)
        encabezado = next(lector, [])
        if not encabezado or fold(encabezado[0]) != "correo":
            raise Falla("El registro %s no tiene el formato esperado (la primera columna se "
                        "llama Correo). Si lo editaste, revisa que siga así; si no es el "
                        "registro, pasa la ruta correcta con --registro." % ruta)
        return {correo_limpio(f[0]) for f in lector if f and f[0].strip()}


def agregar_registro(ruta, renglones):
    nuevo = not ruta.is_file() or ruta.stat().st_size == 0
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with ruta.open("a", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        if nuevo:
            w.writerow(REG_COLUMNAS)
        for fila in renglones:
            w.writerow(fila)


# ------------------------------------------------------------------ clasificar
def clasificar(renglones, atendidos):
    """Un estado por renglon: Por revisar, Ya atendido, Repetido en este archivo o Borrador."""
    vistos = {}
    for g in renglones:
        motivo = por_que_no_sirve(g["correo"])
        if not g["nombre"]:
            motivo = "sin nombre" + ("; " + motivo if motivo else "")
        if motivo:
            g["estado"], g["motivo"] = REVISAR, motivo
        elif g["correo"] in atendidos:
            g["estado"], g["motivo"] = YA_ATENDIDO, "el correo ya está en el registro de atendidos"
        elif g["correo"] in vistos:
            g["estado"] = REPETIDO
            g["motivo"] = "mismo correo que el renglón %d" % vistos[g["correo"]]
        else:
            g["estado"], g["motivo"] = BORRADOR, ""
            vistos[g["correo"]] = g["renglon"]
    return renglones


# ------------------------------------------------------------------ escritura xlsx
# Las funciones de escritura de comparar_versiones.py (Sesion 7).
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
def carpeta_salida(archivo, salida):
    if salida:
        return Path(salida).expanduser()
    if fold(archivo.parent.name) in ("entrada", "versiones"):
        return archivo.parent.parent / "salida"
    return archivo.parent / "salida"


def principal(args):
    ruta = Path(args.archivo).expanduser()
    asunto_p, cuerpo_p, origen_p = leer_plantilla(args.plantilla)
    campos = entre(asunto_p + "\n" + cuerpo_p, "{", "}")
    por_llenar = [c for c in entre(asunto_p + "\n" + cuerpo_p, "[", "]")]
    forms = leer_forms(ruta, args.columna_nombre, args.columna_correo, campos)

    out = carpeta_salida(ruta, args.salida)
    ruta_reg = Path(args.registro).expanduser() if args.registro else out / REGISTRO
    atendidos = leer_registro(ruta_reg)
    renglones = clasificar(forms["renglones"], atendidos)

    # Los nombres de la plantilla se llenan con la columna que nombran
    nombre_real = {fold(c): c for c in campos}
    borradores = []
    for g in renglones:
        if g["estado"] != BORRADOR:
            continue
        valores = {}
        for c in campos:
            v = g["valores"][next(k for k in g["valores"] if fold(k) == fold(c))]
            valores[nombre_real[fold(c)]] = g["correo"] if fold(c) == fold(args.columna_correo) else v
        borradores.append(dict(g, asunto=llenar(asunto_p, valores), cuerpo=llenar(cuerpo_p, valores)))

    cuenta = {e: sum(1 for g in renglones if g["estado"] == e)
              for e in (BORRADOR, YA_ATENDIDO, REPETIDO, REVISAR)}
    out.mkdir(parents=True, exist_ok=True)
    base = ruta.stem
    ruta_xlsx = ruta_libre(out / ("borradores-%s.xlsx" % base))
    ruta_txt = ruta_libre(out / ("borradores-%s.txt" % base))
    ruta_res = ruta_libre(out / ("resumen-%s.txt" % base))

    revisar = [g for g in renglones if g["estado"] == REVISAR]
    l1 = "Borradores de correo de JobsTec a partir de %s." % ruta.name
    l2 = "Hay %s, %s, %s y %s (de %s)." % (
        cuantos(cuenta[BORRADOR], "borradores nuevos", "borrador nuevo"),
        cuantos(cuenta[YA_ATENDIDO], "renglones ya atendidos", "renglón ya atendido"),
        cuantos(cuenta[REPETIDO], "repetidos en el mismo archivo", "repetido en el mismo archivo"),
        cuantos(cuenta[REVISAR], "por revisar", "por revisar"),
        cuantos(len(renglones), "renglones", "renglón"))
    l3 = "Detalle: %s." % ruta_xlsx.name
    if revisar:
        l3 += " Por revisar: %s %s." % ("renglón" if len(revisar) == 1 else "renglones",
                                        y_lista(str(g["renglon"]) for g in revisar))
    if por_llenar:
        l3 += " Aviso: la plantilla todavía es de ejemplo; no mandes estos borradores."
    lineas = [l1, l2, l3]

    control = [
        ("Archivo del Forms", ruta.name),
        ("Renglón del encabezado", forms["encabezado_renglon"]),
        ("Columnas usadas", "%s y %s" % (forms["nombre_col"], forms["correo_col"])),
        ("Plantilla", origen_p),
        ("Registro de atendidos", str(ruta_reg)),
        ("Correos en el registro antes de esta corrida", len(atendidos)),
        ("Renglones con datos", len(renglones)),
        ("Borradores nuevos", cuenta[BORRADOR]),
        ("Ya atendidos", cuenta[YA_ATENDIDO]),
        ("Repetidos en el mismo archivo", cuenta[REPETIDO]),
        ("Por revisar", cuenta[REVISAR]),
    ]

    # Excel de detalle
    res = [[(l, NEGRITA)] if i == 0 else [l] for i, l in enumerate(lineas)]
    res += [[], [("Concepto", NEGRITA), ("Valor", NEGRITA)]] + [[k, v] for k, v in control]
    if por_llenar:
        res += [[], [("Aviso", NEGRITA)],
                ["La plantilla todavía trae textos de ejemplo entre corchetes: %s. Sustitúyelos "
                 "en el SKILL.md antes de mandar." % y_lista("[%s]" % p for p in por_llenar)]]
    h_borr = [[(c, NEGRITA) for c in ("Renglón", "Nombre del profesor", "Para", "Asunto", "Cuerpo")]]
    h_borr += [[b["renglon"], b["nombre"], b["correo"], b["asunto"], b["cuerpo"]] for b in borradores]
    h_rev = [[(c, NEGRITA) for c in ("Renglón", "Nombre del profesor", "Correo electrónico", "Motivo")]]
    h_rev += [[g["renglon"], g["nombre"], g["valores"][args.columna_correo], g["motivo"]]
              for g in revisar]
    h_sin = [[(c, NEGRITA) for c in ("Renglón", "Nombre del profesor", "Correo electrónico", "Estado", "Motivo")]]
    h_sin += [[g["renglon"], g["nombre"], g["correo"], g["estado"], g["motivo"]]
              for g in renglones if g["estado"] in (YA_ATENDIDO, REPETIDO)]
    escribir_xlsx(ruta_xlsx, [
        ("Resumen", res, [80, 50], False),
        ("Borradores", h_borr, [9, 32, 36, 44, 80], True),
        ("Por revisar", h_rev, [9, 32, 36, 44], True),
        ("Sin borrador", h_sin, [9, 32, 36, 24, 44], True),
    ])

    # Texto de los borradores
    partes = ["BORRADORES DE CORREO, JobsTec. Archivo: %s. No se ha mandado ninguno." % ruta.name]
    if por_llenar:
        partes.append("AVISO: la plantilla todavía es de ejemplo; no mandes estos borradores.")
    for n, b in enumerate(borradores, start=1):
        partes.append("\n".join([
            "=" * 60,
            "Borrador %d de %d (renglón %d)" % (n, len(borradores), b["renglon"]),
            "Para: %s" % b["correo"],
            "Asunto: %s" % b["asunto"],
            "",
            b["cuerpo"]]))
    ruta_txt.write_text("\n\n".join(partes) + "\n", encoding="utf-8")
    ruta_res.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    hoy = datetime.date.today().isoformat()
    agregar_registro(ruta_reg, [[b["correo"], b["nombre"], hoy, ruta.name, b["renglon"]]
                                for b in borradores])

    if por_llenar:
        print("Aviso: la plantilla todavía trae textos de ejemplo entre corchetes: %s."
              % y_lista("[%s]" % p for p in por_llenar))
        print()
    print("Resumen:")
    for l in lineas:
        print(l)
    print()
    print("Control:")
    for k, v in control:
        print("  %s: %s" % (k, miles(v) if isinstance(v, int) else v))
    print("  Comprobación: %s + %s + %s + %s = %s renglones"
          % (cuenta[BORRADOR], cuenta[YA_ATENDIDO], cuenta[REPETIDO], cuenta[REVISAR],
             sum(cuenta.values())))
    print()
    print("Archivos:")
    for r in (ruta_xlsx, ruta_txt, ruta_res, ruta_reg):
        print("  %s" % r)
    print()
    print("No se mandó ningún correo.")
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Arma un borrador de correo de JobsTec por renglón del Excel del Forms "
                    "(skill correo-jobstec). No manda nada.")
    ap.add_argument("archivo", help="el Excel (.xlsx) o .csv que llega del Forms")
    ap.add_argument("--salida", default=None,
                    help="carpeta de salida (por omisión, salida/ junto a la carpeta del archivo)")
    ap.add_argument("--registro", default=None,
                    help="registro de atendidos (por omisión, salida/%s)" % REGISTRO)
    ap.add_argument("--plantilla", default=None,
                    help="archivo .txt con la plantilla (por omisión, la del SKILL.md)")
    ap.add_argument("--columna-nombre", default=COL_NOMBRE)
    ap.add_argument("--columna-correo", default=COL_CORREO)
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude preparar los borradores: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
