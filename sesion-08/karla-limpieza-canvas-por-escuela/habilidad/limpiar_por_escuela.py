#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
limpiar_por_escuela.py: limpia los archivos de actividad de Canvas que se descargan uno por
escuela y deja, por cada uno, un Excel con el nombre de la escuela que solo trae los grupos
digitales. Es parte del skill limpieza-canvas-por-escuela (curso Ed Digital, propuesta de skill
del grupo, ronda 2). Solo biblioteca estandar: corre con cualquier python3, sin instalar nada.
Si openpyxl esta instalado (TECgpt lo trae), lo usa para leer; si no, lee con la biblioteca
estandar.

La regla (SUPUESTA a partir de lo que se describio en la Sesion 7; ver LEEME.md):
  1. En toda hoja que traiga la columna Rol, se quitan los renglones con Rol = Estudiante.
  2. En toda hoja que traiga la columna SIS ID del curso, se quedan solo los renglones cuyo
     codigo empieza con "Nacional" (sin distinguir mayusculas ni espacios al inicio).
  3. Las demas hojas se copian igual. El nombre de la escuela se lee de la primera hoja, en la
     celda a la derecha de la que dice "Escuela".

Uso:
    python3 limpiar_por_escuela.py CARPETA_DE_DESCARGAS
    python3 limpiar_por_escuela.py ARCHIVO1.xlsx ARCHIVO2.xlsx ... --salida CARPETA
    python3 limpiar_por_escuela.py CARPETA --reemplazar --solo-profesores
    python3 limpiar_por_escuela.py CARPETA --col-rol "Role" --col-codigo "Course SIS ID"

Deja en limpios/ (junto a la carpeta de descargas; o en --salida):
    NOMBRE DE LA ESCUELA.xlsx   uno por archivo, con las mismas hojas ya limpias
    control-limpieza.xlsx       el conteo de control, archivo por archivo y hoja por hoja
    resumen-limpieza.txt        las tres lineas del resumen
Si ya existe el Excel limpio de una escuela, no lo toca y lo avisa (con --reemplazar, lo
reemplaza). No modifica ningun archivo de entrada.
"""
import argparse
import datetime
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

FILAS_ENCABEZADO = 10           # el encabezado se busca en los primeros 10 renglones
ETIQUETA_ESCUELA = "Escuela"    # en la primera hoja, la celda junto a esta lleva el nombre
COL_ROL = "Rol"
COL_CODIGO = "SIS ID del curso"
ROL_QUE_SE_QUITA = "Estudiante"
ROL_PROFESOR = "Profesor"
PREFIJO = "Nacional"


class Falla(Exception):
    """Un problema con un archivo que la persona tiene que resolver; se explica."""


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


def vacia(fila):
    return not any(limpio(c) for c in fila if c is not None)


def nombre_de_archivo(escuela):
    """El nombre de la escuela, sin caracteres que Windows no acepta en un archivo."""
    return "".join("-" if c in '<>:"/\\|?*' else c for c in limpio(escuela)).strip(" .")


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


def _numero(raw):
    if raw is None or raw == "":
        return ""
    try:
        f = float(raw)
    except ValueError:
        return raw
    if f == int(f) and abs(f) < 1e15:
        return int(f)
    return f


def _filas_de_hoja(datos, compartidas):
    filas, ultimo = {}, 0
    for _, row in ET.iterparse(datos):
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
                    val = _numero(raw)
            celdas[j] = val
        fila = [""] * (max(celdas) + 1) if celdas else []
        for j, val in celdas.items():
            fila[j] = val
        filas[r] = fila
        row.clear()
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
            with z.open(parte) as datos:
                hojas.append((e.get("name"), _filas_de_hoja(datos, compartidas)))
    return hojas


def _xlsx_openpyxl(ruta, openpyxl):
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hojas = []
    for ws in wb.worksheets:
        filas = [["" if v is None else v for v in fila] for fila in ws.iter_rows(values_only=True)]
        hojas.append((ws.title, filas))
    wb.close()
    return hojas


def leer_xlsx(ruta):
    """Todas las hojas: con openpyxl si esta instalado, si no con la biblioteca estandar."""
    if not ruta.exists():
        raise Falla("No encontré el archivo %s. Revisa la ruta y el nombre." % ruta)
    if ruta.stat().st_size == 0:
        raise Falla("%s llegó vacío: pesa 0 bytes. Descárgalo otra vez." % ruta.name)
    if ruta.suffix.lower() not in (".xlsx", ".xlsm"):
        raise Falla("No sé leer archivos %s (%s). Guarda la descarga como .xlsx."
                    % (ruta.suffix or "sin extensión", ruta.name))
    try:
        try:
            import openpyxl  # noqa: F401  (TECgpt lo trae; la laptop no)
        except ImportError:
            return _xlsx_biblioteca_estandar(ruta), "biblioteca estándar"
        return _xlsx_openpyxl(ruta, openpyxl), "openpyxl"
    except (zipfile.BadZipFile, ET.ParseError, KeyError, IndexError, ValueError,
            StopIteration, OSError):
        raise Falla("No pude leer %s: no parece un .xlsx válido o está dañado. Ábrelo en Excel "
                    "y guárdalo otra vez como .xlsx." % ruta.name)


def buscar_encabezado(filas, buscada):
    """Indice de la primera fila (de las primeras 10) que trae la columna buscada, y su indice."""
    b = fold(buscada)
    for i, f in enumerate(filas[:FILAS_ENCABEZADO]):
        nombres = [fold(x) for x in f]
        if b in nombres:
            return i, nombres.index(b)
    return None, None


def columnas_que_hay(hojas):
    partes = []
    for nombre, filas in hojas:
        candidatas = [f for f in filas[:FILAS_ENCABEZADO] if not vacia(f)]
        if candidatas:
            mejor = max(candidatas, key=lambda f: sum(1 for x in f if limpio(x)))
            partes.append("%s: %s" % (nombre, " | ".join(limpio(x) for x in mejor if limpio(x))))
    return "; ".join(partes) or "ninguna"


def leer_escuela(hojas, ruta):
    nombre_hoja, filas = hojas[0]
    etiqueta = fold(ETIQUETA_ESCUELA)
    for f in filas[:FILAS_ENCABEZADO * 3]:
        for j, x in enumerate(f):
            if fold(x) == etiqueta:
                resto = [limpio(v) for v in f[j + 1:] if limpio(v)]
                if resto:
                    return resto[0]
    raise Falla("No encontré el nombre de la escuela en %s: en la primera hoja (%s) no hay una "
                "celda que diga \"%s\" con el nombre a su derecha. Dime en qué hoja y en qué "
                "celda viene la escuela." % (ruta.name, nombre_hoja, ETIQUETA_ESCUELA))


# ------------------------------------------------------------------ limpieza
def limpiar_hoja(nombre, filas, col_rol, col_codigo, solo_profesores):
    """Devuelve (filas_limpias, conteo). Si la hoja no trae el codigo, se copia igual."""
    enc, j_cod = buscar_encabezado(filas, col_codigo)
    conteo = dict(hoja=nombre, regla="se copia igual", leidos=0, fuera_rol=0, fuera_codigo=0,
                  quedan=0, por_rol={}, tiene_rol=False)
    if enc is None:
        datos = [f for f in filas if not vacia(f)]
        conteo["leidos"] = conteo["quedan"] = max(len(datos) - 1, 0)
        return filas, conteo
    nombres = [fold(x) for x in filas[enc]]
    j_rol = nombres.index(fold(col_rol)) if fold(col_rol) in nombres else None
    conteo["tiene_rol"] = j_rol is not None
    conteo["regla"] = ("rol y código" if j_rol is not None else "solo código")
    prefijo = fold(PREFIJO)
    quitar = fold(ROL_QUE_SE_QUITA)
    profe = fold(ROL_PROFESOR)
    salida = filas[:enc + 1]
    for f in filas[enc + 1:]:
        if vacia(f):
            continue
        conteo["leidos"] += 1
        f = list(f) + [""] * (len(filas[enc]) - len(f))
        if j_rol is not None:
            rol = fold(f[j_rol])
            if rol == quitar or (solo_profesores and rol != profe):
                conteo["fuera_rol"] += 1
                continue
        if not fold(f[j_cod]).startswith(prefijo):
            conteo["fuera_codigo"] += 1
            continue
        conteo["quedan"] += 1
        if j_rol is not None:
            r = limpio(f[j_rol]).capitalize() or "(sin rol)"
            conteo["por_rol"][r] = conteo["por_rol"].get(r, 0) + 1
        salida.append(f)
    return salida, conteo


# ------------------------------------------------------------------ escritura xlsx
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _col_letra(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _celda(ref, v, negrita):
    s = ' s="1"' if negrita else ""
    if isinstance(v, bool):
        v = "VERDADERO" if v else "FALSO"
    if isinstance(v, (int, float)):
        return '<c r="%s"%s><v>%s</v></c>' % (ref, s, repr(v))
    if isinstance(v, (datetime.datetime, datetime.date)):
        v = v.isoformat(sep=" ") if isinstance(v, datetime.datetime) else v.isoformat()
    return ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, s, escape(str(v))))


def escribir_xlsx(ruta, hojas):
    """hojas: lista de (nombre, filas, fila_en_negrita o None). Escribe renglon por renglon."""
    n = len(hojas)
    tipos = "".join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/'
                    'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
                    for i in range(n))
    fijos = [
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
                                     % (escape(h[0][:31]), i + 1, i + 1)
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
         '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
         '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
         '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
         '</styleSheet>' % NS_MAIN),
    ]
    temporal = ruta.with_name(ruta.name + ".parcial")
    with zipfile.ZipFile(temporal, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in fijos:
            z.writestr(nombre, contenido.encode("utf-8"))
        for i, (_, filas, negrita) in enumerate(hojas):
            with z.open("xl/worksheets/sheet%d.xml" % (i + 1), "w") as w:
                w.write(('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                         '<worksheet xmlns="%s"><sheetData>' % NS_MAIN).encode("utf-8"))
                for r, fila in enumerate(filas):
                    celdas = "".join(_celda("%s%d" % (_col_letra(j), r + 1), v, r == negrita)
                                     for j, v in enumerate(fila) if v is not None and v != "")
                    if celdas:
                        w.write(('<row r="%d">%s</row>' % (r + 1, celdas)).encode("utf-8"))
                w.write(b"</sheetData></worksheet>")
    temporal.replace(ruta)


# ------------------------------------------------------------------ principal
def archivos_de_entrada(entradas):
    rutas = []
    for e in entradas:
        p = Path(e).expanduser()
        if p.is_dir():
            encontrados = sorted(x for x in p.iterdir() if x.suffix.lower() in (".xlsx", ".xlsm")
                                 and not x.name.startswith((".", "~$")))
            if not encontrados:
                raise Falla("La carpeta %s no trae ningún .xlsx." % p)
            rutas += encontrados
        else:
            rutas.append(p)
    return rutas


def principal(args):
    rutas = archivos_de_entrada(args.entradas)
    if args.salida:
        out = Path(args.salida).expanduser()
    else:
        primera = Path(args.entradas[0]).expanduser()
        out = (primera if primera.is_dir() else primera.parent).parent / "limpios"
    out.mkdir(parents=True, exist_ok=True)

    resultados, avisos, problemas, escuelas_vistas = [], [], [], {}
    como_se_leyo = set()
    for ruta in rutas:
        try:
            hojas, como = leer_xlsx(ruta)
            como_se_leyo.add(como)
            escuela = leer_escuela(hojas, ruta)
            clave = fold(escuela)
            if clave in escuelas_vistas:
                raise Falla("%s es de la misma escuela que %s (%s); no lo limpio para no "
                            "encimar uno sobre otro. Revisa si descargaste dos veces la misma "
                            "escuela." % (ruta.name, escuelas_vistas[clave], escuela))
            escuelas_vistas[clave] = ruta.name
            limpias, conteos = [], []
            for nombre, filas in hojas:
                f2, c = limpiar_hoja(nombre, filas, args.col_rol, args.col_codigo,
                                     args.solo_profesores)
                limpias.append((nombre, f2, None))
                conteos.append(c)
            if not any(c["regla"] != "se copia igual" for c in conteos):
                raise Falla("En %s ninguna hoja trae la columna \"%s\" en sus primeros %d "
                            "renglones, así que no sé cuáles son los grupos digitales. Columnas "
                            "que sí hay: %s." % (ruta.name, args.col_codigo, FILAS_ENCABEZADO,
                                                 columnas_que_hay(hojas)))
            if not any(c["tiene_rol"] for c in conteos):
                raise Falla("En %s ninguna hoja trae la columna \"%s\" junto a \"%s\", así que no "
                            "sé quién es estudiante. Columnas que sí hay: %s."
                            % (ruta.name, args.col_rol, args.col_codigo, columnas_que_hay(hojas)))
            destino = out / (nombre_de_archivo(escuela) + ".xlsx")
            if destino.exists() and not args.reemplazar:
                avisos.append("Ya existe %s; no lo toqué. Si quieres generarlo de nuevo, corre "
                              "con --reemplazar." % destino.name)
                estado = "ya existía, sin tocar"
            else:
                estado = "reemplazado" if destino.exists() else "nuevo"
                escribir_xlsx(destino, limpias)
            resultados.append(dict(archivo=ruta.name, escuela=escuela, destino=destino.name,
                                   conteos=conteos, estado=estado, como=como))
        except Falla as e:
            problemas.append((ruta.name, str(e)))

    if not resultados:
        raise Falla("ningún archivo quedó limpio. " + " ".join(p for _, p in problemas))

    # Totales sobre las hojas con rol (las de actividad) y sobre las de solo codigo
    tot = dict(leidos=0, fuera_rol=0, fuera_codigo=0, quedan=0)
    por_rol = {}
    mat = dict(leidos=0, fuera_codigo=0, quedan=0)
    cuadra = True
    for r in resultados:
        for c in r["conteos"]:
            if c["regla"] == "se copia igual":
                continue
            if c["leidos"] != c["fuera_rol"] + c["fuera_codigo"] + c["quedan"]:
                cuadra = False
            if c["tiene_rol"]:
                for k in tot:
                    tot[k] += c[k]
                for k, v in c["por_rol"].items():
                    por_rol[k] = por_rol.get(k, 0) + v
            else:
                for k in mat:
                    mat[k] += c[k]

    n_ok = len(resultados)
    l1 = ("Limpieza de %d %s de actividad de Canvas: %d %s en %s/%s." % (
        len(rutas), "archivo" if len(rutas) == 1 else "archivos", n_ok,
        "escuela" if n_ok == 1 else "escuelas", out.name,
        "" if not problemas else "; %d con problema" % len(problemas)))
    l2 = ("Actividades: %s renglones leídos; quedan %s de grupos digitales (%s); fuera %s por "
          "rol %s y %s por código." % (
              miles(tot["leidos"]), miles(tot["quedan"]),
              ", ".join("%s %s" % (k, miles(v)) for k, v in sorted(por_rol.items())) or "ninguno",
              miles(tot["fuera_rol"]),
              "distinto de %s" % ROL_PROFESOR if args.solo_profesores else ROL_QUE_SE_QUITA,
              miles(tot["fuera_codigo"])))
    l3 = ("Control: la suma %s en %s; %s. Detalle: control-limpieza.xlsx." % (
        "cuadra" if cuadra else "NO cuadra",
        "el archivo" if n_ok == 1 else "los %d archivos" % n_ok,
        "sin avisos" if not (avisos or problemas) else
        "%d %s" % (len(avisos) + len(problemas),
                   "aviso" if len(avisos) + len(problemas) == 1 else "avisos")))
    lineas = [l1, l2, l3]

    # Excel de control
    neg = None
    ctrl = [["Archivo", "Escuela", "Excel limpio", "Estado", "Hoja", "Regla", "Renglones leídos",
             "Fuera por rol", "Fuera por código", "Quedan", "Quedan por rol"]]
    for r in resultados:
        for c in r["conteos"]:
            ctrl.append([r["archivo"], r["escuela"], r["destino"], r["estado"], c["hoja"],
                         c["regla"], c["leidos"], c["fuera_rol"] if c["tiene_rol"] else "",
                         c["fuera_codigo"] if c["regla"] != "se copia igual" else "",
                         c["quedan"],
                         ", ".join("%s %d" % (k, v) for k, v in sorted(c["por_rol"].items()))])
    res = [[l1], [l2], [l3], [],
           ["Leído con", ", ".join(sorted(como_se_leyo))],
           ["Regla de rol", "solo %s" % ROL_PROFESOR if args.solo_profesores
            else "fuera %s" % ROL_QUE_SE_QUITA],
           ["Regla de código", "%s empieza con %s" % (args.col_codigo, PREFIJO)]]
    if avisos or problemas:
        res += [[], ["Avisos"]] + [[a] for a in avisos] + [["%s: %s" % p] for p in problemas]
    escribir_xlsx(out / "control-limpieza.xlsx", [("Resumen", res, 0), ("Control", ctrl, 0)])
    (out / "resumen-limpieza.txt").write_text("\n".join(lineas) + "\n", encoding="utf-8")

    for a in avisos:
        print("Aviso: %s" % a)
    for nombre, p in problemas:
        print("Problema con %s: %s" % (nombre, p))
    if avisos or problemas:
        print()
    print("Resumen:")
    for l in lineas:
        print(l)
    print()
    print("Control por archivo (hoja de actividades):")
    for r in resultados:
        for c in r["conteos"]:
            if c["tiene_rol"]:
                print("  %s -> %s: %s leídos = %s fuera por rol + %s fuera por código + %s "
                      "quedan (%s)" % (r["archivo"], r["destino"], miles(c["leidos"]),
                                       miles(c["fuera_rol"]), miles(c["fuera_codigo"]),
                                       miles(c["quedan"]),
                                       ", ".join("%s %d" % (k, v)
                                                 for k, v in sorted(c["por_rol"].items()))))
    print("  Materias (solo código): %s leídas, %s fuera por código, %s quedan"
          % (miles(mat["leidos"]), miles(mat["fuera_codigo"]), miles(mat["quedan"])))
    print("  Leído con: %s" % ", ".join(sorted(como_se_leyo)))
    print()
    print("Archivos en %s:" % out)
    for r in resultados:
        print("  %s (%s)" % (r["destino"], r["estado"]))
    print("  control-limpieza.xlsx")
    print("  resumen-limpieza.txt")
    return 2 if problemas else 0


def main():
    ap = argparse.ArgumentParser(
        description="Limpia los archivos de actividad de Canvas por escuela y deja solo los "
                    "grupos digitales (skill limpieza-canvas-por-escuela).")
    ap.add_argument("entradas", nargs="+",
                    help="la carpeta con las descargas, o los .xlsx uno por uno")
    ap.add_argument("--salida", default=None,
                    help="carpeta de salida (por omisión, limpios/ junto a la carpeta de "
                         "descargas)")
    ap.add_argument("--reemplazar", action="store_true",
                    help="reemplazar el Excel limpio de una escuela si ya existe")
    ap.add_argument("--solo-profesores", action="store_true",
                    help="dejar solo el rol Profesor (por omisión solo se quita Estudiante)")
    ap.add_argument("--col-rol", default=COL_ROL,
                    help="nombre de la columna del rol (por omisión: %s)" % COL_ROL)
    ap.add_argument("--col-codigo", default=COL_CODIGO,
                    help="nombre de la columna del código del curso (por omisión: %s)"
                         % COL_CODIGO)
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude limpiar: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
