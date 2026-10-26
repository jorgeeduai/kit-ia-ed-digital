#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validar_programacion.py: cruza el archivo de programacion contra el reporte de grupos de IRIS
y dice, CRN por CRN, que no coincide, agrupado por tipo de discrepancia.
Es parte del skill validar-programacion-iris (curso Ed Digital, borrador del grupo, Sesion 7).
Solo biblioteca estandar: corre con cualquier python3, sin instalar nada (ni openpyxl).

Uso:
    python3 validar_programacion.py PROGRAMACION IRIS
    python3 validar_programacion.py PROGRAMACION IRIS --crn CRN --nomina "Nómina del profesor" --salida CARPETA

PROGRAMACION e IRIS son .xlsx o .csv. Cada renglon se identifica con el CRN mas la nomina del
profesor (un grupo con dos profesores trae dos renglones con el mismo CRN). Las columnas se
reconocen por su nombre, sin importar el orden, las mayusculas, los acentos ni los espacios.

Tipos de discrepancia:
    CRN solo en programación       el CRN no aparece en IRIS
    CRN solo en IRIS               el CRN no aparece en programacion
    Profesor asignado              el CRN esta en los dos, con otra nomina de profesor
    Materia y grupo                Clave, Nombre de la materia, Grupo, PMT
    Días y horarios                columnas con Dia u Horario en el nombre
    Fechas del periodo             columnas con Fecha en el nombre
    Nombre o % del profesor        Nombre del profesor y % de responsabilidad
    Atributos del curso            columnas con Atributo en el nombre
    Otras columnas                 cualquier otra columna que este en los dos archivos

Las diferencias solo de formato no cuentan: fechas como 08/02/2027, 2027-02-08 o fecha de
Excel; horas como 7:30, 07:30 o 0730; porcentajes como 50, 50% o 0.5.

Deja en la carpeta salida/ (junto a la carpeta ficticio/ o entrada/ donde esta el archivo de
programacion; si no, dentro de esa carpeta; o en --salida):
    discrepancias-NOMBRE.xlsx           hojas Resumen, Todas y una por tipo, el CRN primero
    resumen-discrepancias-NOMBRE.txt    las tres lineas del resumen
donde NOMBRE es el nombre del archivo de programacion sin extension. No modifica los archivos.
"""
import argparse
import csv
import datetime
import io
import math
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

FILAS_ENCABEZADO = 10      # el encabezado se busca en los primeros 10 renglones
TOP_TIPOS = 3              # tipos con mas discrepancias que van en la tercera linea
SOLO_PROG, SOLO_IRIS, PROF_ASIGNADO = ("CRN solo en programación", "CRN solo en IRIS",
                                       "Profesor asignado")
MATERIA, HORARIOS, FECHAS, PROF_DATOS, ATRIBUTOS, OTRAS = (
    "Materia y grupo", "Días y horarios", "Fechas del periodo", "Nombre o % del profesor",
    "Atributos del curso", "Otras columnas")
TIPOS = [SOLO_PROG, SOLO_IRIS, PROF_ASIGNADO, MATERIA, HORARIOS, FECHAS, PROF_DATOS,
         ATRIBUTOS, OTRAS]
ETIQUETAS = ["nombre de la materia", "materia", "nombre materia", "curso", "nombre del curso"]


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
# Las funciones de lectura de comparar_versiones.py (Sesion 7).
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


# ------------------------------------------------------------------ formatos que no cuentan
def es_profesor(nf):
    return "profesor" in nf or "nomina" in nf or "responsabilidad" in nf


def tipo_de(nf):
    """El tipo de discrepancia de una columna, por su nombre ya sin acentos ni mayusculas."""
    if es_profesor(nf):
        return PROF_DATOS
    if "fecha" in nf:
        return FECHAS
    if nf.startswith("dia") or "horario" in nf or "hora" in nf.split():
        return HORARIOS
    if "atributo" in nf:
        return ATRIBUTOS
    if any(p in nf for p in ("crn", "clave", "materia", "grupo", "pmt")):
        return MATERIA
    return OTRAS


def _fecha(s):
    """08/02/2027, 8/2/27, 2027-02-08 o una fecha de Excel (46426) quedan como 2027-02-08."""
    t = s.replace(" 00:00:00", "")
    try:
        f = float(t)
        if 20000 <= f <= 80000 and f == int(f):
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(f))).isoformat()
    except ValueError:
        pass
    for sep in ("/", "-"):
        p = t.split(sep)
        if len(p) == 3 and all(x.isdigit() for x in p):
            try:
                if len(p[0]) == 4:
                    a, m, d = int(p[0]), int(p[1]), int(p[2])
                else:
                    d, m, a = int(p[0]), int(p[1]), int(p[2])
                    if a < 100:
                        a += 2000
                return datetime.date(a, m, d).isoformat()
            except ValueError:
                return s
    return s


def _hora(s):
    """7:30, 07:30, 07:30:00, 0730 o una hora de Excel (0.3125) quedan como 07:30."""
    try:
        f = float(s)
        if 0 <= f < 1 and "." in s:
            minutos = int(round(f * 1440))
            return "%02d:%02d" % (minutos // 60, minutos % 60)
    except ValueError:
        pass
    if ":" in s:
        p = s.split(":")
        if p[0].isdigit() and p[1][:2].isdigit():
            return "%02d:%02d" % (int(p[0]), int(p[1][:2]))
    if s.isdigit() and len(s) in (3, 4):
        return "%02d:%02d" % (int(s[:-2]), int(s[-2:]))
    return s


def _porcentaje(s):
    """50, 50%, 50.0 o 0.5 quedan como 50."""
    t = s.replace("%", "").strip()
    try:
        f = float(t)
    except ValueError:
        return s
    if "%" not in s and 0 < f <= 1:
        f = f * 100
    f = round(f, 4)
    return str(int(f)) if f == int(f) else repr(f)


def normal(nf, v):
    """La forma con la que se comparan dos celdas de la columna nf."""
    s = limpio(v)
    if not s:
        return s
    if "fecha" in nf or nf.startswith("dia"):
        return _fecha(s)
    if "horario" in nf or "hora" in nf.split():
        return _hora(s)
    if "%" in nf or "responsabilidad" in nf or "porcentaje" in nf:
        return _porcentaje(s)
    return canonico(s)


def mostrar(nf, v):
    """Como se ve un valor en el reporte: fechas 08/02/2027, horas 07:30, % con su signo y las
    celdas vacias como (vacío), sin importar el formato con que llegaron."""
    s = limpio(v)
    if not s:
        return "(vacío)"
    n = normal(nf, s)
    if "fecha" in nf or nf.startswith("dia"):
        try:
            return datetime.date.fromisoformat(n).strftime("%d/%m/%Y")
        except ValueError:
            return s
    if "horario" in nf or "hora" in nf.split():
        return n
    if "%" in nf or "responsabilidad" in nf or "porcentaje" in nf:
        return n + "%" if n.replace(".", "").isdigit() else s
    return s


# ------------------------------------------------------------------ un archivo
def leer_archivo(ruta, crn, nomina, lado):
    """Lee un archivo: encuentra el encabezado (la fila que trae CRN y nomina), arma un renglon
    por CRN y nomina, y avisa si falta una columna o si un renglon se repite."""
    filas = leer_filas(ruta)
    if not any(limpio(c) for f in filas for c in f):
        raise Falla("El archivo de %s (%s) llegó vacío: no trae ningún renglón." % (lado, ruta.name))
    buscadas = [fold(crn), fold(nomina)]
    enc = None
    for i, f in enumerate(filas[:FILAS_ENCABEZADO]):
        nombres = [fold(x) for x in f]
        if all(k in nombres for k in buscadas):
            enc = i
            break
    if enc is None:
        candidatas = [f for f in filas[:FILAS_ENCABEZADO] if any(limpio(x) for x in f)]
        mejor = max(candidatas, key=lambda f: sum(1 for x in f if limpio(x))) if candidatas else []
        nombres = [fold(x) for x in mejor]
        faltan = [k for k in (crn, nomina) if fold(k) not in nombres]
        raise Falla("%s %s en el archivo de %s (%s): no %s encontré en los primeros %d renglones. "
                    "Columnas que sí hay: %s. Si en este archivo se llama distinto, pásala con "
                    "--crn o --nomina."
                    % ("Falta la columna" if len(faltan) == 1 else "Faltan las columnas",
                       y_lista('"%s"' % k for k in faltan), lado, ruta.name,
                       "la" if len(faltan) == 1 else "las", FILAS_ENCABEZADO,
                       " | ".join(limpio(x) for x in mejor if limpio(x)) or "ninguna"))
    encabezado = [limpio(x) for x in filas[enc]]
    columnas, avisos, vistas = [], [], {}
    for j, nombre in enumerate(encabezado):
        if not nombre:
            continue
        if fold(nombre) in vistas:
            avisos.append("La columna %s aparece dos veces en el archivo de %s; uso solo la "
                          "primera." % (nombre, lado))
            continue
        vistas[fold(nombre)] = j
        columnas.append((nombre, j))
    jc, jn = vistas[fold(crn)], vistas[fold(nomina)]
    registros, por_crn, repetidas, sin_crn = {}, {}, {}, 0
    for r, f in enumerate(filas[enc + 1:], start=enc + 2):
        if not any(limpio(c) for c in f):
            continue
        f = list(f) + [""] * (len(encabezado) - len(f))
        c = canonico(f[jc])
        if c == "":
            sin_crn += 1
            continue
        clave = (c, canonico(f[jn]))
        if clave in registros:
            repetidas.setdefault(clave, [registros[clave]["renglon"]]).append(r)
            continue
        registros[clave] = dict(renglon=r, fila=f)
        por_crn.setdefault(c, []).append(clave)
    if repetidas:
        ejemplos = ["CRN %s con nómina %s (renglones %s)" % (k[0], k[1] or "vacía",
                                                             y_lista(str(x) for x in rs))
                    for k, rs in sorted(repetidas.items(), key=lambda x: orden_natural(x[0]))[:5]]
        raise Falla("En el archivo de %s (%s) hay %d %s el mismo CRN y la misma nómina: %s. "
                    "Cada renglón tiene que ser un profesor de un grupo. Si son duplicados, "
                    "quítalos; si no, dime qué otra columna los distingue."
                    % (lado, ruta.name, len(repetidas),
                       "renglón repetido con" if len(repetidas) == 1 else "renglones repetidos con",
                       "; ".join(ejemplos)))
    if not registros:
        raise Falla("El archivo de %s (%s) trae el encabezado (renglón %d) pero ningún renglón "
                    "de datos." % (lado, ruta.name, enc + 1))
    return dict(ruta=ruta, lado=lado, encabezado_renglon=enc + 1, columnas=columnas,
                jc=jc, jn=jn, registros=registros, por_crn=por_crn, sin_crn=sin_crn,
                avisos=avisos)


# ------------------------------------------------------------------ el cruce
def cruzar(prog, iris, crn, nomina):
    col_p = {fold(n): (n, j) for n, j in prog["columnas"]}
    col_i = {fold(n): (n, j) for n, j in iris["columnas"]}
    llaves = (fold(crn), fold(nomina))
    avisos = prog["avisos"] + iris["avisos"]
    solo_p = [n for n, _ in prog["columnas"] if fold(n) not in col_i]
    solo_i = [n for n, _ in iris["columnas"] if fold(n) not in col_p]
    for solo, en, no_en in ((solo_p, "programación", "IRIS"), (solo_i, "IRIS", "programación")):
        if solo:
            uno = len(solo) == 1
            avisos.append("%s %s %s en %s y no en %s: %s." % (
                "La columna" if uno else "Las columnas", y_lista(solo),
                "está" if uno else "están", en, no_en,
                "no se compara" if uno else "no se comparan"))
    comunes = [(n, fold(n), j, col_i[fold(n)][1]) for n, j in prog["columnas"]
               if fold(n) in col_i and fold(n) not in llaves]
    del_grupo = [c for c in comunes if not es_profesor(c[1])]
    del_profesor = [c for c in comunes if es_profesor(c[1])]
    nf_prof = next((nf for nf in ("nombre del profesor", "profesor", "nombre")
                    if nf in col_p and nf in col_i), None)
    nf_etq = next((e for e in ETIQUETAS if e in col_p), None)
    nombre_nomina = col_p[fold(nomina)][0]

    def materia(fila):
        return limpio(fila[col_p[nf_etq][1]]) if nf_etq else ""

    def quien(fila, lado_cols, jn):
        nom = limpio(fila[jn]) or "sin nómina"
        if nf_prof:
            nombre = limpio(fila[lado_cols[nf_prof][1]])
            return "%s (%s)" % (nom, nombre) if nombre else nom
        return nom

    pc, ic = prog["por_crn"], iris["por_crn"]
    rp, ri = prog["registros"], iris["registros"]
    filas, solo_formato = [], 0

    def agrega(c, tipo, columna, nom, v_p, v_i, mat):
        filas.append(dict(crn=c, tipo=tipo, columna=columna, nomina=nom, prog=v_p, iris=v_i,
                          materia=mat))

    for c in sorted(set(pc) - set(ic), key=lambda x: orden_natural((x,))):
        f = rp[pc[c][0]]["fila"]
        agrega(c, SOLO_PROG, "(todo el grupo)", "", "está", "no está", materia(f))
    for c in sorted(set(ic) - set(pc), key=lambda x: orden_natural((x,))):
        f = ri[ic[c][0]]["fila"]
        mat = limpio(f[col_i[nf_etq][1]]) if nf_etq and nf_etq in col_i else ""
        agrega(c, SOLO_IRIS, "(todo el grupo)", "", "no está", "está", mat)

    for c in sorted(set(pc) & set(ic), key=lambda x: orden_natural((x,))):
        fp0, fi0 = rp[pc[c][0]]["fila"], ri[ic[c][0]]["fila"]
        mat = materia(fp0)
        # Datos del grupo: se comparan una vez por CRN, con el primer renglon de cada archivo
        for nombre, nf, jp, ji in del_grupo:
            a, b = limpio(fp0[jp]), limpio(fi0[ji])
            if normal(nf, a) != normal(nf, b):
                agrega(c, tipo_de(nf), nombre, "", mostrar(nf, a), mostrar(nf, b), mat)
            elif canonico(a) != canonico(b):
                solo_formato += 1
        # Un grupo con dos profesores: sus renglones deben traer los mismos datos del grupo
        for lado, por, reg, pos in (("programación", pc, rp, 2), ("IRIS", ic, ri, 3)):
            if len(por[c]) < 2:
                continue
            f0 = reg[por[c][0]]["fila"]
            otras = [reg[k]["fila"] for k in por[c][1:]]
            if any(normal(col[1], f0[col[pos]]) != normal(col[1], fk[col[pos]])
                   for fk in otras for col in del_grupo):
                avisos.append("En el archivo de %s, los renglones del CRN %s no traen los mismos "
                              "datos del grupo; comparo el primero (renglón %d)."
                              % (lado, c, reg[por[c][0]]["renglon"]))
        # Profesores: la nomina tiene que ser la misma en los dos archivos
        noms_p = [k[1] for k in pc[c]]
        noms_i = [k[1] for k in ic[c]]
        if set(noms_p) != set(noms_i):
            v_p = "; ".join(quien(rp[(c, n)]["fila"], col_p, prog["jn"])
                            for n in noms_p if n not in noms_i)
            v_i = "; ".join(quien(ri[(c, n)]["fila"], col_i, iris["jn"])
                            for n in noms_i if n not in noms_p)
            agrega(c, PROF_ASIGNADO, nombre_nomina, "", v_p or "(ninguno más)",
                   v_i or "(ninguno más)", mat)
        for n in noms_p:
            if n not in noms_i:
                continue
            fp, fi = rp[(c, n)]["fila"], ri[(c, n)]["fila"]
            for nombre, nf, jp, ji in del_profesor:
                a, b = limpio(fp[jp]), limpio(fi[ji])
                if normal(nf, a) != normal(nf, b):
                    agrega(c, PROF_DATOS, nombre, n, mostrar(nf, a), mostrar(nf, b), mat)
                elif canonico(a) != canonico(b):
                    solo_formato += 1

    por_tipo = {}
    for f in filas:
        por_tipo[f["tipo"]] = por_tipo.get(f["tipo"], 0) + 1
    top = sorted(por_tipo, key=lambda t: (-por_tipo[t], TIPOS.index(t)))
    crn_inc = {f["crn"] for f in filas}
    con_datos = {f["crn"] for f in filas if f["tipo"] not in (SOLO_PROG, SOLO_IRIS, PROF_ASIGNADO)}
    celdas = sum(1 for f in filas if f["tipo"] not in (SOLO_PROG, SOLO_IRIS, PROF_ASIGNADO))
    return dict(filas=filas, por_tipo=por_tipo, top=top, crn_inc=crn_inc, con_datos=con_datos,
                celdas=celdas, avisos=avisos, solo_formato=solo_formato,
                en_ambos=len(set(pc) & set(ic)), comparadas=[n for n, _, _, _ in comunes])


def lineas_resumen(prog, iris, r, nombre_detalle, nombre_nomina):
    def n_tipo(t):
        return r["por_tipo"].get(t, 0)
    celdas = r["celdas"]
    l1 = "Validación de %s contra %s, por CRN y %s." % (prog["ruta"].name, iris["ruta"].name,
                                                       nombre_nomina)
    if not r["crn_inc"]:
        l2 = "No hay discrepancias: los %s CRN coinciden en todas las columnas que se comparan." \
             % miles(r["en_ambos"])
    else:
        l2 = "Hay %s CRN con incidencia: %s solo en programación, %s solo en IRIS, %s con otro " \
             "profesor asignado y %s con datos distintos (%s %s)." % (
                 miles(len(r["crn_inc"])), miles(n_tipo(SOLO_PROG)), miles(n_tipo(SOLO_IRIS)),
                 miles(n_tipo(PROF_ASIGNADO)), miles(len(r["con_datos"])), miles(celdas),
                 "celda" if celdas == 1 else "celdas")
    mas = y_lista("%s (%s)" % (t, miles(r["por_tipo"][t])) for t in r["top"][:TOP_TIPOS]) \
        or "ninguno"
    l3 = "Tipos con más discrepancias: %s. Detalle: %s." % (mas, nombre_detalle)
    return [l1, l2, l3]


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
def carpeta_salida(prog, salida):
    if salida:
        return Path(salida).expanduser()
    if fold(prog.parent.name) in ("ficticio", "entrada", "versiones"):
        return prog.parent.parent / "salida"
    return prog.parent / "salida"


def principal(args):
    ruta_p = Path(args.programacion).expanduser()
    ruta_i = Path(args.iris).expanduser()
    prog = leer_archivo(ruta_p, args.crn, args.nomina, "programación")
    iris = leer_archivo(ruta_i, args.crn, args.nomina, "IRIS")
    r = cruzar(prog, iris, args.crn, args.nomina)

    out = carpeta_salida(ruta_p, args.salida)
    out.mkdir(parents=True, exist_ok=True)
    base = ruta_p.stem
    nombre_detalle = "discrepancias-%s.xlsx" % base
    ruta_detalle = out / nombre_detalle
    ruta_resumen = out / ("resumen-discrepancias-%s.txt" % base)
    nombre_crn = dict((j, n) for n, j in prog["columnas"])[prog["jc"]]
    nombre_nomina = dict((j, n) for n, j in prog["columnas"])[prog["jn"]]
    lineas = lineas_resumen(prog, iris, r, nombre_detalle, nombre_nomina)

    control = [
        ("Archivo de programación", ruta_p.name),
        ("Reporte de IRIS", ruta_i.name),
        ("Clave", "%s y %s" % (nombre_crn, nombre_nomina)),
        ("Renglón del encabezado (programación, IRIS)", "%d, %d" % (prog["encabezado_renglon"],
                                                                    iris["encabezado_renglon"])),
        ("Renglones leídos en programación", len(prog["registros"])),
        ("Renglones leídos en IRIS", len(iris["registros"])),
        ("CRN en programación", len(prog["por_crn"])),
        ("CRN en IRIS", len(iris["por_crn"])),
        ("CRN en los dos", r["en_ambos"]),
        ("CRN con incidencia", len(r["crn_inc"])),
        ("Celdas con datos distintos", r["celdas"]),
        ("Celdas que solo cambian de formato (no cuentan)", r["solo_formato"]),
    ]
    if prog["sin_crn"] or iris["sin_crn"]:
        control.append(("Renglones sin CRN, que no se comparan (programación, IRIS)",
                        "%d, %d" % (prog["sin_crn"], iris["sin_crn"])))

    # Hoja Resumen
    res = [[(lineas[0], NEGRITA)], [lineas[1]], [lineas[2]], [],
           [("Concepto", NEGRITA), ("Valor", NEGRITA)]]
    res += [[k, v] for k, v in control]
    res += [[], [("Tipo de discrepancia", NEGRITA), ("Renglones", NEGRITA)]]
    res += [[t, r["por_tipo"][t]] for t in TIPOS if t in r["por_tipo"]] or [["ninguno", 0]]
    res += [[], [("Columnas que se comparan", NEGRITA)]] + [[n] for n in r["comparadas"]]
    if r["avisos"]:
        res += [[], [("Avisos", NEGRITA)]] + [[a] for a in r["avisos"]]

    # Hoja Todas y una hoja por tipo: siempre empiezan con el CRN
    cab = [nombre_crn, "Tipo", "Columna", nombre_nomina, "En programación", "En IRIS",
           "Nombre de la materia"]
    anchos = [10, 24, 26, 16, 34, 34, 36]

    def hoja(filas):
        h = [[(c, NEGRITA) for c in cab]]
        for f in filas:
            h.append([a_numero(f["crn"]), f["tipo"], f["columna"], f["nomina"],
                      f["prog"], f["iris"], f["materia"]])
        return h

    ordenadas = sorted(r["filas"], key=lambda f: (orden_natural((f["crn"],)),
                                                  TIPOS.index(f["tipo"])))
    hojas = [("Resumen", res, [70, 40], False), ("Todas", hoja(ordenadas), anchos, True)]
    for t in TIPOS:
        if t in r["por_tipo"]:
            hojas.append((t, hoja([f for f in ordenadas if f["tipo"] == t]), anchos, True))
    escribir_xlsx(ruta_detalle, hojas)
    ruta_resumen.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    for a in r["avisos"]:
        print("Aviso: %s" % a)
    if r["avisos"]:
        print()
    print("Resumen:")
    for l in lineas:
        print(l)
    print()
    print("Control del cruce:")
    for k, v in control:
        print("  %s: %s" % (k, miles(v) if isinstance(v, int) else v))
    if r["por_tipo"]:
        print("  Por tipo: %s" % ", ".join("%s %s" % (t, miles(r["por_tipo"][t]))
                                          for t in TIPOS if t in r["por_tipo"]))
    print()
    print("Archivos:")
    print("  %s" % ruta_detalle)
    print("  %s" % ruta_resumen)
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Cruza la programación contra el reporte de IRIS por CRN y nómina "
                    "(skill validar-programacion-iris).")
    ap.add_argument("programacion", help="archivo de programación (.xlsx o .csv)")
    ap.add_argument("iris", help="reporte de grupos de IRIS (.xlsx o .csv)")
    ap.add_argument("--crn", default="CRN", help="nombre de la columna del CRN (CRN)")
    ap.add_argument("--nomina", default="Nómina del profesor",
                    help="nombre de la columna de la nómina (Nómina del profesor)")
    ap.add_argument("--salida", default=None,
                    help="carpeta de salida (por omisión, salida/ junto a la carpeta del archivo "
                         "de programación)")
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude validar: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
