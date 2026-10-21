#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
limpiar_reporte.py: limpia el reporte de inscritos y arma la tabla y el resumen para Teams.
Es parte del skill limpiar-reporte (curso Ed Digital, Sesion 6). Solo biblioteca estandar:
corre con cualquier python3, sin instalar nada.

Uso:
    python3 limpiar_reporte.py                  busca el reporte mas reciente en ./entrada o aqui
    python3 limpiar_reporte.py RUTA             RUTA es un .xlsx, un .csv o una carpeta
    python3 limpiar_reporte.py RUTA --salida CARPETA --corte-anterior AAAA-MM-DD

Deja en la carpeta salida/ (junto a entrada/, o junto al reporte):
    tabla-inscritos-AAAA-MM-DD.xlsx   hojas Tabla, Peticiones y Control
    resumen-teams-AAAA-MM-DD.txt      las tres lineas para el grupo de Teams
e imprime el resumen y el control de la limpieza. No modifica el reporte original.
"""
import argparse
import csv
import io
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta
from pathlib import Path
from xml.sax.saxutils import escape

CAMPUS = ["Monterrey", "Guadalajara", "Ciudad de México", "Querétaro", "Puebla",
          "Estado de México"]
ESTATUS = ["Nueva", "En proceso", "Resuelta", "Cancelada"]
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]
DIAS = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
MES_POR_ABR = {m[:3]: i + 1 for i, m in enumerate(MESES)}
MES_POR_ABR.update({"jan": 1, "apr": 4, "aug": 8, "dec": 12})


# ------------------------------------------------------------------ utilidades
def fold(s):
    """Para comparar: sin acentos, sin mayusculas y sin espacios de mas."""
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s))
                if not unicodedata.combining(c))
    return " ".join(s.casefold().split())


def limpio(s):
    return " ".join(str(s).split())


def miles(n):
    return "{:,}".format(n)


def fecha_larga(d, con_anio=True):
    s = "%s %d de %s" % (DIAS[d.weekday()], d.day, MESES[d.month - 1])
    return s + (" de %d" % d.year if con_anio else "")


def corte_anterior(d):
    """El corte anterior del flujo: el lunes o miercoles mas reciente antes de d."""
    a = d - timedelta(days=1)
    while a.weekday() not in (0, 2):
        a -= timedelta(days=1)
    return a


def lineas_teams(corte, hora, periodo, n_pet, n_alum, por_estatus, ant, nuevas,
                 campus_top, n_top, nombre_tabla):
    a_las = " a las %s" % hora if hora else ""
    l1 = "Reporte de inscritos, corte del %s%s (%s): %s peticiones de %s alumnos." % (
        fecha_larga(corte), a_las, periodo, miles(n_pet), miles(n_alum))
    l2 = "Por estatus: %s nuevas, %s en proceso, %s resueltas y %s canceladas." % (
        miles(por_estatus.get("Nueva", 0)), miles(por_estatus.get("En proceso", 0)),
        miles(por_estatus.get("Resuelta", 0)), miles(por_estatus.get("Cancelada", 0)))
    l3 = ("Desde el corte del %s llegaron %s peticiones; el campus con más peticiones "
          "es %s (%s). Tabla completa: %s.") % (
        fecha_larga(ant, con_anio=False), miles(nuevas), campus_top, miles(n_top),
        nombre_tabla)
    return [l1, l2, l3]


def leer_fecha(s):
    """Fechas en texto, dia antes que mes: 14/10/2026, 2026-10-14, 14-oct-2026,
    14/10/2026 09:32. Tambien un numero de serie de Excel. None si no se reconoce."""
    s = str(s).strip()
    if not s:
        return None
    try:
        x = float(s)
        if 20000 < x < 80000:
            return date(1899, 12, 30) + timedelta(days=int(x))
        return None
    except ValueError:
        pass
    parte = s.split()[0]
    if "T" in parte and parte[:4].isdigit():
        parte = parte.split("T")[0]
    trozos = None
    for sep in ("/", "-", "."):
        if len(parte.split(sep)) == 3:
            trozos = parte.split(sep)
            break
    if not trozos:
        return None
    a, b, c = trozos
    try:
        if len(a) == 4 and a.isdigit():
            return date(int(a), int(b), int(c))
        mes = int(b) if b.isdigit() else MES_POR_ABR.get(fold(b)[:3])
        anio = int(c)
        if anio < 100:
            anio += 2000
        return date(anio, mes, int(a))
    except (ValueError, TypeError):
        return None


def fecha_en_nombre(nombre):
    base = nombre.rsplit(".", 1)[0]
    for i in range(len(base) - 9):
        trozo = base[i:i + 10]
        if trozo[4:5] == "-" and trozo[7:8] == "-":
            try:
                return date.fromisoformat(trozo)
            except ValueError:
                continue
    return None


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


def buscar_reporte(arg):
    if arg:
        p = Path(arg).expanduser()
        if p.is_file():
            return p
        if not p.is_dir():
            return None
        carpetas = [p / "entrada", p]
    else:
        carpetas = [Path("entrada"), Path(".")]
    for carpeta in carpetas:
        if not carpeta.is_dir():
            continue
        cands = [x for x in carpeta.iterdir()
                 if x.is_file() and x.suffix.lower() in (".xlsx", ".csv")
                 and "inscritos" in fold(x.name)
                 and not x.name.startswith(("~$", "."))
                 and not fold(x.name).startswith(("tabla", "resumen"))]
        if cands:
            cands.sort(key=lambda x: (fecha_en_nombre(x.name) or date.min,
                                      x.suffix.lower() == ".xlsx", x.stat().st_mtime))
            return cands[-1]
    return None


# ------------------------------------------------------------------ limpieza
def buscar_col(nombres, *palabras):
    for j, n in enumerate(nombres):
        if all(p in n for p in palabras):
            return j
    return None


def limpiar(filas, ruta, corte_ant_forzado=None):
    enc = None
    for i, f in enumerate(filas[:40]):
        nombres = [fold(x) for x in f]
        if any("matricula" in n for n in nombres) and "crn" in nombres:
            enc = i
            break
    if enc is None:
        raise SystemExit("No encontré el renglón de encabezados (con Matrícula y CRN) en "
                         "los primeros 40 renglones de %s." % ruta.name)
    encabezado = [limpio(x) for x in filas[enc]]
    nombres = [fold(x) for x in filas[enc]]
    ncol = len(encabezado)
    col = dict(campus=buscar_col(nombres, "campus"),
               mat=buscar_col(nombres, "matricula"),
               crn=nombres.index("crn"),
               est=buscar_col(nombres, "estatus", "peticion"),
               fecha=buscar_col(nombres, "fecha", "solicitud"),
               periodo=buscar_col(nombres, "periodo"))
    faltan = [k for k in ("campus", "mat", "est", "fecha") if col[k] is None]
    if faltan:
        raise SystemExit("Faltan columnas en el encabezado (%s). Columnas encontradas: %s"
                         % (", ".join(faltan), " | ".join(encabezado)))

    # Fecha y hora de corte, y periodo, de los renglones de titulo
    corte, hora, periodo_titulo = None, "", ""
    for f in filas[:enc]:
        for celda in f:
            t = fold(celda)
            if t.startswith("fecha de corte") and ":" in celda:
                resto = celda.split(":", 1)[1].strip()
                corte = leer_fecha(resto)
                hora = resto.split()[1] if len(resto.split()) > 1 else ""
            elif t.startswith("periodo") and ":" in celda:
                periodo_titulo = celda.split(":", 1)[1].strip()
    if corte is None:
        corte = fecha_en_nombre(ruta.name)

    st = dict(hoja=len(filas), encabezado=enc + 1, datos=0, duplicados=0, cotitular=0,
              campus_var=0, estatus_var=0, fechas_mal=0)
    vistos, por_clave, orden = set(), {}, []
    campus_desconocidos, estatus_desconocidos = set(), set()
    mapa_campus = {fold(c): c for c in CAMPUS}
    mapa_estatus = {fold(e): e for e in ESTATUS}
    for f in filas[enc + 1:]:
        f = list(f) + [""] * (ncol - len(f))
        mat, crn = limpio(f[col["mat"]]), limpio(f[col["crn"]])
        if not mat or not crn:
            continue                      # titulo, seccion, subtotal o renglon vacio
        st["datos"] += 1
        c_raw, e_raw = f[col["campus"]], f[col["est"]]
        c = mapa_campus.get(fold(c_raw))
        if c is None:
            c = limpio(c_raw)
            campus_desconocidos.add(c)
        e = mapa_estatus.get(fold(e_raw))
        if e is None:
            e = limpio(e_raw) or "(vacío)"
            estatus_desconocidos.add(e)
        st["campus_var"] += c_raw != c
        st["estatus_var"] += e_raw != e
        llave_exacta = tuple(f[:ncol])
        if llave_exacta in vistos:
            st["duplicados"] += 1
            continue
        vistos.add(llave_exacta)
        clave = (mat.upper(), crn)
        if clave in por_clave:
            st["cotitular"] += 1
            continue
        fecha = leer_fecha(f[col["fecha"]])
        if fecha is None:
            st["fechas_mal"] += 1
        por_clave[clave] = dict(fila=f[:ncol], campus=c, estatus=e, fecha=fecha)
        orden.append(clave)

    peticiones = [por_clave[k] for k in orden]
    if corte is None:
        fechas = [p["fecha"] for p in peticiones if p["fecha"]]
        corte = (max(fechas) + timedelta(days=1)) if fechas else date.today()
    ant = corte_ant_forzado or corte_anterior(corte)
    periodos = {}
    if col["periodo"] is not None:
        for p in peticiones:
            v = limpio(p["fila"][col["periodo"]])
            periodos[v] = periodos.get(v, 0) + 1
    periodo = max(periodos, key=periodos.get) if periodos else periodo_titulo

    campus_orden = CAMPUS + sorted(campus_desconocidos)
    estatus_orden = ESTATUS + sorted(estatus_desconocidos)
    tabla = {c: {e: 0 for e in estatus_orden} for c in campus_orden}
    for p in peticiones:
        tabla[p["campus"]][p["estatus"]] += 1
    campus_orden = [c for c in campus_orden if sum(tabla[c].values()) > 0]

    st.update(peticiones=len(peticiones),
              alumnos=len({k[0] for k in orden}),
              nuevas=sum(1 for p in peticiones if p["fecha"] and p["fecha"] >= ant),
              corte=corte, hora=hora, anterior=ant, periodo=periodo,
              no_datos=len(filas) - st["datos"],
              campus_desconocidos=sorted(campus_desconocidos),
              estatus_desconocidos=sorted(estatus_desconocidos))
    return dict(st=st, encabezado=encabezado, col=col, peticiones=peticiones,
                tabla=tabla, campus_orden=campus_orden, estatus_orden=estatus_orden)


# ------------------------------------------------------------------ escritura xlsx
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NORMAL, NEGRITA, FECHA, MILES, NEGRITA_MILES = 0, 1, 2, 3, 4


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
            if isinstance(v, date):
                if estilo == NORMAL:
                    estilo = FECHA
                v = (v - date(1899, 12, 30)).days
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
         '<cellXfs count="5">'
         '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
         '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
         '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
         '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
         '<xf numFmtId="3" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyNumberFormat="1"/>'
         '</cellXfs>'
         '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
         '</styleSheet>' % NS_MAIN),
    ]
    for i, (_, filas, anchos, congelar) in enumerate(hojas):
        archivos.append(("xl/worksheets/sheet%d.xml" % (i + 1), _hoja_xml(filas, anchos, congelar)))
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in archivos:
            z.writestr(nombre, contenido.encode("utf-8"))


def a_numero(s):
    s = limpio(s)
    if s.isdigit() and len(s) < 16 and (s == "0" or not s.startswith("0")):
        return int(s)
    if s.count(".") == 1 and s.replace(".", "").isdigit() and not s.startswith("."):
        return float(s)
    return s


# ------------------------------------------------------------------ principal
def main():
    ap = argparse.ArgumentParser(description="Limpia el reporte de inscritos (skill limpiar-reporte).")
    ap.add_argument("ruta", nargs="?", default=None,
                    help="reporte .xlsx o .csv, o carpeta donde buscar el más reciente")
    ap.add_argument("--salida", default=None, help="carpeta de salida (por omisión, salida/)")
    ap.add_argument("--corte-anterior", default=None,
                    help="fecha AAAA-MM-DD del corte anterior (por omisión, el lunes o miércoles previo)")
    args = ap.parse_args()

    rep = buscar_reporte(args.ruta)
    if rep is None:
        print("No encontré el reporte. Pásame la ruta del archivo o de la carpeta "
              "(busco archivos .xlsx o .csv con 'inscritos' en el nombre, primero en entrada/).",
              file=sys.stderr)
        return 2
    filas = leer_xlsx(rep) if rep.suffix.lower() == ".xlsx" else leer_csv(rep)
    ant = date.fromisoformat(args.corte_anterior) if args.corte_anterior else None
    r = limpiar(filas, rep, ant)
    st, tabla = r["st"], r["tabla"]

    if args.salida:
        out = Path(args.salida).expanduser()
    elif rep.parent.name.lower() == "entrada":
        out = rep.parent.parent / "salida"
    else:
        out = rep.parent / "salida"
    out.mkdir(parents=True, exist_ok=True)
    iso = st["corte"].isoformat()
    nombre_tabla = "tabla-inscritos-%s.xlsx" % iso
    ruta_tabla, ruta_resumen = out / nombre_tabla, out / ("resumen-teams-%s.txt" % iso)

    por_estatus = {e: sum(tabla[c][e] for c in r["campus_orden"]) for e in r["estatus_orden"]}
    totales = {c: sum(tabla[c].values()) for c in r["campus_orden"]}
    top = max(r["campus_orden"], key=lambda c: (totales[c], -r["campus_orden"].index(c)))
    lineas = lineas_teams(st["corte"], st["hora"], st["periodo"], st["peticiones"],
                          st["alumnos"], por_estatus, st["anterior"], st["nuevas"], top,
                          totales[top], nombre_tabla)

    # Hoja Tabla
    t = [[(h, NEGRITA) for h in ["Campus"] + r["estatus_orden"] + ["Total"]]]
    for c in r["campus_orden"]:
        t.append([c] + [(tabla[c][e], MILES) for e in r["estatus_orden"]] + [(totales[c], MILES)])
    t.append([("Total", NEGRITA)] + [(por_estatus[e], NEGRITA_MILES) for e in r["estatus_orden"]]
             + [(st["peticiones"], NEGRITA_MILES)])
    t.append([])
    t.append(["Corte", fecha_larga(st["corte"]) + (", %s" % st["hora"] if st["hora"] else "")])
    t.append(["Periodo", st["periodo"]])
    t.append(["Fuente", rep.name])
    t.append(["Peticiones únicas (alumno + CRN)", (st["peticiones"], MILES)])
    t.append(["Alumnos distintos", (st["alumnos"], MILES)])
    t.append(["Peticiones desde el corte del %s" % fecha_larga(st["anterior"], False),
              (st["nuevas"], MILES)])

    # Hoja Peticiones: una por alumno + CRN, sin columnas vacias
    col = r["col"]
    usadas = [j for j in range(len(r["encabezado"]))
              if r["encabezado"][j] and any(limpio(p["fila"][j]) for p in r["peticiones"])]
    quitadas = [r["encabezado"][j] for j in range(len(r["encabezado"]))
                if j not in usadas and r["encabezado"][j]]
    pet = [[(r["encabezado"][j], NEGRITA) for j in usadas]]
    for p in r["peticiones"]:
        fila = []
        for j in usadas:
            if j == col["campus"]:
                fila.append(p["campus"])
            elif j == col["est"]:
                fila.append(p["estatus"])
            elif j == col["fecha"]:
                fila.append(p["fecha"] if p["fecha"] else limpio(p["fila"][j]))
            else:
                fila.append(a_numero(p["fila"][j]))
        pet.append(fila)

    # Hoja Control
    control = [
        ("Reporte leído", rep.name),
        ("Renglones en la hoja (título, encabezado, secciones, subtotales y datos)", st["hoja"]),
        ("Renglón del encabezado", st["encabezado"]),
        ("Renglones que no son datos", st["no_datos"]),
        ("Renglones de datos", st["datos"]),
        ("Duplicados exactos quitados", st["duplicados"]),
        ("Renglones repetidos por co-titular unidos", st["cotitular"]),
        ("Peticiones únicas (alumno + CRN)", st["peticiones"]),
        ("Alumnos distintos", st["alumnos"]),
        ("Campus escritos con variante (corregidos)", st["campus_var"]),
        ("Estatus escritos con variante (corregidos)", st["estatus_var"]),
        ("Fechas de solicitud sin reconocer", st["fechas_mal"]),
        ("Corte", fecha_larga(st["corte"]) + (", %s" % st["hora"] if st["hora"] else "")),
        ("Corte anterior", fecha_larga(st["anterior"])),
        ("Peticiones desde el corte anterior", st["nuevas"]),
        ("Columnas vacías quitadas", ", ".join(quitadas) or "ninguna"),
    ]
    if st["campus_desconocidos"]:
        control.append(("Campus no reconocidos (revisar)", ", ".join(st["campus_desconocidos"])))
    if st["estatus_desconocidos"]:
        control.append(("Estatus no reconocidos (revisar)", ", ".join(st["estatus_desconocidos"])))
    ctl = [[("Concepto", NEGRITA), ("Valor", NEGRITA)]]
    ctl += [[k, (v, MILES) if isinstance(v, int) else v] for k, v in control]

    escribir_xlsx(ruta_tabla, [
        ("Tabla", t, [34, 11, 12, 11, 12, 11], False),
        ("Peticiones", pet, [18] * len(usadas), True),
        ("Control", ctl, [70, 40], False),
    ])
    ruta_resumen.write_text("\n".join(lineas) + "\n", encoding="utf-8")

    print("Resumen para Teams:")
    for l in lineas:
        print(l)
    print()
    print("Control de la limpieza:")
    for k, v in control:
        print("  %s: %s" % (k, miles(v) if isinstance(v, int) else v))
    print()
    print("Archivos:")
    print("  %s" % ruta_tabla)
    print("  %s" % ruta_resumen)
    if st["campus_desconocidos"] or st["estatus_desconocidos"] or st["fechas_mal"]:
        print("\nAviso: hay valores que no reconocí; están en la hoja Control. Revísalos antes "
              "de mandar la tabla.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
