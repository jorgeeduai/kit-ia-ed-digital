#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
copiar_tabla_recursos.py: toma las dos primeras columnas de la tabla "Recursos que requieren
una version en ingles" del Reporte_Traduccion_[CLAVE]_[NombreUF] y las deja listas para pegar
en las dos primeras columnas de la misma tabla del ProfesorRevisor_PL2026_UF. Es parte del
skill tabla-recursos-ingles (curso Ed Digital, propuesta de skill del grupo, ronda 2).
Solo biblioteca estandar: corre con cualquier python3, sin instalar nada.

El script NO escribe en el Word del profesor: deja la tabla en CSV y en Excel, con las ligas
de cada celda y las capturas de pantalla guardadas en una carpeta. El pegado al Word lo hace
TECgpt con una instruccion (ver LEEME.md). Si los documentos son de Google, se descargan antes
como Word (Archivo, Descargar, Microsoft Word).

Uso:
    python3 copiar_tabla_recursos.py REPORTE.docx PROFESOR.docx
    python3 copiar_tabla_recursos.py REPORTE.docx PROFESOR.docx --salida CARPETA
    python3 copiar_tabla_recursos.py REPORTE.docx PROFESOR.docx --titulo "Otra tabla"

Deja en salida/ (junto al reporte; o en --salida):
    tabla-para-profesor-CLAVE_NombreUF.csv    la tabla lista para pegar (abre en Excel)
    tabla-para-profesor-CLAVE_NombreUF.xlsx   la misma tabla, con una hoja de control
    imagenes/renglon-NN-col-C-K.png           las capturas de pantalla de cada celda
    resumen-tabla-recursos.txt                las tres lineas del resumen
No modifica ningun archivo de entrada.
"""
import argparse
import csv
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape

TITULO_TABLA = "Recursos que requieren una versión en inglés"
COLUMNAS_QUE_SE_COPIAN = 2

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
V = "{urn:schemas-microsoft-com:vml}"


class Falla(Exception):
    """Un problema con los documentos que la persona tiene que resolver; se explica y se sale."""


# ------------------------------------------------------------------ utilidades
def fold(s):
    s = "".join(c for c in unicodedata.normalize("NFKD", str(s))
                if not unicodedata.combining(c))
    return " ".join(s.casefold().split())


def limpio(s):
    return " ".join(str(s).split())


def plural(n, uno, varios):
    return "%d %s" % (n, uno if n == 1 else varios)


def ligas_en_texto(texto):
    """Ligas escritas como texto (sin hipervinculo): palabras que empiezan con http."""
    salida = []
    for palabra in texto.split():
        p = palabra.strip("()[]<>.,;\"'")
        if p.startswith("http://") or p.startswith("https://"):
            salida.append(p)
    return salida


# ------------------------------------------------------------------ lectura del .docx
class Word:
    def __init__(self, ruta, que):
        self.ruta, self.que = ruta, que
        if not ruta.exists():
            raise Falla("No encontré %s: %s. Revisa la ruta y el nombre." % (que, ruta))
        if ruta.stat().st_size == 0:
            raise Falla("%s (%s) llegó vacío: pesa 0 bytes." % (que.capitalize(), ruta.name))
        if ruta.suffix.lower() != ".docx":
            raise Falla("%s (%s) no es un documento de Word (.docx). Si es un documento de Google, "
                        "ábrelo y descárgalo como Word: Archivo, Descargar, Microsoft Word (.docx)."
                        % (que.capitalize(), ruta.name))
        try:
            self.zip = zipfile.ZipFile(ruta)
            self.doc = ET.fromstring(self.zip.read("word/document.xml"))
            rels = {}
            if "word/_rels/document.xml.rels" in self.zip.namelist():
                for e in ET.fromstring(self.zip.read("word/_rels/document.xml.rels")):
                    rels[e.get("Id")] = e.get("Target")
            self.rels = rels
        except (zipfile.BadZipFile, KeyError, ET.ParseError):
            raise Falla("No pude leer %s: no parece un .docx válido o está dañado. Ábrelo en Word y "
                        "guárdalo otra vez como .docx." % ruta.name)

    def bloques(self):
        """Parrafos y tablas del cuerpo, en orden (entra a los controles de contenido)."""
        def recorrer(nodo):
            for h in nodo:
                if h.tag in (W + "p", W + "tbl"):
                    yield h
                elif h.tag in (W + "sdt", W + "sdtContent", W + "customXml"):
                    yield from recorrer(h)
        body = self.doc.find(W + "body")
        return list(recorrer(body)) if body is not None else []

    @staticmethod
    def texto_parrafo(p):
        partes = []
        for e in p.iter():
            if e.tag == W + "t":
                partes.append(e.text or "")
            elif e.tag == W + "tab":
                partes.append(" ")
            elif e.tag in (W + "br", W + "cr"):
                partes.append("\n")
        return "".join(partes)

    def celda(self, tc):
        """Texto (un renglon por parrafo), ligas y referencias de imagen de una celda."""
        textos, ligas, imagenes = [], [], []
        for p in tc.iter(W + "p"):
            textos.append(self.texto_parrafo(p).strip())
        for e in tc.iter():
            if e.tag == W + "hyperlink":
                rid = e.get(R + "id")
                if rid and rid in self.rels:
                    ligas.append(self.rels[rid])
            elif e.tag == W + "instrText" and e.text and "HYPERLINK" in e.text:
                partes = e.text.split('"')
                if len(partes) >= 2:
                    ligas.append(partes[1])
            elif e.tag == A + "blip":
                rid = e.get(R + "embed")
                if rid and rid in self.rels:
                    imagenes.append(self.rels[rid])
            elif e.tag == V + "imagedata":
                rid = e.get(R + "id")
                if rid and rid in self.rels:
                    imagenes.append(self.rels[rid])
        texto = "\n".join(t for t in textos if t)
        for l in ligas_en_texto(texto):
            if l not in ligas:
                ligas.append(l)
        return texto, ligas, imagenes

    def filas_de_tabla(self, tbl):
        filas = []
        for tr in tbl.findall(W + "tr"):
            filas.append([self.celda(tc) for tc in tr.findall(W + "tc")])
        return filas

    def tablas_con_titulo(self):
        """Lista de (titulo, tabla): el titulo es el ultimo parrafo con texto antes de la tabla."""
        salida, ultimo = [], ""
        for b in self.bloques():
            if b.tag == W + "p":
                t = limpio(self.texto_parrafo(b))
                if t:
                    ultimo = t
            else:
                salida.append((ultimo, b))
                ultimo = ""
        return salida

    def buscar_tabla(self, titulo):
        """Filas de la tabla con ese titulo (encima de ella o en su primer renglon)."""
        buscado = fold(titulo)
        tablas = self.tablas_con_titulo()
        for t, tbl in tablas:
            if buscado in fold(t):
                return self.filas_de_tabla(tbl)
        for _, tbl in tablas:
            filas = self.filas_de_tabla(tbl)
            if filas and buscado in fold(" ".join(c[0] for c in filas[0])):
                return filas[1:]
        vistos = [t or "(sin título)" for t, _ in tablas]
        raise Falla("No encontré la tabla \"%s\" en %s (%s). Tablas que sí hay, con el texto que "
                    "tienen encima: %s. Revisa el nombre de la tabla o corre con --titulo."
                    % (titulo, self.que, self.ruta.name,
                       "; ".join(vistos) if vistos else "ninguna"))

    def bytes_de(self, destino):
        parte = destino.lstrip("/") if destino.startswith("/") else "word/" + destino
        try:
            return self.zip.read(parte)
        except KeyError:
            return None


# ------------------------------------------------------------------ escritura xlsx
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _col_letra(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _hoja_xml(filas, anchos):
    partes = []
    for i, fila in enumerate(filas):
        celdas = []
        for j, v in enumerate(fila):
            if v is None or v == "":
                continue
            ref = "%s%d" % (_col_letra(j), i + 1)
            s = ' s="1"' if i == 0 else ' s="2"'
            if isinstance(v, int):
                celdas.append('<c r="%s"%s><v>%d</v></c>' % (ref, s, v))
            else:
                celdas.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                              % (ref, s, escape(str(v))))
        if celdas:
            partes.append('<row r="%d">%s</row>' % (i + 1, "".join(celdas)))
    cols = "".join('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (j + 1, j + 1, w)
                   for j, w in enumerate(anchos))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s"><cols>%s</cols><sheetData>%s</sheetData></worksheet>'
            % (NS_MAIN, cols, "".join(partes)))


def escribir_xlsx(ruta, hojas):
    """hojas: lista de (nombre, filas, anchos). Primera fila en negrita; texto con ajuste."""
    n = len(hojas)
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
         '<workbook xmlns="%s" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
         '<sheets>%s</sheets></workbook>'
         % (NS_MAIN, "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
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
         '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
         '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
         '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">'
         '<alignment wrapText="1" vertical="top"/></xf></cellXfs>'
         '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
         '</styleSheet>' % NS_MAIN),
    ]
    for i, (_, filas, anchos) in enumerate(hojas):
        archivos.append(("xl/worksheets/sheet%d.xml" % (i + 1), _hoja_xml(filas, anchos)))
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in archivos:
            z.writestr(nombre, contenido.encode("utf-8"))


# ------------------------------------------------------------------ principal
def base_de_nombre(ruta):
    """'Reporte_Traducción_FIC1001_NarrativaFicticia' -> 'FIC1001_NarrativaFicticia'."""
    stem = ruta.stem
    prefijo = fold("Reporte_Traducción_")
    if fold(stem).startswith(prefijo):
        return stem[len("Reporte_Traducción_"):] or stem
    return stem


def principal(args):
    rep = Word(Path(args.reporte).expanduser(), "el reporte de traducción")
    prof = Word(Path(args.profesor).expanduser(), "el reporte del profesor revisor")
    filas_rep = rep.buscar_tabla(args.titulo)
    filas_prof = prof.buscar_tabla(args.titulo)
    if not filas_rep:
        raise Falla("La tabla \"%s\" de %s no trae renglones." % (args.titulo, rep.ruta.name))
    enc_rep = [limpio(c[0]) for c in filas_rep[0]]
    if len(enc_rep) < COLUMNAS_QUE_SE_COPIAN:
        raise Falla("La tabla \"%s\" de %s tiene %d columna(s); necesito al menos %d."
                    % (args.titulo, rep.ruta.name, len(enc_rep), COLUMNAS_QUE_SE_COPIAN))
    enc_prof = [limpio(c[0]) for c in filas_prof[0]] if filas_prof else []
    if len(enc_prof) < COLUMNAS_QUE_SE_COPIAN:
        raise Falla("La tabla \"%s\" de %s tiene %d columna(s); necesito al menos %d para pegar."
                    % (args.titulo, prof.ruta.name, len(enc_prof), COLUMNAS_QUE_SE_COPIAN))
    destino = enc_prof[:COLUMNAS_QUE_SE_COPIAN]
    avisos = []
    llenos_prof = sum(1 for f in filas_prof[1:]
                      if any(c[0] for c in f[:COLUMNAS_QUE_SE_COPIAN]))
    if len(enc_prof) != 4:
        avisos.append("La tabla de %s tiene %d columnas; esperaba 4." % (prof.ruta.name,
                                                                         len(enc_prof)))
    if llenos_prof:
        avisos.append("La tabla de %s ya trae %s en las dos primeras columnas: al pegar, "
                      "decide si los nuevos van abajo o reemplazan."
                      % (prof.ruta.name, plural(llenos_prof, "renglón lleno", "renglones llenos")))

    out = Path(args.salida).expanduser() if args.salida else rep.ruta.parent / "salida"
    carpeta_img = out / "imagenes"
    out.mkdir(parents=True, exist_ok=True)
    base = base_de_nombre(rep.ruta)

    tabla, vacios, n_ligas, n_img = [], 0, 0, 0
    guardadas = []
    for i, fila in enumerate(filas_rep[1:], start=2):
        celdas = (fila + [("", [], [])] * COLUMNAS_QUE_SE_COPIAN)[:COLUMNAS_QUE_SE_COPIAN]
        if not any(t or l or im for t, l, im in celdas):
            vacios += 1
            continue
        ligas, imgs = [], []
        for col, (texto, ls, ims) in enumerate(celdas, start=1):
            for l in ls:
                ligas.append("col. %d: %s" % (col, l))
            for k, ref in enumerate(ims, start=1):
                datos = rep.bytes_de(ref)
                ext = Path(ref).suffix.lower() or ".png"
                nombre = "renglon-%02d-col-%d-%d%s" % (i, col, k, ext)
                if datos is None:
                    avisos.append("La imagen del renglón %d, columna %d no está dentro del documento "
                                  "(¿es una liga a una imagen?); no la pude guardar." % (i, col))
                    continue
                carpeta_img.mkdir(parents=True, exist_ok=True)
                (carpeta_img / nombre).write_bytes(datos)
                imgs.append(nombre)
                guardadas.append(nombre)
        n_ligas += len(ligas)
        n_img += len(imgs)
        tabla.append([i, celdas[0][0], celdas[1][0], "\n".join(ligas), "\n".join(imgs)])

    encabezado = ["Renglón en el reporte", destino[0], destino[1], "Ligas", "Capturas de pantalla"]
    ruta_csv = out / ("tabla-para-profesor-%s.csv" % base)
    ruta_xlsx = out / ("tabla-para-profesor-%s.xlsx" % base)
    with open(ruta_csv, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(encabezado)
        w.writerows(tabla)

    l1 = ("De %s se tomaron %s de la tabla %s (columnas 1 y 2: %s y %s; %s)." % (
        rep.ruta.name, plural(len(tabla), "renglón", "renglones"), args.titulo,
        enc_rep[0], enc_rep[1],
        "ningún renglón vacío" if not vacios else
        plural(vacios, "renglón vacío omitido", "renglones vacíos omitidos")))
    l2 = ("Traen %s y %s%s." % (
        plural(n_ligas, "liga", "ligas"), plural(n_img, "captura de pantalla",
                                                  "capturas de pantalla"),
        "; las capturas quedaron en imagenes/ con el renglón y la columna en el nombre"
        if n_img else ""))
    l3 = ("La tabla de %s tiene %d columnas y %s: %s en %s y %s. Detalle: %s." % (
        prof.ruta.name, len(enc_prof), plural(llenos_prof, "renglón lleno", "renglones llenos"),
        "lista para pegar" if not llenos_prof else "revisa antes de pegar",
        destino[0], destino[1], ruta_xlsx.name))
    lineas = [l1, l2, l3]

    control = [
        ("Reporte de traducción", rep.ruta.name),
        ("Reporte del profesor", prof.ruta.name),
        ("Tabla", args.titulo),
        ("Columnas en el reporte", "%d: %s" % (len(enc_rep), " | ".join(enc_rep))),
        ("Columnas en la tabla del profesor", "%d: %s" % (len(enc_prof), " | ".join(enc_prof))),
        ("Renglones de datos en el reporte", len(filas_rep) - 1),
        ("Renglones copiados", len(tabla)),
        ("Renglones vacíos omitidos", vacios),
        ("Ligas", n_ligas),
        ("Capturas de pantalla guardadas", n_img),
        ("Renglones ya llenos en la tabla del profesor", llenos_prof),
    ]
    hoja_control = [["Concepto", "Valor"]] + [[k, v] for k, v in control]
    hoja_control += [[], ["Resumen", ""]] + [[l, ""] for l in lineas]
    if avisos:
        hoja_control += [[], ["Avisos", ""]] + [[a, ""] for a in avisos]
    escribir_xlsx(ruta_xlsx, [("Tabla para pegar", [encabezado] + tabla, [12, 40, 55, 60, 28]),
                              ("Control", hoja_control, [45, 80])])
    (out / "resumen-tabla-recursos.txt").write_text("\n".join(lineas) + "\n", encoding="utf-8")

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
        print("  %s: %s" % (k, v))
    print()
    print("Renglones para pegar (renglón del reporte: columna 1 | ligas | capturas):")
    for f in tabla:
        print("  %d: %s | %d liga(s) | %s" % (f[0], f[1].replace("\n", " / ") or "(vacío)",
                                             len(f[3].split("\n")) if f[3] else 0,
                                             f[4].replace("\n", ", ") or "sin captura"))
    print()
    print("Archivos:")
    print("  %s" % ruta_csv)
    print("  %s" % ruta_xlsx)
    for g in guardadas:
        print("  %s" % (carpeta_img / g))
    print("  %s" % (out / "resumen-tabla-recursos.txt"))
    return 0


def main():
    ap = argparse.ArgumentParser(
        description="Deja listas para pegar las dos primeras columnas de la tabla de recursos "
                    "en inglés (skill tabla-recursos-ingles).")
    ap.add_argument("reporte", help="Reporte_Traducción_[CLAVE]_[NombreUF].docx")
    ap.add_argument("profesor", help="ProfesorRevisor_PL2026_UF.docx")
    ap.add_argument("--titulo", default=TITULO_TABLA,
                    help="título de la tabla en los dos documentos (por omisión: %s)"
                         % TITULO_TABLA)
    ap.add_argument("--salida", default=None,
                    help="carpeta de salida (por omisión, salida/ junto al reporte)")
    args = ap.parse_args()
    try:
        return principal(args)
    except Falla as e:
        print("No pude copiar la tabla: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
