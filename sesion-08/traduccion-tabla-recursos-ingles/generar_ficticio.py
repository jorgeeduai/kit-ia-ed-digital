#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera la copia FICTICIA de los dos documentos de Word del flujo de la
tabla "Recursos que requieren una version en ingles" (curso Ed Digital, propuesta de skill del
grupo, ronda 2, equipo de traduccion):

  1. Reporte_Traduccion_[CLAVE]_[NombreUF].docx   el reporte de traduccion. Trae tres tablas;
     la de "Recursos que requieren una version en ingles" tiene tres columnas, con ligas y
     capturas de pantalla dentro de las celdas.
  2. ProfesorRevisor_PL2026_UF.docx               el reporte para el profesor. La misma tabla,
     con cuatro columnas y vacia, lista para recibir las dos primeras columnas del reporte.

Los encabezados de las tablas son SUPUESTOS armados con lo que se leyo en voz en la Sesion 7;
estan en las constantes de abajo para cambiarlos en cuanto se confirmen.

Como correrlo (solo biblioteca estandar, Python 3.8 o mas reciente), desde esta carpeta:
    python3 -B generar_ficticio.py
    python3 -B generar_ficticio.py --semilla 22 --salida ficticio

Salen, dentro de --salida (ficticio/ por omision):
    Reporte_Traducción_FIC1001_NarrativaFicticia.docx
    ProfesorRevisor_PL2026_UF.docx
    NOTAS-ficticio.txt                lo que se sembro y la respuesta correcta

Todo es inventado: la unidad de formacion, los recursos, las ligas (dominio example.org, que no
existe) y las capturas (rectangulos de color hechos por el programa).
"""
import argparse
import random
import struct
import sys
import zipfile
import zlib
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

# ------------------------------------------------------------------ supuestos (cambiar aqui)
TITULO_TABLA = "Recursos que requieren una versión en inglés"
ENC_REPORTE = ["Nombre del recurso en español",
               "Comentarios y enlace donde aparece en Canvas en inglés",
               "Estatus de la revisión"]
ENC_PROFESOR = ["Recurso en español", "Comentarios y enlace en Canvas",
                "Respuesta del profesor", "Comentarios del profesor"]
CLAVE, NOMBRE_UF = "FIC1001", "NarrativaFicticia"
NOMBRE_REPORTE = "Reporte_Traducción_%s_%s.docx" % (CLAVE, NOMBRE_UF)
NOMBRE_PROFESOR = "ProfesorRevisor_PL2026_UF.docx"
BASE_LIGA = "https://canvas.example.org/courses/0001/"

# Renglones de la tabla del reporte. Cada celda es una lista de "piezas":
#   ("t", texto)  ("l", texto, ruta)  ("img", ancho, alto)  ("p",) nuevo parrafo
# El tercer valor es la columna 3 (no se copia).
RENGLONES = [
    ([("t", "Lectura 1. Orígenes del relato (PDF)")],
     [("t", "No encontramos versión en inglés. Aparece en "), ("l", "Módulo 1, Lectura 1",
                                                                "pages/lectura-1")],
     "Aprobado", "liga dentro del texto"),
    ([("t", "Video: cómo se construye un personaje")],
     [("t", "Hay una versión con subtítulos en inglés; falta confirmar. "),
      ("l", "Ver en Canvas", "pages/video-personaje")],
     "Aprobado", "liga dentro del texto"),
    ([("t", "Infografía de la estructura en tres actos")],
     [("img", 120, 60), ("p",), ("t", "Captura de la infografía como aparece en Canvas.")],
     "Aprobado", "captura de pantalla en la celda"),
    ([], [], "", "renglón vacío: no se copia"),
    ([("t", "Rúbrica del reto final")],
     [("t", "Requiere versión en inglés."), ("p",), ("t", "Aparece en el Módulo 4: "),
      ("l", "Reto final", "assignments/reto-final")],
     "Pendiente", "dos párrafos en la celda y una liga"),
    ([("t", "Podcast: voces de la narrativa")],
     [("t", "Solo existe en español: " + BASE_LIGA + "files/22")],
     "Aprobado", "liga escrita como texto, sin hipervínculo"),
    ([("t", "Mapa conceptual del curso")],
     [("img", 90, 90), ("p",), ("l", "Página de inicio", "pages/inicio")],
     "Aprobado", "captura de pantalla y liga en la misma celda"),
    ([("t", "Guía de estilo para ensayos (Word)")],
     [("t", "Se envía la alternativa encontrada para revisión del profesor (ver tabla 3).")],
     "Aprobado", "texto con acentos y paréntesis"),
]

DISTRACTOR_ANTES = ("Recursos traducidos sin observaciones",
                    ["Recurso", "Tipo", "Observación"],
                    [["Bienvenida al curso", "Página", "Sin cambios"],
                     ["Calendario de actividades", "Página", "Sin cambios"],
                     ["Foro de dudas", "Foro", "Sin cambios"]])
DISTRACTOR_DESPUES = ("Alternativas encontradas para revisión del profesor",
                      ["Recurso original", "Alternativa en inglés", "Motivo"],
                      [["Guía de estilo para ensayos", "Essay style guide (ficticia)",
                        "No es copia exacta del original"]])


# ------------------------------------------------------------------ imagen PNG sin instalar nada
def png(ancho, alto, color):
    fila = b"\x00" + bytes(color) * ancho
    crudo = fila * alto

    def trozo(tipo, datos):
        return (struct.pack(">I", len(datos)) + tipo + datos
                + struct.pack(">I", zlib.crc32(tipo + datos) & 0xFFFFFFFF))
    return (b"\x89PNG\r\n\x1a\n" + trozo(b"IHDR", struct.pack(">IIBBBBB", ancho, alto, 8, 2, 0, 0, 0))
            + trozo(b"IDAT", zlib.compress(crudo, 9)) + trozo(b"IEND", b""))


# ------------------------------------------------------------------ armado del .docx
class Documento:
    def __init__(self):
        self.cuerpo = []
        self.rels = []          # (id, tipo, destino, externo)
        self.medios = []        # (nombre, bytes)
        self.n_img = 0

    def _rid(self, tipo, destino, externo=False):
        rid = "rId%d" % (len(self.rels) + 10)
        self.rels.append((rid, tipo, destino, externo))
        return rid

    def _run(self, texto, liga=False):
        rpr = ('<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>' if liga else "")
        return '<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, escape(texto))

    def _imagen(self, ancho, alto, color):
        self.n_img += 1
        nombre = "image%d.png" % self.n_img
        self.medios.append((nombre, png(ancho, alto, color)))
        rid = self._rid("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "media/" + nombre)
        cx, cy = ancho * 9525, alto * 9525
        return ('<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
                '<wp:extent cx="%d" cy="%d"/><wp:docPr id="%d" name="Captura %d"/>'
                '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                '<pic:nvPicPr><pic:cNvPr id="%d" name="%s"/><pic:cNvPicPr/></pic:nvPicPr>'
                '<pic:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
                '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'
                % (cx, cy, self.n_img, self.n_img, self.n_img, nombre, rid, cx, cy))

    def parrafos_de(self, piezas, rnd):
        """Convierte piezas en uno o varios <w:p>."""
        parrafos, actual = [], []
        for p in piezas:
            if p[0] == "t":
                actual.append(self._run(p[1]))
            elif p[0] == "l":
                rid = self._rid("http://schemas.openxmlformats.org/officeDocument/2006/"
                                "relationships/hyperlink", BASE_LIGA + p[2], externo=True)
                actual.append('<w:hyperlink r:id="%s">%s</w:hyperlink>'
                              % (rid, self._run(p[1], liga=True)))
            elif p[0] == "img":
                color = (rnd.randint(60, 200), rnd.randint(60, 200), rnd.randint(60, 200))
                actual.append(self._imagen(p[1], p[2], color))
            elif p[0] == "p":
                parrafos.append(actual)
                actual = []
        parrafos.append(actual)
        return "".join("<w:p>%s</w:p>" % "".join(x) for x in parrafos)

    def titulo(self, texto, nivel=1):
        tam = 32 if nivel == 1 else 26
        self.cuerpo.append('<w:p><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>'
                           '<w:r><w:rPr><w:b/><w:sz w:val="%d"/></w:rPr><w:t xml:space="preserve">'
                           '%s</w:t></w:r></w:p>' % (tam, escape(texto)))

    def parrafo(self, texto):
        self.cuerpo.append("<w:p>%s</w:p>" % self._run(texto))

    def tabla(self, encabezado, renglones_xml):
        n = len(encabezado)
        ancho = 9000 // n
        borde = ('<w:tblBorders>%s</w:tblBorders>' % "".join(
            '<w:%s w:val="single" w:sz="4" w:space="0" w:color="808080"/>' % b
            for b in ("top", "left", "bottom", "right", "insideH", "insideV")))
        celda = lambda contenido: ('<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>%s</w:tc>'
                                   % (ancho, contenido or "<w:p/>"))
        filas = ['<w:tr>%s</w:tr>' % "".join(
            celda('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">%s</w:t></w:r></w:p>'
                  % escape(h)) for h in encabezado)]
        for r in renglones_xml:
            filas.append('<w:tr>%s</w:tr>' % "".join(celda(c) for c in r))
        self.cuerpo.append('<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/>%s</w:tblPr>'
                           '<w:tblGrid>%s</w:tblGrid>%s</w:tbl>'
                           % (borde, ('<w:gridCol w:w="%d"/>' % ancho) * n, "".join(filas)))
        self.cuerpo.append("<w:p/>")

    def guardar(self, ruta):
        doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
               'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
               '<w:body>%s<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
               '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" '
               'w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>'
               % "".join(self.cuerpo))
        rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '%s</Relationships>' % "".join(
                    '<Relationship Id="%s" Type="%s" Target=%s%s/>'
                    % (rid, tipo, quoteattr(dest), ' TargetMode="External"' if ext else "")
                    for rid, tipo, dest, ext in self.rels))
        tipos = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                 '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                 '<Default Extension="xml" ContentType="application/xml"/>'
                 '<Default Extension="png" ContentType="image/png"/>'
                 '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                 '</Types>')
        raiz = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '</Relationships>')
        partes = [("[Content_Types].xml", tipos.encode("utf-8")),
                  ("_rels/.rels", raiz.encode("utf-8")),
                  ("word/document.xml", doc.encode("utf-8")),
                  ("word/_rels/document.xml.rels", rels.encode("utf-8"))]
        partes += [("word/media/" + n, b) for n, b in self.medios]
        with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
            for nombre, datos in partes:
                info = zipfile.ZipInfo(nombre, date_time=(2026, 10, 22, 0, 0, 0))
                info.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(info, datos)


def tabla_simple(doc, titulo, encabezado, filas):
    doc.titulo(titulo, 2)
    doc.tabla(encabezado, [["<w:p>%s</w:p>" % doc._run(c) for c in f] for f in filas])


def principal():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia de los dos documentos de "
                                             "Word del flujo de la tabla de recursos en inglés.")
    ap.add_argument("--semilla", type=int, default=22, help="fija el azar (22)")
    ap.add_argument("--salida", default="ficticio", help="carpeta de salida (ficticio)")
    args = ap.parse_args()
    rnd = random.Random(args.semilla)
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)

    # 1. Reporte de traduccion
    rep = Documento()
    rep.titulo("Reporte de traducción · %s %s" % (CLAVE, "Narrativa ficticia"))
    rep.parrafo("Unidad de formación ficticia, generada por un programa para practicar.")
    tabla_simple(rep, *DISTRACTOR_ANTES)
    rep.titulo(TITULO_TABLA, 2)
    filas = []
    for c1, c2, c3, _ in RENGLONES:
        filas.append([rep.parrafos_de(c1, rnd) if c1 else "",
                      rep.parrafos_de(c2, rnd) if c2 else "",
                      "<w:p>%s</w:p>" % rep._run(c3) if c3 else ""])
    rep.tabla(ENC_REPORTE, filas)
    tabla_simple(rep, *DISTRACTOR_DESPUES)
    rep.guardar(out / NOMBRE_REPORTE)

    # 2. Reporte para el profesor revisor: misma tabla, cuatro columnas, vacia
    prof = Documento()
    prof.titulo("Reporte para el profesor revisor · PL2026")
    prof.parrafo("Documento ficticio. La tabla de recursos llega vacía.")
    tabla_simple(prof, "Datos de la unidad de formación", ["Dato", "Valor"],
                 [["Clave", CLAVE], ["Unidad de formación", "Narrativa ficticia"]])
    prof.titulo(TITULO_TABLA, 2)
    prof.tabla(ENC_PROFESOR, [["", "", "", ""]])
    tabla_simple(prof, "Alternativas para revisión", ["Recurso original", "Alternativa",
                                                       "Respuesta del profesor"], [["", "", ""]])
    prof.guardar(out / NOMBRE_PROFESOR)

    # Respuesta correcta
    copiados = [r for r in RENGLONES if r[0] or r[1]]
    ligas = sum(1 for r in copiados for p in r[0] + r[1]
                if p[0] == "l" or (p[0] == "t" and "https://" in p[1]))
    imagenes = sum(1 for r in copiados for p in r[0] + r[1] if p[0] == "img")
    lineas = [
        "NOTAS-ficticio.txt: lo que sembró generar_ficticio.py (semilla %d)" % args.semilla,
        "Todo es inventado: la unidad de formación, los recursos, las ligas (example.org) y las",
        "capturas (rectángulos de color). Los encabezados son SUPUESTOS (ver LEEME.md).",
        "",
        "%s: tres tablas, en este orden:" % NOMBRE_REPORTE,
        "  1. %s (3 columnas, distractor: no se copia)" % DISTRACTOR_ANTES[0],
        "  2. %s (3 columnas: %s)" % (TITULO_TABLA, " | ".join(ENC_REPORTE)),
        "  3. %s (3 columnas, distractor: no se copia)" % DISTRACTOR_DESPUES[0],
        "%s: la tabla %s con 4 columnas (%s) y un renglón vacío; antes, una tabla de datos"
        % (NOMBRE_PROFESOR, TITULO_TABLA, " | ".join(ENC_PROFESOR)),
        "  de la unidad, y después, una de alternativas.",
        "",
        "Renglones de la tabla del reporte (renglón de la tabla, contando el encabezado como 1):",
    ]
    for i, (c1, c2, c3, que) in enumerate(RENGLONES, start=2):
        texto = "".join(p[1] for p in c1 if p[0] in ("t", "l")) or "(vacío)"
        lineas.append("  %d. %s | %s" % (i, texto, que))
    lineas += [
        "",
        "Respuesta correcta:",
        "  Renglones que se copian: %d (el renglón 5 está vacío y no se copia)" % len(copiados),
        "  Ligas en las columnas 1 y 2: %d (4 hipervínculos y 1 escrita como texto)" % ligas,
        "  Capturas de pantalla en las columnas 1 y 2: %d (renglones 4 y 8)" % imagenes,
        "  La columna 3 (%s) no se copia." % ENC_REPORTE[2],
        "  La tabla del profesor tiene 4 columnas y 0 renglones llenos.",
        "  Las tablas 1 y 3 del reporte no se tocan.",
    ]
    (out / "NOTAS-ficticio.txt").write_text("\n".join(lineas) + "\n", encoding="utf-8")

    print("Listo. Archivos en %s/:" % out)
    print("  %s (tres tablas; la de recursos con %d renglones)" % (NOMBRE_REPORTE, len(RENGLONES)))
    print("  %s (la tabla de recursos vacía, cuatro columnas)" % NOMBRE_PROFESOR)
    print("  NOTAS-ficticio.txt")
    print()
    print("Sembrado: %d renglones por copiar, %d ligas y %d capturas de pantalla; un renglón vacío "
          "y dos tablas que no se tocan." % (len(copiados), ligas, imagenes))
    return 0


if __name__ == "__main__":
    sys.exit(principal())
