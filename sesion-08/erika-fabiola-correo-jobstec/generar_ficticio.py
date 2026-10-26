#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_ficticio.py: genera dos copias FICTICIAS del Excel que llega del Forms de altas de
profesores (flujo de Erika Fabiola, curso Ed Digital, borrador post-S7). Imita el archivo que
se descarga del Forms dos veces: el lunes 26 y el miercoles 28 de octubre de 2026. El del
miercoles trae los mismos renglones del lunes y, debajo, las respuestas nuevas, porque el
Forms acumula.

Como correrlo (solo biblioteca estandar, Python 3.8 o mas reciente), desde esta carpeta:
    python3 -B generar_ficticio.py

Parametros:
    --semilla  numero que fija el azar (26). Misma semilla, mismos archivos.
    --salida   carpeta donde se escriben los dos archivos (entrada).
    --notas    ruta de NOTAS-ficticio.txt (por omision, NOTAS-ficticio.txt en esta carpeta).

Salen:
    entrada/respuestas-forms-2026-10-26.xlsx   20 renglones
    entrada/respuestas-forms-2026-10-28.xlsx   26 renglones (los 20 del lunes y 6 nuevos)
    NOTAS-ficticio.txt                         lo que se sembro y la respuesta correcta

Columnas, tal como las escribio Erika Fabiola en el muro: Nombre del profesor, Correo electronico.
Todos los nombres son inventados y todos los correos usan el dominio ejemplo.com: ninguno
corresponde a una persona real.
"""
import argparse
import random
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

COLUMNAS = ["Nombre del profesor", "Correo electrónico"]
ANCHOS = [36, 40]
DOMINIO = "ejemplo.com"
ARCHIVOS = ["respuestas-forms-2026-10-26.xlsx", "respuestas-forms-2026-10-28.xlsx"]

NOMBRES = ["Andrea", "Luis", "Mónica", "Ricardo", "Paola", "Héctor", "Daniela", "Arturo",
           "Gabriela", "Fernando", "Lucía", "Óscar", "Valeria", "Sergio", "Adriana", "Tomás",
           "Renata", "Julián", "Marisol", "Emilio", "Silvia", "Raúl", "Claudia", "Iván"]
APELLIDOS = ["Treviño", "Garza", "Villarreal", "Cantú", "Salazar", "Ibarra", "Montemayor",
             "Zambrano", "Elizondo", "Guajardo", "Leal", "Quiroga", "Benavides", "Saldívar",
             "Cavazos", "Rendón", "Olivares", "Escamilla", "Lozano", "Barragán"]

N_LUNES_VALIDOS = 15     # personas con nombre y correo correctos en el archivo del lunes
N_MIERCOLES_NUEVOS = 5   # personas nuevas en el del miercoles


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def correo_de(nombre, apellido):
    return "%s.%s@%s" % (sin_acentos(nombre).lower(), sin_acentos(apellido).lower(), DOMINIO)


def personas(rnd, n):
    """n personas inventadas, sin repetir correo."""
    vistas, salida = set(), []
    while len(salida) < n:
        nom, ap1, ap2 = rnd.choice(NOMBRES), rnd.choice(APELLIDOS), rnd.choice(APELLIDOS)
        if ap1 == ap2:
            continue
        correo = correo_de(nom, ap1)
        if correo in vistas:
            continue
        vistas.add(correo)
        salida.append(("%s %s %s" % (nom, ap1, ap2), correo))
    return salida


# ------------------------------------------------------------------ escritura xlsx
# Las mismas funciones de escritura de comparar_versiones.py (Sesion 7).
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _col_letra(i):
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def escribir_xlsx(ruta, nombre_hoja, filas, anchos):
    celdas_xml = []
    for i, fila in enumerate(filas):
        celdas = []
        for j, v in enumerate(fila):
            if v is None or v == "":
                continue
            ref = "%s%d" % (_col_letra(j), i + 1)
            s = ' s="1"' if i == 0 else ""
            celdas.append('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                          % (ref, s, escape(str(v))))
        if celdas:
            celdas_xml.append('<row r="%d">%s</row>' % (i + 1, "".join(celdas)))
    cols = "".join('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (j + 1, j + 1, w)
                   for j, w in enumerate(anchos))
    hoja = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="%s" xmlns:r="%s"><sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '</sheetView></sheetViews><cols>%s</cols><sheetData>%s</sheetData></worksheet>'
            % (NS_MAIN, NS_REL, cols, "".join(celdas_xml)))
    archivos = [
        ("[Content_Types].xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
         '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
         '<Default Extension="xml" ContentType="application/xml"/>'
         '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
         '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
         '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
         '</Types>'),
        ("_rels/.rels",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
         '</Relationships>'),
        ("xl/workbook.xml",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="%s" sheetId="1" r:id="rId1"/>'
         '</sheets></workbook>' % (NS_MAIN, NS_REL, escape(nombre_hoja))),
        ("xl/_rels/workbook.xml.rels",
         '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
         '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
         '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
         '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
         '</Relationships>'),
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
        ("xl/worksheets/sheet1.xml", hoja),
    ]
    with zipfile.ZipFile(ruta, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, contenido in archivos:
            zi = zipfile.ZipInfo(nombre, date_time=(2026, 10, 26, 9, 0, 0))
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, contenido.encode("utf-8"))


# ------------------------------------------------------------------ principal
def main():
    ap = argparse.ArgumentParser(description="Genera la copia ficticia del Excel del Forms.")
    ap.add_argument("--semilla", type=int, default=26)
    ap.add_argument("--salida", default=None)
    ap.add_argument("--notas", default=None)
    args = ap.parse_args()

    aqui = Path(__file__).resolve().parent
    out = Path(args.salida).expanduser() if args.salida else aqui / "entrada"
    notas = Path(args.notas).expanduser() if args.notas else aqui / "NOTAS-ficticio.txt"
    out.mkdir(parents=True, exist_ok=True)
    rnd = random.Random(args.semilla)

    gente = personas(rnd, N_LUNES_VALIDOS + 3 + N_MIERCOLES_NUEVOS)
    validos = gente[:N_LUNES_VALIDOS]
    p_sin_arroba, p_sin_correo, p_sin_nombre = gente[N_LUNES_VALIDOS:N_LUNES_VALIDOS + 3]
    nuevos = gente[N_LUNES_VALIDOS + 3:]

    # Lunes: 15 correctos en orden de llegada, con un correo escrito con espacios y mayusculas
    lunes = [list(p) for p in validos]
    con_espacios = lunes[6]
    con_espacios[1] = "  " + con_espacios[1].title().replace("@Ejemplo.Com", "@Ejemplo.com") + " "
    # Dos personas que llenaron el Forms dos veces (una con el correo en mayusculas)
    orig_1, orig_2 = lunes[1], lunes[8]
    dup_1, dup_2 = [orig_1[0], orig_1[1]], [orig_2[0], orig_2[1].upper()]
    lunes.insert(10, dup_1)
    lunes.insert(15, dup_2)
    # Tres renglones que no se pueden atender
    sin_arroba = [p_sin_arroba[0], p_sin_arroba[1].replace("@", " ")]
    sin_correo = [p_sin_correo[0], ""]
    sin_nombre = ["", p_sin_nombre[1]]
    lunes.insert(4, sin_arroba)
    lunes.insert(12, sin_correo)
    lunes.append(sin_nombre)

    # Miercoles: lo del lunes y debajo 6 respuestas: 5 nuevas y una persona del lunes que
    # vuelve a llenar el Forms (ya tiene borrador)
    orig_r = lunes[3] if lunes[3] is not sin_arroba else lunes[4]
    reenvio = [orig_r[0], orig_r[1].strip().lower()]
    mie_extra = [list(p) for p in nuevos[:3]] + [reenvio] + [list(p) for p in nuevos[3:]]
    miercoles = list(lunes) + mie_extra

    for nombre, filas in zip(ARCHIVOS, (lunes, miercoles)):
        escribir_xlsx(out / nombre, "Respuestas de formulario", [COLUMNAS] + filas, ANCHOS)

    def renglon(lista, fila):
        """Renglon de Excel de esa fila (por identidad, no por valor): +1 base cero, +1 encabezado."""
        return next(i for i, f in enumerate(lista) if f is fila) + 2

    n_lun, n_mie = len(lunes), len(miercoles)
    revisar = 3
    repetidos = 2
    borr_lunes = n_lun - revisar - repetidos
    ya_atendidos = borr_lunes + repetidos + 1
    borr_mie = n_mie - revisar - ya_atendidos
    lineas = [
        "NOTAS DE LA COPIA FICTICIA · correo de JobsTec (flujo de Erika Fabiola)",
        "",
        "Todo es inventado: ningún nombre corresponde a una persona real y todos los correos",
        "usan el dominio %s. Este archivo lo escribe generar_ficticio.py." % DOMINIO,
        "Semilla: %d. Comando, desde esta carpeta: python3 -B generar_ficticio.py" % args.semilla,
        "Columnas, en este orden: %s." % ", ".join(COLUMNAS),
        "",
        "ARCHIVOS",
        "  entrada/%s  %d renglones de datos (lunes 26-oct)" % (ARCHIVOS[0], n_lun),
        "  entrada/%s  %d renglones de datos (miércoles 28-oct): los %d del lunes y %d nuevos"
        % (ARCHIVOS[1], n_mie, n_lun, n_mie - n_lun),
        "",
        "LO QUE SE SEMBRÓ EN EL DEL LUNES",
        "  Personas con nombre y correo correctos: %d" % N_LUNES_VALIDOS,
        "  De esas, una con el correo escrito con espacios y mayúsculas (renglón %d): se limpia"
        % renglon(lunes, con_espacios),
        "    y SÍ lleva borrador.",
        "  Personas que llenaron el Forms dos veces: 2. El renglón %d repite al %d y el %d"
        % (renglon(lunes, dup_1), renglon(lunes, orig_1), renglon(lunes, dup_2)),
        "    repite al %d con el correo en mayúsculas. Cada una lleva un solo borrador."
        % renglon(lunes, orig_2),
        "  Renglones que no se pueden atender: 3",
        "    renglón %d: correo sin arroba" % renglon(lunes, sin_arroba),
        "    renglón %d: sin correo" % renglon(lunes, sin_correo),
        "    renglón %d: sin nombre" % renglon(lunes, sin_nombre),
        "",
        "LO QUE SE SEMBRÓ EN EL DEL MIÉRCOLES",
        "  Personas nuevas: %d" % N_MIERCOLES_NUEVOS,
        "  Una persona del lunes que vuelve a llenar el Forms (renglón %d, repite el renglón %d)"
        % (renglon(miercoles, reenvio), renglon(lunes, orig_r)),
        "",
        "RESPUESTA CORRECTA",
        "  1. Con el del lunes y el registro de atendidos vacío:",
        "     %d renglones: %d borradores, %d repetidos en el mismo archivo, %d por revisar."
        % (n_lun, borr_lunes, repetidos, revisar),
        "  2. Después, con el del miércoles y el registro que dejó la corrida 1:",
        "     %d renglones: %d borradores, %d ya atendidos, %d por revisar."
        % (n_mie, borr_mie, ya_atendidos, revisar),
        "     (los %d ya atendidos son los %d del lunes, sus 2 repeticiones y el que volvió a"
        % (ya_atendidos, borr_lunes),
        "     llenar el Forms)",
        "  3. Si se corre otra vez con el del miércoles: 0 borradores, %d ya atendidos, %d por revisar."
        % (ya_atendidos + borr_mie, revisar),
    ]
    notas.write_text("\n".join(lineas) + "\n", encoding="utf-8")
    print("\n".join(lineas))
    return 0


if __name__ == "__main__":
    sys.exit(main())
