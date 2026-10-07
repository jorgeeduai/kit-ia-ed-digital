"""alumnos_por_campus.py · script de referencia de la Sesión 5 (datos ficticios).

Qué hace:
  1. Lee todos los archivos inscripciones-*.csv de una carpeta.
  2. Normaliza la columna Campus: quita espacios al inicio y al final y la escribe
     con mayúscula inicial ("MONTERREY" y "Monterrey " cuentan como "Monterrey").
  3. Cuenta alumnos distintos (por Matrícula alumno) por campus y por periodo.
     Un alumno se cuenta una sola vez aunque aparezca en varios grupos, en renglones
     duplicados o repetido porque su grupo tiene co-titular. Se cuentan todos los
     estatus de inscripción.
  4. Guarda alumnos-por-campus.xlsx con la tabla y una gráfica de barras por campus.

Cómo se usa (necesita Python con openpyxl), desde la carpeta sesion-05:
  python script-de-referencia/alumnos_por_campus.py        lee los CSV de la subcarpeta csv
  python script-de-referencia/alumnos_por_campus.py RUTA   lee los CSV de la carpeta RUTA

El Excel queda en la carpeta donde se ejecuta (sesion-05), no junto a los CSV. Si se
ejecuta desde otra carpeta que no tiene subcarpeta csv, usa la carpeta sesion-05 donde
vive el script. El script no modifica los CSV.
"""

import csv
import glob
import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font

MINUSCULAS = {"de", "del", "la", "las", "los", "y"}
ORDEN_PERIODO = {"Febrero-Junio": 0, "Agosto-Diciembre": 1}


def normaliza_campus(texto):
    """'CIUDAD DE MÉXICO' o 'Ciudad de México ' -> 'Ciudad de México'."""
    palabras = texto.strip().lower().split()
    salida = []
    for i, p in enumerate(palabras):
        salida.append(p if (i > 0 and p in MINUSCULAS) else p[:1].upper() + p[1:])
    return " ".join(salida)


def clave_periodo(periodo):
    """Ordena 'Agosto-Diciembre 2022' antes que 'Febrero-Junio 2023'."""
    partes = periodo.split()
    return (int(partes[-1]), ORDEN_PERIODO.get(partes[0], 9))


def carpetas(argumento):
    """Devuelve (carpeta de los CSV, carpeta donde se guarda el Excel)."""
    aqui = os.getcwd()
    if argumento:
        return argumento, aqui
    if glob.glob(os.path.join(aqui, "csv", "inscripciones-*.csv")):
        return os.path.join(aqui, "csv"), aqui
    # Banco de archivos de un proyecto: los CSV pueden quedar junto al script, sin subcarpeta
    if glob.glob(os.path.join(aqui, "inscripciones-*.csv")):
        return aqui, aqui
    junto = os.path.dirname(os.path.abspath(__file__))
    if glob.glob(os.path.join(junto, "inscripciones-*.csv")):
        return junto, junto
    # Respaldo: la carpeta sesion-05 que contiene a script-de-referencia/
    sesion = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(sesion, "csv"), sesion


def main():
    carpeta, salida = carpetas(sys.argv[1] if len(sys.argv) > 1 else None)
    archivos = sorted(glob.glob(os.path.join(carpeta, "inscripciones-*.csv")))
    if not archivos:
        print("No encontré archivos inscripciones-*.csv en", os.path.abspath(carpeta))
        sys.exit(1)

    por_celda = {}      # (campus, periodo) -> set de matrículas
    por_campus = {}     # campus -> set de matrículas (todos los periodos)
    renglones = 0
    for ruta in archivos:
        # utf-8-sig: los archivos traen la marca BOM al inicio
        with open(ruta, newline="", encoding="utf-8-sig") as f:
            for fila in csv.DictReader(f):
                renglones += 1
                campus = normaliza_campus(fila["Campus"])
                periodo = fila["Periodo"].strip()
                matricula = fila["Matrícula alumno"].strip()
                por_celda.setdefault((campus, periodo), set()).add(matricula)
                por_campus.setdefault(campus, set()).add(matricula)

    periodos = sorted({p for _, p in por_celda}, key=clave_periodo)
    campus_lista = sorted(por_campus, key=lambda c: -len(por_campus[c]))
    todos = set().union(*por_campus.values())

    # ---- Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Alumnos por campus"
    encabezado = ["Campus"] + periodos + ["Total (alumnos distintos)"]
    ws.append(encabezado)
    for celda in ws[1]:
        celda.font = Font(bold=True)
    for c in campus_lista:
        ws.append([c] + [len(por_celda.get((c, p), ())) for p in periodos] + [len(por_campus[c])])
    totales_periodo = []
    for p in periodos:
        s = set()
        for c in campus_lista:
            s |= por_celda.get((c, p), set())
        totales_periodo.append(len(s))
    ws.append(["Todos los campus"] + totales_periodo + [len(todos)])
    for celda in ws[ws.max_row]:
        celda.font = Font(bold=True)
    ws.column_dimensions["A"].width = 22
    for i in range(2, len(encabezado) + 1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 14

    nota = ws.max_row + 2
    ws.cell(row=nota, column=1, value=(
        "Alumnos distintos por matrícula; cada alumno cuenta una vez por campus y periodo, "
        "sin renglones duplicados ni repetición por co-titular. Datos ficticios."))
    ws.cell(row=nota + 1, column=1, value=f"Archivos leídos: {len(archivos)} · renglones leídos: {renglones}")

    grafica = BarChart()
    grafica.type = "col"
    grafica.title = "Alumnos distintos por campus (todos los periodos)"
    grafica.y_axis.title = "Alumnos"
    grafica.x_axis.title = "Campus"
    col_total = len(encabezado)
    datos = Reference(ws, min_col=col_total, min_row=1, max_row=1 + len(campus_lista))
    categorias = Reference(ws, min_col=1, min_row=2, max_row=1 + len(campus_lista))
    grafica.add_data(datos, titles_from_data=True)
    grafica.set_categories(categorias)
    grafica.legend = None
    grafica.height, grafica.width = 9, 18
    ws.add_chart(grafica, f"A{nota + 3}")

    destino = os.path.join(salida, "alumnos-por-campus.xlsx")
    wb.save(destino)

    # ---- Resumen en pantalla
    print(f"Leí {len(archivos)} archivos y {renglones:,} renglones.")
    ancho = max(len(c) for c in campus_lista + ["Todos los campus"])
    print("Campus".ljust(ancho), *[p.split()[0][:3] + p.split()[-1][2:] for p in periodos], "Total", sep="\t")
    for c in campus_lista:
        print(c.ljust(ancho), *[len(por_celda.get((c, p), ())) for p in periodos], len(por_campus[c]), sep="\t")
    print("Todos los campus".ljust(ancho), *totales_periodo, len(todos), sep="\t")
    print("Guardé", os.path.abspath(destino))


if __name__ == "__main__":
    main()
