#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_csv_ed_digital.py: genera archivos CSV FICTICIOS que imitan registros
academicos de una universidad (cursos-profesores e inscripciones por periodo).

Como correrlo (solo biblioteca estandar, Python 3):
    python generar_csv_ed_digital.py --escala 0.05 --salida csv-chico
    python generar_csv_ed_digital.py --salida csv-ed-digital

Parametros: --semilla (42), --escala (1.0), --salida (csv-ed-digital).
Todos los datos son inventados; ningun nombre corresponde a una persona real.
Al final imprime una tabla por periodo y la guarda en NOTAS-generacion.md.
"""
import argparse
import csv
import random
import time
import unicodedata
from pathlib import Path

PERIODOS = [
    "Agosto-Diciembre 2022", "Febrero-Junio 2023",
    "Agosto-Diciembre 2023", "Febrero-Junio 2024",
    "Agosto-Diciembre 2024", "Febrero-Junio 2025",
    "Agosto-Diciembre 2025", "Febrero-Junio 2026",
]
PERIODO_ASISTENCIA_TEXTO = 3  # Febrero-Junio 2024


def codigo_periodo(p):
    return ("AD" if p.startswith("Agosto") else "FJ") + p.split()[-1]


CAMPUS = ["Monterrey", "Guadalajara", "Ciudad de México", "Querétaro",
          "Puebla", "Estado de México"]
CAMPUS_PESOS = [0.34, 0.20, 0.20, 0.10, 0.09, 0.07]

# Constantes de volumen (a escala 1.0)
N_PROFESORES = 1800
N_ALUMNOS = 60000
TAM_GRUPO = 20.5          # alumnos promedio por grupo
P_COTITULAR = 0.08
P_DUPLICADO = 0.015
P_CAMPUS_VARIANTE = 0.02
P_CORREO_MALO = 0.005
DOMINIO_ALUMNO = "@alumnos.ejemplo.edu.mx"

NIVELES = ["I", "II", "III", "IV", "V", "VI"]

ESCUELAS = [
    dict(nombre="Negocios", letra="N", peso=0.24, n_mat=96,
         programas=["LAE", "LAF", "LCDE", "LMC", "LEM"],
         deptos={
             "Finanzas": ["Finanzas Corporativas", "Mercados de Capitales", "Análisis de Inversiones", "Contabilidad Financiera", "Administración de Riesgos"],
             "Mercadotecnia": ["Mercadotecnia Digital", "Comportamiento del Consumidor", "Investigación de Mercados", "Marca y Comunicación", "Ventas y Negociación"],
             "Administración": ["Administración de Operaciones", "Gestión de Proyectos", "Liderazgo Organizacional", "Emprendimiento", "Cadena de Suministro"],
             "Economía": ["Microeconomía", "Macroeconomía", "Econometría", "Economía Internacional", "Desarrollo Económico"]}),
    dict(nombre="Ingeniería", letra="I", peso=0.30, n_mat=120,
         programas=["IMT", "ISC", "ITC", "IME", "IC", "IQ", "IBT", "IID"],
         deptos={
             "Mecatrónica": ["Circuitos Eléctricos", "Control Automático", "Robótica", "Sistemas Embebidos", "Mecánica de Materiales"],
             "Sistemas Computacionales": ["Programación", "Estructuras de Datos", "Bases de Datos", "Redes de Computadoras", "Ingeniería de Software"],
             "Ingeniería Civil": ["Estructuras", "Hidráulica", "Geotecnia", "Topografía", "Materiales de Construcción"],
             "Ingeniería Química": ["Termodinámica", "Química Orgánica", "Procesos de Separación", "Reactores Químicos", "Química Analítica"],
             "Ingeniería Industrial": ["Calidad", "Logística", "Ergonomía", "Investigación de Operaciones", "Manufactura"]}),
    dict(nombre="Humanidades", letra="H", peso=0.14, n_mat=56,
         programas=["LRI", "LPS", "LLE", "LCC", "LEC"],
         deptos={
             "Ciencias Sociales": ["Sociología", "Antropología", "Ciencia Política", "Relaciones Internacionales", "Metodología de la Investigación"],
             "Letras y Comunicación": ["Literatura Hispanoamericana", "Redacción Académica", "Periodismo", "Producción Audiovisual", "Lingüística"],
             "Psicología": ["Psicología del Desarrollo", "Psicología Social", "Psicometría", "Psicología Clínica", "Neurociencias"]}),
    dict(nombre="Medicina", letra="M", peso=0.10, n_mat=40,
         programas=["MC", "LNB", "LBC", "LFT", "LEN"],
         deptos={
             "Ciencias Básicas": ["Anatomía", "Fisiología", "Bioquímica", "Histología"],
             "Clínica": ["Semiología", "Patología", "Farmacología", "Cirugía"],
             "Salud Pública": ["Epidemiología", "Bioestadística", "Nutrición", "Salud Comunitaria"]}),
    dict(nombre="Arquitectura", letra="A", peso=0.10, n_mat=40,
         programas=["LAR", "LUB", "LDI", "LPD", "LAD"],
         deptos={
             "Diseño": ["Taller de Diseño Arquitectónico", "Dibujo y Representación", "Teoría de la Arquitectura", "Diseño Sustentable"],
             "Urbanismo": ["Diseño Urbano", "Planeación Territorial", "Movilidad", "Paisaje"],
             "Construcción": ["Sistemas Constructivos", "Instalaciones", "Estructuras para Arquitectura", "Costos y Presupuestos"]}),
    dict(nombre="Derecho", letra="D", peso=0.12, n_mat=48,
         programas=["LD", "LDE", "LCP"],
         deptos={
             "Derecho Privado": ["Derecho Civil", "Derecho Mercantil", "Contratos", "Derecho Familiar", "Propiedad Intelectual"],
             "Derecho Público": ["Derecho Constitucional", "Derecho Administrativo", "Derecho Penal", "Derecho Fiscal", "Amparo"],
             "Derecho Internacional": ["Derecho Internacional Público", "Comercio Exterior", "Derechos Humanos", "Arbitraje", "Derecho Migratorio"]}),
]
ESC_PESOS = [e["peso"] for e in ESCUELAS]

NOMBRES = [
    "José", "Juan", "Luis", "Carlos", "Miguel", "Pedro", "Ricardo", "Fernando",
    "Roberto", "Alejandro", "Daniel", "Eduardo", "Javier", "Manuel", "Francisco",
    "Sergio", "Raúl", "Héctor", "Óscar", "Andrés", "Diego", "Gerardo", "Mario",
    "Arturo", "Ángel", "Iván", "Rodrigo", "Emilio", "Gustavo", "Rafael",
    "María", "Guadalupe", "Ana", "Laura", "Patricia", "Claudia", "Gabriela",
    "Verónica", "Sofía", "Fernanda", "Alejandra", "Daniela", "Mariana", "Paola",
    "Karla", "Lucía", "Elena", "Rosa", "Teresa", "Silvia", "Adriana", "Leticia",
    "Mónica", "Beatriz", "Carmen", "Irene", "Julieta", "Valeria", "Renata", "Andrea",
]
APELLIDOS = [
    "Hernández", "García", "Martínez", "López", "González", "Rodríguez", "Pérez",
    "Sánchez", "Ramírez", "Flores", "Gómez", "Díaz", "Reyes", "Morales", "Cruz",
    "Ortiz", "Gutiérrez", "Chávez", "Ruiz", "Jiménez", "Mendoza", "Álvarez",
    "Castillo", "Vázquez", "Romero", "Torres", "Herrera", "Medina", "Aguilar",
    "Vargas", "Castro", "Ramos", "Moreno", "Guerrero", "Salazar", "Domínguez",
    "Rojas", "Silva", "Navarro", "Cortés", "Soto", "Contreras", "Luna", "Estrada",
    "Delgado", "Ibarra", "Espinoza", "Santiago", "Vega", "Campos", "Cabrera",
    "Fuentes", "Rivera", "Mejía", "Carrillo", "Sandoval", "Núñez", "Zamora",
    "Trejo", "Montes", "Peña", "Ávila", "Cervantes", "Ríos", "Lara", "Valdez",
    "Maldonado", "Padilla", "Acosta", "Bravo", "Orozco", "Villarreal", "Leal",
    "Garza", "Treviño", "Elizondo", "Benítez", "Cantú", "Solís", "Macías", "Pacheco",
]
OBSERVACIONES = [
    "Cambio de grupo", "Justificante médico", "Beca deportiva", "Beca académica",
    "Intercambio", "Solicitud de revisión", "Baja temporal", "Examen extemporáneo",
]
HORARIOS = [("LuMiVi", 60, 4), ("MaJu", 90, 4), ("LuMi", 90, 2), ("MiVi", 90, 2),
            ("LuJu", 90, 1), ("Vi", 180, 1), ("Sa", 180, 1)]
HOR_PESOS = [h[2] for h in HORARIOS]


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))


def hhmm(m):
    return "%02d:%02d" % (m // 60, m % 60)


def variante_campus(rng, nombre):
    return nombre.upper() if rng.random() < 0.5 else nombre + " "


# ------------------------------------------------------------------ catalogo
def construir_materias(rng):
    """Devuelve lista de materias (clave, nombre, escuela, depto, creditos)
    y los indices por escuela."""
    materias, por_escuela, usadas = [], [], set()
    for ei, esc in enumerate(ESCUELAS):
        combos = [(d, t, n) for d, ts in esc["deptos"].items()
                  for t in ts for n in range(1, 7)]
        elegidas = rng.sample(combos, esc["n_mat"])
        idx = []
        for depto, tema, nivel in elegidas:
            while True:
                clave = "%s%d%03d" % (esc["letra"], nivel, rng.randrange(1000))
                if clave not in usadas:
                    usadas.add(clave)
                    break
            cred = rng.choices([3, 6, 8], [0.5, 0.35, 0.15])[0]
            idx.append(len(materias))
            materias.append((clave, "%s %s" % (tema, NIVELES[nivel - 1]),
                             ei, depto, cred))
        por_escuela.append(idx)
    return materias, por_escuela


def construir_profesores(rng, escala):
    n = max(24, round(N_PROFESORES * escala))
    if n > len(NOMBRES) * len(APELLIDOS):
        raise ValueError("--escala demasiado alta para la lista de nombres")
    combos = rng.sample(range(len(NOMBRES) * len(APELLIDOS)), n)
    nominas = rng.sample(range(10_000_000, 100_000_000), n + 3)
    pesos_esc = ESC_PESOS
    profs = []
    for i in range(n):
        nom = NOMBRES[combos[i] // len(APELLIDOS)]
        ap1 = APELLIDOS[combos[i] % len(APELLIDOS)]
        ap2 = rng.choice(APELLIDOS)
        esc = rng.choices(range(len(ESCUELAS)), pesos_esc)[0]
        r = rng.random()
        if r < 0.80:
            ini, fin = 0, 7
        elif r < 0.867:
            ini, fin = rng.randint(1, 6), 7
        elif r < 0.933:
            ini, fin = 0, rng.randint(1, 6)
        else:
            ini, fin = rng.randint(1, 3), rng.randint(4, 6)
        profs.append(dict(
            nomina=["L%08d" % nominas[i]],
            cambio=None,
            nombre="%s %s %s" % (nom, ap1, ap2),
            correo="%s.%s@ejemplo.edu.mx" % (sin_acentos(nom).lower(),
                                              sin_acentos(ap1).lower()),
            tipo=rng.choices(["Planta", "Cátedra", "Adjunto"], [0.40, 0.35, 0.25])[0],
            esc=esc, ini=ini, fin=fin))
    # Problema 5: tres profesores cambian de nómina entre periodos
    elegidos = rng.sample(range(n), 3)
    periodos_cambio = sorted(rng.sample(range(1, 8), 3))
    for k, p in enumerate(elegidos):
        profs[p]["ini"], profs[p]["fin"] = 0, 7
        profs[p]["cambio"] = (periodos_cambio[k], "L%08d" % nominas[n + k])
    return profs


def nomina_de(p, t):
    if p["cambio"] and t >= p["cambio"][0]:
        return p["cambio"][1]
    return p["nomina"][0]


def construir_alumnos(rng, escala):
    n = max(300, round(N_ALUMNOS * escala))
    mats = rng.sample(range(1_000_000, 10_000_000), n)
    spans = [2, 3, 4, 5, 6, 7, 8]
    pesos_span = [0.52, 0.30, 0.11, 0.04, 0.015, 0.01, 0.005]
    alumnos = []
    for i in range(n):
        span = rng.choices(spans, pesos_span)[0]
        esc = rng.choices(range(len(ESCUELAS)), ESC_PESOS)[0]
        nom = rng.choice(NOMBRES)
        mat = "A0%07d" % mats[i]
        alumnos.append(dict(
            mat=mat,
            correo=mat.lower() + DOMINIO_ALUMNO,
            nombre="%s %s %s" % (nom, rng.choice(APELLIDOS), rng.choice(APELLIDOS)),
            prog=rng.choice(ESCUELAS[esc]["programas"]),
            esc=esc,
            campus=rng.choices(range(6), CAMPUS_PESOS)[0],
            ini=rng.choices(range(9 - span), [1 + 2.5 * (j == 0) + 2.5 * (j == 8 - span) for j in range(9 - span)])[0],
            span=span,
            sem0=rng.randint(1, 5),
            base=min(97.0, max(66.0, rng.gauss(83, 8)))))
    return alumnos


# ------------------------------------------------------------------ grupos
def asignar_profesores(rng, activos, n_tit, n_co, grupos_idx, grupos_co):
    """Reparte titulares y co-titulares entre profesores activos de una escuela.
    Devuelve (titulares alineados con grupos_idx, co alineados con grupos_co)."""
    N = n_tit + n_co
    if N == 0:
        return [], []
    forzados = [i for i in activos if profs_global[i]["cambio"]]
    if N <= len(activos):
        elegidos = forzados[:N]
        sel = set(elegidos)
        resto = [i for i in activos if i not in sel]
        elegidos += rng.sample(resto, N - len(elegidos))
        carga = {i: 1 for i in elegidos}
    else:
        carga = {i: 1 for i in activos}
        extra = N - len(activos)
        cand = list(activos)
        while extra > 0 and cand:
            j = rng.randrange(len(cand))
            p = cand[j]
            if carga[p] < 5:
                carga[p] += 1
                extra -= 1
            else:
                cand[j] = cand[-1]
                cand.pop()
        while extra > 0:  # caso extremo: sin tope
            carga[rng.choice(activos)] += 1
            extra -= 1
    slots = [p for p, c in carga.items() for _ in range(c)]
    rng.shuffle(slots)
    tit, co = slots[:n_tit], slots[n_tit:]
    pos = {g: i for i, g in enumerate(grupos_idx)}
    copos = {g: i for i, g in enumerate(grupos_co)}
    for i in range(len(co)):
        tg = pos[grupos_co[i]]
        if co[i] != tit[tg]:
            continue
        v = co[i]
        for _ in range(200):
            k = rng.randrange(len(tit))
            u = tit[k]
            if u == v:
                continue
            # intercambio co[i] <-> tit[k]: tit[k] pasa a v, co[i] pasa a u
            gk = grupos_idx[k]
            if gk in copos and co[copos[gk]] == v:
                continue
            tit[k], co[i] = v, u
            break
        else:
            co[i] = rng.choice([x for x in activos if x != v])
    return tit, co


def construir_grupos(rng, t, G, materias, mat_por_esc):
    g_campus, g_esc, g_mat, g_num, g_co = [], [], [], [], []
    contador = {}
    for g in range(G):
        c = g % 6 if g < 72 else rng.choices(range(6), CAMPUS_PESOS)[0]
        for _ in range(50):
            e = rng.choices(range(len(ESCUELAS)), ESC_PESOS)[0]
            m = rng.choice(mat_por_esc[e])
            if contador.get((c, m), 0) < 12:
                break
        contador[(c, m)] = contador.get((c, m), 0) + 1
        g_campus.append(c)
        g_esc.append(materias[m][2])
        g_mat.append(m)
        g_num.append(contador[(c, m)])
        g_co.append(rng.random() < P_COTITULAR)
    crns = rng.sample(range(10000, 100000), G)
    mods = [rng.choices(["Presencial", "Híbrida", "Remota"], [0.60, 0.25, 0.15])[0]
            for _ in range(G)]
    horarios, salones = [], []
    for g in range(G):
        dias, dur, _ = rng.choices(HORARIOS, HOR_PESOS)[0]
        ini = rng.randrange(7 * 60, 21 * 60 - dur + 1, 30)
        horarios.append("%s %s-%s" % (dias, hhmm(ini), hhmm(ini + dur)))
        if mods[g] == "Remota":
            salones.append("")
        else:
            salones.append("%s%d-%d%02d" % (rng.choice("ABCDEFG"), rng.randint(1, 4),
                                             rng.randint(1, 4), rng.randint(1, 30)))
    # profesores por escuela
    g_tit, g_cop = [-1] * G, [-1] * G
    for e in range(len(ESCUELAS)):
        idx = [g for g in range(G) if g_esc[g] == e]
        idx_co = [g for g in idx if g_co[g]]
        activos = [i for i, p in enumerate(profs_global)
                   if p["esc"] == e and p["ini"] <= t <= p["fin"]]
        if not idx:
            continue
        if not activos:
            activos = [i for i, p in enumerate(profs_global) if p["esc"] == e] or [0]
        tit, co = asignar_profesores(rng, activos, len(idx), len(idx_co), idx, idx_co)
        for g, p in zip(idx, tit):
            g_tit[g] = p
        for g, p in zip(idx_co, co):
            g_cop[g] = p
    return dict(campus=g_campus, esc=g_esc, mat=g_mat, num=g_num, crn=crns,
                mod=mods, hor=horarios, salon=salones, tit=g_tit, co=g_cop)


# ------------------------------------------------------------------ principal
profs_global = []

COLS_CURSOS = ["Periodo", "Campus", "Escuela", "Departamento", "Clave materia",
               "Nombre materia", "Grupo", "CRN", "Modalidad", "Horario", "Salón",
               "Nómina profesor", "Nombre profesor", "Correo profesor",
               "Tipo de profesor", "Rol en grupo", "Alumnos inscritos", "Créditos"]
COLS_INSC = ["Periodo", "Campus", "Clave materia", "Nombre materia", "Grupo", "CRN",
             "Modalidad", "Nómina profesor", "Nombre profesor", "Matrícula alumno",
             "Nombre alumno", "Correo alumno", "Programa", "Semestre alumno",
             "Estatus inscripción", "Tipo de inscripción", "Créditos",
             "Promedio previo", "Calificación parcial 1", "Asistencia %",
             "Observaciones"]


def main():
    global profs_global
    ap = argparse.ArgumentParser(description="Genera CSV ficticios (Ed Digital, S5).")
    ap.add_argument("--semilla", type=int, default=42)
    ap.add_argument("--escala", type=float, default=1.0)
    ap.add_argument("--salida", default="csv-ed-digital")
    args = ap.parse_args()

    t0 = time.time()
    rng = random.Random(args.semilla)
    rnd = rng.random
    out = Path(args.salida)
    out.mkdir(parents=True, exist_ok=True)

    materias, mat_por_esc = construir_materias(rng)
    profs_global = construir_profesores(rng, args.escala)
    alumnos = construir_alumnos(rng, args.escala)
    n_per = len(PERIODOS)
    activos_t = [[i for i, a in enumerate(alumnos) if a["ini"] <= t < a["ini"] + a["span"]]
                 for t in range(n_per)]

    filas = []
    for t, periodo in enumerate(PERIODOS):
        cod = codigo_periodo(periodo)
        act = activos_t[t]
        G = max(300, round(5.5 * len(act) / TAM_GRUPO))
        gr = construir_grupos(rng, t, G, materias, mat_por_esc)
        pool_all = [[] for _ in range(6)]
        pool_esc = {}
        for g in range(G):
            pool_all[gr["campus"][g]].append(g)
            pool_esc.setdefault((gr["campus"][g], gr["esc"][g]), []).append(g)
        conteo = [0] * G
        st = dict(rows=0, dup=0, cotit=0, camp=0, mail=0, asist=0, reales=0, alumnos=0)

        # Fragmentos fijos por grupo para armar filas rapido
        frag = []
        for g in range(G):
            m = materias[gr["mat"][g]]
            pt = profs_global[gr["tit"][g]]
            pc = profs_global[gr["co"][g]] if gr["co"][g] >= 0 else None
            frag.append((CAMPUS[gr["campus"][g]], m[0], m[1], gr["num"][g], gr["crn"][g],
                         gr["mod"][g], nomina_de(pt, t), pt["nombre"],
                         (nomina_de(pc, t), pc["nombre"]) if pc else None, m[4]))

        f_ins = open(out / ("inscripciones-%s.csv" % cod), "w", newline="", encoding="utf-8-sig")
        w = csv.writer(f_ins)
        w.writerow(COLS_INSC)
        wr = w.writerow
        texto_asist = (t == PERIODO_ASISTENCIA_TEXTO)
        for a in act:
            al = alumnos[a]
            k = rng.randint(4, 7)
            propios = pool_esc.get((al["campus"], al["esc"]))
            todos = pool_all[al["campus"]]
            elegidos, mats_el = [], set()
            intentos = 0
            while len(elegidos) < k and intentos < 400:
                intentos += 1
                pool = propios if (propios and rnd() < 0.8) else todos
                g = pool[int(rnd() * len(pool))]
                if g in elegidos or gr["mat"][g] in mats_el:
                    continue
                elegidos.append(g)
                mats_el.add(gr["mat"][g])
            st["alumnos"] += 1
            sem = al["sem0"] + (t - al["ini"])
            prom = "%.1f" % min(100.0, max(60.0, al["base"] + rng.gauss(0, 2)))
            for g in elegidos:
                (camp, clave, nmat, num, crn, mod, nomT, nombT, co, cred) = frag[g]
                conteo[g] += 1
                st["reales"] += 1
                cv = 1 if rnd() < P_CAMPUS_VARIANTE else 0
                mal = 1 if rnd() < P_CORREO_MALO else 0
                correo = al["correo"]
                if mal:
                    mm = al["mat"].lower()
                    correo = rng.choice([mm + "@alumnos..ejemplo.edu.mx",
                                         mm + "@alumnos.ejemplo",
                                         mm + "@alumnos"])
                r = rnd()
                est = "Inscrito" if r < 0.92 else ("Baja" if r < 0.98 else "Oyente")
                tipo = "Ordinaria" if rnd() < 0.90 else "Recursamiento"
                cal = "" if rnd() < 0.05 else int(min(100, max(0, rng.gauss(al["base"] - 3, 12))))
                asis = int(min(100, max(40, rng.gauss(88, 10))))
                if texto_asist:
                    asis = "%d%%" % asis
                obs = rng.choice(OBSERVACIONES) if rnd() < 0.03 else ""
                row = [periodo, variante_campus(rng, camp) if cv else camp, clave, nmat,
                       num, crn, mod, nomT, nombT, al["mat"], al["nombre"], correo,
                       al["prog"], sem, est, tipo, cred, prom, cal, asis, obs]
                filas_emit = [(row, False)]
                if co:
                    row2 = row[:]
                    row2[7], row2[8] = co
                    filas_emit.append((row2, True))
                for rw, es_co in filas_emit:
                    wr(rw)
                    n = 1
                    if es_co:
                        st["cotit"] += 1
                    if rnd() < P_DUPLICADO:
                        wr(rw)
                        n = 2
                        st["dup"] += 1
                    st["rows"] += n
                    st["camp"] += cv * n
                    st["mail"] += mal * n
                    if texto_asist:
                        st["asist"] += n
        f_ins.close()

        # cursos-profesores
        f_cur = open(out / ("cursos-profesores-%s.csv" % cod), "w", newline="", encoding="utf-8-sig")
        wc = csv.writer(f_cur)
        wc.writerow(COLS_CURSOS)
        n_cur, camp_cur = 0, 0
        usadas, profs_t = set(), set()
        cambios_t = 0
        for g in range(G):
            m = materias[gr["mat"][g]]
            esc = ESCUELAS[m[2]]["nombre"]
            camp = CAMPUS[gr["campus"][g]]
            lista = [(gr["tit"][g], "Titular")]
            if gr["co"][g] >= 0:
                lista.append((gr["co"][g], "Co-titular"))
            for pi, rol in lista:
                p = profs_global[pi]
                nom = nomina_de(p, t)
                cvc = rnd() < P_CAMPUS_VARIANTE
                camp_cur += 1 if cvc else 0
                wc.writerow([periodo, variante_campus(rng, camp) if cvc else camp, esc,
                             m[3], m[0], m[1], gr["num"][g], gr["crn"][g], gr["mod"][g],
                             gr["hor"][g], gr["salon"][g], nom, p["nombre"], p["correo"],
                             p["tipo"], rol, conteo[g], m[4]])
                n_cur += 1
                profs_t.add(nom)
                if p["cambio"] and p["cambio"][0] == t and nom == p["cambio"][1] and pi not in usadas:
                    usadas.add(pi)
                    cambios_t += 1
        f_cur.close()

        assert sum(conteo) == st["reales"]
        filas.append(dict(periodo=periodo, cursos=n_cur, insc=st["rows"],
                          reales=st["reales"], profs=len(profs_t), alumnos=st["alumnos"],
                          dup=st["dup"], cotit=st["cotit"], camp_i=st["camp"],
                          camp_c=camp_cur, asist=st["asist"], nomina=cambios_t,
                          mail=st["mail"]))
        print("  %s listo (%d inscripciones)" % (periodo, st["rows"]), flush=True)

    # tabla
    cab = ["Periodo", "Renglones cursos", "Renglones inscripciones",
           "Inscripciones reales", "Profesores distintos", "Alumnos distintos",
           "P1 duplicados exactos", "P2 renglones por co-titular",
           "P3 campus con variante (inscripciones)", "P3 campus con variante (cursos)",
           "P4 asistencia como texto", "P5 profesores con nómina nueva",
           "P6 correos de alumno mal formados"]
    claves = ["periodo", "cursos", "insc", "reales", "profs", "alumnos", "dup",
              "cotit", "camp_i", "camp_c", "asist", "nomina", "mail"]
    tot = dict(periodo="Total")
    for c in claves[1:]:
        if c == "alumnos":
            tot[c] = len({i for lst in activos_t for i in lst})
        elif c == "profs":
            tot[c] = ""
        else:
            tot[c] = sum(f[c] for f in filas)
    lineas = ["| " + " | ".join(cab) + " |", "|" + "---|" * len(cab)]
    for f in filas + [tot]:
        lineas.append("| " + " | ".join(str(f[c]) for c in claves) + " |")
    tabla = "\n".join(lineas)
    seg = time.time() - t0
    texto = ("# Notas de generación\n\n"
             "**Todos los datos de estos archivos son ficticios.** Ningún nombre, "
             "matrícula, nómina ni correo corresponde a una persona real.\n\n"
             "- Semilla: %d\n- Escala: %s\n- Total de renglones de inscripciones: %d\n\n"
             "## Definiciones\n\n"
             "- Inscripciones reales: pares (alumno, grupo) distintos, sin duplicados "
             "ni repetición por co-titular. `Alumnos inscritos` en cursos-profesores "
             "cuenta estas inscripciones reales de cada grupo (todos los estatus).\n"
             "- P1: renglones de inscripciones repetidos de forma exacta.\n"
             "- P2: renglones extra de inscripciones por grupos con co-titular.\n"
             "- P3: renglones con `Campus` en variante (MAYÚSCULAS o espacio al final).\n"
             "- P4: en Febrero-Junio 2024, `Asistencia %%` viene como texto con signo.\n"
             "- P5: tres profesores cambian de nómina conservando nombre y correo; "
             "la columna cuenta cuántos aparecen con nómina nueva en ese periodo.\n"
             "- P6: correos de alumno con doble punto o dominio incompleto.\n\n"
             "## Tabla por periodo\n\n%s\n") % (args.semilla, args.escala,
                                               tot["insc"], tabla)
    (out / "NOTAS-generacion.md").write_text(texto, encoding="utf-8")
    print()
    print(tabla)
    print("\nTotal inscripciones: %d | tiempo: %.1f s | salida: %s"
          % (tot["insc"], seg, out.resolve()))


if __name__ == "__main__":
    main()
