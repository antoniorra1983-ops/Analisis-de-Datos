#!/usr/bin/env python3
"""
planilla_maniobras.py
=====================

Convierte la salida de un simulador de operación ferroviaria (CSV) en una
planilla horaria con maniobras, en formato de **varias terminales** lado a lado
(estilo "Planilla + Maniobras").

Cada viaje (tripID) se resume en una sola fila y se ubica en la columna de su
**estación de origen**. Siempre se muestran las dos terminales de cabecera
(Puerto y Limache); las terminales intermedias (El Belloto, Sargento Aldea)
aparecen solo si hay servicios que parten desde ellas.

Columnas de cada terminal:

    Viaje | Tren | Partida | N° | Inter. | Man. | Destino | M | Obs.

Maniobras (Man.), derivadas de la cadena cronológica de cada tren:

    * EV  (Entrada a Vía)  -> primer viaje del día de ese tren
    * RET (Retorno)        -> cada viaje siguiente (da vuelta en la terminal)
    * SV  (Sale de Vía)    -> fila al terminar la jornada (estaciona)

Uso básico:

    python planilla_maniobras.py Planilla_Simulador.csv -o salida.xlsx

Ver todas las opciones:

    python planilla_maniobras.py --help
"""

from __future__ import annotations

import argparse
import os
from dataclasses import dataclass, field
from datetime import datetime, time

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# --------------------------------------------------------------------------- #
# Configuración
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    """Parámetros que controlan la conversión."""

    sep: str = ";"

    # Terminales de cabecera (siempre presentes). El Puerto es el extremo "sur":
    # los trenes que parten de Puerto van hacia Limache; todos los demás orígenes
    # van hacia Puerto.
    cod_puerto: str = "PUE"
    nombre_puerto: str = "Terminal Puerto"
    cod_limache: str = "LIM"
    nombre_limache: str = "Terminal Limache"

    # Terminales intermedias (puntos de retorno). Solo aparecen si hay servicios
    # que parten (o terminan) en ellas. Orden geográfico Puerto -> Limache.
    intermedios: list[tuple[str, str]] = field(
        default_factory=lambda: [("BTO", "El Belloto"), ("SGA", "Sargento Aldea")]
    )

    multiple_threshold: int = 0         # capacidad >= esto -> "Múltiple" (0 = automático)
    renumerar_viajes: bool = True       # N° de viaje: pares vía 1, impares vía 2
    filas_de_paso: bool = False         # listar trenes que solo pasan por una intermedia
    round_minutes: bool = False         # redondear horas al minuto
    maniobras: bool = True              # derivar EV / RET / SV
    train_prefix: str = ""              # prefijo para renumerar trenes
    titulo: str = "Planilla Horaria + Maniobras — Simulador"
    fuente: str = "Arial"
    fecha_thdr: str = ""                # etiqueta de fecha en las hojas THDR (ej. "010526")

    # Nombres de columnas esperados en el CSV
    col_trip: str = "tripID"
    col_train: str = "trainID"
    col_cap: str = "trainTotalCapacity"
    col_track: str = "trackID"
    col_station: str = "stationName"
    col_arrive: str = "arriveTime"
    col_leave: str = "leaveTime"

    # Columnas de pasajeros (para las hojas de Carga de Pasajeros V1/V2).
    # Si no están en el CSV, esas hojas simplemente no se generan.
    col_arrive_pax: str = "arrivePassengers"
    col_arrive_est_pax: str = "arriveStationPassengers"
    col_leave_pax: str = "leavePassengers"
    col_leave_est_pax: str = "leaveStationPassengers"

    linea_carga: str = "Línea: Valparaíso-Limache"
    contrato_carga: str = "Todos los Contratos"


# Encabezados de cada terminal, en orden (9 columnas; la capacidad se refleja
# en la columna "M" como "Múltiple").
COLUMNAS = ["Viaje", "Tren", "Partida", "N°", "Inter.",
            "Man.", "Destino", "M", "Obs."]


# --------------------------------------------------------------------------- #
# Hojas THDR (una por vía) — formato de los archivos THDR_viaN
# --------------------------------------------------------------------------- #
# Código de estación del CSV -> nombre completo usado en el THDR.
NOMBRES_ESTACIONES = {
    "PUE": "Puerto", "BEL": "Bellavista", "FRA": "Francia", "BAR": "Baron",
    "POR": "Portales", "REC": "Recreo", "MIR": "Miramar", "VIN": "Vina del mar",
    "HOS": "Hospital", "CHO": "Chorrillos", "SLT": "El Salto", "VAL": "Valencia",
    "QUI": "Quilpue", "SOL": "El Sol", "BTO": "El Belloto", "AME": "Americas",
    "CON": "La Concepcion", "VAM": "Villa Alemana", "SGA": "Sargento Aldea",
    "PEN": "Penablanca", "LIM": "Limache",
}

# Campos de cabecera del THDR (columnas A..I).
THDR_CAMPOS = ["Viaje", "Tren", "Hora Salida Programada", "Motriz 1", "Motriz 2",
               "Unidad", "Maquinista", "Controlador", "Observación"]

THDR_COL_INICIO = 12   # primera columna de estaciones (A..I + 2 de separación)
THDR_FILA_DATOS = 6    # primera fila de datos


# --------------------------------------------------------------------------- #
# Utilidades de tiempo
# --------------------------------------------------------------------------- #
def hms_a_segundos(value):
    """Segundos desde medianoche, o None si la celda está vacía o no es una hora.

    Acepta 'H:MM:SS' y 'H:MM'. Tolera celdas vacías, NaN y texto no numérico,
    que aparecen cuando el CSV trae filas incompletas o líneas en blanco.
    """
    if value is None:
        return None
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "nat", "none", "-"):
        return None
    partes = s.split(":")
    if not 1 <= len(partes) <= 3:
        return None
    try:
        nums = [int(float(p)) for p in partes]
    except ValueError:
        return None
    while len(nums) < 3:
        nums.append(0)
    return nums[0] * 3600 + nums[1] * 60 + nums[2]


def hms_to_seconds(value) -> int:
    """Igual que `hms_a_segundos`, pero falla si el valor no es una hora válida."""
    seg = hms_a_segundos(value)
    if seg is None:
        raise ValueError(f"Hora no válida: {value!r}")
    return seg


def numerar_viajes(registros) -> dict[int, int]:
    """Numeración de viajes según la convención de la planilla.

    Vía 1 (Puerto -> Limache, track 0): pares 2, 4, 6... por hora de SALIDA.
    Vía 2 (hacia Puerto): impares 1, 3, 5... por hora de LLEGADA a Puerto.
    (Verificado contra el archivo laboral: 135/135 y 131/131.)

    `registros`: lista de tuplas (trip, track, salida_s, llegada_s).
    Devuelve {trip original: número de viaje}.
    """
    registros = list(registros)
    mapa: dict[int, int] = {}
    via1 = sorted((x for x in registros if x[1] == 0), key=lambda x: (x[2], x[0]))
    via2 = sorted((x for x in registros if x[1] != 0), key=lambda x: (x[3], x[0]))
    for i, x in enumerate(via1):
        mapa[x[0]] = 2 + 2 * i
    for i, x in enumerate(via2):
        mapa[x[0]] = 1 + 2 * i
    return mapa


def umbral_multiple(capacidades, cfg: Config) -> int:
    """Capacidad a partir de la cual un tren se considera «Múltiple» (doble).

    Si `cfg.multiple_threshold` es mayor que 0 se respeta ese valor. Si es 0
    (automático) se deduce del propio archivo: el doble de la capacidad simple,
    que es la menor del archivo. Así, con capacidades 450/900 el simple es 450 y
    el doble 900; con 200/400, el simple es 200 y el doble 400.
    """
    if cfg.multiple_threshold and cfg.multiple_threshold > 0:
        return int(cfg.multiple_threshold)
    validas = [int(c) for c in capacidades if _entero(c) > 0]
    return 2 * min(validas) if validas else 10 ** 9


def _entero(valor, defecto: int = 0) -> int:
    """Convierte a int tolerando vacíos, NaN y texto. Devuelve `defecto` si no se puede."""
    try:
        f = float(valor)
    except (TypeError, ValueError):
        return defecto
    return defecto if f != f else int(f)      # f != f  ->  NaN


def _texto(valor) -> str:
    """Texto limpio de una celda; cadena vacía si viene vacía o NaN."""
    if valor is None:
        return ""
    s = str(valor).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s


def seconds_to_time(total: float, round_minutes: bool = False) -> time:
    total = int(round(total))
    if round_minutes:
        total = (total + 30) // 60 * 60
    h, m, s = total // 3600, (total % 3600) // 60, total % 60
    return time(h % 24, m, 0 if round_minutes else s)


def _hhmmss(total: float) -> str:
    """Hora como texto 'HH:MM:SS' a partir de segundos."""
    t = int(round(total))
    return f"{t // 3600:02d}:{(t % 3600) // 60:02d}:{t % 60:02d}"


# --------------------------------------------------------------------------- #
# 1) Resumir el CSV a una fila por viaje
# --------------------------------------------------------------------------- #
def _paradas_validas(g, cfg: Config) -> list[dict]:
    """Paradas utilizables de un viaje: descarta filas sin estación o sin ninguna hora.

    Si falta una de las dos horas, se usa la otra (una parada sin llegada es el
    origen; sin salida, el destino final).
    """
    paradas = []
    for _, r in g.iterrows():
        est = _texto(r[cfg.col_station])
        lleg = hms_a_segundos(r[cfg.col_arrive])
        sal = hms_a_segundos(r[cfg.col_leave])
        if not est or (lleg is None and sal is None):
            continue
        paradas.append({
            "est": est,
            "lleg": lleg if lleg is not None else sal,
            "sal": sal if sal is not None else lleg,
            "fila": r,
        })
    # El origen, el destino y las horas salen del recorrido ORDENADO por hora,
    # no del orden en que vengan las filas en el archivo.
    paradas.sort(key=lambda p: (p["lleg"], p["sal"]))
    return paradas


def cargar_viajes(csv_path, cfg: Config) -> pd.DataFrame:
    """Lee el CSV (ruta o buffer) y devuelve un DataFrame con una fila por tripID.

    Las filas incompletas (sin estación o sin horas) se descartan; el total
    descartado queda en `viajes.attrs["descartadas"]`.
    """
    df = pd.read_csv(csv_path, sep=cfg.sep)

    faltan = [c for c in (cfg.col_trip, cfg.col_train, cfg.col_cap, cfg.col_track,
                          cfg.col_station, cfg.col_arrive, cfg.col_leave)
              if c not in df.columns]
    if faltan:
        raise ValueError(
            f"El CSV no tiene las columnas esperadas: {faltan}\n"
            f"Columnas encontradas: {list(df.columns)}"
        )

    viajes, descartadas = [], 0
    for trip_id, g in df.groupby(cfg.col_trip, sort=True):
        g = g.reset_index(drop=True)
        paradas = _paradas_validas(g, cfg)
        descartadas += len(g) - len(paradas)
        if not paradas:
            continue
        primera, ultima = paradas[0], paradas[-1]
        viajes.append({
            "trip": _entero(trip_id),
            "train": _entero(primera["fila"][cfg.col_train]),
            "cap": _entero(primera["fila"][cfg.col_cap]),
            "track": _entero(primera["fila"][cfg.col_track]),
            "orig": primera["est"],
            "dep": _hhmmss(primera["sal"]),
            "dep_s": primera["sal"],
            "dest": ultima["est"],
            "arr_s": ultima["lleg"],
        })

    if not viajes:
        raise ValueError("El CSV no tiene viajes con estación y hora válidas.")

    if cfg.renumerar_viajes:
        mapa = numerar_viajes([(v["trip"], v["track"], v["dep_s"], v["arr_s"]) for v in viajes])
        for v in viajes:
            v["trip"] = mapa.get(v["trip"], v["trip"])

    out = pd.DataFrame(viajes).sort_values("dep_s").reset_index(drop=True)
    out.attrs["descartadas"] = descartadas
    return out


# --------------------------------------------------------------------------- #
# 2) Maniobras (EV / RET) y fines de servicio (SV)
# --------------------------------------------------------------------------- #
def asignar_maniobras(viajes: pd.DataFrame) -> pd.DataFrame:
    """EV para el primer viaje de cada tren (orden cronológico), RET para el resto."""
    primer = {tren: viajes[viajes["train"] == tren].sort_values("dep_s").index[0]
              for tren in viajes["train"].unique()}
    viajes = viajes.copy()
    viajes["man"] = ["EV" if idx == primer[row.train] else "RET"
                     for idx, row in viajes.iterrows()]
    return viajes


def fines_de_servicio(viajes: pd.DataFrame) -> dict[str, list[dict]]:
    """Devuelve, por código de estación, la lista de trenes que terminan su jornada allí."""
    fines: dict[str, list[dict]] = {}
    for tren in viajes["train"].unique():
        ultimo = viajes[viajes["train"] == tren].sort_values("dep_s").iloc[-1]
        fines.setdefault(ultimo["dest"], []).append(
            {"train": tren, "cap": int(ultimo["cap"]), "end_s": ultimo["arr_s"]}
        )
    return fines


# --------------------------------------------------------------------------- #
# 3) Definir y construir las columnas (terminales)
# --------------------------------------------------------------------------- #
def columnas_visibles(viajes: pd.DataFrame, cfg: Config,
                      orden_estaciones=None) -> list[dict]:
    """Lista ordenada de terminales a mostrar: Puerto, intermedias con actividad, Limache.

    Las intermedias NO están fijas: es cualquier estación donde empiece o termine
    algún servicio y que no sea uno de los dos extremos. El orden geográfico sale
    del recorrido del propio archivo (o del catálogo de estaciones si no se pasa).
    """
    presentes = set(viajes["orig"]) | set(viajes["dest"])
    extremos = (cfg.cod_puerto, cfg.cod_limache)
    orden = list(orden_estaciones or NOMBRES_ESTACIONES)
    nombres = dict(cfg.intermedios)          # nombres preferidos, si se configuraron

    intermedias = [e for e in orden if e in presentes and e not in extremos]
    intermedias += [e for e in sorted(presentes)          # por si alguna no está en el orden
                    if e not in extremos and e not in intermedias]

    cols = [dict(code=cfg.cod_puerto, nombre=cfg.nombre_puerto, implied=cfg.cod_limache)]
    for code in intermedias:
        cols.append(dict(code=code,
                         nombre=nombres.get(code) or NOMBRES_ESTACIONES.get(code, code),
                         implied=cfg.cod_puerto))
    cols.append(dict(code=cfg.cod_limache, nombre=cfg.nombre_limache, implied=cfg.cod_puerto))
    return cols


def _fmt_tren(num: int, cfg: Config):
    return f"{cfg.train_prefix}{num}" if cfg.train_prefix else num


def _columna_de(orig: str, track: int, cols: list[dict], cfg: Config) -> int:
    """Índice de la columna que corresponde a un viaje según su origen (con respaldo por sentido)."""
    for i, c in enumerate(cols):
        if c["code"] == orig:
            return i
    return 0 if track == 0 else len(cols) - 1  # respaldo: track 0 -> Puerto, 1 -> Limache


def _pasos_por_terminal(col: dict, paradas_viajes, cfg: Config) -> list[tuple]:
    """Trenes que PASAN por una terminal intermedia camino a Puerto, sin salir de ella.

    En la planilla aparecen como filas sin N° de viaje, solo con tren, hora e
    intervalo; sirven para ver la frecuencia de paso por ese punto.
    """
    if (not cfg.filas_de_paso or not paradas_viajes
            or col["code"] in (cfg.cod_puerto, cfg.cod_limache)):
        return []
    pasos = []
    for v in paradas_viajes:
        if v["track"] == 0:                             # solo sentido hacia Puerto
            continue
        if v["paradas"][0]["est"] == col["code"]:       # sale de aquí: ya es un servicio
            continue
        for p in v["paradas"][:-1]:                     # el destino final no es "paso"
            if p["est"] == col["code"]:
                pasos.append((p["sal"], v["train"]))
                break
    return pasos


def construir_columna(viajes: pd.DataFrame, fines: dict, col: dict, cfg: Config,
                      paradas_viajes=None) -> list[list]:
    """Filas (listas de 9 celdas) de una terminal: salidas desde col['code'],
    trenes que pasan (si es intermedia) y SV que terminan allí."""
    sub = viajes[viajes["_col_code"] == col["code"]].sort_values("dep_s").reset_index(drop=True)
    umbral = umbral_multiple(viajes["cap"], cfg)

    eventos = [(x.dep_s, True, x) for x in sub.itertuples()]
    eventos += [(h, False, tren) for h, tren in _pasos_por_terminal(col, paradas_viajes, cfg)]
    eventos.sort(key=lambda e: (e[0], not e[1]))

    filas: list[list] = []
    prev_s = None
    n = 0
    for hora_s, es_servicio, dato in eventos:
        inter = seconds_to_time(hora_s - prev_s, cfg.round_minutes) if prev_s is not None else None
        prev_s = hora_s
        if es_servicio:
            n += 1
            x = dato
            destino = "" if x.dest == col["implied"] else x.dest
            multiple = "Múltiple" if x.cap >= umbral else ""
            maniobra = x.man if cfg.maniobras else ""
            filas.append([x.trip, _fmt_tren(x.train, cfg),
                          seconds_to_time(hora_s, cfg.round_minutes),
                          n, inter, maniobra, destino, multiple, ""])
        else:                                            # fila de paso: sin N° de viaje
            filas.append(["", _fmt_tren(dato, cfg),
                          seconds_to_time(hora_s, cfg.round_minutes),
                          "", inter, "", "", "", ""])

    if cfg.maniobras:
        for r in sorted(fines.get(col["code"], []), key=lambda d: d["end_s"]):
            multiple = "Múltiple" if r["cap"] >= umbral else ""
            filas.append(["", _fmt_tren(r["train"], cfg), None, "", None,
                          "SV", "", multiple, "Estaciona"])
    return filas


def construir_tablas(viajes: pd.DataFrame, cfg: Config,
                     paradas_viajes=None) -> tuple[list[dict], list[list[list]]]:
    """Devuelve (columnas, tablas) en paralelo: una lista de filas por terminal."""
    viajes = asignar_maniobras(viajes) if cfg.maniobras else viajes.assign(man="")
    fines = fines_de_servicio(viajes) if cfg.maniobras else {}
    orden = _orden_estaciones(paradas_viajes, 0) if paradas_viajes else None
    cols = columnas_visibles(viajes, cfg, orden)

    # asignar cada viaje a su columna por origen
    viajes = viajes.copy()
    viajes["_col_code"] = [cols[_columna_de(r.orig, r.track, cols, cfg)]["code"]
                           for r in viajes.itertuples()]

    tablas = [construir_columna(viajes, fines, c, cfg, paradas_viajes) for c in cols]
    return cols, tablas


# --------------------------------------------------------------------------- #
# 4) Escribir el Excel con formato
# --------------------------------------------------------------------------- #
def cargar_paradas(csv_path, cfg: Config) -> list[dict]:
    """Lee el CSV y devuelve, por viaje, el detalle de cada parada.

    Cada elemento: {trip, train, cap, track, paradas: [{est, lleg, sal}, ...]}
    Es el detalle que necesitan las hojas THDR (llegada/salida por estación).
    """
    df = pd.read_csv(csv_path, sep=cfg.sep)
    faltan = [c for c in (cfg.col_trip, cfg.col_train, cfg.col_cap, cfg.col_track,
                          cfg.col_station, cfg.col_arrive, cfg.col_leave)
              if c not in df.columns]
    if faltan:
        raise ValueError(f"El CSV no tiene las columnas esperadas: {faltan}")

    viajes = []
    tiene_pax = all(c in df.columns for c in
                    (cfg.col_arrive_pax, cfg.col_arrive_est_pax,
                     cfg.col_leave_pax, cfg.col_leave_est_pax))
    for trip_id, g in df.groupby(cfg.col_trip, sort=True):
        g = g.reset_index(drop=True)
        validas = _paradas_validas(g, cfg)
        if not validas:
            continue
        paradas = []
        for p in validas:
            r = p["fila"]
            parada = {"est": p["est"], "lleg": p["lleg"], "sal": p["sal"]}
            if tiene_pax:
                # carga = pasajeros a bordo al salir; suben = los que abordan aquí
                parada["carga"] = _entero(r[cfg.col_leave_pax])
                parada["suben"] = (_entero(r[cfg.col_arrive_est_pax])
                                   - _entero(r[cfg.col_leave_est_pax]))
            paradas.append(parada)
        viajes.append({
            "trip": _entero(trip_id),
            "train": _entero(validas[0]["fila"][cfg.col_train]),
            "cap": _entero(validas[0]["fila"][cfg.col_cap]),
            "track": _entero(validas[0]["fila"][cfg.col_track]),
            "pax": tiene_pax,
            "paradas": paradas,
        })

    if cfg.renumerar_viajes:
        mapa = numerar_viajes([(v["trip"], v["track"], v["paradas"][0]["sal"],
                                v["paradas"][-1]["lleg"]) for v in viajes])
        for v in viajes:
            v["trip"] = mapa.get(v["trip"], v["trip"])
    return viajes


def _orden_estaciones(paradas_viajes: list[dict], track: int) -> list[str]:
    """Orden de estaciones de una vía: el recorrido más largo de esa vía."""
    recorridos = [[p["est"] for p in v["paradas"]]
                  for v in paradas_viajes if v["track"] == track]
    if not recorridos:
        return []
    base = max(recorridos, key=len)
    # Completar con estaciones que aparezcan en otros recorridos de la misma vía
    for rec in recorridos:
        for i, est in enumerate(rec):
            if est not in base:
                anterior = rec[i - 1] if i else None
                pos = base.index(anterior) + 1 if anterior in base else len(base)
                base.insert(pos, est)
    return base


def agregar_hojas_thdr(wb: Workbook, paradas_viajes: list[dict], cfg: Config) -> list[str]:
    """Añade al libro una hoja THDR por vía (llegada/salida por estación).

    Vía 1 = trenes que van de Puerto a Limache (track 0);
    Vía 2 = sentido contrario (track 1).
    """
    NAVY = "1F3864"
    thin = Side(style="thin", color="B0B0B0")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    f_cell = Font(name=cfg.fuente, size=9)
    f_hdr = Font(name=cfg.fuente, bold=True, color="FFFFFF", size=9)
    f_est = Font(name=cfg.fuente, bold=True, size=10, color=NAVY)
    center = Alignment(horizontal="center", vertical="center")
    wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fmt_time = "h:mm" if cfg.round_minutes else "h:mm:ss"
    umbral_thdr = umbral_multiple([v["cap"] for v in paradas_viajes], cfg)

    creadas = []
    for via, track in ((1, 0), (2, 1)):
        orden = _orden_estaciones(paradas_viajes, track)
        if not orden:
            continue
        ws = wb.create_sheet(f"THDR Vía {via}")

        # Fila 1: fecha (si se indicó) y nombres de estación
        if cfg.fecha_thdr:
            c = ws.cell(1, 1, cfg.fecha_thdr)
            c.font = Font(name=cfg.fuente, bold=True, size=10)
        col_de_est = {}
        for i, est in enumerate(orden):
            base = THDR_COL_INICIO + i * 2
            col_de_est[est] = base
            c = ws.cell(1, base, NOMBRES_ESTACIONES.get(est, est))
            c.font, c.alignment = f_est, center

        # Fila 2: encabezados
        for j, campo in enumerate(THDR_CAMPOS):
            c = ws.cell(2, 1 + j, campo)
            c.font, c.alignment, c.border = f_hdr, wrap, borde
            c.fill = PatternFill("solid", fgColor=NAVY)
        ultimo = orden[-1]
        for est in orden:
            base = col_de_est[est]
            etiquetas = ["Hora Llegada"] if est == ultimo else ["Hora Llegada", "Hora Salida"]
            for k, etq in enumerate(etiquetas):
                c = ws.cell(2, base + k, etq)
                c.font, c.alignment, c.border = f_hdr, wrap, borde
                c.fill = PatternFill("solid", fgColor=NAVY)

        # Datos (ordenados por hora de salida del origen)
        # Vía 1 se ordena por salida; vía 2 por llegada a Puerto (igual que la numeración)
        orden_t = ((lambda v: v["paradas"][0]["sal"]) if track == 0
                   else (lambda v: v["paradas"][-1]["lleg"]))
        viajes_via = sorted((v for v in paradas_viajes if v["track"] == track), key=orden_t)
        for i, v in enumerate(viajes_via):
            r = THDR_FILA_DATOS + i
            ws.cell(r, 1, v["trip"]).alignment = center
            ws.cell(r, 2, v["train"]).alignment = center
            prog = ws.cell(r, 3, seconds_to_time(v["paradas"][0]["sal"], True))
            prog.number_format, prog.alignment = "h:mm", center
            if v["cap"] >= umbral_thdr:
                ws.cell(r, 6, "M").alignment = center     # Unidad: M = múltiple
            for j in range(1, len(THDR_CAMPOS) + 1):
                cc = ws.cell(r, j)
                cc.font, cc.border = f_cell, borde

            primera, ultima_p = v["paradas"][0], v["paradas"][-1]
            for p in v["paradas"]:
                base = col_de_est.get(p["est"])
                if base is None:
                    continue
                # El origen no tiene llegada; el destino final no tiene salida.
                if p is not primera:
                    c = ws.cell(r, base, seconds_to_time(p["lleg"], cfg.round_minutes))
                    c.number_format, c.alignment, c.font, c.border = fmt_time, center, f_cell, borde
                if p is not ultima_p and p["est"] != ultimo:
                    c = ws.cell(r, base + 1, seconds_to_time(p["sal"], cfg.round_minutes))
                    c.number_format, c.alignment, c.font, c.border = fmt_time, center, f_cell, borde

        # Anchos y vista
        for j, w in enumerate([6, 6, 11, 8, 8, 7, 26, 14, 14]):
            ws.column_dimensions[get_column_letter(1 + j)].width = w
        ws.column_dimensions[get_column_letter(10)].width = 2
        ws.column_dimensions[get_column_letter(11)].width = 2
        for est in orden:
            base = col_de_est[est]
            for k in range(2 if est != ultimo else 1):
                ws.column_dimensions[get_column_letter(base + k)].width = 10
        ws.freeze_panes = ws.cell(THDR_FILA_DATOS, 3)
        ws.sheet_view.showGridLines = False
        creadas.append(ws.title)
    return creadas


# --------------------------------------------------------------------------- #
# Hojas V1 / V2 — Carga de Pasajeros (formato EXPORT Carga Pasajeros)
# --------------------------------------------------------------------------- #
CARGA_CAMPOS = ["N° THDR", "N° Viaje", "Tren", "Fecha", "Hora Origen", "Hora Fin",
                "Motriz 1", "Motriz 2"]
CARGA_FINALES = ["Total a Bordo", "Carga Máxima", "Estación Máxima"]
CARGA_TITULO = "CARGA DE PASAJEROS POR SERVICIO para Filial EFE-VALPO"
CARGA_FILA_ENC = 10     # fila de encabezados
CARGA_FILA_DATOS = 11   # primera fila de datos


def agregar_hojas_carga(wb: Workbook, paradas_viajes: list[dict], cfg: Config) -> list[str]:
    """Añade las hojas V1 y V2 con la carga de pasajeros por servicio.

    Replica el formato de los archivos *EXPORT Carga Pasajeros*: una fila por
    viaje, con los pasajeros a bordo al salir de cada estación, más el total
    transportado, la carga máxima y en qué estación se alcanza.
    """
    if not any(v.get("pax") for v in paradas_viajes):
        return []                      # el CSV no trae columnas de pasajeros

    NAVY = "1F3864"
    thin = Side(style="thin", color="B0B0B0")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    f_cell = Font(name=cfg.fuente, size=9)
    f_hdr = Font(name=cfg.fuente, bold=True, color="FFFFFF", size=9)
    f_bold = Font(name=cfg.fuente, bold=True, size=9)
    center = Alignment(horizontal="center", vertical="center")
    wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    creado_el = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    creadas = []
    for via, track in ((1, 0), (2, 1)):
        orden = _orden_estaciones(paradas_viajes, track)
        if not orden:
            continue
        ws = wb.create_sheet(f"V{via}")

        # Cabecera del reporte
        ws.cell(1, 1, CARGA_TITULO).font = Font(name=cfg.fuente, bold=True, size=11, color=NAVY)
        meta = [(3, "Fecha creación archivo ", creado_el),
                (4, "Desde:", cfg.fecha_thdr), (5, "Hasta:", cfg.fecha_thdr),
                (6, "Línea:", cfg.linea_carga), (7, "Vía:", f"Vía {via}"),
                (8, "Contrato:", cfg.contrato_carga)]
        for fila, etq, val in meta:
            ws.cell(fila, 1, etq).font = f_bold
            ws.cell(fila, 2, val).font = f_cell

        # Encabezados: campos + estaciones (códigos) + totales
        encabezados = CARGA_CAMPOS + orden + CARGA_FINALES
        for j, h in enumerate(encabezados):
            c = ws.cell(CARGA_FILA_ENC, 1 + j, h)
            c.font, c.alignment, c.border = f_hdr, wrap, borde
            c.fill = PatternFill("solid", fgColor=NAVY)

        col_de_est = {est: len(CARGA_CAMPOS) + 1 + i for i, est in enumerate(orden)}
        col_total = len(CARGA_CAMPOS) + len(orden) + 1

        # Vía 1 se ordena por salida; vía 2 por llegada a Puerto (igual que la numeración)
        orden_t = ((lambda v: v["paradas"][0]["sal"]) if track == 0
                   else (lambda v: v["paradas"][-1]["lleg"]))
        viajes_via = sorted((v for v in paradas_viajes if v["track"] == track), key=orden_t)
        for i, v in enumerate(viajes_via):
            r = CARGA_FILA_DATOS + i
            primera, ultima = v["paradas"][0], v["paradas"][-1]
            ws.cell(r, 1, v["trip"])                                  # N° THDR
            ws.cell(r, 2, "")                                         # N° Viaje (no viene en el CSV)
            ws.cell(r, 3, v["train"])
            ws.cell(r, 4, cfg.fecha_thdr)
            ws.cell(r, 5, _hhmmss(primera["sal"]))                    # Hora Origen (texto)
            ws.cell(r, 6, _hhmmss(ultima["lleg"]))                    # Hora Fin (texto)
            ws.cell(r, 7, "")                                         # Motriz 1
            ws.cell(r, 8, "")                                         # Motriz 2

            cargas = {}
            for p in v["paradas"]:
                col = col_de_est.get(p["est"])
                if col is not None:
                    ws.cell(r, col, p.get("carga", 0))
                    cargas[p["est"]] = p.get("carga", 0)
            maxima = max(cargas.values()) if cargas else 0
            ws.cell(r, col_total, sum(p.get("suben", 0) for p in v["paradas"]))
            ws.cell(r, col_total + 1, maxima)
            ws.cell(r, col_total + 2,
                    " - ".join(e for e in orden if cargas.get(e) == maxima) if maxima else "")

            for j in range(1, col_total + 3):
                cc = ws.cell(r, j)
                cc.font, cc.border = f_cell, borde
                cc.alignment = center if j != col_total + 2 else Alignment(
                    horizontal="left", vertical="center")

        # Anchos y vista
        for j, w in enumerate([9, 10, 7, 11, 11, 10, 9, 9]):
            ws.column_dimensions[get_column_letter(1 + j)].width = w
        for est in orden:
            ws.column_dimensions[get_column_letter(col_de_est[est])].width = 6
        ws.column_dimensions[get_column_letter(col_total)].width = 12
        ws.column_dimensions[get_column_letter(col_total + 1)].width = 12
        ws.column_dimensions[get_column_letter(col_total + 2)].width = 24
        ws.freeze_panes = ws.cell(CARGA_FILA_DATOS, 4)
        ws.sheet_view.showGridLines = False
        creadas.append(ws.title)
    return creadas


def construir_workbook(cols: list[dict], tablas: list[list[list]],
                       cfg: Config, hora_inicio: str, origen_archivo: str,
                       paradas_viajes: list[dict] | None = None) -> Workbook:
    """Arma el libro Excel (en memoria) con N terminales lado a lado.

    Si se pasa `paradas_viajes` (de `cargar_paradas`), añade además una hoja
    THDR por vía, quedando 3 hojas en total.
    """
    NAVY, GRIS, VERDE = "1F3864", "F2F2F2", "E2EFDA"
    NCOL = len(COLUMNAS)            # 9 columnas por bloque
    PASO = NCOL + 1                 # + 1 columna separadora
    total_cols = PASO * len(cols) - 1

    thin = Side(style="thin", color="B0B0B0")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    f_cell = Font(name=cfg.fuente, size=9)
    f_hdr = Font(name=cfg.fuente, bold=True, color="FFFFFF", size=9)
    f_term = Font(name=cfg.fuente, bold=True, color="FFFFFF", size=11)
    center = Alignment(horizontal="center", vertical="center")
    left_al = Alignment(horizontal="left", vertical="center")
    fmt_time = "h:mm" if cfg.round_minutes else "h:mm:ss"

    wb = Workbook()
    ws = wb.active
    ws.title = "Planilla + Maniobras"

    # Título
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    c = ws.cell(1, 1, cfg.titulo)
    c.font = Font(name=cfg.fuente, bold=True, size=13, color=NAVY)
    c.alignment = left_al

    # Metadatos
    ws.cell(3, 1, "Inicio").font = Font(name=cfg.fuente, bold=True, size=9)
    hi = ws.cell(3, 2, seconds_to_time(hms_to_seconds(hora_inicio), True))
    hi.font, hi.number_format = f_cell, "h:mm"
    ws.cell(3, 6, "Origen:").font = Font(name=cfg.fuente, bold=True, size=9)
    ws.cell(3, 7, origen_archivo).font = f_cell
    ws.cell(4, 1, "Versión").font = Font(name=cfg.fuente, bold=True, size=9)
    ws.cell(4, 2, "Simulador").font = f_cell

    bases = [1 + i * PASO for i in range(len(cols))]

    # Rótulos de terminal
    for base, col in zip(bases, cols):
        ws.merge_cells(start_row=5, start_column=base, end_row=5, end_column=base + NCOL - 1)
        t = ws.cell(5, base, col["nombre"])
        t.font, t.alignment = f_term, center
        t.fill = PatternFill("solid", fgColor=NAVY)

    # Encabezados
    for base in bases:
        for j, h in enumerate(COLUMNAS):
            cc = ws.cell(6, base + j, h)
            cc.font, cc.alignment, cc.border = f_hdr, center, borde
            cc.fill = PatternFill("solid", fgColor=NAVY)

    # Datos
    def volcar(filas: list[list], base: int) -> None:
        for i, fila in enumerate(filas):
            r = 7 + i
            for j, val in enumerate(fila):
                cc = ws.cell(r, base + j, val if val != "" else None)
                cc.border, cc.font = borde, f_cell
                if j in (2, 4):                       # Partida, Inter. -> hora
                    cc.number_format = fmt_time
                    cc.alignment = center
                elif j in (0, 1, 3, 5, 6, 7):
                    cc.alignment = center
                else:
                    cc.alignment = left_al
            man = fila[5]
            relleno = VERDE if man == "EV" else GRIS if man == "SV" else None
            if relleno:
                for j in range(NCOL):
                    ws.cell(r, base + j).fill = PatternFill("solid", fgColor=relleno)
            ws.cell(r, base + 5).font = Font(
                name=cfg.fuente, size=9, bold=True,
                color="375623" if man == "EV" else "7F7F7F" if man == "SV" else "000000",
            )

    for base, filas in zip(bases, tablas):
        volcar(filas, base)

    # Anchos / vista
    anchos = [6, 6, 9, 5, 8, 6, 7, 9, 15]
    for base in bases:
        for j, w in enumerate(anchos):
            ws.column_dimensions[get_column_letter(base + j)].width = w
        if base + NCOL <= total_cols:
            ws.column_dimensions[get_column_letter(base + NCOL)].width = 2  # separador
    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False

    # Hojas THDR (una por vía) y hojas V1/V2 de carga de pasajeros
    if paradas_viajes:
        agregar_hojas_thdr(wb, paradas_viajes, cfg)
        agregar_hojas_carga(wb, paradas_viajes, cfg)
    return wb


def escribir_excel(cols, tablas, out_path: str, cfg: Config,
                   hora_inicio: str, origen_archivo: str,
                   paradas_viajes: list[dict] | None = None) -> None:
    wb = construir_workbook(cols, tablas, cfg, hora_inicio, origen_archivo, paradas_viajes)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)


# --------------------------------------------------------------------------- #
# Orquestador
# --------------------------------------------------------------------------- #
def convertir(csv_path: str, out_path: str, cfg: Config) -> dict:
    """Pipeline completo CSV -> XLSX en disco. Devuelve un resumen."""
    viajes = cargar_viajes(csv_path, cfg)
    paradas_viajes = cargar_paradas(csv_path, cfg)
    cols, tablas = construir_tablas(viajes, cfg, paradas_viajes)
    hora_inicio = viajes["dep"].iloc[0]
    escribir_excel(cols, tablas, out_path, cfg, hora_inicio,
                   os.path.basename(csv_path), paradas_viajes)

    return {
        "viajes_total": len(viajes),
        "trenes": int(viajes["train"].nunique()),
        "terminales": {c["nombre"]: len(t) for c, t in zip(cols, tablas)},
        "thdr": {f"Vía {v}": sum(1 for x in paradas_viajes if x["track"] == t)
                 for v, t in ((1, 0), (2, 1))},
        "salida": out_path,
    }


# =========================================================================== #
# CONVERSIÓN INVERSA
# Planilla + Maniobras (.xls)  ->  entrada del simulador (.xls, estilo Planillaprueba2)
#
# Formato de salida (9 columnas, una fila por servicio, ordenadas por hora):
#   hora | origen | unidades | destino | unidades | destino | "servicio" | tren | 406
# (xlrd y xlwt se importan de forma diferida para no exigirlos en el flujo CSV.)
# =========================================================================== #

# Respaldo para los códigos de zona de la columna Destino, por si no se pueden
# deducir del archivo (ver _pm_zonas_a_terminal, que los infiere de los datos).
ZONA_A_ESTACION = {6: "LIM", 4: "SGA", 5: "BTO"}
_ENCABEZADOS_PM = ["Viaje", "Tren", "Partida", "N°", "Inter.", "Man.", "Destino", "M", "Obs.", "Capacidad"]


def _pm_norm(s) -> str:
    return str(s).strip().lower()


def _pm_entero_pos(v) -> bool:
    return isinstance(v, (int, float)) and float(v) == int(v) and int(v) > 0


def _pm_val(sh, r, c):
    if c is None:
        return ""
    v = sh.cell_value(r, c)
    return int(v) if isinstance(v, float) and v == int(v) else v


def _pm_hora_seg(sh, wb, r, c):
    import xlrd
    if c is None:
        return None
    if sh.cell_type(r, c) == xlrd.XL_CELL_DATE:
        t = xlrd.xldate_as_tuple(sh.cell_value(r, c), wb.datemode)
        return t[3] * 3600 + t[4] * 60 + t[5]
    v = sh.cell_value(r, c)
    if isinstance(v, str) and ":" in v:
        p = [int(x) for x in v.split(":")]
        return p[0] * 3600 + p[1] * 60 + (p[2] if len(p) > 2 else 0)
    return None


def _codigo_terminal(etiqueta: str, usados) -> str:
    """Código de estación a partir del rótulo de la columna ('Terminal Puerto' -> PUE).

    Usa el catálogo de estaciones; si el rótulo no está, arma un código con sus
    iniciales, de modo que funcione con terminales que no conozcamos.
    """
    txt = _pm_norm(etiqueta).replace("terminal", "").strip(" .:-")
    for code, nombre in NOMBRES_ESTACIONES.items():
        if _pm_norm(nombre) == txt or _pm_norm(code) == txt:
            return code
    for code, nombre in NOMBRES_ESTACIONES.items():      # coincidencia parcial
        if txt and (_pm_norm(nombre) in txt or txt in _pm_norm(nombre)):
            return code
    palabras = [p for p in txt.split() if p not in ("de", "la", "el", "los", "las")]
    base = ("".join(p[0] for p in palabras) if len(palabras) > 1 else txt[:3]).upper()[:4]
    base = base or "TER"
    code, n = base, 2
    while code in usados:                                 # evitar duplicados
        code, n = f"{base}{n}", n + 1
    return code


def _pm_detectar(sh):
    """Localiza la fila de encabezados y los bloques (terminal -> columnas).

    Los bloques se detectan por el FORMATO —cada bloque empieza en una columna
    cuyo encabezado es 'Viaje'—, no por los nombres de las terminales, así que
    sirve para cualquier planilla con esta estructura, tenga las terminales que
    tenga y en la cantidad que sea.
    """
    fila_enc = None
    for r in range(min(15, sh.nrows)):
        textos = [_pm_norm(sh.cell_value(r, c)) for c in range(sh.ncols)]
        if any("viaje" in t for t in textos) and any("partida" in t for t in textos):
            fila_enc = r
            break
    if fila_enc is None:
        raise ValueError("No se encontró la fila de encabezados (con 'Viaje'/'Partida').")

    inicios = [c for c in range(sh.ncols) if _pm_norm(sh.cell_value(fila_enc, c)) == "viaje"]
    if not inicios:
        raise ValueError("No se encontró ninguna columna 'Viaje' en los encabezados.")

    # Rótulo de cada bloque: la fila no vacía más cercana por encima del encabezado
    def rotulo(c0, c1):
        for r in range(fila_enc - 1, max(-1, fila_enc - 4), -1):
            for c in range(c0, c1):
                txt = str(sh.cell_value(r, c)).strip()
                if txt:
                    return txt
        return ""

    bloques, usados = [], []
    for i, c0 in enumerate(inicios):
        c1 = inicios[i + 1] if i + 1 < len(inicios) else sh.ncols
        colmap = {}
        for c in range(c0, c1):
            h = _pm_norm(sh.cell_value(fila_enc, c))
            for nombre in _ENCABEZADOS_PM:
                if _pm_norm(nombre) == h or (nombre == "N°" and h in ("n°", "n")):
                    colmap.setdefault(nombre, c)
        code = _codigo_terminal(rotulo(c0, c1), usados)
        usados.append(code)
        bloques.append((code, colmap))
    return fila_enc, bloques


def _pm_via(origen: str, extremo_puerto: str) -> int:
    """Vía del viaje: 1 = sentido Puerto -> Limache, 2 = sentido hacia Puerto.

    El extremo "Puerto" es el primer bloque de la planilla, no un código fijo.
    Verificado contra las hojas V1/V2 del archivo laboral: V1 son las salidas
    desde Puerto (135) y V2 las que van hacia Puerto (131).
    """
    return 1 if origen == extremo_puerto else 2


def _pm_zonas_a_terminal(sh, fila_enc, bloques, zona_por_tren) -> dict:
    """Qué terminal representa cada código de la columna Destino (6, 4, ...).

    No usa una tabla fija: mira desde qué terminal vuelve cada tren que lleva
    ese código, así que se adapta a la planilla que sea.
    """
    from collections import Counter, defaultdict
    vuelve = defaultdict(Counter)
    extremo_puerto = bloques[0][0]
    for code, colmap in bloques:
        if code == extremo_puerto:
            continue
        for r in range(fila_enc + 1, sh.nrows):
            tren = _pm_a_entero(_pm_val(sh, r, colmap.get("Tren")))
            if tren is None or tren <= 0:
                continue
            zona = zona_por_tren.get(tren)
            if zona is not None:
                vuelve[zona][code] += 1
    return {z: c.most_common(1)[0][0] for z, c in vuelve.items() if c}


def _pm_destino(origen: str, zona, raw, zonas: dict,
                extremo_puerto: str, extremo_limache: str) -> str:
    """Estación destino del viaje.

    Prioridad: código de estación escrito en la celda -> terminal deducida del
    código de zona -> extremo opuesto. Las salidas desde cualquier terminal que
    no sea el extremo Puerto van hacia Puerto.
    """
    if isinstance(raw, str) and raw.strip().upper() in set(NOMBRES_ESTACIONES):
        return raw.strip().upper()
    if origen != extremo_puerto:
        return extremo_puerto
    if zona is None:
        return extremo_limache
    return zonas.get(zona) or ZONA_A_ESTACION.get(zona, extremo_limache)


def _pm_a_entero(v):
    """Convierte a int un valor que puede venir como número o como texto ('6')."""
    try:
        if isinstance(v, str):
            v = v.strip()
            if v == "":
                return None
        return int(float(v))
    except (TypeError, ValueError):
        return None


def _pm_fila_multiple(sh, r, colmap) -> bool:
    """¿Esta fila (este servicio) va en composición múltiple?

    La marca «Múltiple» se escribe en la fila del servicio que la usa — en estos
    archivos en la columna Obs. — y NO se aplica al resto de los viajes del tren:
    un mismo tren puede correr múltiple en unos servicios y simple en otros.
    """
    c_ini, c_fin = min(colmap.values()), max(colmap.values())
    return any("ltiple" in str(_pm_val(sh, r, c)).lower() for c in range(c_ini, c_fin + 1))


def _pm_mapas_por_tren(sh, fila_enc, bloques):
    """Recorre toda la hoja y arma dos mapas por tren:
       - zona_por_tren: valor (fijo por tren) de la columna Destino -> zona de
         destino (6 = Limache, 4 = Sargento Aldea). NO es la vía.
       - cap_por_tren: valor de la columna Capacidad, si el archivo la trae."""
    zona_por_tren: dict[int, int] = {}
    cap_por_tren: dict[int, int] = {}
    for _code, colmap in bloques:
        for r in range(fila_enc + 1, sh.nrows):
            tren = _pm_a_entero(_pm_val(sh, r, colmap.get("Tren")))
            if tren is None or tren <= 0:
                continue
            zona = _pm_a_entero(_pm_val(sh, r, colmap.get("Destino")))
            if tren not in zona_por_tren and zona is not None:
                zona_por_tren[tren] = zona
            cap = _pm_a_entero(_pm_val(sh, r, colmap.get("Capacidad")))
            if tren not in cap_por_tren and cap is not None and cap > 0:
                cap_por_tren[tren] = cap
    return zona_por_tren, cap_por_tren


def _pm_abrir(origen):
    """Abre un .xls desde ruta o bytes y devuelve el workbook (xlrd)."""
    import xlrd
    if isinstance(origen, (bytes, bytearray)):
        return xlrd.open_workbook(file_contents=bytes(origen))
    return xlrd.open_workbook(origen)


def _pm_contar_salidas(sh, wb, fila_enc, bloques) -> int:
    """Cuenta filas que son salidas reales (con N° de viaje y hora)."""
    n = 0
    for _code, colmap in bloques:
        for r in range(fila_enc + 1, sh.nrows):
            if _pm_entero_pos(_pm_val(sh, r, colmap.get("Viaje"))) and \
               _pm_hora_seg(sh, wb, r, colmap.get("Partida")) is not None:
                n += 1
    return n


def _elegir_hoja(wb, hoja_preferida=None) -> str:
    """Nombre de la hoja con formato Planilla + Maniobras (sin depender del nombre).

    Prioridad: nombre pedido si existe -> hoja con estructura válida y MÁS salidas
    (desempata el nombre que contenga 'maniobra') -> heurística de nombre -> 1ª hoja.
    """
    nombres = wb.sheet_names()
    if hoja_preferida and hoja_preferida in nombres:
        return hoja_preferida
    candidatas = []  # (0 si nombre tiene 'maniobra' si no 1, -salidas, nombre)
    for n in nombres:
        sh = wb.sheet_by_name(n)
        try:
            fila_enc, bloques = _pm_detectar(sh)
        except Exception:
            continue
        ndep = _pm_contar_salidas(sh, wb, fila_enc, bloques)
        if ndep > 0:
            candidatas.append((0 if "maniobra" in n.lower() else 1, -ndep, n))
    if candidatas:
        candidatas.sort()
        return candidatas[0][2]
    for clave in ("maniobra", "planilla"):
        for n in nombres:
            if clave in n.lower():
                return n
    return nombres[0]


def elegir_hoja_pm(origen, hoja_preferida=None) -> str:
    """Devuelve el nombre de la hoja Planilla + Maniobras de un archivo (ruta o bytes)."""
    return _elegir_hoja(_pm_abrir(origen), hoja_preferida)


def listar_hojas(origen) -> list[str]:
    """Lista los nombres de hoja de un archivo .xls (ruta o bytes)."""
    return _pm_abrir(origen).sheet_names()


def leer_planilla_maniobras(origen, hoja=None) -> list[dict]:
    """Extrae los servicios de una hoja Planilla + Maniobras. `origen` puede ser ruta o bytes.

    Si `hoja` es None (o no existe en el archivo), la hoja correcta se detecta
    automáticamente por su estructura, sin importar cómo se llame.

    La vía (columnas C/E del simulador) es el sentido del viaje: 1 = Puerto ->
    Limache, 2 = hacia Puerto. La columna Destino de la planilla (6/4) da la
    estación de destino, no la vía. Un tren es doble si aparece 'Múltiple' en
    alguna de sus filas.
    """
    wb = _pm_abrir(origen)
    sh = wb.sheet_by_name(_elegir_hoja(wb, hoja))
    fila_enc, bloques = _pm_detectar(sh)
    zona_por_tren, cap_por_tren = _pm_mapas_por_tren(sh, fila_enc, bloques)
    # Extremos de la línea: primer y último bloque de la planilla
    extremo_puerto, extremo_limache = bloques[0][0], bloques[-1][0]
    zonas = _pm_zonas_a_terminal(sh, fila_enc, bloques, zona_por_tren)

    salidas = []
    for code, colmap in bloques:
        for r in range(fila_enc + 1, sh.nrows):
            viaje = _pm_val(sh, r, colmap.get("Viaje"))
            hora = _pm_hora_seg(sh, wb, r, colmap.get("Partida"))
            tren = _pm_val(sh, r, colmap.get("Tren"))
            if not _pm_entero_pos(viaje) or hora is None or not _pm_entero_pos(tren):
                continue  # filas de paso (sin viaje) o SV (sin hora) se omiten
            tren = int(tren)
            destino = _pm_destino(code, zona_por_tren.get(tren),
                                  _pm_val(sh, r, colmap.get("Destino")),
                                  zonas, extremo_puerto, extremo_limache)
            # Capacidad: la de la fila; si no, la del tren; si no, queda None
            cap = _pm_a_entero(_pm_val(sh, r, colmap.get("Capacidad")))
            if cap is None or cap <= 0:
                cap = cap_por_tren.get(tren)
            salidas.append({
                "hora": hora, "origen": code, "destino": destino,
                "via": _pm_via(code, extremo_puerto), "capacidad": cap,
                "tren": tren, "unidades": 2 if _pm_fila_multiple(sh, r, colmap) else 1,
            })
    salidas.sort(key=lambda d: (d["hora"], d["origen"]))
    return salidas


def escribir_simulador_xls(salidas: list[dict], destino, constante: int = 450) -> None:
    """Escribe el formato plano del simulador como .xls. `destino` puede ser ruta o un buffer.

    La columna I lleva la cantidad de pasajeros que soporta el tren: se toma de la
    columna **Capacidad** de la Planilla + Maniobras y, si esa columna no existe o
    viene vacía, se usa `constante` por unidad: 450 si es simple y 900 si el tren
    aparece como «Múltiple».

    Formato de celdas:
      * Columna A (hora): TEXTO con la forma "HH:MM:SS".
      * Columnas numéricas (vía, tren, capacidad): formato número.
      * Resto (origen, destino, "servicio"): texto.
    """
    import xlwt
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Hoja1")
    estilo_texto = xlwt.easyxf(num_format_str="@")    # formato texto
    estilo_numero = xlwt.easyxf(num_format_str="0")   # formato número (entero)
    for i, s in enumerate(salidas):
        h = int(s["hora"])
        hora_txt = f"{h // 3600:02d}:{(h % 3600) // 60:02d}:{h % 60:02d}"
        ws.write(i, 0, hora_txt, estilo_texto)                       # A · hora (texto HH:MM:SS)
        ws.write(i, 1, s["origen"])                                  # B · texto
        ws.write(i, 2, s["via"], estilo_numero)                      # C · vía (número)
        ws.write(i, 3, s["destino"])                                 # D · texto
        ws.write(i, 4, s["via"], estilo_numero)                      # E · vía (número)
        ws.write(i, 5, s["destino"])                                 # F · texto
        ws.write(i, 6, "servicio")                                   # G · texto
        ws.write(i, 7, s["tren"], estilo_numero)                     # H · tren (número)
        cap = s.get("capacidad") or constante * s.get("unidades", 1)
        ws.write(i, 8, cap, estilo_numero)                           # I · capacidad (número)
    wb.save(destino)


def simulador_a_bytes(salidas: list[dict], constante: int = 450) -> bytes:
    """Devuelve el .xls del simulador como bytes (para descargar en la web)."""
    import io
    buf = io.BytesIO()
    escribir_simulador_xls(salidas, buf, constante)
    return buf.getvalue()


def convertir_a_simulador(entrada_xls: str, salida_xls: str,
                          hoja=None, constante: int = 450) -> dict:
    """Pipeline: Planilla + Maniobras (.xls) -> entrada del simulador (.xls).

    Si `hoja` es None, se detecta automáticamente la hoja con formato
    Planilla + Maniobras (sin importar su nombre).
    """
    from collections import Counter
    hoja_usada = elegir_hoja_pm(entrada_xls, hoja)
    salidas = leer_planilla_maniobras(entrada_xls, hoja_usada)
    escribir_simulador_xls(salidas, salida_xls, constante)
    return {
        "servicios": len(salidas),
        "por_origen": dict(Counter(s["origen"] for s in salidas)),
        "hoja": hoja_usada,
        "salida": salida_xls,
    }


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Convierte el CSV del simulador a una planilla horaria con "
                    "maniobras en formato de varias terminales (.xlsx).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("input", help="Ruta del CSV del simulador.")
    p.add_argument("-o", "--output", default="Planilla_Maniobras.xlsx",
                   help="Ruta del archivo .xlsx de salida.")
    p.add_argument("--sep", default=";", help="Separador de columnas del CSV.")

    g = p.add_argument_group("Terminales")
    g.add_argument("--cod-puerto", default="PUE", help="Código de la Terminal Puerto.")
    g.add_argument("--cod-limache", default="LIM", help="Código de la Terminal Limache.")
    g.add_argument("--nombre-puerto", default="Terminal Puerto")
    g.add_argument("--nombre-limache", default="Terminal Limache")

    o = p.add_argument_group("Opciones de formato")
    o.add_argument("--multiple-threshold", type=int, default=400,
                   help="Capacidad mínima para marcar el tren como 'Múltiple'.")
    o.add_argument("--round-minutes", action="store_true",
                   help="Redondear las horas al minuto (descarta los segundos).")
    o.add_argument("--no-maniobras", action="store_true",
                   help="No derivar EV/RET/SV (genera solo la planilla básica).")
    o.add_argument("--train-prefix", default="",
                   help="Prefijo para renumerar los trenes (ej. '60').")
    o.add_argument("--titulo", default="Planilla Horaria + Maniobras — Simulador")
    o.add_argument("--fuente", default="Arial")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    cfg = Config(
        sep=args.sep,
        cod_puerto=args.cod_puerto, cod_limache=args.cod_limache,
        nombre_puerto=args.nombre_puerto, nombre_limache=args.nombre_limache,
        multiple_threshold=args.multiple_threshold,
        round_minutes=args.round_minutes, maniobras=not args.no_maniobras,
        train_prefix=args.train_prefix, titulo=args.titulo, fuente=args.fuente,
    )
    resumen = convertir(args.input, args.output, cfg)
    print("Planilla generada correctamente.")
    print(f"  Viajes procesados : {resumen['viajes_total']}")
    print(f"  Trenes            : {resumen['trenes']}")
    for nombre, n in resumen["terminales"].items():
        print(f"  {nombre:<18}: {n} filas")
    print(f"  Archivo           : {resumen['salida']}")
    return 0


def _corriendo_en_streamlit() -> bool:
    try:
        from streamlit.runtime import exists
        return exists()
    except Exception:
        try:
            from streamlit.runtime.scriptrunner import get_script_run_ctx
            return get_script_run_ctx() is not None
        except Exception:
            return False


if __name__ == "__main__":
    if _corriendo_en_streamlit():
        import streamlit_app  # noqa: F401  -> renderiza la app si apuntan aquí por error
    else:
        raise SystemExit(main())
