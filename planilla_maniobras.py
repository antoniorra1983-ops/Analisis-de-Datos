#!/usr/bin/env python3
"""
planilla_maniobras.py
=====================

Convierte la salida de un simulador de operación ferroviaria (CSV) en una
planilla horaria con maniobras, en formato de dos terminales lado a lado
(estilo "Planilla + Maniobras").

El CSV de entrada tiene una fila por cada parada de cada viaje, con columnas:

    tripID;trainID;trainTotalCapacity;trackID;stationName;
    arriveTime;arrivePassengers;arriveStationPassengers;
    leaveTime;leavePassengers;leaveStationPassengers

Cada viaje (tripID) se resume en una sola fila de salida. Los viajes se reparten
en dos tablas según el sentido (trackID):

    * track 0  -> Terminal Puerto   (sentido Puerto -> Limache)
    * track 1  -> Terminal Limache  (sentido Limache -> Puerto)

Para cada terminal se generan las columnas:

    Viaje | Tren | Partida | N° | Inter. | Man. | Destino | M | Obs. | Capacidad

Las maniobras (Man.) se derivan rastreando la cadena cronológica de cada tren:

    * EV  (Entrada a Vía)  -> primer viaje del día de ese tren
    * RET (Retorno)        -> cada viaje siguiente (da vuelta en la terminal)
    * SV  (Sale de Vía)    -> fila adicional al terminar la jornada (estaciona)

Uso básico:

    python planilla_maniobras.py Planilla_Simulador.csv -o salida.xlsx

Ver todas las opciones:

    python planilla_maniobras.py --help

Requisitos: pandas, openpyxl  (ver requirements.txt)
"""

from __future__ import annotations

import argparse
import os
from dataclasses import dataclass, field
from datetime import time

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# --------------------------------------------------------------------------- #
# Configuración
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    """Parámetros que controlan la conversión (todos ajustables por CLI)."""

    sep: str = ";"                      # separador del CSV
    track_puerto: int = 0               # trackID del sentido "Terminal Puerto"
    track_limache: int = 1              # trackID del sentido "Terminal Limache"
    cod_puerto: str = "PUE"             # código de estación de la terminal Puerto
    cod_limache: str = "LIM"            # código de estación de la terminal Limache
    nombre_puerto: str = "Terminal Puerto"
    nombre_limache: str = "Terminal Limache"
    multiple_threshold: int = 400       # capacidad >= esto  ->  "Múltiple"
    round_minutes: bool = False         # redondear horas al minuto (descarta segundos)
    maniobras: bool = True              # derivar EV / RET / SV
    train_prefix: str = ""              # prefijo para renumerar trenes (ej. "60")
    titulo: str = "Planilla Horaria + Maniobras — Simulador"
    fuente: str = "Arial"

    # Nombres de columnas esperados en el CSV (cambiar solo si el simulador difiere)
    col_trip: str = "tripID"
    col_train: str = "trainID"
    col_cap: str = "trainTotalCapacity"
    col_track: str = "trackID"
    col_station: str = "stationName"
    col_arrive: str = "arriveTime"
    col_leave: str = "leaveTime"


# Encabezados de cada tabla, en orden.
COLUMNAS = ["Viaje", "Tren", "Partida", "N°", "Inter.",
            "Man.", "Destino", "M", "Obs.", "Capacidad"]


# --------------------------------------------------------------------------- #
# Utilidades de tiempo
# --------------------------------------------------------------------------- #
def hms_to_seconds(value: str) -> int:
    """'6:15:29' -> 22529 (segundos desde medianoche)."""
    h, m, s = (int(p) for p in str(value).split(":"))
    return h * 3600 + m * 60 + s


def seconds_to_time(total: float, round_minutes: bool = False) -> time:
    """Segundos -> datetime.time (envuelve a 24 h por seguridad)."""
    total = int(round(total))
    if round_minutes:
        total = (total + 30) // 60 * 60  # redondeo al minuto más cercano
    h, m, s = total // 3600, (total % 3600) // 60, total % 60
    return time(h % 24, m, s if not round_minutes else 0)


# --------------------------------------------------------------------------- #
# 1) Resumir el CSV a una fila por viaje
# --------------------------------------------------------------------------- #
def cargar_viajes(csv_path: str, cfg: Config) -> pd.DataFrame:
    """Lee el CSV y devuelve un DataFrame con una fila por tripID."""
    df = pd.read_csv(csv_path, sep=cfg.sep)

    faltan = [c for c in (cfg.col_trip, cfg.col_train, cfg.col_cap, cfg.col_track,
                          cfg.col_station, cfg.col_arrive, cfg.col_leave)
              if c not in df.columns]
    if faltan:
        raise ValueError(
            f"El CSV no tiene las columnas esperadas: {faltan}\n"
            f"Columnas encontradas: {list(df.columns)}\n"
            "Ajusta los nombres con las opciones --col-* o revisa el separador (--sep)."
        )

    viajes = []
    for trip_id, g in df.groupby(cfg.col_trip, sort=True):
        g = g.reset_index(drop=True)
        primera, ultima = g.iloc[0], g.iloc[-1]
        viajes.append({
            "trip": int(trip_id),
            "train": int(primera[cfg.col_train]),
            "cap": int(primera[cfg.col_cap]),
            "track": int(primera[cfg.col_track]),
            "orig": primera[cfg.col_station],
            "dep": primera[cfg.col_leave],
            "dep_s": hms_to_seconds(primera[cfg.col_leave]),
            "dest": ultima[cfg.col_station],
            "arr_s": hms_to_seconds(ultima[cfg.col_arrive]),
        })

    return pd.DataFrame(viajes).sort_values("dep_s").reset_index(drop=True)


# --------------------------------------------------------------------------- #
# 2) Derivar maniobras (EV / RET) y fines de servicio (SV)
# --------------------------------------------------------------------------- #
def asignar_maniobras(viajes: pd.DataFrame) -> pd.DataFrame:
    """Agrega la columna 'man': EV para el primer viaje de cada tren, RET para el resto."""
    primer_viaje = {
        tren: viajes[viajes["train"] == tren].sort_values("dep_s").index[0]
        for tren in viajes["train"].unique()
    }
    viajes = viajes.copy()
    viajes["man"] = [
        "EV" if idx == primer_viaje[row.train] else "RET"
        for idx, row in viajes.iterrows()
    ]
    return viajes


def fines_de_servicio(viajes: pd.DataFrame, cfg: Config) -> pd.DataFrame:
    """Para cada tren, calcula dónde termina su jornada (fila SV)."""
    def lado(estacion: str) -> str:
        # Puerto si termina en la estación Puerto; en cualquier otro caso, lado Limache.
        return "P" if estacion == cfg.cod_puerto else "L"

    filas = []
    for tren in viajes["train"].unique():
        ultimo = viajes[viajes["train"] == tren].sort_values("dep_s").iloc[-1]
        filas.append({
            "train": tren,
            "cap": int(ultimo["cap"]),
            "endst": ultimo["dest"],
            "end_s": ultimo["arr_s"],
            "side": lado(ultimo["dest"]),
        })
    return pd.DataFrame(filas)


# --------------------------------------------------------------------------- #
# 3) Construir cada tabla de terminal
# --------------------------------------------------------------------------- #
def _fmt_tren(num: int, cfg: Config) -> str | int:
    return f"{cfg.train_prefix}{num}" if cfg.train_prefix else num


def construir_tabla(viajes: pd.DataFrame, sv: pd.DataFrame, track: int,
                    cod_terminal: str, cod_opuesto: str, cfg: Config) -> list[list]:
    """Devuelve la lista de filas (listas de 10 celdas) para una terminal."""
    sub = viajes[viajes["track"] == track].sort_values("dep_s").reset_index(drop=True)
    filas: list[list] = []
    prev_s = None
    n = 0

    for x in sub.itertuples():
        n += 1
        inter = seconds_to_time(x.dep_s - prev_s, cfg.round_minutes) if prev_s is not None else None
        prev_s = x.dep_s
        destino = "" if x.dest == cod_opuesto else x.dest          # solo excepciones
        obs = "" if x.orig == cod_terminal else f"Sale {x.orig}"   # inicios fuera de terminal
        multiple = "Múltiple" if x.cap >= cfg.multiple_threshold else ""
        maniobra = x.man if cfg.maniobras else ""
        filas.append([
            x.trip, _fmt_tren(x.train, cfg),
            seconds_to_time(x.dep_s, cfg.round_minutes),
            n, inter, maniobra, destino, multiple, obs, x.cap,
        ])

    # Filas SV (sin hora de partida), al final, ordenadas por hora de término.
    if cfg.maniobras:
        side = "P" if cod_terminal == cfg.cod_puerto else "L"
        for r in sv[sv["side"] == side].sort_values("end_s").itertuples():
            multiple = "Múltiple" if r.cap >= cfg.multiple_threshold else ""
            filas.append(["", _fmt_tren(r.train, cfg), None, "", None,
                          "SV", "", multiple, f"Estaciona {r.endst}", r.cap])

    return filas


# --------------------------------------------------------------------------- #
# 4) Escribir el Excel con formato
# --------------------------------------------------------------------------- #
def construir_workbook(izq: list[list], der: list[list],
                       cfg: Config, hora_inicio: str, origen_archivo: str) -> Workbook:
    """Arma el libro Excel (en memoria) con las dos terminales y estilo profesional.

    Devuelve un openpyxl.Workbook; quien llame decide si guardarlo en disco
    (escribir_excel) o en un buffer en memoria (p. ej. para descargar en Streamlit).
    """
    NAVY, GRIS, VERDE = "1F3864", "F2F2F2", "E2EFDA"
    L0, R0 = 1, 12  # columna inicial izquierda / derecha (col 11 = separador)

    thin = Side(style="thin", color="B0B0B0")
    borde = Border(left=thin, right=thin, top=thin, bottom=thin)
    f_cell = Font(name=cfg.fuente, size=9)
    f_hdr = Font(name=cfg.fuente, bold=True, color="FFFFFF", size=9)
    f_term = Font(name=cfg.fuente, bold=True, color="FFFFFF", size=11)
    center = Alignment(horizontal="center", vertical="center")
    left_al = Alignment(horizontal="left", vertical="center")

    wb = Workbook()
    ws = wb.active
    ws.title = "Planilla + Maniobras"

    # --- Título ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=21)
    c = ws.cell(1, 1, cfg.titulo)
    c.font = Font(name=cfg.fuente, bold=True, size=13, color=NAVY)
    c.alignment = left_al

    # --- Metadatos ---
    ws.cell(3, 1, "Inicio").font = Font(name=cfg.fuente, bold=True, size=9)
    hi = ws.cell(3, 2, seconds_to_time(hms_to_seconds(hora_inicio), True))
    hi.font, hi.number_format = f_cell, "h:mm"
    ws.cell(3, 9, "Origen:").font = Font(name=cfg.fuente, bold=True, size=9)
    ws.cell(3, 10, origen_archivo).font = f_cell
    ws.cell(4, 1, "Versión").font = Font(name=cfg.fuente, bold=True, size=9)
    ws.cell(4, 2, "Simulador").font = f_cell

    # --- Rótulos de terminal ---
    for base, nombre in ((L0, cfg.nombre_puerto), (R0, cfg.nombre_limache)):
        ws.merge_cells(start_row=5, start_column=base, end_row=5, end_column=base + 9)
        t = ws.cell(5, base, nombre)
        t.font, t.alignment = f_term, center
        t.fill = PatternFill("solid", fgColor=NAVY)

    # --- Encabezados de columnas ---
    for base in (L0, R0):
        for j, h in enumerate(COLUMNAS):
            cc = ws.cell(6, base + j, h)
            cc.font, cc.alignment, cc.border = f_hdr, center, borde
            cc.fill = PatternFill("solid", fgColor=NAVY)

    # --- Datos ---
    def volcar(filas: list[list], base: int) -> None:
        for i, fila in enumerate(filas):
            r = 7 + i
            for j, val in enumerate(fila):
                cc = ws.cell(r, base + j, val if val != "" else None)
                cc.border, cc.font = borde, f_cell
                if j in (2, 4):                      # Partida, Inter. -> hora
                    cc.number_format = "h:mm:ss" if not cfg.round_minutes else "h:mm"
                    cc.alignment = center
                elif j in (0, 1, 3, 5, 6, 7, 9):
                    cc.alignment = center
                else:
                    cc.alignment = left_al
            man = fila[5]
            if man == "EV":
                relleno = VERDE
            elif man == "SV":
                relleno = GRIS
            else:
                relleno = None
            if relleno:
                for j in range(10):
                    ws.cell(r, base + j).fill = PatternFill("solid", fgColor=relleno)
            ws.cell(r, base + 5).font = Font(
                name=cfg.fuente, size=9, bold=True,
                color="375623" if man == "EV" else "7F7F7F" if man == "SV" else "000000",
            )

    volcar(izq, L0)
    volcar(der, R0)

    # --- Anchos de columna / vista ---
    anchos = [6, 6, 9, 5, 8, 6, 7, 9, 15, 8]
    for base in (L0, R0):
        for j, w in enumerate(anchos):
            ws.column_dimensions[get_column_letter(base + j)].width = w
    ws.column_dimensions[get_column_letter(11)].width = 2  # separador
    ws.freeze_panes = "A7"
    ws.sheet_view.showGridLines = False

    return wb


def escribir_excel(izq: list[list], der: list[list], out_path: str,
                   cfg: Config, hora_inicio: str, origen_archivo: str) -> None:
    """Guarda el libro en disco (versión de línea de comandos)."""
    wb = construir_workbook(izq, der, cfg, hora_inicio, origen_archivo)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.save(out_path)


def construir_tablas(viajes: pd.DataFrame, cfg: Config) -> tuple[list[list], list[list]]:
    """A partir del DataFrame de viajes, arma las filas de ambas terminales."""
    if cfg.maniobras:
        viajes = asignar_maniobras(viajes)
    else:
        viajes = viajes.assign(man="")
    sv = fines_de_servicio(viajes, cfg) if cfg.maniobras else pd.DataFrame(columns=["side"])
    izq = construir_tabla(viajes, sv, cfg.track_puerto, cfg.cod_puerto, cfg.cod_limache, cfg)
    der = construir_tabla(viajes, sv, cfg.track_limache, cfg.cod_limache, cfg.cod_puerto, cfg)
    return izq, der


# --------------------------------------------------------------------------- #
# Orquestador
# --------------------------------------------------------------------------- #
def convertir(csv_path: str, out_path: str, cfg: Config) -> dict:
    """Pipeline completo CSV -> XLSX en disco. Devuelve un pequeño resumen."""
    viajes = cargar_viajes(csv_path, cfg)
    izq, der = construir_tablas(viajes, cfg)
    hora_inicio = viajes["dep"].iloc[0]
    escribir_excel(izq, der, out_path, cfg, hora_inicio, os.path.basename(csv_path))

    return {
        "viajes_total": len(viajes),
        "filas_puerto": len(izq),
        "filas_limache": len(der),
        "trenes": int(viajes["train"].nunique()),
        "salida": out_path,
    }


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Convierte el CSV del simulador a una planilla horaria con "
                    "maniobras en formato de dos terminales (.xlsx).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("input", help="Ruta del CSV del simulador.")
    p.add_argument("-o", "--output", default="Planilla_Maniobras.xlsx",
                   help="Ruta del archivo .xlsx de salida.")
    p.add_argument("--sep", default=";", help="Separador de columnas del CSV.")

    g = p.add_argument_group("Terminales y sentidos")
    g.add_argument("--track-puerto", type=int, default=0,
                   help="trackID que corresponde a la Terminal Puerto.")
    g.add_argument("--track-limache", type=int, default=1,
                   help="trackID que corresponde a la Terminal Limache.")
    g.add_argument("--cod-puerto", default="PUE", help="Código de estación de la Terminal Puerto.")
    g.add_argument("--cod-limache", default="LIM", help="Código de estación de la Terminal Limache.")
    g.add_argument("--nombre-puerto", default="Terminal Puerto", help="Rótulo de la columna izquierda.")
    g.add_argument("--nombre-limache", default="Terminal Limache", help="Rótulo de la columna derecha.")

    o = p.add_argument_group("Opciones de formato")
    o.add_argument("--multiple-threshold", type=int, default=400,
                   help="Capacidad mínima para marcar el tren como 'Múltiple'.")
    o.add_argument("--round-minutes", action="store_true",
                   help="Redondear las horas al minuto (descarta los segundos).")
    o.add_argument("--no-maniobras", action="store_true",
                   help="No derivar EV/RET/SV (genera solo la planilla básica).")
    o.add_argument("--train-prefix", default="",
                   help="Prefijo para renumerar los trenes (ej. '60' -> 600,601,...).")
    o.add_argument("--titulo", default="Planilla Horaria + Maniobras — Simulador",
                   help="Título del encabezado de la planilla.")
    o.add_argument("--fuente", default="Arial", help="Tipografía de la planilla.")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    cfg = Config(
        sep=args.sep,
        track_puerto=args.track_puerto, track_limache=args.track_limache,
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
    print(f"  Filas Puerto      : {resumen['filas_puerto']}")
    print(f"  Filas Limache     : {resumen['filas_limache']}")
    print(f"  Archivo           : {resumen['salida']}")
    return 0


def _corriendo_en_streamlit() -> bool:
    """True si el módulo se está ejecutando dentro de una app Streamlit."""
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
    # Si alguien apunta Streamlit a este archivo por error, mostramos la app
    # en lugar de fallar con el error de línea de comandos ("required: input").
    if _corriendo_en_streamlit():
        import streamlit_app  # noqa: F401  -> renderiza la interfaz completa
    else:
        raise SystemExit(main())
