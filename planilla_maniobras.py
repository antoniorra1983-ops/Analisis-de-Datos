#!/usr/bin/env python3
"""
planilla_a_simulador.py
=======================

Convierte la hoja **Planilla + Maniobras** de un archivo .xls (laboral o sábado)
al **formato plano de entrada del simulador** (estilo `Planillaprueba2.xls`),
y lo guarda como **.xls**.

Formato de salida (9 columnas, una fila por servicio, ordenadas por hora):

    hora | origen | unidades | destino | unidades | destino | "servicio" | tren | 406

Donde:
  * origen / destino  -> códigos de estación (PUE, LIM, SGA, BTO)
  * unidades          -> 2 si el tren es "Múltiple" (o capacidad >= 400), si no 1
  * "servicio"        -> texto fijo
  * 406               -> valor fijo (parámetro --constante)

Uso:
    python planilla_a_simulador.py Planilla_Laboral_510_minutos_Ver_1_mardic.xls -o salida.xls

Requisitos: xlrd (lectura .xls), xlwt (escritura .xls)
"""
from __future__ import annotations

import argparse

import xlrd
import xlwt

# Etiqueta de terminal -> código de estación
TERMINALES = [
    ("puerto", "PUE"),
    ("belloto", "BTO"),
    ("sargento", "SGA"),
    ("aldea", "SGA"),
    ("limache", "LIM"),
]

# Códigos numéricos de destino que aparecen en la columna "Destino" (zonas).
# 6 = Limache, 4 = Sargento Aldea (verificado por proporción y retornos).
ZONA_A_ESTACION = {6: "LIM", 4: "SGA", 5: "BTO"}

ESTACIONES = {"PUE", "LIM", "SGA", "BTO", "AME", "PEN"}
ENCABEZADOS = ["Viaje", "Tren", "Partida", "N°", "Inter.", "Man.", "Destino", "M", "Obs.", "Capacidad"]


def _norm(s) -> str:
    return str(s).strip().lower()


def _es_entero_pos(v) -> bool:
    return isinstance(v, (int, float)) and float(v) == int(v) and int(v) > 0


def _hora_seg(sh, wb, r, c):
    """Devuelve la hora de la celda en segundos desde medianoche, o None."""
    if c is None:
        return None
    ct = sh.cell_type(r, c)
    v = sh.cell_value(r, c)
    if ct == xlrd.XL_CELL_DATE:
        t = xlrd.xldate_as_tuple(v, wb.datemode)
        return t[3] * 3600 + t[4] * 60 + t[5]
    if isinstance(v, str) and ":" in v:
        p = [int(x) for x in v.split(":")]
        return p[0] * 3600 + p[1] * 60 + (p[2] if len(p) > 2 else 0)
    return None


def _val(sh, r, c):
    if c is None:
        return ""
    v = sh.cell_value(r, c)
    if isinstance(v, float) and v == int(v):
        return int(v)
    return v


def _detectar_estructura(sh):
    """Encuentra la fila de rótulos de terminal, la de encabezados y los bloques."""
    fila_rotulos = fila_encabezados = None
    for r in range(min(15, sh.nrows)):
        textos = [_norm(sh.cell_value(r, c)) for c in range(sh.ncols)]
        if any("viaje" in t for t in textos) and any("partida" in t for t in textos):
            fila_encabezados = r
        if any(any(k in t for k, _ in TERMINALES) for t in textos) and \
           any("terminal" in t or "belloto" in t or "sargento" in t or "aldea" in t for t in textos):
            fila_rotulos = r
    if fila_encabezados is None:
        raise ValueError("No se encontró la fila de encabezados (con 'Viaje'/'Partida').")
    if fila_rotulos is None:
        fila_rotulos = fila_encabezados - 1

    # Detectar inicio de cada bloque por los rótulos de terminal
    bloques = []
    for c in range(sh.ncols):
        etiqueta = _norm(sh.cell_value(fila_rotulos, c))
        if not etiqueta:
            continue
        code = next((cod for clave, cod in TERMINALES if clave in etiqueta), None)
        if code:
            bloques.append((c, code))

    # Para cada bloque, mapear columnas usando la fila de encabezados
    resultado = []
    for i, (c0, code) in enumerate(bloques):
        c1 = bloques[i + 1][0] if i + 1 < len(bloques) else sh.ncols
        colmap = {}
        for c in range(c0, c1):
            h = _norm(sh.cell_value(fila_encabezados, c))
            for nombre in ENCABEZADOS:
                if _norm(nombre) == h or (nombre == "N°" and h in ("n°", "n")):
                    colmap[nombre] = c
        resultado.append((code, colmap))
    return fila_encabezados, resultado


def _decodificar_destino(origen: str, raw, opuesto_de_puerto: str) -> str:
    """Determina el código de estación destino."""
    if isinstance(raw, str):
        s = raw.strip().upper()
        if s in ESTACIONES:
            return s
        if s == "":
            raw = None
    if isinstance(raw, (int, float)) and float(raw) == int(raw):
        return ZONA_A_ESTACION.get(int(raw), opuesto_de_puerto if origen == "PUE" else "PUE")
    # vacío -> extremo opuesto
    return "LIM" if origen == "PUE" else "PUE"


def _unidades(sh, r, colmap, umbral=400) -> int:
    cap = _val(sh, r, colmap.get("Capacidad"))
    if isinstance(cap, (int, float)) and cap:
        return 2 if cap >= umbral else 1
    m = str(_val(sh, r, colmap.get("M"))).lower()
    return 2 if "ltiple" in m else 1


def leer_planilla_maniobras(path: str, hoja: str = "Planilla + Maniobras") -> list[dict]:
    """Extrae las salidas (servicios) de la hoja Planilla + Maniobras."""
    wb = xlrd.open_workbook(path)
    nombres = wb.sheet_names()
    if hoja not in nombres:
        hoja = next((n for n in nombres if "maniobra" in n.lower()), nombres[0])
    sh = wb.sheet_by_name(hoja)
    fila_enc, bloques = _detectar_estructura(sh)

    salidas = []
    for code, colmap in bloques:
        for r in range(fila_enc + 1, sh.nrows):
            viaje = _val(sh, r, colmap.get("Viaje"))
            hora = _hora_seg(sh, wb, r, colmap.get("Partida"))
            if not _es_entero_pos(viaje) or hora is None:
                continue  # filas de paso (sin viaje) o SV (sin hora) se omiten
            tren = _val(sh, r, colmap.get("Tren"))
            if not _es_entero_pos(tren):
                continue
            dest = _decodificar_destino(code, _val(sh, r, colmap.get("Destino")), "LIM")
            salidas.append({
                "hora": hora, "origen": code, "destino": dest,
                "tren": int(tren), "unidades": _unidades(sh, r, colmap),
            })
    salidas.sort(key=lambda d: (d["hora"], d["origen"]))
    return salidas


def escribir_xls(salidas: list[dict], out_path: str, constante: int = 406) -> None:
    """Escribe el formato plano del simulador como .xls."""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Hoja1")
    estilo_hora = xlwt.easyxf(num_format_str="HH:MM:SS")
    for i, s in enumerate(salidas):
        ws.write(i, 0, s["hora"] / 86400.0, estilo_hora)
        ws.write(i, 1, s["origen"])
        ws.write(i, 2, s["unidades"])
        ws.write(i, 3, s["destino"])
        ws.write(i, 4, s["unidades"])
        ws.write(i, 5, s["destino"])
        ws.write(i, 6, "servicio")
        ws.write(i, 7, s["tren"])
        ws.write(i, 8, constante)
    wb.save(out_path)


def convertir(entrada: str, salida: str, hoja: str = "Planilla + Maniobras",
              constante: int = 406) -> dict:
    salidas = leer_planilla_maniobras(entrada, hoja)
    escribir_xls(salidas, salida, constante)
    from collections import Counter
    return {
        "servicios": len(salidas),
        "por_origen": dict(Counter(s["origen"] for s in salidas)),
        "salida": salida,
    }


def main(argv=None) -> int:
    p = argparse.ArgumentParser(
        description="Convierte la hoja Planilla + Maniobras (.xls) al formato plano "
                    "del simulador (estilo Planillaprueba2), en .xls.")
    p.add_argument("input", help="Archivo .xls de entrada (laboral o sábado).")
    p.add_argument("-o", "--output", default="Planilla_Simulador_entrada.xls")
    p.add_argument("--hoja", default="Planilla + Maniobras", help="Nombre de la hoja a leer.")
    p.add_argument("--constante", type=int, default=406, help="Valor fijo de la última columna.")
    args = p.parse_args(argv)
    r = convertir(args.input, args.output, args.hoja, args.constante)
    print("Conversión lista.")
    print(f"  Servicios : {r['servicios']}")
    print(f"  Por origen: {r['por_origen']}")
    print(f"  Archivo   : {r['salida']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
