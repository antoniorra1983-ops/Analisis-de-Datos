"""
streamlit_app.py
================

App web (Streamlit) del conversor **Planilla + Maniobras**.

Sube el CSV del simulador, ajusta las opciones y descarga el Excel con el
formato de varias terminales (Puerto · El Belloto · Sargento Aldea · Limache).
Reutiliza toda la lógica de `planilla_maniobras.py`.

Ejecutar en local:
    pip install -r requirements.txt
    streamlit run streamlit_app.py
"""

import datetime as dt
import io

import pandas as pd
import streamlit as st

from planilla_maniobras import (
    COLUMNAS,
    Config,
    cargar_viajes,
    construir_tablas,
    construir_workbook,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

st.set_page_config(page_title="Planilla + Maniobras", page_icon="🚆", layout="wide")
st.title("🚆 Planilla Horaria + Maniobras")
st.caption("Convierte el CSV del simulador en una planilla de varias terminales (.xlsx).")


# --------------------------------------------------------------------------- #
# Opciones (barra lateral)
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Opciones")
    maniobras = st.checkbox("Incluir maniobras (EV/RET/SV)", value=True)
    round_minutes = st.checkbox("Redondear horas al minuto", value=False)
    multiple_threshold = st.number_input(
        "Capacidad para marcar «Múltiple» ≥", min_value=1, value=400, step=50
    )
    train_prefix = st.text_input("Prefijo de tren", value="", placeholder="(ninguno)")

    with st.expander("Avanzado · terminales y formato"):
        sep = st.text_input("Separador del CSV", value=";")
        cod_puerto = st.text_input("Código Terminal Puerto", value="PUE")
        cod_limache = st.text_input("Código Terminal Limache", value="LIM")
        nombre_puerto = st.text_input("Rótulo Terminal Puerto", value="Terminal Puerto")
        nombre_limache = st.text_input("Rótulo Terminal Limache", value="Terminal Limache")
        st.caption(
            "Las columnas de **El Belloto** (BTO) y **Sargento Aldea** (SGA) "
            "aparecen automáticamente si hay servicios que parten desde ellas."
        )
        titulo = st.text_input("Título de la planilla",
                               value="Planilla Horaria + Maniobras — Simulador")


# --------------------------------------------------------------------------- #
# Utilidades de presentación
# --------------------------------------------------------------------------- #
def _fmt_time(v, redondear: bool) -> str:
    if v is None or v == "":
        return ""
    if isinstance(v, dt.time):
        return v.strftime("%H:%M" if redondear else "%H:%M:%S")
    return str(v)


def filas_a_df(filas: list[list], redondear: bool) -> pd.DataFrame:
    datos = []
    for f in filas:
        fila = list(f)
        fila[2] = _fmt_time(fila[2], redondear)  # Partida
        fila[4] = _fmt_time(fila[4], redondear)  # Inter.
        datos.append(fila)
    return pd.DataFrame(datos, columns=COLUMNAS)


def estilo(df: pd.DataFrame):
    """Colorea EV (verde) y SV (gris), igual que el Excel."""
    def color(row):
        man = row["Man."]
        if man == "EV":
            return ["background-color: #E2EFDA"] * len(row)
        if man == "SV":
            return ["background-color: #F2F2F2; color: #777"] * len(row)
        return [""] * len(row)
    return df.style.apply(color, axis=1)


# --------------------------------------------------------------------------- #
# Carga y procesamiento
# --------------------------------------------------------------------------- #
archivo = st.file_uploader("Sube el CSV del simulador", type=["csv"])

if archivo is None:
    st.info(
        "Sube un archivo CSV para comenzar. El procesamiento ocurre en el "
        "servidor de la app; el archivo no se comparte con terceros."
    )
    with st.expander("¿Qué formato debe tener el CSV?"):
        st.markdown(
            "Separado por `;`, con (al menos) estas columnas:\n\n"
            "`tripID`, `trainID`, `trainTotalCapacity`, `trackID`, "
            "`stationName`, `arriveTime`, `leaveTime`\n\n"
            "Cada fila es una parada de un viaje; el primer y último registro "
            "de cada `tripID` definen origen y destino. Cada viaje se ubica en "
            "la columna de su estación de origen."
        )
    st.stop()

cfg = Config(
    sep=sep,
    cod_puerto=cod_puerto,
    cod_limache=cod_limache,
    nombre_puerto=nombre_puerto,
    nombre_limache=nombre_limache,
    multiple_threshold=int(multiple_threshold),
    round_minutes=round_minutes,
    maniobras=maniobras,
    train_prefix=train_prefix.strip(),
    titulo=titulo,
)

try:
    raw = archivo.getvalue()
    viajes = cargar_viajes(io.BytesIO(raw), cfg)
    cols, tablas = construir_tablas(viajes, cfg)
    hora_inicio = viajes["dep"].iloc[0]
except Exception as exc:  # noqa: BLE001
    st.error(f"No se pudo procesar el archivo: {exc}")
    st.stop()

# --- Resumen ---
metric_cols = st.columns(2 + len(cols))
metric_cols[0].metric("Viajes", len(viajes))
metric_cols[1].metric("Trenes", int(viajes["train"].nunique()))
for i, (c, t) in enumerate(zip(cols, tablas)):
    metric_cols[2 + i].metric(c["nombre"], len(t))

# --- Descarga ---
wb = construir_workbook(cols, tablas, cfg, hora_inicio, archivo.name)
buffer = io.BytesIO()
wb.save(buffer)
buffer.seek(0)

st.download_button(
    label="⬇️  Descargar Excel (.xlsx)",
    data=buffer,
    file_name="Planilla_Maniobras.xlsx",
    mime=XLSX_MIME,
    type="primary",
)

# --- Vista previa (una pestaña por terminal) ---
st.subheader("Vista previa")
pestanas = st.tabs([f"{c['nombre']} ({len(t)})" for c, t in zip(cols, tablas)])
for tab, c, t in zip(pestanas, cols, tablas):
    with tab:
        st.dataframe(estilo(filas_a_df(t, round_minutes)),
                     hide_index=True, use_container_width=True, height=460)

st.caption("EV = entrada a vía · RET = retorno · SV = sale de vía (estaciona)")
