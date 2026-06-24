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

import base64
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
# Procesamiento con caché: se calcula una sola vez por archivo + opciones.
# Esto evita que al pulsar "Descargar" se recalcule todo y la página se caiga.
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner="Procesando…")
def procesar(raw: bytes, nombre: str, sep: str, cod_puerto: str, cod_limache: str,
             nombre_puerto: str, nombre_limache: str, multiple_threshold: int,
             round_minutes: bool, maniobras: bool, train_prefix: str, titulo: str):
    cfg = Config(
        sep=sep, cod_puerto=cod_puerto, cod_limache=cod_limache,
        nombre_puerto=nombre_puerto, nombre_limache=nombre_limache,
        multiple_threshold=multiple_threshold, round_minutes=round_minutes,
        maniobras=maniobras, train_prefix=train_prefix, titulo=titulo,
    )
    viajes = cargar_viajes(io.BytesIO(raw), cfg)
    cols, tablas = construir_tablas(viajes, cfg)
    wb = construir_workbook(cols, tablas, cfg, viajes["dep"].iloc[0], nombre)
    buf = io.BytesIO()
    wb.save(buf)
    resumen = dict(viajes=len(viajes), trenes=int(viajes["train"].nunique()))
    return cols, tablas, buf.getvalue(), resumen


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
        datos.append(["" if x is None else x for x in fila])
    return pd.DataFrame(datos, columns=COLUMNAS)


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
            "de cada `tripID` definen origen y destino."
        )
    st.stop()

try:
    cols, tablas, xlsx_bytes, resumen = procesar(
        archivo.getvalue(), archivo.name, sep, cod_puerto, cod_limache,
        nombre_puerto, nombre_limache, int(multiple_threshold),
        round_minutes, maniobras, train_prefix.strip(), titulo,
    )
except Exception as exc:  # noqa: BLE001
    st.error(f"No se pudo procesar el archivo: {exc}")
    st.stop()

# Detecta un planilla_maniobras.py desactualizado en el repositorio.
if not cols or not isinstance(cols[0], dict):
    st.error(
        "El archivo **planilla_maniobras.py** del repositorio está "
        "desactualizado (no coincide con esta versión de la app). Sube la "
        "última versión de **ambos** archivos `.py` y reinicia la app "
        "(*Manage app → Reboot*)."
    )
    st.stop()

# --- Descarga ---
# En vez de st.download_button (que en Safari de iPhone interrumpe la conexión
# con el servidor y obliga a reiniciar la app), usamos un enlace con el archivo
# incrustado que se abre en una pestaña nueva: la pestaña de la app queda intacta.
b64 = base64.b64encode(xlsx_bytes).decode()
enlace = (
    f'<a href="data:{XLSX_MIME};base64,{b64}" '
    f'download="Planilla_Maniobras.xlsx" target="_blank" rel="noopener" '
    f'style="display:inline-block;padding:0.55rem 1.3rem;background-color:#1F3864;'
    f'color:#ffffff;border-radius:0.5rem;text-decoration:none;font-weight:600;'
    f'font-family:sans-serif;">⬇️  Descargar Excel (.xlsx)</a>'
)
st.markdown(enlace, unsafe_allow_html=True)
st.caption(
    "En iPhone/iPad se abre en una pestaña nueva y el archivo queda en "
    "**Archivos → Descargas** (ábrelo con Numbers, Excel o Google Sheets). "
    "La app no se cierra: vuelve a su pestaña cuando termines."
)

with st.expander("¿El enlace no descarga? (descarga estándar)"):
    st.caption(
        "Este es el botón clásico de descarga. Funciona en computador, pero en "
        "Safari de iPhone puede interrumpir la app (tendrías que recargarla)."
    )
    st.download_button(
        label="Descargar (método estándar)",
        data=xlsx_bytes,
        file_name="Planilla_Maniobras.xlsx",
        mime=XLSX_MIME,
    )

# --- Resumen ---
metric_cols = st.columns(2 + len(cols))
metric_cols[0].metric("Viajes", resumen["viajes"])
metric_cols[1].metric("Trenes", resumen["trenes"])
for i, (c, t) in enumerate(zip(cols, tablas)):
    metric_cols[2 + i].metric(c["nombre"], len(t))

# --- Vista previa (tablas simples, livianas; una pestaña por terminal) ---
st.subheader("Vista previa")
st.caption(
    "EV = entrada a vía · RET = retorno · SV = sale de vía (estaciona). "
    "El Excel descargado incluye el resaltado de color de EV y SV."
)
pestanas = st.tabs([f"{c['nombre']} ({len(t)})" for c, t in zip(cols, tablas)])
for tab, c, t in zip(pestanas, cols, tablas):
    with tab:
        st.dataframe(filas_a_df(t, round_minutes),
                     hide_index=True, use_container_width=True, height=460)
