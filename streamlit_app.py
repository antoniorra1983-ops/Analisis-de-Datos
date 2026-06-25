"""
streamlit_app.py
================

App web (Streamlit) del conversor **Planilla + Maniobras** — una sola
herramienta con dos modos:

  1) Simulador → Planilla + Maniobras   (CSV del simulador → Excel .xlsx)
  2) Planilla + Maniobras → Simulador   (.xls Planilla+Maniobras → .xls entrada)

Toda la lógica vive en `planilla_maniobras.py`.

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
    elegir_hoja_pm,
    leer_planilla_maniobras,
    listar_hojas,
    simulador_a_bytes,
)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLS_MIME = "application/vnd.ms-excel"

MODO_1 = "Simulador → Planilla + Maniobras  (CSV → Excel)"
MODO_2 = "Planilla + Maniobras → Simulador  (.xls → .xls)"

st.set_page_config(page_title="Planilla + Maniobras", page_icon="🚆", layout="wide")
st.title("🚆 Planilla Horaria + Maniobras")


# --------------------------------------------------------------------------- #
# Utilidades comunes
# --------------------------------------------------------------------------- #
def hhmmss(seg: int) -> str:
    return f"{seg // 3600:02d}:{(seg % 3600) // 60:02d}:{seg % 60:02d}"


# --------------------------------------------------------------------------- #
# MODO 1 · Simulador → Planilla + Maniobras
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner="Procesando…", max_entries=3, ttl=3600)
def _procesar_csv(raw, nombre, sep, cod_puerto, cod_limache, nombre_puerto,
                  nombre_limache, multiple_threshold, round_minutes, maniobras,
                  train_prefix, titulo):
    cfg = Config(
        sep=sep, cod_puerto=cod_puerto, cod_limache=cod_limache,
        nombre_puerto=nombre_puerto, nombre_limache=nombre_limache,
        multiple_threshold=multiple_threshold, round_minutes=round_minutes,
        maniobras=maniobras, train_prefix=train_prefix, titulo=titulo,
    )
    viajes = cargar_viajes(io.BytesIO(raw), cfg)
    cols, tablas = construir_tablas(viajes, cfg)
    wb = construir_workbook(cols, tablas, cfg, viajes["dep"].iloc[0], nombre)
    buf = io.BytesIO(); wb.save(buf)
    return cols, tablas, buf.getvalue(), dict(viajes=len(viajes), trenes=int(viajes["train"].nunique()))


def _fmt_time(v, redondear):
    if v is None or v == "":
        return ""
    return v.strftime("%H:%M" if redondear else "%H:%M:%S") if isinstance(v, dt.time) else str(v)


def _filas_a_df(filas, redondear):
    datos = []
    for f in filas:
        fila = list(f)
        fila[2] = _fmt_time(fila[2], redondear)
        fila[4] = _fmt_time(fila[4], redondear)
        datos.append(["" if x is None else x for x in fila])
    return pd.DataFrame(datos, columns=COLUMNAS)


def modo_csv_a_planilla():
    with st.sidebar:
        maniobras = st.checkbox("Incluir maniobras (EV/RET/SV)", value=True)
        round_minutes = st.checkbox("Redondear horas al minuto", value=False)
        multiple_threshold = st.number_input("Capacidad para «Múltiple» ≥", min_value=1, value=400, step=50)
        train_prefix = st.text_input("Prefijo de tren", value="", placeholder="(ninguno)")
        with st.expander("Avanzado · terminales"):
            sep = st.text_input("Separador del CSV", value=";")
            cod_puerto = st.text_input("Código Terminal Puerto", value="PUE")
            cod_limache = st.text_input("Código Terminal Limache", value="LIM")
            nombre_puerto = st.text_input("Rótulo Terminal Puerto", value="Terminal Puerto")
            nombre_limache = st.text_input("Rótulo Terminal Limache", value="Terminal Limache")
            titulo = st.text_input("Título", value="Planilla Horaria + Maniobras — Simulador")

    archivo = st.file_uploader("Sube el CSV del simulador", type=["csv"], key="csv")
    if archivo is None:
        st.info("Sube el CSV del simulador (`Planilla_Simulador.csv`) para generar la planilla.")
        return

    try:
        cols, tablas, xlsx_bytes, resumen = _procesar_csv(
            archivo.getvalue(), archivo.name, sep, cod_puerto, cod_limache,
            nombre_puerto, nombre_limache, int(multiple_threshold),
            round_minutes, maniobras, train_prefix.strip(), titulo,
        )
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo procesar el archivo: {exc}")
        return

    if not cols or not isinstance(cols[0], dict):
        st.error("`planilla_maniobras.py` está desactualizado en el repositorio. "
                 "Sube la última versión de los archivos `.py` y reinicia la app.")
        return

    st.download_button("⬇️  Descargar Excel (.xlsx)", data=xlsx_bytes,
                       file_name="Planilla_Maniobras.xlsx", mime=XLSX_MIME, type="primary")

    mc = st.columns(2 + len(cols))
    mc[0].metric("Viajes", resumen["viajes"])
    mc[1].metric("Trenes", resumen["trenes"])
    for i, (c, t) in enumerate(zip(cols, tablas)):
        mc[2 + i].metric(c["nombre"], len(t))

    if st.checkbox("Ver vista previa de la planilla"):
        st.caption("EV = entrada a vía · RET = retorno · SV = sale de vía. Mostrando hasta 60 filas por terminal.")
        pestanas = st.tabs([f"{c['nombre']} ({len(t)})" for c, t in zip(cols, tablas)])
        for tab, c, t in zip(pestanas, cols, tablas):
            with tab:
                st.table(_filas_a_df(t, round_minutes).head(60))


# --------------------------------------------------------------------------- #
# MODO 2 · Planilla + Maniobras → entrada del simulador
# --------------------------------------------------------------------------- #
@st.cache_data(show_spinner="Convirtiendo…", max_entries=3, ttl=3600)
def _procesar_pm(raw, hoja, constante):
    salidas = leer_planilla_maniobras(raw, hoja)
    return salidas, simulador_a_bytes(salidas, constante)


@st.cache_data(show_spinner=False, max_entries=5, ttl=3600)
def _hojas_archivo(raw):
    """Lista las hojas del archivo y cuál es la Planilla + Maniobras detectada."""
    return listar_hojas(raw), elegir_hoja_pm(raw)


def modo_planilla_a_simulador():
    with st.sidebar:
        constante = st.number_input("Capacidad por unidad (simple)", min_value=0, value=406, step=1,
                                    help="406 por unidad: 406 si es simple, 812 si es doble.")

    archivo = st.file_uploader("Sube la Planilla + Maniobras (.xls)", type=["xls"], key="xls")
    if archivo is None:
        st.info("Sube el `.xls` de la Planilla + Maniobras (laboral o sábado) para generar la "
                "entrada del simulador en formato plano (.xls).")
        return

    raw = archivo.getvalue()

    # Detectar la hoja correcta (por estructura, sin importar el nombre) y dejar elegir
    try:
        hojas, detectada = _hojas_archivo(raw)
    except ModuleNotFoundError as exc:
        st.error(f"Falta una librería en el servidor: **{exc.name}**. Tu `requirements.txt` debe "
                 "incluir `xlrd` y `xlwt`. Súbelo a GitHub y reinicia la app (*Manage app → Reboot*).")
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo abrir el archivo: {exc}")
        return

    if len(hojas) > 1:
        idx = hojas.index(detectada) if detectada in hojas else 0
        hoja = st.selectbox("Hoja a convertir", hojas, index=idx,
                            help="Se detecta automáticamente la hoja con formato Planilla + "
                                 "Maniobras; si no es la correcta, elígela aquí.")
    else:
        hoja = hojas[0]
        st.caption(f"Hoja: **{hoja}**")

    try:
        salidas, xls_bytes = _procesar_pm(raw, hoja, int(constante))
    except ModuleNotFoundError as exc:
        st.error(f"Falta una librería en el servidor: **{exc.name}**. Tu `requirements.txt` debe "
                 "incluir `xlrd` y `xlwt`. Súbelo a GitHub y reinicia la app (*Manage app → Reboot*).")
        return
    except Exception as exc:  # noqa: BLE001
        st.error(f"No se pudo leer la planilla: {exc}")
        return

    if not salidas:
        st.warning(f"La hoja **{hoja}** no tiene servicios con N° de viaje y hora. "
                   "Prueba a elegir otra hoja en el desplegable.")
        return

    st.download_button("⬇️  Descargar entrada del simulador (.xls)", data=xls_bytes,
                       file_name="Entrada_Simulador.xls", mime=XLS_MIME, type="primary")

    from collections import Counter
    por_origen = Counter(s["origen"] for s in salidas)
    mc = st.columns(1 + len(por_origen))
    mc[0].metric("Servicios", len(salidas))
    for i, (org, n) in enumerate(sorted(por_origen.items())):
        mc[1 + i].metric(f"Salen de {org}", n)

    if st.checkbox("Ver vista previa de los servicios"):
        df = pd.DataFrame([{
            "Hora": hhmmss(s["hora"]), "Origen": s["origen"], "Vía": s["via"],
            "Destino": s["destino"], "Tren": s["tren"],
            "Cap.": int(constante) * s["unidades"],
        } for s in salidas])
        st.table(df.head(60))
        st.caption("Primeras 60 filas; el .xls trae todas. La vía va en las columnas C y E; "
                   "la capacidad (406 simple · 812 doble) en la última.")


# --------------------------------------------------------------------------- #
# Selector de modo
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Modo")
    modo = st.radio("¿Qué quieres hacer?", [MODO_1, MODO_2], label_visibility="collapsed")
    st.divider()
    st.subheader("Opciones")

st.caption(MODO_1 if modo == MODO_1 else MODO_2)

if modo == MODO_1:
    modo_csv_a_planilla()
else:
    modo_planilla_a_simulador()
