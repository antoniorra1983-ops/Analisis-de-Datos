import streamlit as st

st.title("Prueba mínima")
st.caption("Si esto NO se va a blanco al interactuar, el problema está en la app grande.")

opcion = st.radio("Elige una opción", ["A", "B"])
st.write("Elegiste:", opcion)

texto = st.text_input("Escribe algo")
st.write("Escribiste:", texto)

archivo = st.file_uploader("Sube cualquier archivo")
if archivo is not None:
    st.success(f"Subiste: {archivo.name} ({len(archivo.getvalue())} bytes)")
