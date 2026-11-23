import glob
import io
import json
import os
import tempfile

import streamlit as st

from casa_inspeccion_a_planilla import detectar_perfil, procesar

BASE = os.path.dirname(os.path.abspath(__file__))
PLANTILLA = os.path.join(BASE, "plantilla", "planilla_base.xlsx")


def perfiles_disponibles():
    opciones = {}
    for ruta in sorted(glob.glob(os.path.join(BASE, "perfiles", "*.json"))):
        with open(ruta, encoding="utf-8") as f:
            nombre = json.load(f).get("nombre", os.path.basename(ruta))
        opciones[nombre] = ruta
    return opciones


st.set_page_config(page_title="Carga de informes de inspección", layout="wide")
st.title("Carga de informes de inspección a planilla Excel")
st.caption("Sube el PDF, elige el formato del informe y descarga la planilla con las observaciones cargadas.")

perfiles = perfiles_disponibles()
col1, col2 = st.columns([2, 1])
with col1:
    archivo = st.file_uploader("Informe en PDF", type=["pdf"])
with col2:
    opciones = ["Detectar automáticamente"] + list(perfiles.keys())
    formato = st.selectbox("Formato del informe", opciones)

if st.button("Generar planilla", type="primary", disabled=archivo is None):
    with st.spinner("Leyendo el informe..."):
        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = os.path.join(tmp, "informe.pdf")
            salida = os.path.join(tmp, "planilla.xlsx")
            with open(pdf_path, "wb") as f:
                f.write(archivo.getbuffer())
            ruta_perfil = None if formato == "Detectar automáticamente" else perfiles[formato]
            try:
                total, pendientes, nombre_formato = procesar(pdf_path, PLANTILLA, salida, ruta_perfil)
            except ValueError as e:
                st.error(str(e))
                st.stop()
            with open(salida, "rb") as f:
                contenido = f.read()

    st.success(f"Listo: {total} observaciones cargadas. {len(pendientes)} requieren revisión.")
    st.caption(f"Formato usado: {nombre_formato}")
    nombre = os.path.splitext(archivo.name)[0] + "_planilla.xlsx"
    st.download_button("Descargar planilla Excel", data=contenido, file_name=nombre,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if pendientes:
        st.subheader("Observaciones para revisar")
        st.dataframe([
            {"Observación": r["etiqueta"], "Faltan": ", ".join(r["faltantes"]), "Texto": r["desc"]}
            for r in pendientes
        ], use_container_width=True)
