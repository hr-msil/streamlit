import streamlit as st
import pandas as pd
import io
import re

from services.descuentos import leer_excel
from services.descuentos import armado_df
from services.descuentos import comparacion
from services.descuentos import imprimir_diferencias
from services.descuentos import unir_diccionarios
from services.descuentos import crear_excel

st.set_page_config(page_title="Descuentos", page_icon="✅",  layout = 'wide')

st.markdown("**Versión beta: cualquier cosa rara que encuentres, no dudes en reportarla!**")

st.markdown("Subí la planilla de la dependencia directo del Google Sheets.")
archivo_izquierda = st.file_uploader("archivo_izquierda", type = 'xlsx', accept_multiple_files = False, key = "archivo_izquierda")
st.markdown("Subí la planilla del segundo cálculo de descuentos.")
archivo_derecha = st.file_uploader("archivo_derecha", type = 'xlsx', accept_multiple_files = False, key = "archivo_derecha")

if archivo_derecha and archivo_izquierda:

    excel_file = pd.ExcelFile(archivo_izquierda)
    nombres_hojas = excel_file.sheet_names
    opciones_hojas = [""] + nombres_hojas

    hoja = st.selectbox(
    "Elegir hoja que se quiere procesar",
    opciones_hojas
    )

    if hoja:
        df_uno = leer_excel(archivo_izquierda, hoja)
        df_dos = leer_excel(archivo_derecha)

        dic_legajos_izq, df_limpio_izq = armado_df(df_uno)
        dic_legajos_der, df_limpio_der = armado_df(df_dos)

        diferencias = comparacion(df_limpio_izq,df_limpio_der)

        imprimir_diferencias(diferencias)
        dict_personas = unir_diccionarios(dic_legajos_izq, dic_legajos_der)
        nombre_dependencia = re.sub(r"\s*\(\d+\)", "", archivo_izquierda.name)
        buffer = io.BytesIO()
        planilla_final = crear_excel(df_limpio_izq, diferencias, dict_personas, buffer)
        buffer.seek(0)
        st.download_button(
            label="Descargar planilla para comparar",
            data=buffer,
            file_name=nombre_dependencia,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
        )