import streamlit as st
from mecanizadas import leer_mecanica
import pandas as pd
import io

st.title('⚙️ Mecanizadas')

archivos_subidos = st.file_uploader("Seleccionar los archivos de Mecanizadas que se quieran procesar", type=["pdf"], key="archivo_i",accept_multiple_files=True)

if archivos_subidos:

    dfs = []
    dfs_datos = []
    nombres_institucion = []

    for i,archivo_subido in enumerate(archivos_subidos):

        df, df_datos, nombre_archivo_res = leer_mecanica(nombre_archivo = archivo_subido)

        dfs.append(df)
        dfs_datos.append(df_datos)
        
    df_global = pd.concat(dfs)
    df_datos_global = pd.concat(dfs_datos)
    output = io.BytesIO()
    with pd.ExcelWriter(output) as writer:
        df_global.to_excel(writer, sheet_name="IMPORTES", index=False)
        df_datos_global.to_excel(writer, sheet_name="DATOS", index=False)

    excel_procesado = output.getvalue()

    st.download_button(
        label = f"Descargar Mecanizada global",
        data = excel_procesado,
        file_name= f"mecanizada.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"descarga_mecanizada"
    )