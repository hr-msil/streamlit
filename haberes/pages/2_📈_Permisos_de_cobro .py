import streamlit as st
import pandas as pd

from cobranProductividades import lectura_archivo_query
from cobranProductividades import no_permitido

from productividadesV2 import lectura_archivo_prod



st.sidebar.header("Permisos de cobro productividades")

st.title("📝 Permisos de cobro")

st.divider()

st.markdown("Subir los archivos de productividades arrojados por el sistema")

archivos_excel = st.file_uploader("Seleccionar archivo", type = "xls",key = "productividades",accept_multiple_files=True)

st.markdown("Subir el archivo correspondiente a la QUERY")

archivo_query = st.file_uploader("Seleccionar archivo", type = "xlsx", key="query")

if archivos_excel and archivo_query:

    df_query = lectura_archivo_query(archivo_query)

    dfs_excel = []
    for archivo_excel in archivos_excel:

            df_excel_i = lectura_archivo_prod(archivo_excel)
            
            dfs_excel.append(df_excel_i)
            
    df_excel_final = pd.concat(dfs_excel,ignore_index = True)

    no_pueden_por_categoria, no_pueden_por_ausencia,repetidos_mu, repetidos_me, repetidos_jo,repetidos_do = no_permitido(df_query,df_excel_final)

    if len(repetidos_mu) > 0:
         st.write("Estas personas tienen más de dos cargos en la partición MU: ")
         
         for repetido in repetidos_mu:
            st.write("""-""", repetido[0]," ",repetido[1])

    if len(repetidos_me) > 0:
        st.write("Estas personas tienen más de dos cargos en la partición ME: ")
        
        for repetido in repetidos_me:
            st.write("""-""", repetido[0]," ",repetido[1])

    if len(repetidos_jo) > 0:
         st.write("Estas personas tienen más de dos cargos en la partición JO: ")
         
         for repetido in repetidos_jo:
            st.write("""-""", repetido[0]," ",repetido[1])

    if len(repetidos_do) > 0:
        st.write("Estas personas tienen más de dos cargos en la partición DO: ")
        
        for repetido in repetidos_do:
            st.write("""-""", repetido[0]," ",repetido[1])

    st.write("Estos legajos no pueden cobrar productividad debido a ausencias: ")
    for legajo_ausencia in no_pueden_por_ausencia:
        st.write("""-""",legajo_ausencia)
    st.write("Estos legajos no pueden cobrar productividad debido a que no les corresponde por categoría")
    for lagajo_categoria in no_pueden_por_categoria:
        st.write("""-""", lagajo_categoria)