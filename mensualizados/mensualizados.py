# Toma un archivo, lo convierte a un dataFrame, elimina las ultimas 3 columnas 
# Si el archivo viene del area de SALUD PUBLICA y AMBIENTE Y ESPACIO PUBLICO subdividir en oficinas
# Y que cada oficina sea un archivo por separado

import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers
import xlwt

#---------- Funciones principales -----------

def borrar_ultimas_columnas(df: pd.DataFrame, n: int) -> pd.DataFrame:
    """
    Borra las últimas n columnas de un dataFrame
    
    :param df: DataFrame completo
    :param n: Cantidad de últimas columnas a borrar
    :return: Devuelve el dataFrame original sin las últimas n columnas.
    """

    
    cant_columnas = df.shape[1]

    columnas_a_eliminar = range(cant_columnas - n, cant_columnas)

    df = df.drop(df.columns[columnas_a_eliminar], axis=1)

    return df


def dividir_oficinas(df: pd.DataFrame) -> tuple[list[pd.DataFrame], list[str]]:
    """
    Precondicion: que el DataFrame pertenezca al área de salud pública/ambiente y espacio público/desarrollo humano y deportes/educación, cultura y trabajo.
    Si cuando filtramos por el nombre de oficina, tenemos alguna fila en la columna de Fecha Egreso Cargo con algún valor nan, no agrega esa oficina a la lista
    de dataFrames resultantes

    :param df: archivo .xlsx subido por la secretaría para ser procesado
    :type df: pd.DataFrame
    :return: lista de dataFrames separados por oficinas con ninguna de sus filas en la columna "Fecha Egreso Cargo" con valor nan. Lista de strings
    con los nombres de las oficinas que tienen algpun valor de "Fecha Egreso Cargo" por completar.
    :rtype: tuple[list[pd.DataFrame],list[str]]
    """

    oficinas_unicas = df["Oficina"].unique()
    df_oficinas = []
    oficinas_nan = []

    for oficina in oficinas_unicas:

        df_oficina = df[df["Oficina"] == oficina]
        df_oficina_na = df_oficina[df_oficina["Fecha Egreso Cargo"].isna()]

        if df_oficina_na.shape[0] == 0:
        #Si no hay ningun na en niguna de las filas del dataFrame filtrado por oficina, lo carga y lo devuleve, en caso contrario no lo carga
            df_oficinas.append(df_oficina)
        else:
            oficinas_nan.append(oficina)

    return df_oficinas,oficinas_nan


#############
##STREAMLIT##
#############


st.title("📝Mensualizados")

st.divider()


opciones = [
    "",
    "AMBIENTE Y ESPACIO PUBLICO",
    "ARSI",
    "CAPITAL HUMANO",
    "DESARROLLO HUMANO Y DEPORTES",
    "EDUCACION, CULTURA Y TRABAJO",
    "GENERAL",
    "GOBIERNO",
    "H.C.D.",
    "HACIENDA Y FINANZAS",
    "JEFATURA DE GABINETE",
    "LEGAL Y TECNICA",
    "PLANEAMIENTO URBANO",
    "PRIVADA",
    "SALUD PUBLICA",
    "SEGURIDAD"
]



st.subheader("Elegir el área del cuál se está subiendo el archivo:")

opcion = st.selectbox(
    "Elegir una opción",
    opciones
)

if opcion == "":
    st.subheader("IMPORTANTE❗: seleccionar el área antes de continuar")

elif opcion == "AMBIENTE Y ESPACIO PUBLICO" or opcion == "SALUD PUBLICA" or opcion == "DESARROLLO HUMANO Y DEPORTES" or opcion == "EDUCACION, CULTURA Y TRABAJO":

    #En caso de ser de alguna de estas oficinas, separamos el archivo por oficinas y retornamos tantos archivos como oficinas COMPLETAS haya.

    st.subheader(f"📂Archivo de mensualizados del área {opcion}")

    st.markdown("Subir el archivo de mensualizados")

    archivo_1 = st.file_uploader("Seleccionar el archivo de mensualizados", type=["xlsx", "xls"], key="archivo1",accept_multiple_files=False)

    if archivo_1:

        excel_file = pd.ExcelFile(archivo_1)

        nombres_hojas = excel_file.sheet_names
        opciones_hojas = [""] + nombres_hojas
        

        hoja = st.selectbox(
        "Elegir hoja que se quiere procesar",
        opciones_hojas
        )

        if hoja == "":

            st.subheader("IMPORTANTE❗: seleccionar la hoja antes de continuar")

        elif hoja == "HOJA":

            st.subheader("Esta hoja no puede ser procesada")


        else:

            df = pd.read_excel(archivo_1,sheet_name=hoja)

            df["Categoría"] = df["Categoría"].replace("NO CATEGORIZADO", 999)

            #Filtrar dataFrame sacando los que tienen en Fecha Egreso Cargo #Enviar nota de designación"
            df = df[df["Fecha Egreso Cargo"] != "Enviar nota de designación"]

            df = borrar_ultimas_columnas(df, 3)

            df_oficinas,oficinas_nan = dividir_oficinas(df)

            
            for df_oficina in df_oficinas:
                

                df_oficina = df_oficina.reset_index(drop=True)


                oficina = df_oficina["Oficina"].unique()  # Array de valores únicos

                outputi = io.BytesIO()

                wb = xlwt.Workbook()
                ws = wb.add_sheet("Sheet1")

                # Estilo para fecha
                estilo_fecha = xlwt.XFStyle()
                estilo_fecha.num_format_str = "DD/MM/YYYY"

                # Escribir encabezados
                for col_idx, col_name in enumerate(df_oficina.columns):
                    ws.write(0, col_idx, col_name)

                # Columnas H e I → índices 7 y 8
                columnas_fecha_idx = [7, 8]

                # Escribir datos
                for row_idx, row in df_oficina.iterrows():
                    for col_idx, value in enumerate(row):
                        if col_idx in columnas_fecha_idx:
                            ws.write(row_idx + 1, col_idx, value, estilo_fecha)
                        else:
                            ws.write(row_idx + 1, col_idx, value)

                wb.save(outputi)
                outputi.seek(0)

                nombre_archivo_i = f"{opcion}_oficina_{oficina[0]}_GEDO.xls"

                st.download_button(
                    label=f"Descargar planilla de la oficina: {oficina[0]}",
                    data=outputi.getvalue(),
                    file_name=nombre_archivo_i,
                    mime="application/vnd.ms-excel"
                )

            if len(oficinas_nan) != 0:
                st.markdown(
                    "Estas son las oficinas que no pueden ser procesadas porque faltan completar "
                    "la fecha de egreso del cargo para algunas evaluaciones. "
                    "Por favor completar y volver a realizar procedimiento."
                )

                st.markdown(
                    "\n".join(f"- {oficina}" for oficina in oficinas_nan)
                )

        
else:
    
    agree = st.checkbox("Por favor, hacer clic en la casilla si se desea separar el archivo por oficinas.")

    st.subheader(f"📂Archivo de mensualizados del área {opcion}")

    st.markdown("Subir el archivo de mensualizados")

    archivo_1 = st.file_uploader("Seleccionar el archivo de mensualizados", type=["xlsx", "xls"], key="archivo1",accept_multiple_files=False)

    if archivo_1:

        excel_file = pd.ExcelFile(archivo_1)

        nombres_hojas = excel_file.sheet_names
        opciones_hojas = [""] + nombres_hojas
        

        hoja = st.selectbox(
        "Elegir hoja que se quiere procesar.",
        opciones_hojas
        )

        if hoja == "":

            st.subheader("IMPORTANTE❗: seleccionar la hoja antes de continuar.")

        elif hoja == "HOJA":

            st.subheader("Esta hoja no puede ser procesada.")


        else:

            df = pd.read_excel(archivo_1,sheet_name=hoja)
                
            if agree:

                df["Categoría"] = df["Categoría"].replace("NO CATEGORIZADO", 999)

                #Filtrar dataFrame sacando los que tienen en Fecha Egreso Cargo #Enviar nota de designación"
                df = df[df["Fecha Egreso Cargo"] != "Enviar nota de designación"]

                df = borrar_ultimas_columnas(df, 3)

                
                df_oficinas,oficinas_nan = dividir_oficinas(df)

                

                for df_oficina in df_oficinas:
                    

                    df_oficina = df_oficina.reset_index(drop=True)


                    oficina = df_oficina["Oficina"].unique()  # Array de valores únicos

                    outputi = io.BytesIO()

                    wb = xlwt.Workbook()
                    ws = wb.add_sheet("Sheet1")

                    # Estilo para fecha
                    estilo_fecha = xlwt.XFStyle()
                    estilo_fecha.num_format_str = "DD/MM/YYYY"

                    # Escribir encabezados
                    for col_idx, col_name in enumerate(df_oficina.columns):
                        ws.write(0, col_idx, col_name)

                    # Columnas H e I → índices 7 y 8
                    columnas_fecha_idx = [7, 8]

                    # Escribir datos
                    for row_idx, row in df_oficina.iterrows():
                        for col_idx, value in enumerate(row):
                            if col_idx in columnas_fecha_idx:
                                ws.write(row_idx + 1, col_idx, value, estilo_fecha)
                            else:
                                ws.write(row_idx + 1, col_idx, value)

                    wb.save(outputi)
                    outputi.seek(0)

                    nombre_archivo_i = f"{opcion}_oficina_{oficina[0]}_GEDO.xls"

                    st.download_button(
                        label=f"Descargar planilla de la oficina: {oficina[0]}",
                        data=outputi.getvalue(),
                        file_name=nombre_archivo_i,
                        mime="application/vnd.ms-excel"
                    )

                if len(oficinas_nan) != 0:
                    st.divider()
                    st.markdown("Estas son las oficinas que no pueden ser procesadas porque faltan completar la fecha de egreso del cargo para algunas evaluaciones. Por favor completar y volver a realizar procedimiento.")

                    for oficina_nan in oficinas_nan:
                        st.write("""-""" + oficina_nan)
                

            else:

                df_nan = df[df["Fecha Egreso Cargo"].isna()]

                if df_nan.shape[0] != 0: 

                    #Si falta completar alguna de las filas de Fecha Egreso Cargo, directamente no devolvemos el archivo .xls
                    st.markdown("No se puede procesar el documento porque faltan completar la fecha de egreso del cargo para algunas evaluaciones. Por favor completar y volver a realizar procedimiento.")
                
                else:

                    df = df.reset_index(drop=True)


                    outputi = io.BytesIO()

                    wb = xlwt.Workbook()
                    ws = wb.add_sheet("Sheet1")

                    # Estilo para fecha
                    estilo_fecha = xlwt.XFStyle()
                    estilo_fecha.num_format_str = "DD/MM/YYYY"

                    # Escribir encabezados
                    for col_idx, col_name in enumerate(df.columns):
                        ws.write(0, col_idx, col_name)

                    # Columnas H e I → índices 7 y 8
                    columnas_fecha_idx = [7, 8]

                    # Escribir datos
                    for row_idx, row in df.iterrows():
                        for col_idx, value in enumerate(row):
                            if col_idx in columnas_fecha_idx:
                                ws.write(row_idx + 1, col_idx, value, estilo_fecha)
                            else:
                                ws.write(row_idx + 1, col_idx, value)

                    wb.save(outputi)
                    outputi.seek(0)

                    nombre_archivo_i = f"{opcion}_oficina_COMPLETA_GEDO.xls"

                    st.download_button(
                        label=f"Descargar planilla",
                        data=outputi.getvalue(),
                        file_name=nombre_archivo_i,
                        mime="application/vnd.ms-excel"
                    )
