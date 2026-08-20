# Toma un archivo, lo convierte a un dataFrame, elimina las ultimas n (parametro que puede ir cambiando) columnas 
# Y que cada oficina sea un archivo por separado

import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers
import xlwt
import re

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
    Precondicion: que el DataFrame pertenezca al área de ambiente y espacio público/desarrollo humano y deportes/educación, cultura y trabajo.
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
        df_oficina_na = df_oficina[(df_oficina["Fecha Egreso Cargo"].isna()) & (df_oficina["Evaluación"] != "Enviar nota de designación") ]

        if df_oficina_na.shape[0] == 0:
        #Si no hay ningun na en niguna de las filas del dataFrame filtrado por oficina, lo carga y lo devuleve, en caso contrario no lo carga
            df_oficinas.append(df_oficina)
        else:
            oficinas_nan.append(oficina)

    return df_oficinas,oficinas_nan

tipos_datos = {
    'Bonificación':str
}

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


#Esto contempla los casos que son norenueva, no renueba, no renueva , etc.
pattern = re.compile(
    r"\bno\s+(?:se\s+)?ren(?:ov|uev|ueb)\w*\b",
    re.IGNORECASE
)


st.subheader("Elegir el área de la cuál se está subiendo el archivo:")

opcion = st.selectbox(
    "Elegir una opción",
    opciones
)

if opcion == "":

    st.subheader("IMPORTANTE❗: seleccionar el área antes de continuar")
        
else:

    st.subheader(f"📂Archivo de mensualizados del área {opcion}")

    st.markdown("Subir el archivo de mensualizados")

    archivo_1 = st.file_uploader("Seleccionar el archivo de mensualizados", type=["xlsx"], key="archivo1",accept_multiple_files=False)

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

            #Filtrar dataFrame sacando los que tienen en Evaluación es distinta a enviar nota de designación

            df_oficinas,oficinas_nan = dividir_oficinas(df)

            

            for df_oficina in df_oficinas:
                

                df_oficina = df_oficina.reset_index(drop=True)
                # Opción B: excluir filas donde se cumple CUALQUIERA de las dos condiciones de fecha
                df_oficina = df_oficina[
                    (df_oficina["Evaluación"] != "Enviar nota de designación") &
                    ~(
                        df_oficina["Fecha Egreso Cargo"].astype(str).str.contains(pattern, na=False) |
                        (df_oficina["Fecha Egreso Cargo"].astype(str) == "Art. 32")
                    )
                ]
                if df_oficina.shape[0] == 0: continue
                oficina = df_oficina["Oficina"].unique()  # Array de valores únicos
                df_oficina = borrar_ultimas_columnas(df_oficina,5)
                df_oficina = df_oficina.reset_index(drop=True)
                with pd.option_context('future.no_silent_downcasting', True):
                    df_oficina = df_oficina.fillna('').infer_objects(copy=False)
                

                fechas_distintas = df_oficina["Fecha Egreso Cargo"].unique()

                for fecha in fechas_distintas:
                    df_oficina_fecha = df_oficina[df_oficina["Fecha Egreso Cargo"] == fecha].reset_index(drop=True)
                    outputi = io.BytesIO()

                    wb = xlwt.Workbook()
                    ws = wb.add_sheet("Sheet1")

                    # Estilo para fecha
                    estilo_fecha = xlwt.XFStyle()
                    estilo_fecha.num_format_str = "DD/MM/YYYY"

                    # Escribir encabezados
                    for col_idx, col_name in enumerate(df_oficina_fecha.columns):
                        ws.write(0, col_idx, col_name)

                    # Columnas H e I → índices 7 y 8
                    columnas_fecha_idx = [7, 8]

                    # Escribir datos
                    for row_idx, row in df_oficina_fecha.iterrows():
                        for col_idx, value in enumerate(row):
                            if col_idx in columnas_fecha_idx:
                                ws.write(row_idx + 1, col_idx, value, estilo_fecha)
                            else:
                                ws.write(row_idx + 1, col_idx, value)

                    for col_idx, col_name in enumerate(df_oficina_fecha.columns):
                    
                        # Largo del encabezado
                        max_length = len(str(col_name))
                    
                        # Largo máximo del contenido
                        for value in df_oficina_fecha.iloc[:, col_idx]:
                            if value is not None:
                                length = len(str(value))
                                if length > max_length:
                                    max_length = length
                    
                            # Ajuste (256 es unidad de xlwt, +2 da un pequeño margen)
                        ws.col(col_idx).width = 256 * (max_length + 2)

                    
                    wb.save(outputi)
                    outputi.seek(0)
                    fecha_normalizada = fecha.date().strftime("%d-%m-%Y")
                    nombre_archivo_i = f"{opcion}_oficina_{oficina[0]}_{fecha_normalizada}_GEDO.xls"

                    if nombre_archivo_i not in st.session_state:
                        st.session_state[nombre_archivo_i] = False

                    def marcar_como_descargado(nombre_archivo):
                        st.session_state[nombre_archivo] = True

                    if st.session_state[nombre_archivo_i]:
                        texto_btn = f"✅ Oficina {oficina[0]} con fecha fin {fecha_normalizada}"
                        tipo_btn = "secondary"  # Se vuelve un botón gris/transparente
                    else:
                        texto_btn = f"⬇️ Oficina {oficina[0]} con fecha fin {fecha_normalizada}"
                        tipo_btn = "primary"

                    st.download_button(
                        label=texto_btn,
                        data=outputi.getvalue(),
                        file_name=nombre_archivo_i,
                        type = tipo_btn,
                        on_click = marcar_como_descargado,
                        args=(nombre_archivo_i,),
                        mime = "application/vnd.ms-excel"
                    )

            if len(oficinas_nan) != 0:

                st.divider()

                st.markdown("Estas son las oficinas que no pueden ser procesadas porque faltan completar la fecha de egreso del cargo para algunas evaluaciones. Por favor completar y volver a realizar procedimiento.")

                for oficina_nan in oficinas_nan:
                    st.write("""-""",oficina_nan)