# Toma un archivo, lo convierte a un dataFrame, elimina las ultimas n (parametro que puede ir cambiando) columnas 
# Y que cada oficina sea un archivo por separado

import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import load_workbook
from openpyxl.styles import numbers
import re
from openpyxl.utils import get_column_letter


#---------- Funciones principales -----------

FORMATOS = {
    'Fecha Ingreso Cargo': 'd/m/yyyy',
    'Fecha Egreso Cargo': 'd/m/yyyy',
    'Bonificación': 'str'
}

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
                df_oficina = borrar_ultimas_columnas(df_oficina,7)
                df_oficina = df_oficina.reset_index(drop=True)
                with pd.option_context('future.no_silent_downcasting', True):
                    df_oficina = df_oficina.fillna('').infer_objects(copy=False)
                
                fechas_distintas = df_oficina["Fecha Egreso Cargo"].unique()
                for fecha in fechas_distintas:

                    df_oficina_fecha = df_oficina[
                        df_oficina["Fecha Egreso Cargo"] == fecha
                    ].reset_index(drop=True)

                    outputi = io.BytesIO()

                    with pd.ExcelWriter(outputi, engine='openpyxl') as writer:

                        df_oficina_fecha.to_excel(
                            writer,
                            index=False,
                            sheet_name='Sheet1'
                        )

                        ws = writer.sheets['Sheet1']

                        # Aplicar formatos
                        for columna, formato in FORMATOS.items():

                            if columna in df_oficina_fecha.columns:

                                col_idx = df_oficina_fecha.columns.get_loc(columna) + 1

                                for fila in range(2, len(df_oficina_fecha) + 2):
                                    ws.cell(
                                        row=fila,
                                        column=col_idx
                                    ).number_format = formato

                        # Ajustar ancho de columnas
                        for col_idx, col_name in enumerate(df_oficina_fecha.columns):

                            max_length = len(str(col_name))

                            for value in df_oficina_fecha.iloc[:, col_idx]:

                                if pd.notna(value):
                                    max_length = max(max_length, len(str(value)))

                            ws.column_dimensions[
                                get_column_letter(col_idx + 1)
                            ].width = max_length + 2

                    
                    excel_bytes = outputi.getvalue()

                    # Verificar que sea un Excel válido
                    from openpyxl import load_workbook

                    test = load_workbook(io.BytesIO(excel_bytes))

                    print(f"Archivo válido: {nombre_archivo_i if 'nombre_archivo_i' in locals() else 'Pendiente'}")
                    print(f"Tamaño: {len(excel_bytes)} bytes")

                    # Nombre del archivo
                    fecha_normalizada = fecha.strftime("%d-%m-%Y")

                    nombre_archivo_i = (
                        f"{opcion}_oficina_{oficina[0]}_"
                        f"{fecha_normalizada}_GDSI.xlsx"
                    )

                    if nombre_archivo_i not in st.session_state:
                        st.session_state[nombre_archivo_i] = False

                    def marcar_como_descargado(nombre_archivo):
                        st.session_state[nombre_archivo] = True

                    if st.session_state[nombre_archivo_i]:
                        texto_btn = f"✅ Oficina {oficina[0]} con fecha fin {fecha_normalizada}"
                        tipo_btn = "secondary"
                    else:
                        texto_btn = f"⬇️ Oficina {oficina[0]} con fecha fin {fecha_normalizada}"
                        tipo_btn = "primary"

                    st.download_button(
                        label=texto_btn,
                        data=excel_bytes,
                        file_name=nombre_archivo_i,
                        type=tipo_btn,
                        on_click=marcar_como_descargado,
                        args=(nombre_archivo_i,),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            if len(oficinas_nan) != 0:

                st.divider()

                st.markdown("Estas son las oficinas que no pueden ser procesadas porque faltan completar la fecha de egreso del cargo para algunas evaluaciones. Por favor completar y volver a realizar procedimiento.")

                for oficina_nan in oficinas_nan:
                    st.write("""-""",oficina_nan)
