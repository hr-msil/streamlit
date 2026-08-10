import numpy as np
import pandas as pd
import streamlit as st
import re

from productividadesV2 import comparar
from productividadesV2 import lectura_archivo_dec
from productividadesV2 import lectura_archivo_prod
from productividadesV2 import limpieza_decreto
from productividadesV2 import obtener_decreto

from cobranProductividades import lectura_archivo_query
from cobranProductividades import no_permitido


#######################
#       STREAMLIT     #
#######################

#--------- STREAMLIT -------------------------------

st.title("📝 PRODUCTIVIDADES")

st.divider()

tab1,tab2,tab3 = st.tabs(["Subir archivos", "Sin comparar","Resultados"])

with tab1:

    #Si dan clic al siguiente checkbox la comparación se hará bajo la suposición de que todos los decretos que se encuentran en el excel tienen 
    #un archivo csv corrrespondiente a ese decreto, en caso contrario avisa que ningún csv correspondiente al decreto fue subido
    agree = st.checkbox("Comparar todos los decretos del archivo excel de productividades")

    st.markdown("Subir los archivos de productividades arrojados por el sistema")

    archivos_excel = st.file_uploader("Seleccionar archivo", type = "xls",key = "productividades",accept_multiple_files=True)
    #Acepta multiples, concatenarlos en ese caso (asumimos que las columnas y los nombres son iguales)
    st.markdown("Subir los archivos .csv que se quieren comparar")

    archivos_csv = st.file_uploader("Seleccionar archivo", type = "csv", key = "decreto",accept_multiple_files=True)
    #Acepta multiples, concatenarlos en ese caso(las columnas y sus nombres son iguales porque se 
    #procesan todos en la misma función)

#-------- LECTURA Y LIMPIEZA de los archivos --------

with tab2:

    #bien, tengo los archivos de excel y csv, primero voy a leer los de excel y concatenarlos

    dicc_decretos = dict() #Variable global para guardar en un diccionario el numero de decreto y todas las variantes
    #que aparezcan en el archivo excel de productividades

    if archivos_excel and archivos_csv:

        dfs_excel = []
        dfs_csv = []
        
        #Como se aceptan múltiples archivos, los concatenamos todos en un archivo final, mismo con los archivos csv
        for archivo_excel in archivos_excel:

            df_excel_i = lectura_archivo_prod(archivo_excel)
            limpieza_decreto(df_excel_i)
            dfs_excel.append(df_excel_i)
        
        df_excel_final = pd.concat(dfs_excel,ignore_index = True)
        df_excel_sin_procesar = df_excel_final[df_excel_final["Nombre original"] == ""] #Filtramos el dataFrame por los decretos que no pudimos obtener
        st.write("Esta es la lista de productividades que no pudo ser procesada debido al formato de las leyendas")
        df_excel_sin_procesar = df_excel_sin_procesar.drop(columns = ["Nula1","Inicio","Nula2","Fin","Cantidad","Base","Porcentaje","Indicativo","Nombre original"])
        st.write(df_excel_sin_procesar)
        
        df_excel_final = df_excel_final[df_excel_final["Nombre original"] != ""]


        for archivo_csv in archivos_csv:

            #Acá hago la lectura del decreto y del nombre original del decreto
            decreto_original = archivo_csv.name
            decreto_split = decreto_original.split(".")
            decreto_original = decreto_split[0] #Este es el nombre original del arhivo, le sacamos la extensión .csv
            decreto_original_split = decreto_original.split(" ")#Este es el numero de la forma num-num
            decreto_con_guion = decreto_original_split[0] #Me quedo con el numero de decreto nada más
            decreto_con_guion = decreto_con_guion.split("-")
            decreto = decreto_con_guion[0] + "/" + decreto_con_guion[1]
            

            df_csv_i = lectura_archivo_dec(archivo_csv,decreto,decreto_original)
            dfs_csv.append(df_csv_i)

        df_csv_final = pd.concat(dfs_csv, ignore_index = True)
        #st.write(df_csv_final)

        dfs_dif_excel = []
        dfs_dif_csv = []
        diferencias = set()
        legajos_comp_excel = [] #Lista de legajos con la comparación de excel a csv
        importes_comp_excel = [] #Lista de importes con la comparación de excel a csv
        legajos_comp_csv = [] #Lista de legajos con la comparación de csv a excel
        importes_comp_csv = [] #Lista de importes con la comparación de csv a excel
        decreto_original_excel = []
        decreto_original_csv = []
        decretos_comp_excel = []
        decretos_comp_csv = []
        no_existe_csv = []


        if agree:
            #Si se piden procesar todos los decretos procedemos de la siguiente manera:
            #obtenemos todos los decretos únicos del excel e iteramos sobre ellos para  hacer la comparación de excel a csv y de csv a excel

            decretos_unicos_excel = df_excel_final["Leyenda"].unique()

            for decreto_excel in decretos_unicos_excel:
            
                df_csv_decreto = df_csv_final[df_csv_final["Decreto"] == decreto_excel]
                df_excel_decreto = df_excel_final[df_excel_final["Leyenda"] == decreto_excel]

                if df_csv_decreto.shape[0] == 0:
                    no_existe_csv.append(decreto_excel)
                    #st.write(f"No fue cargado ningún csv correspondiente al decreto {decreto_excel}")

                else:

                    comparar(df_excel_decreto, df_csv_decreto,legajos_comp_excel, importes_comp_excel,decreto_original_excel,decretos_comp_excel,decreto_excel) #Comparación Excel a CSV

                    comparar(df_csv_decreto,df_excel_decreto,legajos_comp_csv,importes_comp_csv,decreto_original_csv,decretos_comp_csv,decreto_excel) #Comparación CSV a Excel

            df_diferencias_excel = pd.DataFrame({"Legajo": legajos_comp_excel, "Importe": importes_comp_excel,"Decreto":decretos_comp_excel,"Decreto original":decreto_original_excel}) 
            df_diferencias_csv = pd.DataFrame({"Legajo": legajos_comp_csv, "Importe": importes_comp_csv,"Decreto":decretos_comp_csv,"Decreto original":decreto_original_csv})

            st.write("Estos son los legajos e importes que no pudieron ser matcheados debido a que su CSV no fue subido:")

            df_no_subidos = df_excel_final[df_excel_final["Leyenda"].isin(no_existe_csv)]
            df_no_subidos = df_no_subidos.drop(columns = ["Nula1","Inicio","Nula2","Fin","Cantidad","Base","Porcentaje","Indicativo","Nombre original"])

            st.write(df_no_subidos)
                    

                    #df_diferencias_csv = pd.DataFrame({"Legajo": legajos_comp_csv, "Importe" : importes_comp_csv,"Decreto original":decreto_original_csv})

                    #df_diferencias_excel = pd.DataFrame({"Legajo": legajos_comp_excel, "Importe" : importes_comp_excel,"Decreto original":decreto_original_excel})

        else:

            decretos_unicos_csv = df_csv_final["Decreto"].unique()

            for decreto_csv in decretos_unicos_csv:
                
                df_csv_decreto = df_csv_final[df_csv_final["Decreto"] == decreto_csv]
                df_excel_decreto = df_excel_final[df_excel_final["Leyenda"] == decreto_csv]

                comparar(df_excel_decreto, df_csv_decreto,legajos_comp_excel, importes_comp_excel,decreto_original_excel,decretos_comp_excel,decreto_csv)

                comparar(df_csv_decreto,df_excel_decreto,legajos_comp_csv,importes_comp_csv,decreto_original_csv,decretos_comp_csv,decreto_csv)
                
            df_diferencias_excel = pd.DataFrame({"Legajo": legajos_comp_excel, "Importe": importes_comp_excel,"Decreto":decretos_comp_excel,"Decreto original":decreto_original_excel}) 
            df_diferencias_csv = pd.DataFrame({"Legajo": legajos_comp_csv, "Importe": importes_comp_csv,"Decreto":decretos_comp_csv,"Decreto original":decreto_original_csv})

        

        df_diferencias_excel.columns = ["Legajo","Importe","Decreto","Leyenda original"]
        df_diferencias_csv.columns = ["Legajo","Importe","Decreto","Nombre archivo original"]

        with tab3:

            diferencias = pd.concat([df_diferencias_excel["Decreto"], df_diferencias_csv["Decreto"]]).unique().tolist()

            if(len(diferencias)>0):

                tabs = st.tabs(diferencias)

                for i in range(len(diferencias)):

                    with tabs[i]:

                        df_diferencias_excel_dec = df_diferencias_excel[df_diferencias_excel["Decreto"] == diferencias[i]]
                        df_diferencias_csv_dec = df_diferencias_csv[df_diferencias_csv["Decreto"] == diferencias[i]]

                        if df_diferencias_excel_dec.shape[0] != 0:
                            st.write("Los siguientes importes de la planilla del sistema no fueron encontrados en ninguno de los CSVs subidos: ")

                            styler_excel = df_diferencias_excel_dec.style.format({
                            'Importe': lambda x: '{:,.2f}'.format(x).replace(',', 'X').replace('.', ',').replace('X', '.')
                            })
                            st.write(styler_excel)
                        
                        if df_diferencias_csv_dec.shape[0] != 0:
                            st.write("Los siguientes importes de los CSVs subidos no fueron encontrados en ninguna de las planillas del sistema subidas: ")
                            styler_csv = df_diferencias_csv_dec.style.format({
                            'Importe': lambda x: '{:,.2f}'.format(x).replace(',', 'X').replace('.', ',').replace('X', '.')
                            })
                            st.write(styler_csv)

            else:
                st.markdown("No se encontraron diferencias")