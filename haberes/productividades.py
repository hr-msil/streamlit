import numpy as np
import pandas as pd
import streamlit as st


#--------- LECTURA de archivos ---------------------
#PRODUCTIVIDADES -> EXCEL, DECRETO -> CSV

def lectura_archivo_prod(archivo_prod:str):
    '''
    Convierte en DataFrame el excel subido del sistema. Ordena los legajos por orden numérico.
    
    :param archivo_prod: Nombre del archivo .xlsx
    :type archivo_prod: Str
    :return DataFrame. 
    '''

    df_prod = pd.read_excel(archivo_prod)
    df_prod.sort_values(by = "Legajo") #Ordeno según legajo

    return df_prod

def lectura_archivo_dec(archivo_dec: str,decreto, nombre_original):
    '''
    Convierte en DataFrame los archivos subidos con extensión .csv. Si al convertirlo tiene más de 4 columnas, se toman las primeras cuatro.
    
    :param archivo_dec: String. Nombre del archivo .csv
    :type archivo_dec: Str.
    :return DataFrame.
    '''

    #Agregar columna decreto y lista con nombre originales

    df_decreto = pd.read_csv(archivo_dec,header=None)

    if df_decreto.shape[1] > 4:

        df_decreto = df_decreto.iloc[:,:4]

    df_decreto.columns = ["Legajo", "Nula", "Nula2", "Importe"] #Renombro las columnas
    df_decreto = df_decreto.dropna() # Elimino las filas con algún Nan
    df_decreto["Legajo"] = df_decreto["Legajo"].astype('Int64') #Cambio tipo del Legajo para que tipe con df_prod
    df_decreto["Decreto"] = decreto
    df_decreto["Nombre original"] = nombre_original
    df_decreto.sort_values(by = "Legajo") #Ordeno según legajo

    return df_decreto

#--------- LIMPIEZA DE LA COLUMNA DECRETOS -----------



def limpieza_decreto(df: pd.DataFrame) -> dict:
    '''
    Modifica la columna "Leyenda" del dataFrame correspondiente al archivo subido del sistema para que el decreto quede de la forma num/num.
    
    :param df: Description
    :type df: pd.DataFrame
    :return dict: diccionario con decreto y un conjunto con todos sus diccionarios
    '''
    cant_prod = df.shape[0]
    dicc_decretos = dict()

    for i in range(cant_prod):

        leyenda = df.iloc[i]["Leyenda"]
        if pd.isna(leyenda):
            # + 1 por el index de python, 1 por el encabezado de Excel
            st.write("La fila ", i + 2, " no tiene leyenda detallada." )
            st.divider()
        else:
            decreto_prod_split = leyenda.split(" ")
            if len(decreto_prod_split) > 1:
                decreto_prod = decreto_prod_split[1]
                #Agregar un if que si al hacer el split es menor a uno que tire un error y que se reporte
                df.loc[i,"Leyenda"] = decreto_prod

                if decreto_prod not in dicc_decretos:
                    dicc_decretos[decreto_prod] = set()

                dicc_decretos[decreto_prod].add(leyenda)

            else:
                st.write("No se pudo obtener el decreto en la fila", i + 2)
                st.divider()

    df["Leyenda"] = df["Leyenda"].astype('str')

    return dicc_decretos



#--------- OBTENGO nombre del decreto, según nombre del archivo-----

def obtener_decreto(nombre_archivo: str) -> str:
    '''
    Dado el nombre del archivo, le quita la extensión .csv. Si el decreto está separado por "-", lo convierte a la forma num/num.
    
    :param nombre_archivo: Description
    :type nombre_archivo: str
    :return: Devuelve  el nombre del archivo sin la extensión.
    :rtype: str
    '''
    decreto = nombre_archivo.split(".")[0] #Con esto saco la extension .csv
    decreto = decreto.split(" ")
    decreto = decreto[0].split("-")[0] + "/" +decreto[0].split("-")[1] #Lo renombro a tipo num/año

    return decreto




#--------- FUNCION PRINCIPAL -------------------------

def comparar(df_prod_dec: pd.DataFrame, df_dec: pd.DataFrame,legajos:list, importes:list) -> None:
    '''
    Toma el df de productividades filtrado por decreto y se fija si encuentra o no el monto correspondiente a cada legajo

    :param df_prod_dec: DataFrame de productividades filtrado por decreto.
    :type df_prod_dec: pd.DataFrame
    :param df_dec: DataFrame del decreto correspondiente
    :type df_dec: pd.DataFrame
    :param nombre_original: Description
    :type nombre_original: str
    '''

    cant_prod_dec = df_prod_dec.shape[0]
    cant_dec = df_dec.shape[0]

    for i in range(cant_prod_dec):

        legajo = df_prod_dec.iloc[i]["Legajo"]
        importe = df_prod_dec.iloc[i]["Importe"]

        #Busco si existe la fila en el dataFrame correspondiente al decreto

        existe_en_csv = False

        for j in range(cant_dec):

            legajo_dec = df_dec.iloc[j]["Legajo"]
            importe_dec = df_dec.iloc[j]["Importe"]

            if importe_dec == importe and legajo_dec == legajo:
                
                existe_en_csv = True

        if existe_en_csv == False: #Agregar a un dataFrame global que sea el de diferencias

            legajos.append(legajo)
            importes.append(importe)

            



#--------- STREAMLIT -------------------------------

st.title("📝 PRODUCTIVIDADES")

st.divider()

tab1,tab2 = st.tabs(["Subir archivos", "Ver resultados"])

with tab1:

    st.markdown("Subir los archivos de productividades correspondientes a lo arrojado por el sistema")

    archivos_excel = st.file_uploader("Seleccionar archivo", type = "xlsx",key = "productividades",accept_multiple_files=True)
    #Acepta multiples, concatenarlos en ese caso (asumimos que las columnas y los nombres son iguales)
    st.markdown("Subir los archivos .csv que se quieren comparar")

    archivos_csv = st.file_uploader("Seleccionar archivo", type = "csv", key = "decreto",accept_multiple_files=True)
    #Acepta multiples, concatenarlos en ese caso(las columnas y sus nombres son iguales porque se 
    #procesan todos en la misma función)

#-------- LECTURA Y LIMPIEZA de los archivos --------

with tab2:

    decretos_originales_csv = []
    decretos_csv = []

    if archivos_excel and archivos_csv:

        dfs_excel = []

        for archivo_prod in archivos_excel:
            #Iteramos sobre todos los archivos de excel subidos a lo que devuelve el sistema de productividades

            df_prod = lectura_archivo_prod(archivo_prod)
            dicc_decreto = limpieza_decreto(df_prod)
            dfs_excel.append(df_prod)
            

        df_productividades = pd.concat(dfs_excel,ignore_index=True)

        st.write(df_productividades)
        
        sin_diferencias = [] #Lista para guardar los archivos sin diferencias
        dfs_inconsistencias = [] #Lista para guardar los dataFrame con inconsistencias de csv a excel
        nombres_inconsistencias = [] #Nombres de los archivos con inconsistencias
        dfs_csv = []
        legajos_csv = []
        importes_csv = []


        for archivo in archivos_csv:

            #--------- CREO df_diferencias ------------------------
            # Acá vamos a guardar todas las productividades correspondientes al archivo que se carga del SISTEMA
            # que no se encuentren en el csv correspondiente al decreto
            #hago un archivo de diferencias por decreto

      
            legajos = []
            importes = []

            nombre_original = archivo.name.split(".")[0]
            decreto = obtener_decreto(archivo.name)
            
            df_dec = lectura_archivo_dec(archivo,decreto, nombre_original)

            df_prod_excel = df_productividades[df_productividades["Leyenda"] == decreto]
            #Filtamos el df correspondiente al Excel para que coincida con el decreto que estamos trabajandoo
            dfs_csv.append(df_dec)
            comparar(df_dec,df_prod_excel,legajos_csv,importes_csv)

            df_diferencias = pd.DataFrame({"Legajo": legajos_csv, "Importe": importes_csv})

            if df_diferencias.shape[0] == 0:

                sin_diferencias.append(nombre_original)

            else:

                dfs_inconsistencias.append(df_diferencias)
                nombres_inconsistencias.append(nombre_original)

        df_final = pd.concat(dfs_csv, ignore_index=True) #df concatenado por decretos
        st.write(df_final)

        decretos_unicos = df_final["Decreto"].unique()
        st.write("estos son los decretos unicos:")
        st.write(decretos_unicos)

        legajos_excel = []
        importes_excel = []

        for decreto_csv in decretos_unicos:

            decreto_csv_filt = df_final[df_final["Decreto"] == decreto_csv]
            decreto_excel_filt = df_productividades[df_productividades["Leyenda"] == decreto_csv]
            


            comparar(decreto_excel_filt,decreto_csv_filt,legajos_excel,importes_excel)
        df_diferencias = pd.DataFrame({"Legajo": legajos_excel, "Importe": importes_excel})
        st.write("Este es el dataFrame de diferencias correspondientes de excel a csv")
        st.write(df_diferencias)


        

        if len(sin_diferencias) != 0:

            st.write("No se encontraron inconsistencias en los siguientes archivos csv: ")
            
            for no_diferencia in sin_diferencias: 

                st.write(""" - """, no_diferencia)

        if(len(nombres_inconsistencias) > 0):

            
            tabs = st.tabs(nombres_inconsistencias)

            for i, df in enumerate(dfs_inconsistencias):

                with tabs[i]:

                    st.write("En el archivo: ",nombres_inconsistencias[i]," no se encontraron estos importes para estos legajos:  ")

                    st.write(df)








