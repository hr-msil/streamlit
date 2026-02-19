import numpy as np
import pandas as pd
import streamlit as st
import re



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
    df_prod.columns = ["Legajo","Cargo","Apellido y Nombre","Nula1","Leyenda","Inicio","Nula2","Fin","Cantidad","Base","Importe","Porcentaje","Indicativo"]

    return df_prod

def lectura_archivo_dec(archivo_dec: str,decreto, nombre_original):
    '''
    Convierte en DataFrame los archivos subidos con extensión .csv. Si al convertirlo tiene más de 4 columnas, se toman las primeras cuatro.
    
    :param archivo_dec: String. Nombre del archivo .csv
    :type archivo_dec: Str.
    :return DataFrame.
    '''
    df_decreto = pd.read_csv(archivo_dec,header=None)

    if df_decreto.shape[1] > 4: #Si el archivo csv tiene más de 4 columnas, me quedo con las primeras 4

        df_decreto = df_decreto.iloc[:,:4]
    
    df_decreto.columns = ["Legajo", "Nula", "Nula2", "Importe"] #Renombro las columnas
    df_decreto = df_decreto.dropna() # Elimino las filas con algún Nan
    df_decreto["Legajo"] = df_decreto["Legajo"].astype('Int64') #Cambio tipo del Legajo para que tipe con df_prod
    df_decreto["Decreto"] = decreto #Agrego la columna de decreto
    df_decreto["Nombre original"] = nombre_original #Agrego la columna de decreto del nombre original
    df_decreto.sort_values(by = "Legajo") #Ordeno según legajo
    df_decreto["Decreto"] = df_decreto["Decreto"].astype(str)
    df_decreto['Importe'] = df_decreto['Importe'].astype(float).round(2)

    return df_decreto

#--------- LIMPIEZA DE LA COLUMNA DECRETOS -----------

def limpieza_decreto(df: pd.DataFrame) -> dict:
    '''
    Modifica la columna "Leyenda" del dataFrame correspondiente al archivo subido del sistema para que el decreto quede de la forma num/num. Y le agrega
    al dataFrame una columna que sea el nombre original del decreto. Para los decretos donde no se pueda sacar apropiadamente el decreto de la forma num/num, ponemos
    como nombre original : ""
    
    :param df: dataFrame correspondiente a la concatenación de archivos csv
    :type df: pd.DataFrame
    :return dict: diccionario con un decreto, como clave y un conjunto asociado correspondiente a los decretos sin limpiar
    '''

    #TODO sacar el diccionario, al final no lo uso
    patron = r"DTO\.?\s*(\d+/\d+)"


    decretos = []
    cant_prod = df.shape[0]

    for i in range(cant_prod):

        leyenda = df.iloc[i]["Leyenda"]
        
        if pd.isna(leyenda):
            # + 1 por el index de python, 1 por el encabezado de Excel
            decretos.append("")
            #st.write("La fila ", i + 2, " no tiene leyenda detallada." ) TODO me queda ver como imprimir bien esto, podria agregarle una columna de 
            #que archivo fue sacado
            
        else:
            match = re.search(patron, leyenda)
            if match: #Si cuando hacemos el split nos queda de longitud mayor o igual a dos, indexamos en 1 para obtenerlo,
                # en cas contrario, decimos que no podemos extraer el decreto

                decreto_prod = match.group(1)
                decretos.append(leyenda)
                df.loc[i,"Leyenda"] = decreto_prod
                


            else:
                decretos.append("")

    df.sort_values(by="Legajo")
    df["Leyenda"] = df["Leyenda"].astype(str)
    df['Importe'] = df['Importe'].astype(float).round(2) #Redondeo a dos decimales así puede matchear con mejor exactitud

    df["Nombre original"] = decretos


    

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

def comparar(
    df_origen: pd.DataFrame,
    df_destino: pd.DataFrame,
    legajos: list,
    importes: list,
    decretos_originales: list,
    decretos: list,
    decreto: str
) -> None:

    comparacion = df_origen.merge(
        df_destino[["Legajo", "Importe"]],
        on=["Legajo", "Importe"],
        how="left",
        indicator=True
    )

    diferencias = comparacion[comparacion["_merge"] == "left_only"]

    for _, row in diferencias.iterrows():
        legajos.append(row["Legajo"])
        importes.append(row["Importe"])
        decretos_originales.append(row["Nombre original"])
        decretos.append(decreto)








       