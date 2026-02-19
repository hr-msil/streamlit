import numpy as np
import pandas as pd
import streamlit as st


from productividadesV2 import lectura_archivo_prod

def lectura_archivo_query(nombre_archivo):

    df = pd.read_excel(nombre_archivo)
    df.columns = ["Legajo","Nombre","Particion","Nro. cargo del legajo","Categoria",
                  "Oficina", "Diferencia", "Ausencias"]
    
    return df


def no_permitido(df_query, df_produ):

    #La idea es ver cuales de los legajos no tienen permitido cobrar la prodcutividad
    #Lo que tienen columna difernecia mayor a 1600 y son de la particion MU
    #Los que tienen en columna ausencia un 1

    no_pueden_por_categoria = []
    no_pueden_por_ausencia = []
    repetidos_mu = []
    repetidos_me = []
    repetidos_jo = []
    repetidos_do = []


    for index, row in df_produ.iterrows():

        legajo = row["Legajo"]

        df_query_legajo = df_query[df_query["Legajo"] == legajo]
        df_query_legajo_mu = df_query_legajo[df_query_legajo["Particion"] == "MU"]
        df_query_legajo_me = df_query_legajo[df_query_legajo["Particion"] == "ME"]
        df_query_legajo_jo = df_query_legajo[df_query_legajo["Particion"] == "JO"]
        df_query_legajo_do = df_query_legajo[df_query_legajo["Particion"] == "DO"]

        cant_mu = df_query_legajo_mu.shape[0]
        cant_me = df_query_legajo_me.shape[0]
        cant_jo = df_query_legajo_jo.shape[0]
        cant_do = df_query_legajo_do.shape[0]

        if cant_mu > 2 :

            repetidos_mu.append([df_produ.loc[index,"Apellido y Nombre"],df_produ.loc[index,"Legajo"]])

        if cant_me > 2:

            repetidos_me.append([df_produ.loc[index,"Apellido y Nombre"],df_produ.loc[index,"Legajo"]])

        if cant_jo > 2:

            repetidos_jo.append([df_produ.loc[index,"Apellido y Nombre"],df_produ.loc[index,"Legajo"]])

        if cant_do > 2:

            repetidos_do.append([df_produ.loc[index,"Apellido y Nombre"],df_produ.loc[index,"Legajo"]])

        if df_query_legajo_mu.shape[0] == 1:


            if df_query_legajo.loc[df_query_legajo_mu.index[0], "Diferencia"] >= 1600 :

                no_pueden_por_categoria.append(legajo)
            
            if df_query_legajo.loc[df_query_legajo_mu.index[0],"Ausencias"] == 1:

                no_pueden_por_ausencia.append(legajo)

    return no_pueden_por_categoria, no_pueden_por_ausencia,repetidos_mu,repetidos_me,repetidos_jo,repetidos_do

