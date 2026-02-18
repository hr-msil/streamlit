import pandas as pd
import numpy as np
import streamlit as st

###################
# CHEQUEO LEGAJOS #
###################

#Antes de hacer todo los chequeos, hay que chequear si todos los legajos que manda la oficina son efectivamente,
#de la oficina correspondiente

def tipo_de_fila_cl(fila:pd.Series) -> tuple[int,int]:
    '''
    Dado una fila del documento otorgado por las oficinas de HHEE determinamos si las filas nos dicen el número de la oficina, o en caso conttrario
    los  datos de la persona.
    
    :param fila: Fila correspondiente al dataFrame de HHEE
    :type fila: pd.Series
    :return: Una tupla de int, donde el primer int corresponde al tipo de fila (0:oficina, 1:persona, 2:oficina(año != 2026)), el segundo int corresponde al dato de interés
    :rtype: tuple[int, int]
    '''

    if fila["Legajo"] == "OFICINA: ":

        if fila["Nombre"] == 2026:

            return 0, fila["Oficina"]
        
        else:

            return 2,0 #En caso de que no corresponda al año 2026, retornamos 2 y 0
    
    else:

        return 1, fila["Legajo"]
    
    
def crear_df(df: pd.DataFrame) -> pd.DataFrame:
    '''
    Recorremos todas las filas del dataFrame de legajos por oficina para armarlo con el legajo en la primer columna, y en la segunda columna la oficina la cual 
    pertenece ese legajo, luego convertimos a int el número de oficina y el legajo
    
    :param df: dataFrame de legajos por oficina
    :type df: pd.DataFrame
    :return: dataFrame con los legajos en primer columna y las oficinas en la segunda columna
    :rtype: DataFrame
    '''
    cant_filas = df.shape[0]
    
    legajos = []
    oficinas = []
    oficina_actual = 0

    for i in range(cant_filas):

        fila = df.iloc[i]

        tipo, dato = tipo_de_fila_cl(fila)

        if tipo == 0:
            oficina_actual = dato
        elif tipo == 1:
            legajos.append(dato)
            oficinas.append(oficina_actual)
    #Si la fila es de tipo 2 la ignoramos, no corresponde a este año
        
    df_res = pd.DataFrame({"Legajo": legajos, "Oficina": oficinas})
    df_res = df_res[df_res["Oficina"] != 0] #Filtramos porque los legajos no correspondientes al 2026 me quedaron con numero de oficina 0 (VER)
    #Lo que pasa es que sigo iterando sobre las personas, lo unico que ignoro es el numero de  oficina, como todos los años que no son 2026 están al principio
    # del dataFrame, me quedan los legajos con número de oficina 0
    df_res["Legajo"] = df_res["Legajo"].astype('Int64')
    df_res["Oficina"] = df_res["Oficina"].astype('Int64')

    return df_res
    
def leer_archivo_leg_of(nombre_archivo:str) -> pd.DataFrame:
    '''
    Dado el nombre del archivo correspondiente a la lista de legajos por oficina, lo leemos como dataFrame, nos quedamos con las primeras 3 columnas
    y renombramos las columnas
    
    :param nombre_archivo: Description
    :type nombre_archivo: str
    :return: dataFrame con las primeras 3 columnas correspondientes al legajo, nommbre de la persona y oficina a la que pertenece
    :rtype: DataFrame
    '''
    
    legajos_por_oficina = pd.read_excel(nombre_archivo)

    ultima_fila = legajos_por_oficina.shape[0]

    legajos_por_oficina = legajos_por_oficina.iloc[:ultima_fila - 1,:3] #Sacamos la ultima fila ya que corresponde al total de empleados
    #Ademas agarramos las primeras 3 columnas que son las que nos interesan

    legajos_por_oficina.columns = ["Legajo", "Nombre", "Oficina"] #Las renombramos

    df_res = crear_df(legajos_por_oficina)

    return df_res

def leer_archivo_oficina(nombre_archivo_oficina:str) -> pd.DataFrame:
    #No lo usamos

    df = pd.read_excel(nombre_archivo_oficina)

    df = df.iloc[:,0]
    
    df = df.dropna()

    df = df.astype('Int64')

    return df

def buscar_legajos(legajos_a_buscar: pd.DataFrame, legajos_oficina: pd.DataFrame, oficinas) -> list[int]:
    '''
    Dado el dataFrame de HHEE cargados por la oficina, buscamos que efectivamente, todos los legajos pasados correspondan a esa oficina.
    
    :param legajos_a_buscar: dataFrame de HHEE con columna 'legajo'
    :type legajos_a_buscar: pd.DataFrame
    :param legajos_oficina: dataFrame creado por nosotros con columnas 'Legajo', 'Nombre' y 'Oficina'.
    :type legajos_oficina: pd.DataFrame
    :return: Lista de los legajos que no corresponden a la oficina
    :rtype: list[int]
    '''
    legajos_a_buscar = legajos_a_buscar["legajo"].astype('Int64')

    if oficinas:
        oficinas_int = np.array(oficinas, dtype=int)
        legajos_oficina = legajos_oficina[legajos_oficina["Oficina"].isin(oficinas_int)]

    no_encontrados = []

    legajos = legajos_a_buscar.unique()

    for legajo in legajos:

        print("Legajo a buscar: ", legajo)

        legajo_buscado = legajos_oficina[legajos_oficina["Legajo"] == legajo]
        legajo_buscado = legajo_buscado["Legajo"]

        print("Legajo encontrado: ", legajo_buscado)

        s = legajo_buscado == legajo
        
        if s.any(): #Se usa porque es una serie

            continue

        else:
            
            no_encontrados.append(legajo)

    return no_encontrados

def reportar_legajos(df_hhee_norm, df_legajos_oficina_original, oficinas):
    no_encontrados = buscar_legajos(df_hhee_norm, df_legajos_oficina_original, oficinas)
    # REVISAR!!!!!!!!!!
    if len(no_encontrados) > 0:
        st.write("Estos son los legajos que no pertenecen a la oficina de la planilla:")
        for legajo in no_encontrados:
            st.write("""-""" + str(legajo) + " no pertenece a la/s oficina/s que se incluyen en la planilla de horas extras.")
    else:

        st.write("Los legajos coinciden con el número de la oficina correspondiente")
