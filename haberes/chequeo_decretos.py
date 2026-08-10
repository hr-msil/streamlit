import pandas as pd
import sys
import os


nombre_csv_1 = "75-26.csv"
nombre_csv_2 = "87-26.csv"
nombre_csv_3 = "88-26.csv"



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

def listadoPorEmpleados(empleados_por_ofi_excel :str) -> pd.DataFrame:
    '''
    Toma el excel de empleados por oficina y lo procesa para que quede legajo -> oficina
    '''

    df = pd.read_excel(empleados_por_ofi_excel,header=None)
    df = df[df[0].apply(lambda x: str(x).startswith("OFICINA: ") or str(x).replace(".0", "").isnumeric())]
    
    df = df.iloc[:,:3] #Sacamos la ultima fila ya que corresponde al total de empleados
    #Ademas agarramos las primeras 3 columnas que son las que nos interesan

    df.columns = ["Legajo", "Nombre", "Oficina"] #Las renombramos
    df_res = crear_df(df)
    return df_res

def chequear_decreto_oficina_empleado(empleados_por_ofi,archivo_csv,decreto_por_oficina,df_csv):
    
    decreto = archivo_csv.name.split(".csv")[0]

    decreto_por_oficina.columns = ["decreto","oficina"]
    oficinas_decreto = decreto_por_oficina[decreto_por_oficina["decreto"] == decreto]["oficina"].tolist()
    
    legajos_csv =df_csv.iloc[:,0].tolist() #legajos que están en el csv

    legajos_oficina = empleados_por_ofi[empleados_por_ofi["Oficina"].isin(oficinas_decreto)]["Legajo"].tolist() #legajos que se encuentran en la oficina que pueden cobrar esos legajos
    

    legajos_no_encontrados = []

    for legajo in legajos_csv:
        if legajo in legajos_oficina:
            continue
        else:
            legajos_no_encontrados.append(legajo)
    
    return legajos_no_encontrados




    
    

