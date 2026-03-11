import pandas as pd
import streamlit as st
import re
import difflib
import calendar
import json
import numpy as np

from pathlib import Path
from datetime import datetime
from helpers import lista_a_string
from collections import defaultdict
from dateutil.relativedelta import relativedelta


def obtener_dias_feriados(anio: int, mes: int) -> list[int]:
    BASE_DIR = Path(__file__).resolve().parent.parent
    ruta_feriados = BASE_DIR / "feriados.txt"

    with open(ruta_feriados, encoding="utf-8") as f: 
        texto = f.read().replace("NaN", "null")
    
    feriados = json.loads(texto)
    dias = []

    for fdo in feriados:
        fecha = datetime.strptime(fdo["fecha"], "%Y-%m-%d")

        if fecha.year == anio and fecha.month == mes:
            dias.append(fecha.day)

    return sorted(dias)

def obtener_fines_de_semana(anio: int, mes: int):
    primer_dia_mes, cant_dias = calendar.monthrange(anio,mes)
    dias = []
    dia = primer_dia_mes
    for nro_dia in range(1,cant_dias+1):
        if dia % 7 >= 5:
            dias.append(nro_dia)
        dia += 1

    return dias

def buscar_idx_encabezado(df_planilla: pd.DataFrame,
                          keyword: str,
                          nombre_tabla: str = "",
                          col_a_buscar: int = 0) -> int:
    '''
    El df es una lectura de dataframe de un archivo excel o csv que está muy mal formateado.
    Por ejemplo: la tabla del mismo empiece en otra fila distinta de la primera, que tenga anotaciones
    que no correspondan en otras celdas, etc.
    Asume que la keyword que es el encabezado de la tabla está en la columna nro. col_a_buscar del df que es 0-index based
    Devuelve el indice idx_encabezado de donde está el encabezado de la tabla.
    '''
    idx_encabezado = None
    
    keyword = keyword.lower()

    #Si la palabra clave está en el encabezado del df_planilla
    if keyword in df_planilla.columns[col_a_buscar].lower():
        return -1

    #Si la tabla empieza en otra fila distinta de la de 1
    cant_filas = df_planilla.shape[0]
    for i in range(cant_filas):
        cell = str(df_planilla.iat[i, col_a_buscar]).lower()
        if keyword in cell: # LIKE '%keyword%'
            idx_encabezado = i
            break

    if idx_encabezado is None:
        st.error(
            f"❌ No se encontró la tabla para el archivo {nombre_tabla}\n\n"
            f"Se esperaba encontrar un encabezado que contenga la palabra '{keyword}' "
            f"en la columna {col_a_buscar + 1} del archivo."
        )
        st.stop()

    return idx_encabezado

#NO COPIAR
def obtener_hoja_planilla(excel) -> str:
    
    """
    Lee un archivo Excel (.xls o .xlsx) y analiza hoja por hoja
    para hallar la PLANILLA DE HORAS EXTRAS
    El criterio es si tiene más de 30 columnas (que debería ser el caso si tiene
    columnas para cada día del mes)

    Errores que puede levantar:
        -de lectura de archivo u hoja de excel
        -que el excel tenga 1 hoja o más de 2 

    Retorna: el nombre de la hoja que posee la planilla a procesar
    """
    
    # Detecta tipo de archivo automáticamente
    try:
        excel_file = pd.ExcelFile(excel)
    except Exception as e:
        raise ValueError(f"Error al abrir el archivo: {e}")
    
    nombres_hojas = excel_file.sheet_names
    es_planilla = []

    # Iterar por cada hoja
    for hoja in nombres_hojas:
        try:
            df = pd.read_excel(excel, sheet_name=hoja, header=None, dtype=str)
            # Es planilla si tiene más de 30 columnas
            tiene_forma_planilla = df.shape[1] > 30
            es_planilla.append(tiene_forma_planilla)
        except Exception as e:
            st.error(f"⚠️ Error leyendo hoja '{hoja}': {e}. Reportar con el equipo de desarrollo.")
            st.stop()

    if sum(es_planilla) > 1:
        st.error("No se pudo determinar la hoja en la que se encuentra la planilla a procesar. Tratá que la hoja solo tenga dicha planilla.")
        st.stop()
    
    idx_hoja_planilla = es_planilla.index(True)
    hoja_planilla = nombres_hojas[idx_hoja_planilla]

    return hoja_planilla


def normalizar_planilla_viaticos(planilla: pd.DataFrame) -> pd.DataFrame:
    '''
    La planilla de viaticos tiene un mal formato al leerla como dataframe. La procesamos para tener los datos que nos interesan:
    legajo, nombre, cantidad de viajes por día del 1 al 31
    '''

    df = planilla

    # Buscar encabezado
    idx_encabezado = buscar_idx_encabezado(df,"apellido y nombre",col_a_buscar=1) 
    if idx_encabezado != -1: # si la tabla no tiene como encabezados los que queremos
            df.columns = df.iloc[idx_encabezado]
            df = df.iloc[idx_encabezado+1:]
    
    # Damos forma al df
    df = df.reset_index(drop=True)
    df = df.dropna(how="all")
    df = df.iloc[:,:33]
    dias_cols = [i+1 for i in range(31)]
    nuevas_cols = ["legajo", "empleado"] + dias_cols
    df.columns = nuevas_cols

    # Procesamos columna "legajo"
    df = df[df["legajo"].notna()] #1. nos quedamos con las filas que tengan un legajo en la primera fila
    df["legajo"] = (
            df["legajo"]
            .apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x))
            .str.replace(r"[.,\s]", "", regex=True)
    ) #2. quitamos los puntos, comas y espacios que puedan tener los legajos
    df = df[df["legajo"].astype(str).str.isdigit()] #3. nos quedamos aquellos que sean efectivamente un legajo (un numero)
    df["legajo"] = df["legajo"].astype(str) #4. transformamos a string

    #Como las columnas que quedan que podrían tener na son de dias sin viaticos, se ponen en cero
    df[dias_cols] = df[dias_cols].apply(pd.to_numeric, errors = 'coerce', downcast = 'integer')

    df = df.fillna(0)

    return df

def son_similares(nombre_1, nombre_2, umbral=0.8):
    if nombre_1 == nombre_2:
        return True
    # Lo hago así porque no sé si llamar al método SequenceMatcher es costoso
    ratio = difflib.SequenceMatcher(None, nombre_1, nombre_2).ratio()
    return ratio >= umbral

def limpiar_nombre(nombre):
    nombre = nombre.upper() # mayúsculas
    nombre = re.sub(r"['’]", "", nombre) # quitar comas y apóstrofes
    nombre = re.sub(r",\s", " ", nombre)
    nombre = re.sub(r",", " ", nombre)
    reemplazos = {"Á": "A", 
                  "É": "E",
                  "Í": "I", 
                  "Ó": "O",
                  "Ú": "U",
                  "Ü": "U"}
    patron = re.compile("|".join(reemplazos.keys()))
    nombre = patron.sub(lambda m: reemplazos[m.group()], nombre) # reemplazar tildes
    nombre = nombre.strip() # sacar espacios adicionales
    nombre = re.sub(' +',' ',nombre) # idem 
    return nombre.split(" ")

def nombres_coinciden(nombre_1: str, nombre_2: str) -> bool:
    '''
    Chequea que dos nombres coincidan, siguiendo el siguiente criterio:
    -Vamos a considerar como coincidencia matchear al menos dos de los strings
    que componen a nombre_1.
    -Matchear un string es que siga la logica de son_similares y no se pueda matchear con un string
    ya usado de nombre_2.
    Siendo nombre_1 el nombre en la planilla/reporte subido, nombre_2 el nombre en sistema y string una "palabra" dentro de un nombre.
    '''
    nombre_1 = limpiar_nombre(nombre_1)
    nombre_2 = limpiar_nombre(nombre_2)
    palabra_usada_nombre_2 = [False for _ in range(len(nombre_2))]
    coincidencias = 0
    for string_1 in nombre_1:
        for idx,string_2 in enumerate(nombre_2):
            if son_similares(string_1,string_2) and not palabra_usada_nombre_2[idx]:
                coincidencias +=1
                palabra_usada_nombre_2[idx] = True
                break
    return coincidencias >= 2

def validar_legajos_y_nombres(planilla: pd.DataFrame,
                              datos_sistema: pd.DataFrame) -> list[tuple[str,str,str]]:
    '''
    '''
    hay_duplicados = planilla["legajo"].duplicated().any()

    if hay_duplicados:
        st.error("Hay legajos duplicados en la planilla de viáticos. Revisar a mano.")
        st.stop()
    
    # 
    datos_planilla = list(zip(planilla["legajo"],planilla["empleado"]))
    
    datos_sistema = datos_sistema[datos_sistema.columns[:2]]
    datos_sistema.columns = ["legajo","empleado"]
    datos_sistema.dropna(subset=["legajo"])
    datos_sistema["legajo"] = datos_sistema["legajo"].astype(int).astype(str)

    empleados_mal_cargados = []

    for legajo,empleado in datos_planilla:        
        
        # Que empleado_sistema sea None significa que no se halló e legajo en la planilla del sistema
        empleado_sistema = datos_sistema.loc[datos_sistema["legajo"] == legajo, "empleado"]
        
        if empleado_sistema.shape[0] == 0:
            empleado_sistema = None
            empleados_mal_cargados.append((legajo,empleado,empleado_sistema))
        else: 
            if not nombres_coinciden(empleado,empleado_sistema.iloc[0]):
                empleados_mal_cargados.append((legajo,empleado,empleado_sistema.iloc[0]))

    return empleados_mal_cargados
    
def reportar_validacion_legajos(resultados: list[tuple[str,str,str]]):
    '''
    Si reportan resultados de haberlo.
    
    :param resultados: Description
    :type resultados: list[tuple[str, str, str]]
    '''
    if len(resultados) > 0:
        legajos_no_encontrados = {}
        nombres_posiblemente_mal = {}
        
        for legajo,nombre_planilla,nombre_sistema in resultados:
            if nombre_sistema is None:
                legajos_no_encontrados[legajo] = nombre_planilla
            else:
                nombres_posiblemente_mal[legajo] = {"nombre_planilla": nombre_planilla,"nombre_sistema": nombre_sistema}        

        if nombres_posiblemente_mal:
            st.warning("Estos legajos potencialmente tiene un nombre distinto al del sistema:")
            for legajo in nombres_posiblemente_mal.keys():
                st.markdown(" - " + f" {legajo} - Nombre planilla: {nombres_posiblemente_mal[legajo]["nombre_planilla"]} - Nombre sistema: {nombres_posiblemente_mal[legajo]["nombre_sistema"]}.")

        if legajos_no_encontrados:
            st.warning("Estos legajos no fueron hallados en el sistema:")
            for legajo in legajos_no_encontrados.keys():
                st.markdown(" - " + f" {legajo} - {legajos_no_encontrados[legajo]}")
    else:
        st.write("No se encontraron errores de carga de legajos o nombres.")

def modificar_reportar_viaticos_en_ausencia(ausencias,planilla_viaticos):
    '''
    Recibe los diccionarios ausencias y planilla_viaticos
    Para cada día donde se ausentaron e hicieron horas extras,
    segun el tipo de ausencia, se pone en 0 la hora extra en planilla_viaticos
    Dejandolo listo para exportar a csv
    '''
    viaticos = planilla_viaticos
    inconsistencias_ausencias = defaultdict(list)

    legajos_planilla = set(viaticos["legajo"].unique().tolist())
    for legajo in ausencias.keys():
        legajo = str(legajo)
        if legajo in legajos_planilla:
            motivos = ausencias[legajo]["motivos"]
            for idx, dia in enumerate(ausencias[legajo]["dias"]):
                # Para ese legajo si en algun día que estuvo ausente tiene viaticos, ponerlas en 0
                if (viaticos.loc[viaticos["legajo"] == legajo, dia] > 0).any():
                    inconsistencias_ausencias[legajo].append((dia,motivos[idx]))
                    viaticos.loc[viaticos["legajo"] == legajo, dia] = 0
    
    # Para todos los casos donde haya un viático con valor mayor a 20, lo ponemos en vente
    dias = [i for i in range(1,32)]

    for idx, row in viaticos.iterrows():
        legajo = row["legajo"]
        for dia in dias:
            valor = row[dia]
            if valor > 20:
                # corregir dataframe
                viaticos.at[idx, dia] = 20
                # registrar inconsistencia
                inconsistencias_ausencias[legajo].append(
                    (dia, "MAYOR A 20 UNIDADES")
                )

    # Para todos los dias no habiles, poner en 0
    hoy = datetime.today()
    mismo_dia_mes_anterior = hoy - relativedelta(months=1)
    mes = mismo_dia_mes_anterior.month
    anio = mismo_dia_mes_anterior.year
    dias_findes = obtener_fines_de_semana(anio,mes)
    dias_feriados = obtener_dias_feriados(anio,mes)
    dias_no_habiles = sorted(set(
                    dias_findes +
                    dias_feriados
                ))

    for idx, row in viaticos.iterrows():
        legajo = row["legajo"]
        for dia in dias_no_habiles:
            valor = row[dia]
            if valor > 0:
                # corregir dataframe
                viaticos.at[idx, dia] = 0
                # registrar inconsistencia
                inconsistencias_ausencias[legajo].append(
                    (dia, "DIA NO HABIL")
                )

    st.write("Se pusieron en cero todos los viáticos que caen en día no hábil. Estos días son: " + lista_a_string(dias_no_habiles))
    
    dict_legajo_empleado = viaticos.set_index("legajo")["empleado"].to_dict()

    if len(inconsistencias_ausencias) > 0:
        st.write("Se anularon horas extras para los siguientes legajos por motivo de ausencia.")
        for legajo in inconsistencias_ausencias:
            inconsistencias_legajo = inconsistencias_ausencias[legajo]
            st.write(f"\n* **Empleado {dict_legajo_empleado[legajo]} - {legajo}**\n")
            s = ""
            for dia, motivo in inconsistencias_legajo:
                s += f"    - {dia}/{mes}/{anio} | {motivo}\n"   
            st.markdown(s)
    
    return viaticos

def transformar_viaticos_a_csv(viaticos: pd.DataFrame):
    resultados = {}
    for index, row in viaticos.iterrows():
        total_viaticos = sum(row[2:])
        legajo = row["legajo"]
        resultados[legajo] = [0,total_viaticos]
    df = (
        pd.DataFrame.from_dict(resultados, orient = 'index', columns = ["cargo(0)", "monto"])
          .reset_index()
          .rename(columns={"index": "legajo"})
    )

    df.sort_values("legajo", ascending = True, inplace = True)
    st.write(df)
    
    return df

def reportar_diferencias_viaticos(df):
    if df.shape[0] > 0:
        st.write("Estas son las diferencias encontradas en los montos al producirse la comparación con las ausencias.")
        st.write(df)
    else:
        st.write("No se hallaron diferencias para reportar.")