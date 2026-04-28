import pandas as pd
import streamlit as st
from datetime import datetime, timedelta
from collections import defaultdict
from pandas.api.types import is_numeric_dtype
import re
import numpy as np
import json
from pathlib import Path

codigos_ausencias_no_descontables_viaticos = set([0,18,38,64,66,76,78,79,80,81,87,100,101,102,600])
codigos_ausencias_no_descontables_hhee = (set([1,4,5,6,7,8,9,10,11,12,13,
                                          14,15,16,17,18,21,22,25,26,
                                          30,31,32,33,34,35,36,37,38,39,
                                          41,42,43,44,48,49,50,57,58,
                                          59,62,65,70,82,83,84,85,87,89,
                                          90,91,96,100,102,103,104,110,
                                          111,120,121,130,131,140,141,
                                          500,501,502,504,505,506,601,
                                          602,603,772,773,774,777,780,
                                          781,783,784,785,788,791,796,
                                          797,798]))


#################################
# FUNCIONES INTERNAS HELPERS.PY #
#################################

def obtener_dias_feriados(anio: int, mes: int) -> list[int]:
    BASE_DIR = Path(__file__).resolve().parent
    ruta_feriados = BASE_DIR / "feriados.txt"
    #ruta_feriados = "feriados.txt"
    with open(ruta_feriados, encoding="utf-8") as f: 
        texto = f.read().replace("NaN", "null")
    
    feriados = json.loads(texto)
    dias = []

    for fdo in feriados:
        fecha = datetime.strptime(fdo["fecha"], "%Y-%m-%d")

        if fecha.year == anio and fecha.month == mes:
            dias.append(fecha.day)

    return sorted(dias)

# Se usa en chequeo_legajos y en armado_csv
def procesar_oficinas(oficinas):
    res = []

    if len(oficinas) == 0:
        return None
    
    if oficinas.strip().lower() == 'todo':
        return [1,1,1]

    oficinas = oficinas.split(",")
    for ofi in oficinas:
        if len(ofi.split("-")) > 1: # Si es un rango de oficinas, ej: 310-312 = 310,311,312
            rango = ofi.split("-")
            for k in range(int(rango[0]),int(rango[1])+1):
                res.append(k)
        else:
            res.append(ofi)

    # Convertir todo a string
    for i in range(0,len(res)):
        if type(res[i]) is int:
            res[i] = str(res[i])

    return res

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

def buscar_idx_encabezado(df_planilla: pd.DataFrame,
                          keyword: str,
                          nombre_tabla: str = "",
                          col_a_buscar: int = 0) -> int:
    '''
    El df es una lectura de dataframe de un archivo excel o csv que está muy mal formateado.
    Por ejemplo: la tabla del mismo empiece en otra fila distinta de la primera, que tenga anotaciones
    que no correspondan en otras celdas, etc.
    Asume que la keyword que es el encabezado de la tabla está en la columna nro. col_a_buscar del df
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

# Se usa en chequeo_legajos y armado_csv
def normalizar_planilla_hhee(planilla_hhee):
    '''
    Para la planilla de horas extras
    '''
    df = planilla_hhee
    df = df[df.columns[:34]]

    idx = buscar_idx_encabezado(df,"Legajo","Planilla de horas extras")
    if idx != -1: # si la tabla no tiene como encabezados los que queremos
        df.columns = df.iloc[idx]
        df = df.drop(idx)
      
    df = df.reset_index(drop=True)
    df = df.dropna(how="all")
    # Rellenar legajos vacíos con último valor válido
    df["Legajo"] = df["Legajo"].ffill(limit=2)
    df["Apellido y Nombre"] = df["Apellido y Nombre"].ffill(limit=2)
    # Quitar filas donde "Legajo" está vacío
    df = df[df["Legajo"].notna()]

    # Identificar columnas de días → las primeras 31 después de las 3 iniciales
    day_cols = df.columns[3:34]
  
    # forzar los valores a numeric
    df.iloc[:, 3:34] = df.iloc[:, 3:34].apply(pd.to_numeric, errors='coerce')

    # Renombrar las fechas por números 1–31
    df.rename(columns={day_cols[i]: i+1 for i in range(len(day_cols))}, inplace=True)

    df = df.rename(columns={
        df.columns[0]: "legajo",
        df.columns[1]: "nombre",
        df.columns[2]: "tipo_hora"
    }) 


    # Tratamiento del legajo    
    df["legajo"] = (
        df["legajo"]
        .apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x))
        .str.replace(r"[.,\s]", "", regex=True)
    )
    
    df = df[df["legajo"].astype(str).str.isdigit()]
    df["legajo"] = df["legajo"].astype(str)

    # Vemos si están correctos los tipos de hora

    tipo_hora = df["tipo_hora"].astype(str).unique()
    
    st.write(tipo_hora)
    if not np.isin(tipo_hora, ["N", "0.5", "0.1"]).any() or len(tipo_hora) > 3:
        st.error(f"Se encontraron estos tipos de hora {tipo_hora} y por tal motivo no pudieron ser procesados.")
        st.stop()

    #Como las columnas que quedan que podrían tener na son de dias y hrs extras, se ponen en cero
    df = df.fillna(0)

    #Quitar aquellos espacios donde legajo quedó en 0
    df = df[df["legajo"] != '0']
    
    return df

def cambiar_fechas(df):
    '''
    Recibe el dataframe de ausencias
    Lo que hace es cambiar las ausencias de forma tal que no se reemplace en
    la tabla las fechas de dia_inicia y dia_fin por el numero de día que correspondería al
    mes anterior (si es que la fecha es del mes pasado).
    '''
    df["dia_inicio"] = pd.to_datetime(df["dia_inicio"],format="%d/%m/%Y")
    df["dia_fin"] = pd.to_datetime(df["dia_fin"],format="%d/%m/%Y")

    hoy = datetime.today()
    hoy = hoy.replace(hour=0, minute=0, second=0, microsecond=0)

    # Determinar el mes anterior
    primer_dia_mes_anterior = (hoy.replace(day=1) - timedelta(days=1)).replace(day=1)
    ultimo_dia_mes_anterior = hoy.replace(day=1) - timedelta(days=1)

    # Función para acotar el rango al mes anterior (i.e. si es anterior al mes pasado
    # se inicializa en el primer día del mes anterior, análogo a si es un mes posterior).
    def acotar_al_mes_anterior(row):
        nuevo_inicio = row["dia_inicio"]
        nuevo_fin = row["dia_fin"]
        if row["dia_inicio"] < primer_dia_mes_anterior:
            nuevo_inicio = primer_dia_mes_anterior
        if row["dia_fin"] > ultimo_dia_mes_anterior:
            nuevo_fin = ultimo_dia_mes_anterior
        if nuevo_inicio > nuevo_fin:
            return pd.Series([primer_dia_mes_anterior,ultimo_dia_mes_anterior])  # rangos fuera del mes anterior
        return pd.Series([nuevo_inicio, nuevo_fin])

    df[["dia_inicio", "dia_fin"]] = df.apply(acotar_al_mes_anterior, axis=1)

    # Extraemos solo los días
    df["dia_inicio"] = df["dia_inicio"].dt.day
    df["dia_fin"] = df["dia_fin"].dt.day

def transformar_ausencias_a_dict(ausencias,es_viaticos=False) -> dict:
    '''
    A partir de las ausencias se arma un diccionario:
    dict[legajo] = { "empleado": string, "dias": [int] }
    donde dias es una lista de numeros de los dias en 
    que esa persona estuvo ausente.
    '''
    df_raw = pd.read_excel(ausencias)

    oficina = None
    empleado = None
    legajo = None

    rows = []
 
    for _, row in df_raw.iterrows():
        row = row.to_list()
        primera_col = str(row[0]).strip() if pd.notna(row[0]) else ""

        # Detectar inicio de un bloque por Oficina
        if primera_col.startswith("Oficina :"):
            oficina = primera_col.replace("Oficina :", "").strip()
            empleado = None
            legajo = None
            continue

        # Detectar empleado
        if primera_col.startswith("Empleado:"):
            match = re.search(r"Empleado:\s*(.*?)\s*Legajo:\s*0*([0-9]+)", primera_col)
            if match:
                empleado = match.group(1).strip()
                legajo = match.group(2).strip()
            continue

        # Filas de ausencias (requieren oficina + empleado + fechas)
        if oficina and empleado and pd.notna(row[0]) and pd.notna(row[1]):
            primer_dia = row[0]
            ultimo_dia = row[1]
            motivo_raw = row[4] if len(row) > 4 else None
            nro_motivo = (
                motivo_raw.split("-")[0].strip()
                if isinstance(motivo_raw, str) and "-" in motivo_raw
                else None
            )
            motivo = (
                motivo_raw.split("-")[1].strip()
                if isinstance(motivo_raw, str) and "-" in motivo_raw
                else None
            )
            

            rows.append([oficina, legajo, empleado, primer_dia, ultimo_dia, nro_motivo, motivo])
    
    df = pd.DataFrame(rows, columns=["oficina","legajo", "empleado", "dia_inicio", "dia_fin", "nro_motivo", "motivo"])
    df["legajo"] = df["legajo"].astype(str).str.lstrip("0")
    
    cambiar_fechas(df)
    df["nro_motivo"] = df["nro_motivo"].astype(int)
    
    if es_viaticos:
        df = df[~df["nro_motivo"].isin(codigos_ausencias_no_descontables_viaticos)]
    else:
        df = df[~df["nro_motivo"].isin(codigos_ausencias_no_descontables_hhee)]

    legajo_dict = defaultdict(lambda: {"empleado": None, "dias": [], "motivos": []})
    
    for _, row in df.iterrows():
        legajo = str(row["legajo"])
        nombre = row["empleado"]
        oficina = row["oficina"]
        nro_motivo = row["nro_motivo"]
        motivo = row["motivo"]
        #if int(row["dia_fin"]) == 30:
        #    dias = list(range(int(row["dia_inicio"]), int(row["dia_fin"]) + 2))
        #else:
        dias = list(range(int(row["dia_inicio"]), int(row["dia_fin"]) + 1))

        legajo_dict[legajo]["empleado"] = nombre
        legajo_dict[legajo]["oficina"] = oficina
        legajo_dict[legajo]["dias"].extend(dias)
        legajo_dict[legajo]["motivos"].extend([f"{nro_motivo} - {motivo}" for _ in range(len(dias))])

    # opcional: eliminar duplicados y ordenar
    for v in legajo_dict.values():
        v["dias"] = sorted(set(v["dias"]))

    return dict(legajo_dict)

def lista_a_string(lista):
    s = ""
    s += str(lista[0])
    for x in lista[1:]:
        s += ", " + str(x)

    return s 

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


