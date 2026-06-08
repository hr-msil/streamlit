import numpy as np
import pandas as pd
import streamlit as st

def buscar_idx_encabezado(df_planilla: pd.DataFrame,
                          keyword: str,
                          nombre_tabla: str = "",
                          col_a_buscar: int = 0) -> int:
    '''
    El df es una lectura de dataframe de un archivo excel o csv que está muy mal formateado.
    Por ejemplo: la tabla del mismo empieza en otra fila distinta de la primera, que tenga anotaciones
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

def normalizar_hoja_guardia(df, nombre_oficina):
    '''
    Para la planilla de horas extras
    '''
    df = df[df.columns[:36]]

    idx = buscar_idx_encabezado(df,"legajo", nombre_oficina)
    df.columns = df.iloc[idx]
    df = df.iloc[idx + 2:].reset_index(drop=True)

    df = df.dropna(how="all")
    
    #renombro columnas
    df.columns = ["legajo", "nombre", "tipo_guardia", "valor_por_hora"] + [str(i) for i in range(1,32)] + ["total"] 

    # Rellenar legajos vacíos con último valor válido
    df["legajo"] = df["legajo"].ffill(limit=1)
    df["nombre"] = df["nombre"].ffill(limit=1)

    # Quitar filas donde "Legajo" está vacío
    df = df[df["legajo"].notna()]

    # Tratamiento del legajo    
    df["legajo"] = (
        df["legajo"]
        .apply(lambda x: str(int(x)) if isinstance(x, (int, float)) and not pd.isna(x) else str(x))
        .str.replace(r"[.,\s]", "", regex=True)
    )
    
    df = df[df["legajo"].astype(str).str.isdigit()]
    df["legajo"] = df["legajo"].astype(str)

    # Vemos si están correctos los tipos de hora
    df['tipo_guardia'] = df['tipo_guardia'].str.strip()
    tipo_guardia = df["tipo_guardia"].unique()
    valores_guardias = ["SEM","S/D/F"]
    valores_extraños = df[~df['tipo_guardia'].isin(valores_guardias)]

    if len(tipo_guardia) != 2 or valores_extraños.shape[0] > 0:
        st.error("Hay un error en la columna del tipo de guardia que se especifica.")
        st.write(tipo_guardia)

    #Como las columnas que quedan que podrían tener na son de dias y hrs extras, se ponen en cero
    df = df.fillna(0)

    #Quitar aquellos espacios donde legajo quedó en 0
    df = df[df["legajo"] != '0']

    # Por ahora es esto, después hay que procesarlo bien y no tiene que quedar solo estas columnas
    df = df[["legajo", "nombre", "tipo_guardia", "total"]]

    df = df.groupby(["legajo", "nombre", "tipo_guardia"], as_index = False)["total"].sum()

    df_pivot = df.pivot(
        index = "legajo",
        columns = "tipo_guardia",
        values = "total"
    ).reset_index()

    return df_pivot

# def consolidar(df_guardias):

def pivotear_guardias_medicas(archivo_xls):
    guardias = pd.read_excel(archivo_xls, sheet_name=None)
    df_guardias = pd.DataFrame(columns = ["legajo", "SEM", "S/D/F"])
    
    for nombre, df_hoja in guardias.items():
        df_oficina = normalizar_hoja_guardia(df_hoja, nombre)
        df_guardias = pd.concat([df_guardias, df_oficina], ignore_index = True)

    # acá hay que añadir un control más que relacion con tema ausencias, horarios de guardia del agente (no puede hacer hhee en un día que
    # es su guardia)
    # y que, como tenés agentes que hacen guardias en varias oficinas no puede sumar más de 24 horas en un mismo día (que no corra guardia por 
    # dos lugares distintos el mismo día)

    df_guardias = df_guardias.groupby(["legajo"], as_index = True)[["S/D/F","SEM"]].sum()
    st.write(df_guardias)
    #consolidado = consolidar(df_guardias)


    