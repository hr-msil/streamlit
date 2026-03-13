import pandas as pd
import streamlit as st
import numpy as np

from collections import defaultdict
from helpers import lista_a_string

def transformar_hhee_a_csv(df: pd.DataFrame):
    '''
    Arma el csv que se carga al sistema
    Se ejecuta una vez que se compararon las ausencias con la planilla de hhee.
    '''
    # Columnas que representan días/horas (todas menos legajo, nombre, tipo_hora)
    non_day = ["legajo", "nombre", "tipo_hora"]
    day_cols = [c for c in df.columns if c not in non_day]

    # 1) Sumar por legajo, nombre y tipo_hora todos los días
    #    -> queda un DataFrame con la suma por tipo_hora para ese legajo/nombre (aún en columnas de día)
    grouped = df.groupby(["legajo", "nombre", "tipo_hora"])[day_cols].sum()

    # 2) Colapsar las columnas de día a un único total por cada (legajo,nombre,tipo_hora)
    grouped_total = grouped.sum(axis=1).reset_index(name="horas")

    grouped_total["horas"] = np.ceil(grouped_total["horas"])
    # 3) Pivotear para que cada tipo_hora quede en su propia columna
    summary = grouped_total.pivot_table(
        index=["legajo", "nombre"],
        columns="tipo_hora",
        values="horas",
        fill_value=0
    ).reset_index()
  
    # 4) Renombrar las columnas según tu nomenclatura solicitada
    # Identificar los valores únicos en orden de aparición
    unique_types = list(df['tipo_hora'].dropna().unique())
    unique_types = unique_types[0:3]
    # Asegurar que tengamos exactamente 3 tipos
    if len(unique_types) < 3:
        st.error("Advertencia: se esperaban exactamente 3 tipos de hora. Se detectaron menos")
        st.stop()

    # Mapeo universal según orden
    mapping = {
        unique_types[0]: 'horas_normales',
        unique_types[1]: 'horas_50',
        unique_types[2]: 'horas_100'
    }

    summary = summary.rename(columns=mapping)
  
    # 5) Asegurarse de que existan las 3 columnas esperadas
    for col in ["horas_normales", "horas_50", "horas_100"]:
        if col not in summary.columns:
            summary[col] = 0

    # 6) Orden final de columnas: legajo, horas_normales, horas_50, horas_100, nombre
    summary_final = summary[["legajo", "horas_normales", "horas_50", "horas_100", "nombre"]]
    summary_final.insert(1, "columna(0)", 0)

    numeric_cols = ["horas_normales", "horas_50", "horas_100"]
    for col in numeric_cols:
        summary_final[col] = (
            summary_final[col]
                .astype(str)          # por si vienen como object/float/string
                .str.replace(",", ".", regex=False)  # reemplaza coma decimal si aparece
        )
        summary_final[col] = pd.to_numeric(summary_final[col], errors="coerce").fillna(0)


    cols = ["horas_normales", "horas_50", "horas_100"]

    summary_final = summary_final[(summary_final[cols] != 0).all(axis=1)]
    # Lo transformamos a CSV
    return summary_final

def eliminar_legajo_sin_hhee(df: pd.DataFrame) -> pd.DataFrame:
   
   df = df[(df["horas_normales"] != 0) & (df["horas_50"] != 0) & (df["horas_100"] != 0)]

   return df

def anular_unidad_por_ausencias(ausencias_ofi,planilla):
    '''
    Recibe los diccionarios ausencias_ofi y planilla.
    Para cada día donde se ausentaron e hicieron horas extras o viajes,
    segun el tipo de ausencia, se pone en 0 la unidad en el día en la planilla (sea hora extra o viatico)
    Dejandolo listo para exportar a csv.
    Devuelve un diccionario legajo -> lista[tuple[int,string]] de los días y motivos donde halló una inconsistencia.
    '''
    planilla = planilla
    inconsistencias_ausencias = defaultdict(list)

    legajos_planilla = set(planilla["legajo"].unique().tolist())
    for legajo in ausencias_ofi.keys():
        legajo = str(legajo)
        if legajo in legajos_planilla:
            motivos = ausencias_ofi[legajo]["motivos"]
            for idx, dia in enumerate(ausencias_ofi[legajo]["dias"]):
                # Para ese legajo si en algún día que estuvo ausente tiene horas extras, ponerlas en 0
                if (planilla.loc[planilla["legajo"] == legajo, dia] > 0).any():
                    inconsistencias_ausencias[legajo].append((dia,motivos[idx]))
                    planilla.loc[planilla["legajo"] == legajo, dia] = 0
    
    return inconsistencias_ausencias

def reportar_inconsistencias_hhee(inconsistencias_ausencias, ausencias_ofi):
    if len(inconsistencias_ausencias) > 0:
        st.write("Se anularon horas extras para los siguientes legajos por motivo de ausencia.")
        s = ""
        for legajo in inconsistencias_ausencias:
            inconsistencias_legajo = inconsistencias_ausencias[legajo]
            s += f"\n* **Empleado {ausencias_ofi[legajo]['empleado']} - {legajo}**\n"
            for dia, motivo in inconsistencias_legajo:
                s += f"    - {dia} | {motivo}\n"   
        st.markdown(s)

def diferencias_entre_planillas(df_1 : pd.DataFrame, df_2 : pd.DataFrame) -> pd.DataFrame:
   """
   Función que compara dos dataFrame para encontrar inconsistencias entre las planillas de 
   horas extras mandadas por la oficina y la de horas extras generadas por el programa
   :param df_1: dataFrame correspondiente al realizado por el programa 
   :param df_2: dataFrame correspondiente al realizado por la oficina

   """
   df_res = pd.DataFrame(columns= ["Legajo", "Columna(0)","Cant. HN", "Cant. H 50%",
                                   "Cant. H 100%", "Apellido y Nombre"])
   
   legajos = []
   cant_HN = []
   cant_H_50 = []
   cant_H_100 = []
   nombres = []
   columna_cero =[]
   cant_filas = df_1.shape[0] #Se supone que todos los legajos se encuentran en orden de mayor a menor
   # Y que ambos dataFrame cuentas con los mismos legajos

   #Los nombres de las columnas son las siguientes: legajo, horas_normales, horas_50, horas_100, nombre
   for i in range(cant_filas):
    
    fila_1 = df_1.iloc[i]
    fila_2 = df_2.iloc[i]
    legajos.append(fila_1["legajo"])
    nombres.append(fila_1["nombre"])
    cant_HN.append(fila_1["horas_normales"] - fila_2["horas_normales"])
    cant_H_50.append(fila_1["horas_50"] - fila_2["horas_50"])
    cant_H_100.append(fila_1["horas_100"] - fila_2["horas_100"])
    columna_cero.append(0)
        
   df_res["Legajo"] = legajos
   df_res["Cant. HN"] = cant_HN
   df_res["Cant. H 50%"] = cant_H_50
   df_res["Cant. H 100%"] = cant_H_100
   df_res["Apellido y Nombre"] = nombres
   df_res["Columna(0)"] = columna_cero

   df_res = df_res[(df_res["Cant. HN"] != 0) | (df_res["Cant. H 50%"] != 0) | (df_res["Cant. H 100%"] != 0)]
   return df_res


def reportar_diferencias_entre_planillas(df_antes, df_despues):
    df_diferencias = diferencias_entre_planillas(df_despues, df_antes)
    if (df_diferencias.shape[0] > 0):
        st.write("Este es el csv antes de compararlo con las ausencias:")
        st.write(df_antes)
        st.write("Este es el csv después de compararlo con ausencias:")
        st.write(df_despues)
        st.write(f"Esta es la diferencia entre lo mandado por la oficina y lo comparado con los ausencias:")
        st.write(df_diferencias)
    else:
        st.write("No se hallaron diferencias entre las dos planillas luego de comparar ausencias.")
        st.write("Este es el csv final:")
        st.write(df_despues)
