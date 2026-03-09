import pandas as pd
import numpy as np
import streamlit as st
from collections import defaultdict
import re
import xlsxwriter as xl 
import io

#Cuando se habla de df_uno o cant_izq se refiere al primer archivo excel que se extrae del sistema y que cada agrupación debe completar en la columna de Se descuenta?
#en cambio, cuando hablamos de df_dos o cant_der se refiere al segunfo archivo tirado por el  sistema que debe tener todas las correcciones realizada en la primer parte


#Motivos que se tratan de forma especial


motivos_sin_accion = ["RESERVA DE CARGO", "LICENCIA S/GOCE DE SUELDO", "RESERVA DE CARGO POR FUNCION CONCEJAL"] #se ignora la comparación
motivos_copiar_fila = ["SUSPENSION",  "JORNADA REDUCIDA SAYEP"] #si hay diferencias se copia la fila_ws
motivo_presentismo = "PRESENTISMO PUNTUALIDAD (PROCESO DESCUEN" #se hace la comparación solo de izquierda a derecha


def leer_excel(nombre_archivo: str, nombre_hoja: str = None) -> pd.DataFrame:
    '''
    Lee un archivo excel de la hoja nombre_hoja que tenga exactamente 8 columnas 
    y lo convierte a dataframe
    '''

    if nombre_hoja:

        df = pd.read_excel(nombre_archivo, sheet_name = nombre_hoja, header= None)

    else: #no hay que especificar nombe de hoja, es unica

        df = pd.read_excel(nombre_archivo, header=None)

    df = df.iloc[:,:8]
    df.columns = ["col1","col2","col3","col4","col5","col6","col7","col8"]
    return df

def tipo_de_fila(valor_columna) -> int:
    '''
    Dado un valor en una celda en una fila de un dataframe devuelve:
    * 0 si es una fila en blanco
    * 1 si es una fila con encabezados
    * 2 si es un fila con los datos de las personas
    * 3 si es una fila con las ausencias
    '''

    if pd.isna(valor_columna):
        return 0 #es una fila_ws en blanco
    elif valor_columna == "Nombre descuento":
        return 1 #es una fila con los nombres de las variables
    elif valor_columna.split(" ")[0] == "Oficina:": 
        return 2 #es una fila_ws con los datos de las personas
    else:
        return 3 #es una fila_ws con las ausencias
    
def extraer_legajo_persona(datos_persona: str) -> str:
    """
    Dados el string de datos de una persona, devuelve el legajo al que corresponde.
    Con string de datos hablamos del formato: "Oficina: XYZ Legajo: 12345 NroCargo: n Nombre: PEREZ JUAN Partición: -- TipoCargo: XX Categoría: YY"
    """
    fila_split = datos_persona.split(":")[2]
    fila_split_dos = fila_split.split(" ")
    dato = fila_split_dos[1]
    legajo = dato

    return legajo

def armado_df(df: pd.DataFrame) -> tuple[dict,pd.DataFrame]:
    '''
    Dado el dataframe de la planilla crudo, arma un dataframe para cada persona y su motivo de ausencia, 
    además un diccionario legajo -> string con datos de la persona.
    '''

    dic_datos_personas = {}

    #Para el armado del dataFrame:
    legajos = []
    nombres_descuento = []
    llegadas_tarde = []
    salidas_antes = []
    dias_sin_fichadas = []
    dias_ausencia = []
    se_descuenta = []
    observaciones = []
    formulario = []

    for fila in df.itertuples():
        
        #iteramos sobre las filas de la primer columna de nuestro dataframe para ver qué tipo de fila es
        tipo = tipo_de_fila(fila.col1)
        if tipo == 2:
            legajo = extraer_legajo_persona(fila.col1)
            dic_datos_personas[str(legajo)] = fila.col1
        
        if tipo == 0 or tipo == 1:
            continue #la fila es una fila en blanco, o una fila que tiene los nombres de las variables (no nos importa)

        if tipo == 3:
            
            legajos.append(legajo)
            nombres_descuento.append(fila.col1)
            llegadas_tarde.append(fila.col2)
            salidas_antes.append(fila.col3)
            dias_sin_fichadas.append(fila.col4)
            dias_ausencia.append(fila.col5)
            se_descuenta.append(fila.col6)
            observaciones.append(fila.col7)
            formulario.append(fila.col8)
            

    df_res = pd.DataFrame({"Legajo":legajos, "Nombre descuento" : nombres_descuento,
                           "Llegadas tarde":llegadas_tarde, "Salidas antes": salidas_antes, "Días sin fichada" : dias_sin_fichadas,
                           "Días ausencia": dias_ausencia, "Se descuenta?" : se_descuenta,"Observaciones": observaciones,"Finalizada":formulario})
    return dic_datos_personas, df_res


def hay_descuentos_vacios(df : pd.DataFrame) -> bool:
    '''
    Devuelve True si hay descuentos vacíos, duh.
    '''

    #Chequea si a la oficina o agrupación le falta completar algo
    df_filtrado = df[~df["Nombre descuento"].isin(motivos_sin_accion)]
    return df_filtrado["Se descuenta?"].isna().any()

def comparacion(df_uno:pd.DataFrame, df_dos:pd.DataFrame):
    '''
    df_uno es el df de la primera planilla de descuentos con los campos de se_descuenta completados
    df_dos es la segunda planilla de descuentos que sale del sistema.
    El criterio de comparación es sencillo:
    * Para los motivos_sin_accion no se consideran en la comparacion
    * Para los motivos_copiar_fila y motivo_presentismo comparamos que coincidan en ambas planillas. De no coincidir se guarda el dato en df_dos 
    '''

    legajos_chequear = df_uno["Legajo"].unique()

    # obtenemos el conjunto de todos los motivos que aparecen en ambos arrays
    motivos_uno = df_uno["Nombre descuento"].unique()
    motivos_dos = df_dos["Nombre descuento"].unique()
    motivos = np.union1d(motivos_uno,motivos_dos)

    # diccionario de diferencias con claves legajos y valor diccionarios cuyas claves son motivos y valor es un vector de los campos
    # 'Llegadas tarde','Salidas antes','Días sin fichada','Días ausencia'
    diferencias = defaultdict(lambda: defaultdict(list))

    for legajo in legajos_chequear:

        legajo_str = str(legajo)

        for motivo in motivos:

            if motivo in motivos_sin_accion: 
                continue
            
            # para cada legajo y motivo, los filtramos en amos dataframe
            df_uno_mot_leg = df_uno[(df_uno["Legajo"] == legajo) & (df_uno["Nombre descuento"] == motivo) & (df_uno["Se descuenta?"] == "Descontar")]
            df_dos_mot_leg = df_dos[(df_dos["Legajo"] == legajo) & (df_dos["Nombre descuento"] == motivo)]

            cant_izq = df_uno_mot_leg.shape[0]
            cant_der = df_dos_mot_leg.shape[0]

            if cant_izq == 0 and cant_der == 0: # Si no existe el motivo para este legajo en ninguna planilla no hago nada
                    continue
            
            if motivo in motivos_copiar_fila: # Lo  que hacemos es para este caso ver el valor y compararlos 

                valor_izq = df_uno_mot_leg["Días ausencia"].iloc[0] if cant_izq > 0 else 0
                valor_der = df_dos_mot_leg["Días ausencia"].iloc[0] if cant_der > 0 else 0

                if valor_der == 0: # Si no está del lado derecho es porque debe estar del lado izquierdo
                    diferencias[legajo_str][motivo] = [0,0,0,valor_izq ]

                elif valor_izq == 0 and valor_izq != valor_der: # Si no está del lado izquierdo, está del lado derecho o si ambos no son cero, chequear que den lo mismo.
                    diferencias[legajo_str][motivo] = [0,0,0,valor_izq]

            elif motivo == motivo_presentismo:
                if cant_der < cant_izq: # por lo hablado con Azu solo reportamos cuando esta en el de la izquierda y no en el de la derecha

                    diferencias[legajo_str][motivo] = [df_uno_mot_leg.loc[0,"Llegadas tarde"],df_uno_mot_leg.loc[0,"Salidas antes"],df_uno_mot_leg.loc[0,"Días sin fichada"],df_uno_mot_leg[0,"Días ausencia"]]

            else: # para cualquier otro motivo
                if cant_der > cant_izq:

                    diferencias[legajo_str][motivo] = [0,0,0,cant_der]

                elif cant_der < cant_izq:

                    diferencias[legajo_str][motivo] = [0,0,0,cant_izq]

    return diferencias

motivos_sin_accion = ["RESERVA DE CARGO", "LICENCIA S/GOCE DE SUELDO", "RESERVA DE CARGO POR FUNCION CONCEJAL"] #se ignora la comparación
motivos_copiar_fila = ["SUSPENSION",  "JORNADA REDUCIDA SAYEP"] #si hay diferencias se copia la fila_ws
motivo_presentismo = "PRESENTISMO PUNTUALIDAD (PROCESO DESCUEN" #se hace la comparación solo de izquierda a derecha

# CREAR PLANILLA
def escribir_datos(ws: xl.worksheet.Worksheet, fila: int, col_inicial: int, datos_persona: str, formato: xl.format.Format):
    '''
    Escribe para la planilla de la derecha los datos de la persona en la fila indicada.
    '''
    # Escribo datos de persona
    #rango_datos = "A" + str(fila) + ":H" + str(fila)
    ws.merge_range(fila, col_inicial, fila, col_inicial + 7, datos_persona, formato)

def escribir_encabezado(ws: xl.worksheet.Worksheet, fila: int, col_inicial: int, encabezados: list[str], formato: xl.format.Format):
    '''
    Escribe el encabezado para cualquiera de las dos planillas.
    '''
    # Escribo encabezados
    for i in range(col_inicial, len(encabezados) + col_inicial):
        ws.write(fila,i,encabezados[i - col_inicial],formato)

def escribir_fila_tabla_original(ws: xl.worksheet.Worksheet, fila: int, datos: pd.Series): 
    '''
    Escribe los datos de la fila de la tabla original.
    '''
    datos = datos.fillna("")
    datos = datos.array
    ws.write_row(fila, 0, datos)

def escribir_fila_tabla_diferencias(ws: xl.worksheet.Worksheet, fila: int, motivo: str, vector: list[int]):
    '''
    Escribe los datos con diferencias a la derecha
    '''
    IDX_DIF_ENCABEZADOS = [i for i in range(9,16)] 
    
    # Primero copiamos el motivo
    ws.write(fila,IDX_DIF_ENCABEZADOS[0],motivo)
    
    # Después copiamos el vector
    for idx_vector,i in enumerate(IDX_DIF_ENCABEZADOS[1:5]):
        ws.write(fila,i,vector[idx_vector])

    # Agregamos dropdown de validación en la columna observacion
    ws.data_validation(fila, IDX_DIF_ENCABEZADOS[5], fila, IDX_DIF_ENCABEZADOS[5], 
                       {'validate': 'list', 'source': ['Descontar', 'No descontar']})

def unir_diccionarios(d1: defaultdict[str, defaultdict[str, list[str]]], d2: defaultdict[str, defaultdict[str, list[str]]]) -> defaultdict:
    '''
    Se espera que d1 y d2 tengan pares claves,valor repetidos. Entonces añadimos a d1, todos los pares clave,valor que no estén de d2
    Como ambos son defaultdict, simplemente añado todo de d2 a d1
    '''

    for k, v in d2.items():
        d1[k] = v
    
    return d1

def imprimir_diferencias(dd: defaultdict[str, defaultdict[str, list[str]]]):
    st.write("Estas son las diferencias encontradas:")
    for clave_principal, subdict in dd.items():
        st.write(f"\n Legajo {clave_principal}")

        for nombre, numeros in subdict.items():
            if nombre not in motivos_copiar_fila and nombre != motivo_presentismo:
                st.write(f"-    Motivo  {nombre}: {numeros[3]}")
            else:
                st.write(f"-    Motivo  {nombre}")

# A factory for nested defaultdicts
def nested_ddict_factory():
    return defaultdict(nested_ddict_factory)

# Function to convert a nested defaultdict to a nested dict
def default_to_regular(d: defaultdict[str, defaultdict[str, list[str]]]) -> dict[str, dict[str, list[str]]]:
    if isinstance(d, defaultdict):
        d = {k: default_to_regular(v) for k, v in d.items()}
    return d

def crear_excel(df_original: pd.DataFrame, diferencias: defaultdict[str, defaultdict[str, list[str]]], dict_personas: defaultdict, buffer):
    '''
    Crea el excel final que empleará asistencia para hacer comparaciones entre lo descontado en el primer cálculo y lo que sale del segundo.
    El criterio queda TO-DO 
    '''
    wb = xl.Workbook(buffer)
    ws = wb.add_worksheet()

    encabezados = ['Nombre descuento','Llegadas tarde','Salidas antes','Días sin fichada','Días ausencia','Se descuenta?','Observaciones','Finalizada']
    formato_encabezado = wb.add_format()
    formato_encabezado.set_bg_color("#A4C2F4")
    formato_encabezado.set_bold()

    fila_ws = 0
    legajo_actual = None
    motivos_copiar_completo = ["SUSPENSION",  "JORNADA REDUCIDA SAYEP", "PRESENTISMO PUNTUALIDAD (PROCESO DESCUEN"]

    df_original = default_to_regular(df_original)
    # Como queremos poner una fila vacía entre cada tabla de legajo con sus motivos de descuento
    # vamos a contar cuantos motivos tenemos por legajo y vamos disminuyendo el contador para añadir la fila vacía.
    cant_filas_por_legajo = df_original["Legajo"].value_counts(dropna=True).to_dict()
    cant_filas_legajo_df = 0
    legajo_en_diferencias = None
    #print("--------Esta es la planilla original-----------")
    #print(df_original[["Legajo", "Nombre descuento", "Se descuenta?"]])
    #print("-------------------")
    for index, row in df_original.iterrows():
        # TO-DO Como no podemos observar los valores del excel tenemos que ir copiando del df original y copiar al mismo tiempo las diferencias
        # todo esto segun corresponda
        motivo = row["Nombre descuento"]
        #print("----------------------------------")
        #print(f"Vamos por la iteración {index}")
        #print(f"Vamos por legajo {row["Legajo"]} y motivo {row["Nombre descuento"]}")
        #imprimir_diferencias(diferencias)
        if row["Legajo"] != legajo_actual:
            #print("[i] Actualizamos legajo")
            legajo_actual = row["Legajo"]
            cant_filas_legajo_df = cant_filas_por_legajo[legajo_actual]
            legajo_en_diferencias = legajo_actual in diferencias.keys()
            
            # Escribo datos del legajo Oficina: XYZ Legajo: 12345...
            escribir_datos(ws, fila_ws, 0, dict_personas[legajo_actual], formato_encabezado)
            if legajo_en_diferencias:
                escribir_datos(ws, fila_ws, 9, dict_personas[legajo_actual], formato_encabezado)
            fila_ws += 1
            
            # Escribo encabezado Nombre descuento, Llegadas tarde...
            escribir_encabezado(ws, fila_ws, 0, encabezados, formato_encabezado)
            if legajo_en_diferencias: 
                escribir_encabezado(ws, fila_ws, 9, encabezados, formato_encabezado)
            fila_ws += 1
        
        # si el legajo y motivo del original estan en diferencias
        if legajo_en_diferencias:
            if motivo in diferencias[legajo_actual].keys(): # existe el motivo en el diccionario de diferencias
                #print("[i] Existe el motivo en diferencias: copiar a ambos lados")
                escribir_fila_tabla_original(ws,fila_ws,row[encabezados]) # siempre se escribe lo de la izquierda
                
                vector_motivo = diferencias[legajo_actual][motivo] 
                if motivo in motivos_copiar_completo: 
                    escribir_fila_tabla_diferencias(ws, fila_ws, motivo, vector_motivo)
                    diferencias[legajo_actual].pop(motivo)
                else:
                    escribir_fila_tabla_diferencias(ws, fila_ws, motivo, [0,0,0,1])
                    nuevo_vector_motivo = [a-b for a,b in zip(vector_motivo,[0,0,0,1])]
                    diferencias[legajo_actual][motivo] = nuevo_vector_motivo
                    # quito el motivo si ya no hay más en las diferencias
                    if sum(nuevo_vector_motivo) == 0: diferencias[legajo_actual].pop(motivo) 
                
                if not diferencias[legajo_actual]:
                    diferencias.pop(legajo_actual)

                cant_filas_legajo_df -= 1 
                fila_ws += 1
        
            else: # si esta iteracion el legajo y motivo no está en diferencias
                #print("[i] No existe el motivo en diferencias: escribir fila tabla original")
                escribir_fila_tabla_original(ws,fila_ws,row[encabezados])
                cant_filas_legajo_df -= 1
                fila_ws += 1
        else: 
            #print("[i] No existe el motivo en diferencias: escribir fila tabla original")
            escribir_fila_tabla_original(ws,fila_ws,row[encabezados])
            cant_filas_legajo_df -= 1
            fila_ws += 1

        if cant_filas_legajo_df == 0: # si no tengo mas filas para el legajo_actual del df original, pongo las diferencias que faltan
            #print("[i] Nos quedamos sin filas legajos df: ver y copiar diferencias restantes sin matcheo")
            #print(f"Legajo en diferencias {legajo_en_diferencias}")
            if not legajo_en_diferencias: 
                fila_ws += 1
                continue
            motivos_de_legajo_dif = list(diferencias[legajo_actual].keys())
            for motivo in motivos_de_legajo_dif:
                if motivo in motivos_copiar_completo:
                    escribir_fila_tabla_diferencias(ws, fila_ws, motivo, diferencias[legajo_actual][motivo])
                else:
                    cant_motivos = diferencias[legajo_actual][motivo][3]
                    for _ in range(cant_motivos):
                        escribir_fila_tabla_diferencias(ws, fila_ws, motivo, [0,0,0,1])
                        fila_ws += 1
                    diferencias[legajo_actual].pop(motivo)
            diferencias.pop(legajo_actual)
            fila_ws += 1 # Añado fila en blanco

    ws.set_column_pixels(0, 0, 340)
    ws.set_column_pixels(1, 5, 102)
    ws.set_column_pixels(6, 6, 290)

    ws.set_column_pixels(9, 9, 340)
    ws.set_column_pixels(10, 14, 102)
    ws.set_column_pixels(15, 15, 290)
    
    wb.close()
    return wb