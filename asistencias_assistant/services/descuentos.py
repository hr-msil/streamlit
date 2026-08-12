import pandas as pd
import numpy as np
import streamlit as st
from collections import defaultdict
import xlsxwriter as xl 

# Cuando se habla de df_uno o cant_izq se refiere al primer archivo excel que se extrae del sistema y que cada agrupación ya debió haber completado la columna 'Se descuenta?'
# en cambio, cuando hablamos de df_dos o cant_der se refiere al segundo archivo tirado por el sistema que debería tener todas las correcciones realizada en la primer parte.
# La idea del programa es hacer un doble chequeo. 
# 1) Primero se baja la planilla de Descuentos del sistema. Aquella planilla se la manda a cada oficina para que cada una de ellas 
# complete la columna de 'Se descuenta?' en el sheets compartido.
# 2) El área de asistencia es la encargada de que esos cambios se vean reflejados en el sistema. 
# 3) Luego se vuelve a bajar la planilla del sistema (que el propio sistema tiene sus fallas, y a veces alguns ausencias 
# salen por duplicado). Entonces, para esas dos planillas tenemos que ver que diferencias se encuentran.

# Siempre reportamos de la siguiente manera: 
# > Si encontramos que la primer planilla filtrada por 'Descontar' y 'Motivo' tiene más faltas en ese motivo que la segunda planilla 
# -> reportamos la primera planilla
# > Si ocurre al revés 
# -> reportamos la segunda planilla.


#Motivos que se tratan de forma especial
# Se ignora la comparación
motivos_sin_accion = ["RESERVA DE CARGO", "LICENCIA S/GOCE DE SUELDO", "RESERVA DE CARGO POR FUNCION CONCEJAL"] 
# Si hay diferencias se copia el ultimo valor porque este en la planilla no se desglosa en filas por cantidad de días.
motivos_no_desglaseados = ["SUSPENSION",  "JORNADA REDUCIDA SAYEP", "EN PROCESO ART. 32 / 70"]
# Para este único motivo, solo nos interesa hacer la comparación de izquierda a derecha
motivo_presentismo = "PRESENTISMO PUNTUALIDAD (PROCESO DESCUEN" 

def suma_vectores(vec1: list, vec2: list) -> list:
    '''
    Devuelve una lista donde cada posición es la suma de los elementos en la misma posición de ambos vectores
    '''
    res = [sum(x) for x in zip(vec1,vec2)]
    return res


def leer_excel(nombre_archivo: str, nombre_hoja: str = None) -> pd.DataFrame:
    '''
    Lee un archivo excel de la hoja nombre_hoja que tenga exactamente 8 columnas 
    y lo convierte a dataframe
    '''

    if nombre_hoja:
        df = pd.read_excel(nombre_archivo, sheet_name = nombre_hoja, header= None)
    else: #no hay que especificar nombre de hoja, es unica
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
    Dados el string de datos_persona, devuelve el legajo al que corresponde.
    Con string de datos_persona hablamos del formato: "Oficina: XYZ Legajo: 12345 NroCargo: n Nombre: PEREZ JUAN Partición: -- TipoCargo: XX Categoría: YY"
    """
    fila_split = datos_persona.split(":")[2]
    fila_split_dos = fila_split.split(" ")
    dato = fila_split_dos[1]
    legajo = dato

    return legajo

def armado_df(df: pd.DataFrame) -> tuple[dict,pd.DataFrame]:
    '''
    Ya no se usa en el código, pero queda guardado para el futuro

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

def comparacion(df_uno:pd.DataFrame, df_dos:pd.DataFrame):
    '''
    df_uno es el df de la primera planilla de descuentos con los campos de se_descuenta completados
    df_dos es la segunda planilla de descuentos que sale del sistema.
    El criterio de comparación es sencillo:
    * Para los motivos_sin_accion no se consideran en la comparacion
    * Para los motivos_no_desglaseados y motivo_presentismo comparamos que coincidan en ambas planillas. De no coincidir se guarda el dato en un nuevo diccionario 
    * Para los otros motivos, nos quedamos con aquel que tenga mayor cantidad de apariciones de las planillas.
    Para más especificidad consultar la documentación en Docs de "Documentación - Asistencias" > Descuentos
    '''
    # Obtenemos el conjunto de todos los legajos que aparecen en ambos arrays
    legajos_uno = df_uno["Legajo"].unique()
    legajos_dos = df_dos["Legajo"].unique()
    legajos_chequear = np.union1d(legajos_uno, legajos_dos)

    # obtenemos el conjunto de todos los motivos que aparecen en ambos arrays
    motivos_uno = df_uno["Nombre descuento"].unique()
    motivos_dos = df_dos["Nombre descuento"].unique()
    motivos = np.union1d(motivos_uno, motivos_dos)

    # diccionario de diferencias con claves legajos y valor diccionarios cuyas claves son motivos y valor es un vector de los campos
    # 'Llegadas tarde','Salidas antes','Días sin fichada','Días ausencia'
    diferencias = defaultdict(lambda: defaultdict(list))

    for legajo in legajos_chequear: #Recorremos el conjunto de legajos que aparecen en ambos dfs, así somos capaces de completar todos los casos

        legajo_str = str(legajo)

        for motivo in motivos: #Recorremos el conjunto de motivos que aparecen en ambos dfs

            if motivo in motivos_sin_accion: 
                continue
            
            # Para cada legajo y motivo, los filtramos en ambos dataFrame. Además, solo filtramos los que se descuentan ya que son los casos que tendrían que aparecer en df_dos           
            df_uno_mot_leg = df_uno[(df_uno["Legajo"] == legajo) & (df_uno["Nombre descuento"] == motivo) & (df_uno["Se descuenta?"] == "Descontar")].reset_index(drop=True)
            df_dos_mot_leg = df_dos[(df_dos["Legajo"] == legajo) & (df_dos["Nombre descuento"] == motivo)].reset_index(drop=True)

            df_uno_mot_leg = df_uno_mot_leg.fillna(0)
            df_dos_mot_leg = df_dos_mot_leg.fillna(0)

            cant_izq = df_uno_mot_leg.shape[0]
            cant_der = df_dos_mot_leg.shape[0]

            if cant_izq == 0 and cant_der == 0: # Si no existe el motivo para este legajo en ninguna planilla no hago nada
                continue
            
            if motivo in motivos_no_desglaseados: # Lo que hacemos es para este caso ver el valor y compararlos 

                valor_izq = df_uno_mot_leg["Días ausencia"].iloc[0] if cant_izq > 0 else 0  
                valor_der = df_dos_mot_leg["Días ausencia"].iloc[0] if cant_der > 0 else 0

                if valor_der == 0 and valor_izq != 0: # Si no está del lado derecho es porque debe estar del lado izquierdo, y reportamos lo del lado izquierdo
                    diferencias[legajo_str][motivo] = [0,0,0,valor_izq]

                elif valor_izq == 0 and valor_der != 0: # Caso contrario, reportamos el lado derecho
                    diferencias[legajo_str][motivo] = [0,0,0,valor_der]


            elif motivo == motivo_presentismo:
                if cant_der < cant_izq: #por lo hablado con Azu solo reportamos cuando esta en el de la izquierda y no en el de la derecha, además, nos guardamos toda la fila de datos
                    diferencias[legajo_str][motivo] = [df_uno_mot_leg.loc[0,"Llegadas tarde"],df_uno_mot_leg.loc[0,"Salidas antes"],df_uno_mot_leg.loc[0,"Días sin fichada"],df_uno_mot_leg.loc[0,"Días ausencia"]]

            else: # para cualquier otro motivo, nos guardamos la cantidad que sea mayor

                if cant_der > cant_izq:
                    diferencias[legajo_str][motivo] = [0,0,0,cant_der]

                elif cant_der < cant_izq:
                    diferencias[legajo_str][motivo] = [0,0,0,cant_izq]

    # Para los casos donde la Oficina NO COMPLETO la columna 'Descontar', los reportamos para que esos casos sean vistos devuelta
    df_na = df_uno[df_uno["Se descuenta?"].isna()].reset_index(drop=True)
    df_na = df_na.fillna(0)
    legajos_na_chequear = df_na["Legajo"].unique()
    
    for legajo in legajos_na_chequear:

        legajo_str = str(legajo)
        df_na_leg = df_na[df_na["Legajo"] == legajo_str].reset_index(drop=True)
        motivos_na_leg = df_na_leg["Nombre descuento"].unique()

        for motivo in motivos_na_leg:

            if motivo in motivos_sin_accion: 
                continue

            df_na_mot_leg = df_na_leg[df_na_leg["Nombre descuento"] == motivo].reset_index(drop=True)
            cant_filas = df_na_leg.shape[0]

            #Es [] si ese legajo no estaba en el diccionario, es un lista con 4 elemntos si ya existe
            vector_en_diferencias = diferencias[legajo_str][motivo] 
            #Si ya existía el vector, nos quedamos con ese, sino, lo creamos con 4 elementos en 0
            vector_en_diferencias = vector_en_diferencias if len(vector_en_diferencias) > 0 else [0,0,0,0] 

            if motivo in motivos_no_desglaseados:
                valor = df_na_mot_leg["Días ausencia"].iloc[0]
                nuevo_vector = [0,0,0,valor]

            elif motivo == motivo_presentismo:
                nuevo_vector = [df_na_mot_leg.loc[0,"Llegadas tarde"],df_na_mot_leg.loc[0,"Salidas antes"],df_na_mot_leg.loc[0,"Días sin fichada"],df_na_mot_leg.loc[0,"Días ausencia"]]
                
            else:
                nuevo_vector = [0,0,0,cant_filas]
            
            diferencias[legajo_str][motivo] = suma_vectores(nuevo_vector,vector_en_diferencias)

    return diferencias

# CREAR PLANILLA
def escribir_datos(ws: xl.worksheet.Worksheet, fila: int, col_inicial: int, datos_persona: str, formato: xl.format.Format):
    '''
    Escribe en ws la fila indicada los datos datos_persona a partir de la col col_inicial con el formato dado.
    '''
    # Escribo datos de persona
    #rango_datos = "A" + str(fila) + ":H" + str(fila)
    ws.merge_range(fila, col_inicial, fila, col_inicial + 7, datos_persona, formato)

def escribir_encabezado(ws: xl.worksheet.Worksheet, fila: int, col_inicial: int, encabezados: list[str], formato: xl.format.Format):
    '''
    Escribe el encabezado en la fila y columna inicial indicadas, con el formato dado
    '''
    # Escribo encabezados
    for i in range(col_inicial, len(encabezados) + col_inicial):
        ws.write(fila,i,encabezados[i - col_inicial],formato)

def escribir_fila_tabla_original(ws: xl.worksheet.Worksheet, fila: int, datos: pd.Series): 
    '''
    Escribe en ws la fila indicada los datos de la primera planilla o tabla_original a partir de la col col_inicial con el formato dado.
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
            if nombre not in motivos_no_desglaseados and nombre != motivo_presentismo:
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
    El criterio está escrito en la documentación en Docs de "Documentación - Asistencias" > Descuentos
    '''
    wb = xl.Workbook(buffer)
    ws = wb.add_worksheet()

    encabezados = ['Nombre descuento','Llegadas tarde','Salidas antes','Días sin fichada','Días ausencia','Se descuenta?','Observaciones','Finalizada']
    formato_encabezado = wb.add_format()
    formato_encabezado.set_bg_color("#A4C2F4")
    formato_encabezado.set_bold()

    fila_ws = 0
    legajo_actual = None
    motivos_copiar_completo = motivos_no_desglaseados + [motivo_presentismo]

    df_original = default_to_regular(df_original)

    # Como queremos poner una fila vacía entre cada tabla de legajo con sus motivos de descuento
    # vamos a contar cuantos motivos tenemos por legajo y vamos disminuyendo el contador para añadir la fila vacía.
    cant_filas_por_legajo = df_original["Legajo"].value_counts(dropna=True).to_dict()
    cant_filas_legajo_df = 0
    legajo_en_diferencias = None
    
    for index, row in df_original.iterrows():
        # Como no podemos observar los valores del excel tenemos que ir copiando del df original y copiar al mismo tiempo las diferencias
        # todo esto segun corresponda
        motivo = row["Nombre descuento"]
        if row["Legajo"] != legajo_actual:
            
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
                escribir_fila_tabla_original(ws,fila_ws,row[encabezados])
                cant_filas_legajo_df -= 1
                fila_ws += 1
        else: 
            escribir_fila_tabla_original(ws,fila_ws,row[encabezados])
            cant_filas_legajo_df -= 1
            fila_ws += 1

        if cant_filas_legajo_df == 0: # si no tengo mas filas para el legajo_actual del df original, pongo las diferencias que faltan
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
            fila_ws += 1 # Añado fila en blanco entre personas

    # Si no quedó nada en el df_original pero sí hay algo nuevo no reportado antes en el sistema, añadirlo
    for legajo, descuentos in diferencias.items():
        escribir_datos(ws, fila_ws, 9, dict_personas[legajo], formato_encabezado)
        fila_ws += 1
        escribir_encabezado(ws, fila_ws, 9, encabezados, formato_encabezado)
        fila_ws += 1
        
        for motivo, vector_motivo in descuentos.items():
            if motivo is None:
                continue
                
            # Validación de seguridad frente a None
            if motivos_copiar_completo and motivo in motivos_copiar_completo:
                escribir_fila_tabla_diferencias(ws, fila_ws, motivo, vector_motivo)
                fila_ws += 1
            else:
                for _ in range(0, vector_motivo[3]):
                    escribir_fila_tabla_diferencias(ws, fila_ws, motivo, [0, 0, 0, 1])
                    fila_ws += 1
        fila_ws += 1 # Añado fila en blanco entre personas

    ws.set_column_pixels(0, 0, 340)
    ws.set_column_pixels(1, 5, 102)
    ws.set_column_pixels(6, 6, 290)

    ws.set_column_pixels(9, 9, 340)
    ws.set_column_pixels(10, 14, 102)
    ws.set_column_pixels(15, 15, 290)
    
    wb.close()
    return wb
