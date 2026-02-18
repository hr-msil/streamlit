from datetime import datetime,timedelta
import pandas as pd
import numpy as np
from collections import defaultdict

######################
# Calcular variación #
######################  
# 
#Recibe dos (o más) planillas de horas extras y liquidación de dos meses distintos, el actual y el anterior

#------- Funciones auxiliares -----------
def esta_en_columna(df: pd.DataFrame, columna: str, valor):
  """
  Docstring for esta_en_columna
  
  :param df: dataFrame que tenga la columna que queremos ver
  :param columna: Nombre de la columna en donde se quiere ver si existe el valor
  :param valor: valor puede ser un string o un número que queremos ver si pertenece a la columna dada.
  """
  return df[columna].isin([valor]).any()

def tiene_guion(valor) -> bool:  
    # Verifica si el valor es una cadena y si contiene un guion
    if isinstance(valor, str) and ' - ' in valor:
        return True
    return False

def tipoDeFila(fila) -> int:
  '''
  Si es una persona devuelve un 0
  Si es un tipo de hora devuelve un 1
  Si es una oficina devuelve un 2
  '''
  tipo_fila = -1

  if tiene_guion(fila["Muni"]):

    palabra_dividida = fila['Muni'].split("-")

    primer_codigo = palabra_dividida[0].strip()

    nombre_value = fila['Nombre']

    if not pd.isna(nombre_value):
      tipo_fila = 0
    elif primer_codigo == "MU":
      tipo_fila = 1
    else:
      tipo_fila = 2

  else:
    tipo_fila = 0

  return tipo_fila

def obtener_tipo_de_hora(fila) -> int:
  '''
  Precondición: tipoDeFila(fila) == 1
  '''
  palabra_dividida = fila['Muni'].split("-")

  codigo_hora = palabra_dividida[1].strip()

  return codigo_hora

def obtener_oficina(fila) -> str:
  '''
  Precondición: tipoDeFila(fila) == 2
  '''
  if 'Muni' not in fila or pd.isna(fila['Muni']):
    return ""
  else:
    codigo_oficina = fila['Muni'].split("-")[0]
    return codigo_oficina.strip()

 #--------- Limpieza de datos originales ---------------

def limpiar(data: pd.DataFrame) -> defaultdict:
  """
  Arregla el dataSet para que sea de tipo Nombre, Legajo, Horas_86, Horas_87, Horas_89, valor_86,valor_87, valor_89.
  
  :param data: archivo en formato .xls pasado a dataFrame
  :return: Diccionario con legajo como clave y una lista de valores de la forma [nombre, oficina, legajo, cant_horas_86, valor_horas_86, cant_horas_87...]
  """
  d = defaultdict(list)

  cant_filas = data.shape[0]
  oficina = ""
  tipo_hora_extra = ""

  for i in range(0, cant_filas):

    fila = data.iloc[i]
    tipo_de_fila = tipoDeFila(fila)

    if tipo_de_fila == 1:

      tipo_hora_extra = obtener_tipo_de_hora(fila)

    elif tipo_de_fila == 2:

      oficina = obtener_oficina(fila)

    elif tipo_de_fila == 0:

      legajo = fila['Legajo']
      cant_horas = fila['Cant horas'] # Keep as is, will convert if not NaN
      valor_por_hora = fila['Valor por hora'] # Keep as is, will convert if not NaN
      valor = fila['Valor total'] # Keep as is, will convert if not NaN
      nombre = fila['Nombre']

      # A veces estos valores pueden ser nulos y no aportan significado
      # a la hora de calcular variaciones mes a mes (según nos comentaron
      # en Asistencias), por eso se ignoran
      if np.isnan(cant_horas) or np.isnan(valor_por_hora):
        continue

      # Convert to int/float after checking for NaN
      cant_horas = int(cant_horas)
      valor = float(valor)

      if legajo in d:

        if tipo_hora_extra == "A0786":

          d[legajo][0][3] = int(cant_horas)
          d[legajo][0][4] = float(valor)

        elif tipo_hora_extra == "A0787":

          d[legajo][0][5] = int(cant_horas)
          d[legajo][0][6] = float(valor)

        else:

          d[legajo][0][7] = int(cant_horas)
          d[legajo][0][8] = float(valor)

      else:

        if tipo_hora_extra == "A0786":

          valores = [nombre, oficina, legajo, int(cant_horas), float(valor), 0, 0, 0, 0]
          d[legajo].append(valores)

        elif tipo_hora_extra == "A0787":

          valores = [nombre, oficina, legajo, 0, 0 ,int(cant_horas), float(valor), 0, 0]
          d[legajo].append(valores)

        else:

          valores = [nombre, oficina, legajo, 0, 0, 0, 0, int(cant_horas), float(valor)]
          d[legajo].append(valores)

  return d

#-------- En caso de tener dos liquidaciones -------------

def agregar_liquidacion_extra(d: defaultdict, data: pd.DataFrame) -> None:
  """
  Precondición: tener más de una liquidación en un mes
  d (defaultDict): Diccionario creado con limpiar(data)
  data (DataFrame): DataFrame del archivo .xls correspondiente a la liquidación extra del mes
  """

  cant_filas = data.shape[0]
  tipo_hora_extra = ""
  oficina = ""


  for i in range(0, cant_filas):

    fila = data.iloc[i]

    if tipoDeFila(fila) == 1:

      tipo_hora_extra = obtener_tipo_de_hora(fila)

    elif tipoDeFila(fila) == 2:

      oficina = obtener_oficina(fila)

    elif tipoDeFila(fila) == 0:

      legajo = fila['Legajo']
      cant_horas = fila["Cant horas"]
      valor = fila["Valor total"]
      nombre = fila["Nombre"]
      valor_por_hora = fila['Valor por hora']
      
      #Fijarse si ese leajo es parte del diccionario del primer excel procesado

      if np.isnan(cant_horas) or np.isnan(valor_por_hora):
        continue

      # Convert to int/float after checking for NaN
      cant_horas = int(cant_horas)
      valor = float(valor)

      if legajo in d:
        if tipo_hora_extra == "A0786":
          
          d[legajo][0][3] = d[legajo][0][3] + cant_horas
          d[legajo][0][4] = d[legajo][0][4] + valor

        elif tipo_hora_extra == "A0787":
          
          d[legajo][0][5] = d[legajo][0][5] + cant_horas
          d[legajo][0][6] = d[legajo][0][6] + valor

        elif tipo_hora_extra == "A0789":
          
          d[legajo][0][7] = d[legajo][0][7] + cant_horas
          d[legajo][0][8] = d[legajo][0][8] + valor

      else:

          if tipo_hora_extra == "A0786":

            valores = [nombre, oficina, legajo, cant_horas, valor, 0, 0, 0, 0]
            d[legajo].append(valores)

          elif tipo_hora_extra == "A0787":

            valores = [nombre, oficina, legajo, 0, 0 ,cant_horas, valor, 0, 0]
            d[legajo].append(valores)

          else:

            valores = [nombre, oficina, legajo, 0, 0, 0, 0, cant_horas, valor]
            d[legajo].append(valores)

#------------- Creación de dataset --------------------

def armar_data_set(d: defaultdict) -> pd.DataFrame:
  """
  Pasar de un diccionario con legajo como clave, y una lista de valores a un dataFrame
  
  :param d: Diccionario con clave "legajo" y valores, los datos recolectados 
  :return: Dataframe resumen del diccionario
  """

  dataSetLimpio = pd.DataFrame(columns = ["Nombre", "Oficina", "Legajo",
                                        "HorasExtra_86", "Valor_86",
                                        "HorasExtra_87", "Valor_87",
                                        "HorasExtra_89", "Valor_89"])

  legajos = []
  nombres = []
  oficinas = []
  horas_86 = []
  horas_87 = []
  horas_89 = []
  valor_86 = []
  valor_87 = []
  valor_89 = []

  for key in d.keys():

    valores = d[key][0]

    legajos.append(key)
    nombres.append(valores[0])
    oficinas.append(valores[1])
    horas_86.append(valores[3])
    horas_87.append(valores[5])
    horas_89.append(valores[7])
    valor_86.append(valores[4])
    valor_87.append(valores[6])
    valor_89.append(valores[8])

  dataSetLimpio["Nombre"] = nombres
  dataSetLimpio["Oficina"] = oficinas
  dataSetLimpio["Legajo"] = legajos
  dataSetLimpio["HorasExtra_86"] = horas_86
  dataSetLimpio["HorasExtra_87"] = horas_87
  dataSetLimpio["HorasExtra_89"] = horas_89
  dataSetLimpio["Valor_86"] = valor_86
  dataSetLimpio['Valor_87'] = valor_87
  dataSetLimpio['Valor_89'] = valor_89

  return dataSetLimpio

def agregar_total(df_anterior: pd.DataFrame, df_actual:pd.DataFrame) -> None:
  """
  Agrega las columnas de totales, donde para su respectiva columna se suma la respectiva columna del dataframe anterior y del actual
  
  :param df_anterior: dataframe del mes anterior
  :param df_actual: dataframe del mes actual
  """

  df_anterior["Total horas"] = df_anterior["HorasExtra_86"] + df_anterior["HorasExtra_87"] + df_anterior["HorasExtra_89"]
  df_actual["Total horas"] = df_actual["HorasExtra_86"] + df_actual["HorasExtra_87"] + df_actual["HorasExtra_89"]

  df_anterior["Total valor"] = df_anterior["Valor_86"] + df_anterior["Valor_87"] + df_anterior["Valor_89"]
  df_actual["Total valor"] = df_actual["Valor_86"] + df_actual["Valor_87"] + df_actual["Valor_89"]


# ----------- Archivos necesarios ------------------
# Son los archivos que van a ser devueltos en formato .xlsx

def unir_oficinas(df_anterior: pd.DataFrame, df_actual: pd.DataFrame) -> pd.DataFrame:
  df_area = pd.DataFrame(columns = ["Oficina", "Dif. horas extras normales", "Dif. liquidado normales",
                                  "Dif. horas extras al 50", "Dif. liquidado al 50",
                                  "Dif.horas extras al 100", "Dif. liquidado al 100","Dif. porcentual",
                                  "Dif. porcentual ponderado"])

  oficinas = []
  dif_horas_86 = []
  dif_valor_86 = []
  dif_horas_87 = []
  dif_valor_87 = []
  dif_horas_89 = []
  dif_valor_89 = []
  dif_horas_total = []
  dif_valor_total = []
  dif_porcentual_horas = []
  dif_porcentual_ponderado = []

  oficinas_mes_ant = df_anterior["Oficina"].unique()
  oficinas_mes_act = df_actual["Oficina"].unique()

  oficinas_unicas = np.union1d(oficinas_mes_ant, oficinas_mes_act)

  for i,oficina in enumerate(oficinas_unicas):

    if (esta_en_columna(df_anterior,"Oficina", oficina)) and (not esta_en_columna(df_actual,"Oficina", oficina)):

      # La oficina solo se encuentra en un archivo .xls
      df_oficinas_ant = df_anterior[df_anterior["Oficina"] == oficina]
      cant_oficina_ant = df_oficinas_ant.shape[0]

      cant_oficina_act = 0

      valor_86_ant = df_oficinas_ant["Valor_86"].sum()
      valor_86_act = 0

      valor_87_ant = df_oficinas_ant["Valor_87"].sum()
      valor_87_act = 0

      valor_89_ant = df_oficinas_ant["Valor_89"].sum()
      valor_89_act = 0

      horas_86_ant = df_oficinas_ant["HorasExtra_86"].sum()
      horas_86_act = 0

      horas_87_ant = df_oficinas_ant["HorasExtra_87"].sum()
      horas_87_act = 0

      horas_89_ant = df_oficinas_ant["HorasExtra_89"].sum()
      horas_89_act = 0

      #Total
      horas_total_ant = df_oficinas_ant["Total horas"].sum()
      horas_total_act = 0
      valor_total_ant = df_oficinas_ant["Total valor"].sum()
      valor_total_act = 0

    elif (not esta_en_columna(df_anterior,"Oficina", oficina)) and (esta_en_columna(df_actual,"Oficina", oficina)):

      # La oficina solo se encuentra en un archivo .xls
      df_oficinas_act = df_actual[df_actual["Oficina"] == oficina]
      cant_oficina_act = df_oficinas_act.shape[0]
      cant_oficina_ant = 0

      valor_86_ant = 0
      valor_86_act = df_oficinas_act["Valor_86"].sum()

      valor_87_ant = 0
      valor_87_act = df_oficinas_act["Valor_87"].sum()

      valor_89_ant = 0
      valor_89_act = df_oficinas_act["Valor_89"].sum()

      horas_86_ant = 0
      horas_86_act = df_oficinas_act["HorasExtra_86"].sum()

      horas_87_ant = 0
      horas_87_act = df_oficinas_act["HorasExtra_87"].sum()

      horas_89_ant = 0
      horas_89_act = df_oficinas_act["HorasExtra_89"].sum()

      #Total
      horas_total_ant = 0
      horas_total_act = df_oficinas_act["Total horas"].sum()
      valor_total_ant = 0
      valor_total_act = df_oficinas_act["Total valor"].sum()

    elif esta_en_columna(df_anterior, "Oficina", oficina) and esta_en_columna(df_actual, "Oficina", oficina):
      # La oficina se encuentra en ambos archivos .xls

      df_oficinas_act = df_actual[df_actual["Oficina"] == oficina]
      cant_oficina_act = df_oficinas_act.shape[0]

      df_oficinas_ant = df_anterior[df_anterior["Oficina"] == oficina]
      cant_oficina_ant = df_oficinas_ant.shape[0]

      valor_86_ant = df_oficinas_ant["Valor_86"].sum()
      valor_86_act = df_oficinas_act["Valor_86"].sum()

      valor_87_ant = df_oficinas_ant["Valor_87"].sum()
      valor_87_act = df_oficinas_act["Valor_87"].sum()

      valor_89_ant = df_oficinas_ant["Valor_89"].sum()
      valor_89_act = df_oficinas_act["Valor_89"].sum()

      horas_86_ant = df_oficinas_ant["HorasExtra_86"].sum()
      horas_86_act = df_oficinas_act["HorasExtra_86"].sum()

      horas_87_ant = df_oficinas_ant["HorasExtra_87"].sum()
      horas_87_act = df_oficinas_act["HorasExtra_87"].sum()

      horas_89_ant = df_oficinas_ant["HorasExtra_89"].sum()
      horas_89_act = df_oficinas_act["HorasExtra_89"].sum()

      #Total
      horas_total_ant = df_oficinas_ant["Total horas"].sum()
      horas_total_act = df_oficinas_act["Total horas"].sum()
      valor_total_ant = df_oficinas_ant["Total valor"].sum()
      valor_total_act = df_oficinas_act["Total valor"].sum()

    oficinas.append(oficina)
    #Tipo de hora: 86
    dif_horas_86.append(horas_86_act - horas_86_ant)
    dif_valor_86.append(valor_86_act - valor_86_ant)
    #Tipo de hora:87
    dif_horas_87.append(horas_87_act - horas_87_ant)
    dif_valor_87.append(valor_87_act - valor_87_ant)
    #Tipo de hora:89
    dif_horas_89.append(horas_89_act - horas_89_ant)
    dif_valor_89.append(valor_89_act - valor_89_ant)
    #Total
    dif_horas_total.append(horas_total_act - horas_total_ant)
    dif_valor_total.append(valor_total_act - valor_total_ant)
    #Dif. porcentual
    if horas_total_ant == 0:
      dif_porcentual_horas.append(np.nan)
    else:
      dif_porcentual_horas.append(horas_total_act/horas_total_ant - 1)
    #Dif. porcentual ponderado
    if horas_total_ant == 0:
      dif_porcentual_ponderado.append(np.nan)
    elif horas_total_act == 0:
      dif_porcentual_ponderado.append(np.nan)
    else:
      dif_porcentual_ponderado.append((horas_total_act/cant_oficina_act - horas_total_ant/cant_oficina_ant)/(horas_total_act/cant_oficina_act))

  df_area["Oficina"] = oficinas
  df_area["Dif. horas extras normales"] = dif_horas_86
  df_area["Dif. liquidado normales"] = dif_valor_86
  df_area["Dif. horas extras al 50"] = dif_horas_87
  df_area["Dif. liquidado al 50"] = dif_valor_87
  df_area["Dif.horas extras al 100"] = dif_horas_89
  df_area["Dif. liquidado al 100"] = dif_valor_89
  df_area["Dif. horas total"] = dif_horas_total
  df_area["Dif. valor total"] = dif_valor_total
  df_area["Dif. porcentual"] = dif_porcentual_horas
  df_area["Dif. porcentual ponderado"] = dif_porcentual_ponderado

  return df_area

def unir_personas(df_1, df_2):

  df_personas = pd.DataFrame(columns = ["Legajo", "Nombre", "Oficina","Dif. horas extras", "Dif. liquidado"])

  dif_horas = []
  dif_valor = []
  nombres = []
  legajos = []
  oficinas = []

  legajos_mes_1 = df_1["Legajo"].unique()
  legajos_mes_2 = df_2["Legajo"].unique()

  # Filtra valores NaN y se consideran legajos como valor numérico
  legajos_mes_1 = legajos_mes_1[~pd.isna(legajos_mes_1)]
  legajos_mes_2 = legajos_mes_2[~pd.isna(legajos_mes_2)]

  legajos_unicos = np.union1d(legajos_mes_1, legajos_mes_2)

  for legajo in legajos_unicos:
    legajos.append(legajo)

    #Una persona puede no estar en ambos meses

    if esta_en_columna(df_1, "Legajo", legajo) and esta_en_columna(df_2, "Legajo", legajo):
      #La persona hizo horas extras ambos meses
      fila_persona_1 = df_1[df_1["Legajo"] == legajo].iloc[0]
      fila_persona_2 = df_2[df_2["Legajo"] == legajo].iloc[0]
      oficina_persona = fila_persona_1["Oficina"]

      dif_horas.append(fila_persona_2["Total horas"] - fila_persona_1["Total horas"])

      dif_valor.append(fila_persona_2["Total valor"] - fila_persona_1["Total valor"])

      nombres.append(fila_persona_1["Nombre"])

    elif esta_en_columna(df_1, "Legajo", legajo) and (not esta_en_columna(df_2, "Legajo", legajo)):
      #La persona hizo horas extras solo el mes 1
      fila_persona_1 = df_1[df_1["Legajo"] == legajo].iloc[0]
      oficina_persona = fila_persona_1["Oficina"]

      dif_horas.append(-fila_persona_1["Total horas"])

      dif_valor.append(-fila_persona_1["Total valor"])
      nombres.append(fila_persona_1["Nombre"])

    else:
      # La persona hizo horas extras solo el mes 2
      fila_persona_2 = df_2[df_2["Legajo"] == legajo].iloc[0]
      oficina_persona = fila_persona_2["Oficina"]

      dif_horas.append(fila_persona_2["Total horas"])

      dif_valor.append(fila_persona_2["Total valor"])

      nombres.append(fila_persona_2["Nombre"])
    oficinas.append(oficina_persona)

  df_personas["Legajo"] = legajos
  df_personas["Nombre"] = nombres
  df_personas["Dif. horas extras"] = dif_horas
  df_personas["Dif. liquidado"] = dif_valor
  df_personas["Oficina"] = oficinas

  return df_personas

def resumen_oficinas(df):
  """
  Solo para el mes actual
  df (pd.DataFrame): dataFrame del mes actual
  """

  oficinas_unicas = df["Oficina"].unique()

  df_area_total = pd.DataFrame(columns = ["Oficina", "Total horas normales", "Total liquidado normales",
                                          "Total horas al 50", "Total liquidado al 50", "Total horas al 100",
                                          "Total liquidado al 100"])

  oficinas = []
  total_hora_86 = []
  total_hora_87 = []
  total_hora_89 = []

  total_valor_86 = []
  total_valor_87 = []
  total_valor_89 = []

  for oficina in oficinas_unicas:

    oficinas.append(oficina) 

    df_oficina = df[df["Oficina"] == oficina]

    hora_86 = df_oficina["HorasExtra_86"].sum()
    hora_87 = df_oficina["HorasExtra_87"].sum()
    hora_89 = df_oficina["HorasExtra_89"].sum()

    valor_86 = df_oficina["Valor_86"].sum()
    valor_87 = df_oficina["Valor_87"].sum()
    valor_89 = df_oficina["Valor_89"].sum()

    total_hora_86.append(hora_86)
    total_hora_87.append(hora_87)
    total_hora_89.append(hora_89)

    total_valor_86.append(valor_86)
    total_valor_87.append(valor_87)
    total_valor_89.append(valor_89)

  df_area_total["Oficina"] = oficinas
  df_area_total["Total horas normales"] = total_hora_86
  df_area_total["Total liquidado normales"] = total_valor_86
  df_area_total["Total horas al 50"] = total_hora_87
  df_area_total["Total liquidado al 50"] = total_valor_87
  df_area_total["Total horas al 100"] = total_hora_89
  df_area_total["Total liquidado al 100"] = total_valor_89

  return df_area_total  


def obtener_mes_anterior():
    ahora = datetime.today()

    # Obtener el primer día del mes actual
    primer_dia_mes = ahora.replace(day=1)

    # Restar un día para ir al último día del mes anterior
    mes_anterior_fecha = primer_dia_mes - timedelta(days=1)

    # Formatear como "YYYY-MM"
    mes_ant = mes_anterior_fecha.strftime("%Y-%m")

    return mes_ant