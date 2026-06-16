#------Importación de librerpias---------
import pandas as pd
import sys
import os
from datetime import datetime, date, timedelta
import re
from collections import defaultdict
import streamlit as st
import calendar
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))#Esto es porque crear_df del script chequeo_legajos.py usa helpers.py que no esta en la misma carpeta
from helpers import crear_df
from helpers import obtener_dias_feriados
from pathlib import Path



motivos_no_descontar = ["88 - LICENCIA ANUAL 1ª FRACCION","82 - LICENCIA ANUAL","11 - MATERNIDAD (ARTICULO 42°)","56 - ACCIDENTE DE TRABAJO"]

#Estos son los días de guardia con su respectivo código
dias_guardias = {700: 'Monday', 699 : 'Tuesday', 701: 'Wednesday', 702: 'Thursday', 703: 'Friday', 704: 'Saturday', 705: 'Sunday'}
dias_semana = {'Monday': 0, 'Tuesday':1, 'Wednesday': 2, 'Thursday':3, 'Friday':4, 'Saturday':5, 'Sunday':6}

#siempre sacamos el listado con respecto al mes pasado, es por eso que es ese mes elque nos resulta de interés
hoy = datetime.today()
hoy = hoy.replace(hour=0, minute=0, second=0, microsecond=0)

# Determinar el mes anterior
primer_dia_mes_anterior = (hoy.replace(day=1) - timedelta(days=1)).replace(day=1)
mes_anterior = primer_dia_mes_anterior.month
anio_anterior = primer_dia_mes_anterior.year

dias_feriados_mes = obtener_dias_feriados(primer_dia_mes_anterior.year,primer_dia_mes_anterior.month)

#-------------Funciones auxiliares------------

#Cargamos el txt con los motivos en que el BAP descuenta solo
def leer_motivos_descuento() -> list:
    BASE_DIR = Path(__file__).resolve().parent.parent
    ruta_motivos = BASE_DIR/"motivos_no_descuentan.txt"
    with open(ruta_motivos, "r", encoding="utf-8") as f:
        motivos_que_descuenta = [linea.strip() for linea in f]
    return motivos_que_descuenta

def leer_motivos_enf_propia() -> list:
    BASE_DIR = Path(__file__).resolve().parent.parent
    ruta_motivos = BASE_DIR/"motivos_enfermedad_propia.txt"
    with open(ruta_motivos, "r", encoding="utf-8") as f:
        motivos_enfermedad_propia = [linea.strip() for linea in f]
    return motivos_enfermedad_propia


def obtener_dia_semana(dia: int, mes: int, anio: int) -> str:
    '''
    Dado el día, el mes y el anio, devuelve qué día de la semana es (en inglés)
    '''
    # Crear el objeto fecha
    fecha = date(anio, mes, dia)
    # por defecto en inglés si no se configura
    nombre_dia = fecha.strftime("%A")
    return nombre_dia

def contar_dias(año: int , mes: int)-> list[int]:
    '''
    Dado un anio y mes, devuelve un array contando cuántos lunes, martes, miércoles, etc. hay en la semana
    '''
    # monthrange devuelve: (primer_dia_semana, num_dias)
    # primer_dia_semana: 0=Lunes, 6=Domingo
    _, num_dias = calendar.monthrange(año, mes)
    count_dias_semana = [0,0,0,0,0,0,0]
    dias = [0,1,2,3,4,5,6]

    for dia_num in dias:
        for dia in range(1, num_dias + 1):
            # 0 representa el Lunes
            if calendar.weekday(año, mes, dia) == dia_num:
                count_dias_semana[dia_num] += 1
    return count_dias_semana

def cambiar_fechas(df: pd.DataFrame):
    '''
    Recibe el dataframe de ausencias
    Lo que hace es cambiar las ausencias de forma tal que no se reemplace en
    la tabla las fechas de dia_inicia y dia_fin por el numero de día que correspondería al
    mes anterior (si es que la fecha es del mes pasado).
    '''
    df["dia_inicio"] = pd.to_datetime(df["dia_inicio"],format="%d/%m/%Y")
    df["dia_fin"] = pd.to_datetime(df["dia_fin"],format="%d/%m/%Y")

    # Determinar el mes anterior
    primer_dia_mes_anterior = (hoy.replace(day=1) - timedelta(days=1)).replace(day=1)
    ultimo_dia_mes_anterior = hoy.replace(day=1) - timedelta(days=1)

    # Mes anterior al anterior
    primer_dia_dos_meses = (primer_dia_mes_anterior - timedelta(days=1)).replace(day=1)
    ultimo_dia_dos_meses = primer_dia_mes_anterior - timedelta(days=1)

    def clasificar_mes(fecha):
        if primer_dia_mes_anterior <= fecha <= ultimo_dia_mes_anterior:
            return 1
        elif primer_dia_dos_meses <= fecha <= ultimo_dia_dos_meses:
            return 0
        else:
            return 0

    df["clasificacion_mes"] = df["dia_inicio"].apply(clasificar_mes)

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

#----------- Funciones para leer archivos ---------------

def leer_novedades(novedades_excel:str) -> pd.DataFrame:
    '''
    Se espera que todos los empleados que se encuentran en este archivo son partición ME, tiene @A_TITULO  valor 8 o 9
    Limpia la columna de legajo y lo pasa a int, mismo con la oficina
    '''
    novedades = pd.read_excel(novedades_excel)
    novedades["legajo_limpio"] = novedades["LEGAJO"].str.split("-").str[0].astype('int')
    novedades["oficina_limpia"] = novedades["OFICINA"].str.split("-").str[1].astype('int')

    return novedades

def leer_horarios(horarios_excel: str) -> pd.DataFrame:
    '''
    Devuelve los horarios de los medicos que hacen guardia
    '''
    df = pd.read_excel(horarios_excel)
    df.columns = ["Legajo", "Nombre", "Desde", "Hasta", "Tipo","Ficha", "Descripción", "unnamed"]
      
    df["legajo_limpio"] = df["Legajo"].str.split("-").str[0].astype('int')

    df["codigo_horario"] = df["Descripción"].str.split(" - ").str[0].astype('int')
    
    return df

def transformar_ausencias_a_dict(ausencias : str) -> dict:
    '''
    A partir de las ausencias se arma un diccionario:
    dict[legajo] = { "empleado": string, "dias": [int] }
    donde dias es una lista de numeros de los dias en 
    que esa persona estuvo ausente.
    '''
    df_raw = pd.read_excel(ausencias)

    motivos_que_descuenta = leer_motivos_descuento()

    #hoy = datetime.today()
    #hoy = hoy.replace(hour=0, minute=0, second=0, microsecond=0)

    # Determinar el mes anterior
    primer_dia_mes_anterior = (hoy.replace(day=1) - timedelta(days=1)).replace(day=1)

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
    

    legajo_dict = defaultdict(lambda: {"empleado": None, "dias": [], "motivos": [],"licencia primera frac":0,"dias_escritos":[],"motivo descuenta solo":0,"cambio_de_guardia":0})
    
    for _, row in df.iterrows():
        legajo = str(row["legajo"])
        nombre = row["empleado"]
        oficina = row["oficina"]
        nro_motivo = row["nro_motivo"]
        motivo = row["motivo"]
        dias = list(range(int(row["dia_inicio"]), int(row["dia_fin"]) + 1))

        #ignoramos los motivos LSGS
        if(nro_motivo == 40 or nro_motivo == 803):continue

        empezo_este_mes = row["clasificacion_mes"]

        legajo_dict[legajo]["empleado"] = nombre
        legajo_dict[legajo]["oficina"] = oficina
        
        legajo_dict[legajo]["dias"].extend(dias)
        legajo_dict[legajo]["motivos"].extend([f"{nro_motivo} - {motivo}" for _ in range(len(dias))])
        if motivo in motivos_que_descuenta:
            legajo_dict[legajo]["motivo descuenta solo"] = 1 if legajo_dict[legajo]["motivo descuenta solo"] == 0 else legajo_dict[legajo]["motivo descuenta solo"]
        if nro_motivo == 88:
            if empezo_este_mes == 1:
                legajo_dict[legajo]["licencia primera frac"] = empezo_este_mes if legajo_dict[legajo]["licencia primera frac"] == 0 else legajo_dict[legajo]["licencia primera frac"]
        if nro_motivo == 19:
            legajo_dict[legajo]["cambio_de_guardia"] = 1 if legajo_dict[legajo]["cambio_de_guardia"] == 0 else legajo_dict[legajo]["cambio_de_guardia"]

    for v in legajo_dict.values():
        v["dias"] = sorted(set(v["dias"]))
        dias_escritos = [obtener_dia_semana(d, primer_dia_mes_anterior.month, primer_dia_mes_anterior.year) for d in v["dias"]]
        v["dias_escritos"].extend(dias_escritos)

        
    return dict(legajo_dict)

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

#--------Chequeo de bonificaciones--------

def chequear_BAP_legajo(ausencias: dict, legajo: str, legajos_a_descontar_BAP: list):
    '''
    Dado un legajo se fija si tiene como ausencia el motivo: 88 - LICENCIA ANUAL 1ª FRACCION, además, para ser descontada la bonificación se fija que no tenga alguno de los motivos que descuenta solo.
    '''
    if legajo in ausencias.keys():
        if ausencias[legajo]["licencia primera frac"] == 1:
            if ausencias[legajo]["motivo descuenta solo"] == 0:
                legajos_a_descontar_BAP.append(legajo)


def chequeo_PG_quinta_guardia_legajo(ausencias: dict, cant_dias_en_mes: list, legajo: str, codigo_horario_leg: str, legajos_a_descontar: list,legajosME_que_cobran_quinta_guardia: list):

    if codigo_horario_leg in dias_guardias.keys():
            dia_horario_leg = dias_guardias[codigo_horario_leg]
            dia_horario_leg_num = dias_semana[dia_horario_leg]
            cant_dias_en_mes_horario = cant_dias_en_mes[dia_horario_leg_num]
            if legajo in ausencias.keys():
                if ausencias[legajo]["cambio_de_guardia"] == 0:
                    dias_faltas = ausencias[legajo]["dias_escritos"]
                    cant_faltas_en_guardia = dias_faltas.count(dia_horario_leg)
                    
                    if cant_dias_en_mes_horario == 5:
                        if cant_faltas_en_guardia == 0:
                            legajosME_que_cobran_quinta_guardia.append(legajo)
                        if cant_faltas_en_guardia > 1:
                            legajos_a_descontar.append(legajo)
                    if cant_dias_en_mes_horario == 4:
                        if cant_faltas_en_guardia > 0:
                            legajos_a_descontar.append(legajo)
            else:
                if cant_dias_en_mes_horario == 5: legajosME_que_cobran_quinta_guardia.append(legajo)


def chequeo_decreto_398_y_guardias_feriados_legajo(ausencias:dict, legajo: str, codigo_horario_legajo: str,legajos_a_descontar: list,porcentajes_a_descontar: list,legajosME_que_cobran_feriado: list):
    
    motivos_enfermedad_propia = leer_motivos_enf_propia()
    if codigo_horario_legajo in dias_guardias.keys():

        dia_horario_leg = dias_guardias[codigo_horario_legajo]
        hace_guardia_en_feriado = False
        for feriado in dias_feriados_mes:
                if obtener_dia_semana(feriado, primer_dia_mes_anterior.month, primer_dia_mes_anterior.year) == dia_horario_leg:
                    hace_guardia_en_feriado = True

        if legajo in ausencias.keys():
            if ausencias[legajo]["cambio_de_guardia"] == 0:

                motivos_ausencias_leg = ausencias[legajo]["motivos"]
                dias_ausencia_leg = ausencias[legajo]["dias_escritos"]
                dias_ausencia_num_legajo = ausencias[legajo]["dias"]
                indices_ausencia_por_guardia = [i for i, dia in enumerate(dias_ausencia_leg) if dia == dia_horario_leg]
                falta_en_guardia_feriado = False
                for idx_ausencia_guardia in indices_ausencia_por_guardia:
                    for feriado in dias_feriados_mes:
                        if dias_ausencia_num_legajo[idx_ausencia_guardia] == feriado:
                            falta_en_guardia_feriado = True
                if hace_guardia_en_feriado == True and not falta_en_guardia_feriado:
                    legajosME_que_cobran_feriado.append(legajo)

                
                ausencias_totales_en_guardia = len(indices_ausencia_por_guardia)
                contar_ausencias_sin_descontar = 0
                contar_ausencias_enf_propia = 0
                porcentaje_max = 0
                
                for idx_ausencia in indices_ausencia_por_guardia:

                    motivo = motivos_ausencias_leg[idx_ausencia]
                    
                    if motivo in motivos_no_descontar:
                        contar_ausencias_sin_descontar += 1
                    elif motivo in motivos_enfermedad_propia:
                        contar_ausencias_enf_propia += 1
        
                if contar_ausencias_enf_propia == 2:
                    porcentaje_max = 25 if porcentaje_max < 25 else porcentaje_max
                elif ausencias_totales_en_guardia - contar_ausencias_enf_propia - contar_ausencias_sin_descontar > 0:
                    porcentaje_max = 50 if porcentaje_max < 50 else porcentaje_max
                elif contar_ausencias_enf_propia == 3:
                    porcentaje_max = 50 if porcentaje_max < 50 else porcentaje_max
                elif contar_ausencias_enf_propia == 4:
                    porcentaje_max = 75 if porcentaje_max < 75 else porcentaje_max
                if porcentaje_max > 0:
                    legajos_a_descontar.append(legajo)
                    porcentajes_a_descontar.append(porcentaje_max)

        else:
            if hace_guardia_en_feriado == True: legajosME_que_cobran_feriado.append(legajo)



def armar_listados(archivo_novedades: str,archivo_horarios_por_ofi:str, archivo_ausencias:str,archivo_listado_empleados:str):

    novedades = leer_novedades(archivo_novedades)
    horarios_guardias = leer_horarios(archivo_horarios_por_ofi)
    ausencias = transformar_ausencias_a_dict(archivo_ausencias)
    empleados_por_ofi = listadoPorEmpleados(archivo_listado_empleados)
    

    cant_dias_en_mes = contar_dias(primer_dia_mes_anterior.year,primer_dia_mes_anterior.month)
    
    df_horarios_ME = pd.merge(novedades,horarios_guardias[["legajo_limpio", "codigo_horario"]],how='left',left_on=["legajo_limpio"],right_on=["legajo_limpio"])

    legajos_descontar_BAP = []
    cambios_de_guardia = []
    legajos_a_descontar_PG = []
    legajos_ME_quinta_guardia = []
    legajos_sin_horario = []
    legajos_a_descontar_398 = []
    porcentajes_a_descontar = []
    legajos_ME_feriado = []
    legajos_con_cambio_de_guardia = []

    for row in df_horarios_ME.itertuples():
        legajo = str(row.legajo_limpio)
        codigo_horario_legajo = row.codigo_horario

        if codigo_horario_legajo in dias_guardias.keys():
            
            if (legajo in ausencias.keys()) and (ausencias[legajo]["cambio_de_guardia"] == 1):
                legajos_con_cambio_de_guardia.append(legajo)

            else:

                chequear_BAP_legajo(ausencias, legajo, legajos_descontar_BAP)
                chequeo_PG_quinta_guardia_legajo(ausencias, cant_dias_en_mes, legajo, codigo_horario_legajo,legajos_a_descontar_PG, legajos_ME_quinta_guardia)
                chequeo_decreto_398_y_guardias_feriados_legajo(ausencias, legajo, codigo_horario_legajo, legajos_a_descontar_398, porcentajes_a_descontar, legajos_ME_feriado)

        else:
            legajos_sin_horario.append(legajo)

    listas = {
        'BAP': legajos_descontar_BAP,
        'PG': legajos_a_descontar_PG,
        '5taGUARDIA': legajos_ME_quinta_guardia,
        'DTO398': [legajos_a_descontar_398,porcentajes_a_descontar],
        'GUARDIA_FERIADO': legajos_ME_feriado,
        'LEGAJOS SIN HORARIO': legajos_sin_horario,
        'CAMBIO_GUARDIA': cambios_de_guardia
    }

    legajos_a_reportar = reportar_legajos_sin_horario(legajos_sin_horario, df_horarios_ME)

    return listas,legajos_a_reportar

def reportar_legajos_sin_horario(legajos:list, df_horarios_ME: pd.DataFrame):
    lista_outputs = []
    for legajo in legajos:
        resultado = df_horarios_ME[df_horarios_ME["legajo_limpio"] == int(legajo)]["codigo_horario"]
        if not resultado.empty and pd.notna(resultado.iloc[0]):
            horario = int(resultado.iloc[0])
        else:
            horario = "No definido"
        reporte = f"El legajo: {str(legajo)} tiene un horario que no es posible procesar. Este es el código: {horario}"
        lista_outputs.append(reporte)
    return lista_outputs




if __name__ == "__main__":

    novedades = leer_novedades(r"archivos_junio\novedades me.xls")
    horarios_guardias = leer_horarios(r"archivos_junio\horariosporoficina.xls")
    ausencias = transformar_ausencias_a_dict(r"archivos_junio\ausencias mili.xls")
    empleados_por_ofi = listadoPorEmpleados(r"archivos_junio\listemploficina (1).xls")
    

    cant_dias_en_mes = contar_dias(primer_dia_mes_anterior.year,primer_dia_mes_anterior.month)
    
    df_horarios_ME = pd.merge(novedades,horarios_guardias[["legajo_limpio", "codigo_horario"]],how='left',left_on=["legajo_limpio"],right_on=["legajo_limpio"])

    df_horarios_ME.to_excel("horarios_medicos_guardia.xlsx")



    

