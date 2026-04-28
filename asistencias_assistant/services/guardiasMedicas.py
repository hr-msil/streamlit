#------Importación de librerpias---------
import numpy as np
import pandas as pd
import sys
import os
from datetime import datetime, date, timedelta
import re
from collections import defaultdict
import streamlit as st

import calendar
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
#Esto es porque crear_df del script chequeo_legajos.py usa helpers.py que no esta en la misma carpeta
from helpers import crear_df
from helpers import obtener_dias_feriados

import json


#------ Variables globales-----------------
oficinas_tecnicos = [724, 714, 743] #cobran por 5tas guardias, y por hacer guardias días feriados

oficinas_MU = [721, 711, 741, 722, 712, 742, 723, 713, 744, 724, 714, 743] #Cobran BAP
motivos_que_descuenta = ['1 - CITACIÓN JUNTA MEDICA','2 - LIC. DEPORTIVA','3 - LIC. X PROFILAXIS','4 - ENFERMEDAD JUSTIFICADA','5 - ENFERMEDAD SIN JUSTIFICAR',
                             '6 - FAMILIAR ENFERMO SIN JUSTIFICAR','7 - INUNDACION','8 - FALTA CON AVISO','9 - FALTA SIN AVISO','10 - SUSPENSION',
                             '11 - MATERNIDAD (ARTICULO 42°)','13 - FRANCO COMPENSATORIO SIN JUSTIFICAR','14 - PARO','15 - ESTUDIO JUSTIFICADO',
                             '17 - JUNTA MEDICA (ARTICULO 45°)','21 - EXAMEN JUSTIFICADO','22 - EXAMEN SIN JUSTIFICAR','23 - ARTICULO 45° (PARTICULAR)',
                             '26 - LICENCIA POR MATRIMONIO','28 - AUSENCIA P/OTRA RELIGION','30 - LICENCIA COMPL.-CARRERA MEDICA','31 - FAMILIAR ENFERMO JUSTIFICADO',
                             '32 - ESTUDIOS SIN JUSTIFICAR','33 - ARTICULO 62 °  (NACIMIENTO / ADOPCION)','34 - ARTICULO 64 ° (PREMATRIMONIAL)',
                             '35 - DONACION DE SANGRE (ART. 64)','36 - ARTICULO 45 ° (LIC.ADOPCION)','37 - FALTA SIN AVISO + 5 DIAS','39 - FALTA SIN AVISO + 10 DIAS',
                             '41 - LICENCIA C/GOCE SDO.','42 - DUELO (CONYUGE,HIJO,HIJASTRO)','43 - DUELO (PADRES,HERMANOS, PADRASTROS)',
                             '44 - DUELO (ABUELOS,NIETOS,ETC)','46 - PRESIDIR MESA EXAMENES','48 - ASISTE A CONGRESO/JORNADA/TALLER',
                             '50 - ENFERMO SIN JUSTIFICAR + 5 DIAS','51 - ACCIDENTE DE TRABAJO SIN JUSTIFICAR','52 - CAMBIO A TAREAS LIVIANAS',
                             '53 - ABANDONO DE SERVICIO','54 - ART.32° (11757) LIC.P/JUB.','56 - ACCIDENTE DE TRABAJO','57 - ENFERMO SIN JUSTIFICAR + 10 DIAS',
                             '58 - AUSENTE SIN AVISO +','59 - PARO DE TRANSPORTE','61 - LICENCIA GREMIAL','62 - DONACION SANGRE S/JUSTIFICAR',
                             '63 - DUELOS VARIOS S/ JUSTIFICAR','65 - LICENCIA POR MATRIMONIO SIN JUSTIFICAR','75 - INCENDIO','84 - ENFERMEDAD COMPENSADA',
                             '100 - LLEGADA TARDE + DE 1 HORA','101 - SALIDA ANTICIPADA + DE 1 HORA','104 - PARO SIN JUSTIFICAR',
                             '106 - SALIDAS EDUCATIVAS  ART. 5 RES. 498/10','107 - POSTERGA EXAMEN','109 - SUMARIO CON GOCE DE SUELDO','110 - ART. 38 CUMPLEAÑOS',
                             '111 - ART 63 ASUNTOS PARTICULARES','119 - ART 48 J.M. LA PLATA ORDENANZA 8850/15','122 - ART. 62 NACIMIENTO PERSONAL MASCULINO',
                             '125 - ENFERMO MAS 10 DIAS','127 - LIC.COM.MEDICOS COVID 19','129 - DIA GREMIAL SIN JUSTIFICAR','160 - EN PROCESO ART32/70','502 - NUEVO EXAMEN JUNTA MEDICA',
                             '509 - NUEVO EXAMEN POR PROFILAXIS','601 - NUEVO EXAMEN SIN JUSTIFICAR','603 - ENVIO DE NOTIFICACION RRHH/ MEDICINA LAB','772 - TRAMITES VARIOS SIN JUSTIFICAR',
                             '773 - SUSPENSON DISCIPLINARIA','774 - CONGRESO/JORNADA /TALLER SIN JUSTIFICAR','775 - PERMISO ENTRADA/SALIDA C/ POE S/ JUST.','776 - EXEDIDO/A EN HORAS GREMIALES',
                             '777 - LIC. POR CUIDADO DE RECIEN NACIDO (DO)','778 - ARTICULO 114 O.4 DOCENTE','780 - ART. 114 A (ENFERMO SIN JUSTIFICAR)','781 - ART. 114 F. FAMILIAR ENFERMO SIN JUST.',
                             '782 - ART. 114 LL  EXAMEN SIN JUSTIFICAR','783 - ART- 114 LL EXAMEN JUSTIFICADO','784 - ART. 114 LL ESTUDIO JUSTIFICADO','785 - ART. 114 LL ESTUDIO SIN  JUSTIFICAR',
                             '786 - ADHESION PARO DOCENTE','787 - ART. 114 J (DUELO SIN JUSTIFICAR)','788 - ART. 114 J (DUELO)','789 - ART. 115 B3 (REPRE. GREMIAL JUSTIFICADO',
                             '790 - ART.115 B3 (REPRE. GREMIAL S/JUSTIFICAR','791 - ART. 114 G DONACION DE SANGRE','793 - ART 114 M CITACION AUTORIDAD SIN JUST.','794 - ART.114 LL 1.3 PRACTICAS DO OBLIG JUST.',
                             '795 - ART.114 LL 1.3 PRACTICA DO OBLIG S/JUST.','796 - ART. 114 M PLENARIA JUSTIFICADA','797 - ART. 114 M PLENARIA SIN JUSTIFICAR','800 - PARO DOCENTE DESCUENTO','802 - ART. 114 C LICENCIA POR MATRIMONIO',
                             '900 - PRESENTISMO PUNTUALIDAD (PROCESO DESCUEN','990 - AUSENCIA JUSTIFICADA POR EL MEDICO']
motivo_primera_fracion = "88 - LICENCIA ANUAL 1ª FRACCION"
licencia_matrimonio = "26 - LICENCIA POR MATRIMONIO"
motivos_enfermedad_propia = ['4 - ENFERMEDAD JUSTIFICADA','5 - ENFERMEDAD SIN JUSTIFICAR','17 - JUNTA MEDICA (ARTICULO 45°)',
                             '119 - ART 48 J.M. LA PLATA ORDENANZA 8850/15','120 - JUNTA MEDICA AL 50%',
                             '121 - JUNTA MEDICA 100% DESCUENTO','140 - JUNTA MEDICA 50% SIN JUSTIFICAR','141 - JUNTA MEDICA 100% SIN JUSTIFICAR']

#TODO chequear si me falta alguno
motivos_no_descontar = ["88 - LICENCIA ANUAL 1ª FRACCION","82 - LICENCIA ANUAL","11 - MATERNIDAD (ARTICULO 42°)","56 - ACCIDENTE DE TRABAJO"]
#Para cada código de guardia, me fijo a que día corresponde
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

#lectura de archivo json con códigos de guardias de técnicos
#with open(r"C:\Users\mmaurer\Desktop\Proyectos Python\asistencias_assistant\horarios_tecnicos.json", "r", encoding = "utf-8") as f:
#    dicc_horarios_tecnicos = json.load(f)



#-------------Funciones auxiliares------------

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

def leer_horarios(horarios_excel: str, df_novedades) -> pd.DataFrameataFrame:
    '''
    Devuelve los horarios de los medicos que hacen guardia
    '''
    df = pd.read_excel(horarios_excel)
    df["legajo_limpio"] = df["LEGAJO"].str.split("-").str[0].astype('int')

    df["codigo_horario"] = df["Descripción del Horario"].str.split(" - ").str[0].astype('int')
    
    #df = df[df["legajo_limpio"].isin(df_novedades["legajo_limpio"])]

    return df

def transformar_ausencias_a_dict(ausencias : str) -> dict:
    '''
    A partir de las ausencias se arma un diccionario:
    dict[legajo] = { "empleado": string, "dias": [int] }
    donde dias es una lista de numeros de los dias en 
    que esa persona estuvo ausente.
    '''
    df_raw = pd.read_excel(ausencias)

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
    

    legajo_dict = defaultdict(lambda: {"empleado": None, "dias": [], "motivos": [],"licencia primera frac":0,"dias_escritos":[],"motivo descuenta solo":0})
    
    for _, row in df.iterrows():
        legajo = str(row["legajo"])
        nombre = row["empleado"]
        oficina = row["oficina"]
        nro_motivo = row["nro_motivo"]
        motivo = row["motivo"]
        dias = list(range(int(row["dia_inicio"]), int(row["dia_fin"]) + 1))

        empezo_este_mes = row["clasificacion_mes"]

        legajo_dict[legajo]["empleado"] = nombre
        legajo_dict[legajo]["oficina"] = oficina
        
        legajo_dict[legajo]["dias"].extend(dias)
        legajo_dict[legajo]["motivos"].extend([f"{nro_motivo} - {motivo}" for _ in range(len(dias))])
        if motivo in motivos_que_descuenta:
            legajo_dict[legajo]["motivo descuenta solo"] = 1 if legajo_dict[legajo]["motivo descuenta solo"] == 0 else legajo_dict[legajo]["motivo descuenta solo"]
        if(nro_motivo == 88):
            if empezo_este_mes == 1:
                legajo_dict[legajo]["licencia primera frac"] = empezo_este_mes if legajo_dict[legajo]["licencia primera frac"] == 0 else legajo_dict[legajo]["licencia primera frac"]

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

def chequear_BAP(empleados_por_ofi:pd.DataFrame,novedades:pd.DataFrame,ausencias:dict) -> list:
    '''
    Dado el listado de empleados por oficina, las novedades que nos dicen los médicos con partición ME que hacen guardis, y el listado de ausencias, no dicen que 
    empleados tanto ME como MU tienen motivo de ausencia "LIC. ANUAL 1ERA FRACCION" en el mes que se está liquidando. Hay algunos motivos que no descuentan
    Devuelven los legajos a los que se les descuenta

    '''
    
    empleados_por_ofi_MU = empleados_por_ofi[empleados_por_ofi["Oficina"].isin(oficinas_MU)]

    #Tomamos los MU que pueden cobrar BAP
    legajos_a_descontar = [] #Armamos una lista con los legajos que se pueden descontar

    for _,row in empleados_por_ofi_MU.iterrows():
        #Si la persona no esta en el diccionario  de ausencias entonces está bien que lo cobre, ahora, si tiene, me tengo que fijar si "licencia primera frac" es 1
        legajo = str(row["Legajo"])
        if legajo in ausencias.keys():
            #Tiene alguna ausencia
            motivos_ausencia = list(set(ausencias[legajo]["motivos"]))
            if ausencias[legajo]["licencia primera frac"] == 1:
                if ausencias[legajo]["motivo descuenta solo"] == 0: #No tiene niingún otro motivo que se descuente solo
                    legajos_a_descontar.append(legajo)

    for _,row in novedades.iterrows():
        legajo = str(row["legajo_limpio"])
        if legajo in ausencias.keys():
            motivos_ausencia = list(set(ausencias[legajo]["motivos"]))
            if ausencias[legajo]["licencia primera frac"] == 1:
                if ausencias[legajo]["motivo descuenta solo"] == 0: #Si no hay ningún motivo que sea razón de descontar, se descuenta
                    legajos_a_descontar.append(legajo)

    return legajos_a_descontar


def chequeo_PG_y_quinta_guardia(df_horarios_ME:pd.DataFrame,ausencias:dict,cant_dias_en_mes:list) :
    '''
    Dado un dataFrame con todos los medicos ME que hacen guardias, el listado de ausencias, y un array de cant de días por mes. Si el mes tiene 5 días en los que 
    hacer guardia y hace menos de 4, se descuenta, si el mes tiene 4 días y hace menos de 4, se descuenta también. También, aquellos que hacen 5ta guardia, 
    los agregamos a una lista de legajos que cobran 5ta guardia.
    Devuelve los legajos a los que se les descuenta bonificación PG
    Devuelve los legajos que cobran por 5ta guardia
    '''
   
    legajos_a_descontar = []
    legajosME_que_cobran_quinta_guardia = []
    legajos_sin_horario = []
    #Acá tengo todos los médicos que hacen guardias con sus respectivas oficinas
    for row in df_horarios_ME.itertuples():
        #TODO agregar que si no encuentra el horario o no lo puede procesar diga legajo - horario

        legajo = str(row.legajo_limpio)
        
        
        codigo_horario_leg = row.codigo_horario
        if codigo_horario_leg in dias_guardias.keys():
            dia_horario_leg = dias_guardias[codigo_horario_leg]
            dia_horario_leg_num = dias_semana[dia_horario_leg]
            cant_dias_en_mes_horario = cant_dias_en_mes[dia_horario_leg_num]
            if legajo in ausencias.keys():
                dias_faltas = ausencias[legajo]["dias_escritos"]
                cant_faltas_en_guardia = dias_faltas.count(dia_horario_leg)
                if legajo == '65830': 
                    #Tiene una llegada tarde el día que hace guardia
                    st.write(f"Día de guardia del legajo: {dia_horario_leg}")
                    st.write(f"Los días en que faltó el legajo son: {dias_faltas}")
                    st.write(f"La cant de días que  el legajo faltó en su día de guardia: {cant_faltas_en_guardia}")

                
                if cant_dias_en_mes_horario == 5:
                    if cant_faltas_en_guardia == 0:
                        legajosME_que_cobran_quinta_guardia.append(legajo)
                    
                    if cant_faltas_en_guardia > 1:
                        legajos_a_descontar.append(legajo)

                if cant_dias_en_mes_horario == 4:
                    if cant_faltas_en_guardia > 0:
                        legajos_a_descontar.append(legajo)
    

            else:
                #TODO tendría que ver a cuáles agrega acá porque les falta horario de guardia
                if legajo == '65830': 
                    st.write(f"Día de guardia del legajo: {dia_horario_leg}")
                    st.write(f"Los días en que faltó el legajo son: {dias_faltas}")
                    st.write(f"La cant de días que  el legajo faltó en su día de guardia: {cant_faltas_en_guardia}")
                if cant_dias_en_mes_horario == 5: legajosME_que_cobran_quinta_guardia.append(legajo)
        else:
            legajos_sin_horario.append(legajo)

    return legajos_a_descontar,legajosME_que_cobran_quinta_guardia,legajos_sin_horario

def chequeo_decreto_398_y_guardias_feriados(df_horarios_ME:pd.DataFrame,ausencias:dict):
    '''
    Devuelve una lista de los legajos a los que se les descuenta el Dto 398
    Devuelve una lista de los legajos que cobran bonificación por hacer guardias en feriados
    '''
    #df_horarios_ME = pd.merge(novedades,horarios_guardias[["legajo_limpio", "codigo_horario"]],how='left',left_on=["legajo_limpio"],right_on=["legajo_limpio"])
    legajos_a_descontar = []
    porcentajes_a_descontar = []
    legajosME_que_cobran_feriado = []
    for row in df_horarios_ME.itertuples():

        legajo = str(row.legajo_limpio)
        codigo_horario_legajo = row.codigo_horario

        if codigo_horario_legajo in dias_guardias.keys():

            dia_horario_leg = dias_guardias[codigo_horario_legajo]
            hace_guardia_en_feriado = False
            for feriado in dias_feriados_mes:
                    if obtener_dia_semana(feriado, primer_dia_mes_anterior.month, primer_dia_mes_anterior.year) == dia_horario_leg:
                        hace_guardia_en_feriado = True

            if legajo in ausencias.keys():

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
                    
                    if motivo == licencia_matrimonio:
                        porcentaje_max = 100
                        continue
                    
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
                elif contar_ausencias_enf_propia == 5:
                    porcentaje_max = 100 if porcentaje_max < 100 else porcentaje_max
                if porcentaje_max > 0:
                    legajos_a_descontar.append(legajo)
                    porcentajes_a_descontar.append(porcentaje_max)

            else:
                if hace_guardia_en_feriado == True: legajosME_que_cobran_feriado.append(legajo)
                
    return legajos_a_descontar,porcentajes_a_descontar,legajosME_que_cobran_feriado

def armar_listados(archivo_novedades,archivo_horarios_por_ofi, archivo_ausencias,archivo_listado_empleados):

    novedades = leer_novedades(archivo_novedades)
    #listado de todos los empleados partición ME que hacen guardias
    horarios_guardias = leer_horarios(archivo_horarios_por_ofi,novedades)
    ausencias = transformar_ausencias_a_dict(archivo_ausencias)
    empleados_por_ofi = listadoPorEmpleados(archivo_listado_empleados)
    df_tecnicos = empleados_por_ofi[empleados_por_ofi["Oficina"].isin(oficinas_tecnicos)]

    df_horarios_tecnicos = pd.merge(df_tecnicos,horarios_guardias[["Descripción del Horario","legajo_limpio","codigo_horario"]],how='left',left_on=["Legajo"], right_on=["legajo_limpio"])

    cant_dias_en_mes = contar_dias(primer_dia_mes_anterior.year,primer_dia_mes_anterior.month)
    #Si a horarios guardias comento la lista que filtra por novedades me da lo mismo que hacer el merge
    df_horarios_ME = pd.merge(novedades,horarios_guardias[["legajo_limpio", "codigo_horario"]],how='left',left_on=["legajo_limpio"],right_on=["legajo_limpio"])

    legajos_descontar_BAP = chequear_BAP(empleados_por_ofi, novedades, ausencias)
    legajos_a_descontar_PG,legajosME_que_cobran_quinta_guardia,legajos_sin_horario = chequeo_PG_y_quinta_guardia(df_horarios_ME, ausencias, cant_dias_en_mes)
    legajos_a_descontar_398,porcentajes_a_descontar,legajosME_que_cobran_feriado = chequeo_decreto_398_y_guardias_feriados(df_horarios_ME,ausencias)

    listas = {
        'BAP': legajos_descontar_BAP,
        'PG': legajos_a_descontar_PG,
        '5taGUARDIA': legajosME_que_cobran_quinta_guardia,
        'DTO398': [legajos_a_descontar_398,porcentajes_a_descontar],
        'GUARDIA_FERIADO': legajosME_que_cobran_feriado,
        'LEGAJOS SIN HORARIO': legajos_sin_horario
    }

    legajos_a_reportar = reportar_legajos_sin_horario(legajos_sin_horario, df_horarios_ME)

    return listas,legajos_a_reportar


def reportar_legajos_sin_horario(legajos, df_horarios_ME):
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
    
def leer_archivo_horarios_tecnicos(archivo_horarios_tecnicos):
    horarios_tecnicos_guardia = pd.read_excel(archivo_horarios_tecnicos)
    horarios_tecnicos_guardia = horarios_tecnicos_guardia.iloc[:,0:2]
    return horarios_tecnicos_guardia

def leer_novedades_tecnicos_guardia(archivo_novedades_tecnicos_guardia):
    #Estos son los técnicos que reciben las bonificaciones
    novedades_tecnicos_guardia = pd.read_excel(archivo_novedades_tecnicos_guardia)
    novedades_tecnicos_guardia["oficina_limpio"] = novedades_tecnicos_guardia["OFICINA"].str.split(" - ").str[1].astype('int')
    novedades_tecnicos_guardia["legajo_limpio"] = novedades_tecnicos_guardia["LEGAJO"].str.split("-") .str[0].astype('int')

    novedades_tecnicos_guardia_oficinas = novedades_tecnicos_guardia[novedades_tecnicos_guardia["oficina_limpio"].isin(oficinas_tecnicos)]
    return novedades_tecnicos_guardia_oficinas

def bonif_tecnicos(ausencias):
    novedades_tecnicos = leer_novedades_tecnicos_guardia("novedades_tecnicos_mu.xls")
    horarios_tecnicos = leer_archivo_horarios_tecnicos("horariosTecnicosDeGuardia.xlsx")
    data_tecnicos = pd.merge(novedades_tecnicos, horarios_tecnicos, how="left", left_on=["legajo_limpio"], right_on=["labo_Codigo"])

    # Obtener feriados del mes anterior
    hoy = date.today()
    primer_dia_mes_anterior = (hoy.replace(day=1) - timedelta(days=1)).replace(day=1)
    dias_feriados_mes = obtener_dias_feriados(primer_dia_mes_anterior.year, primer_dia_mes_anterior.month)

    # Convertir feriados a día de semana (0=lunes, 6=domingo)
    feriados_como_fechas = [
        obtener_dia_semana(d,primer_dia_mes_anterior.month, primer_dia_mes_anterior.year)
        for d in dias_feriados_mes
    ]
    print(f"Los días feriados del mes corresponden a los días {feriados_como_fechas}")

    cobra_guardias_feriados = []

    for row in data_tecnicos.itertuples():
        legajo = str(row.legajo_limpio)
        horario_legajo = str(row.hora_Codigo)

        try:
            horario_legajo = str(int(float(row.hora_Codigo))).strip()
        except (ValueError, TypeError):
            horario_legajo = None
        
        dias_que_trabaja = dicc_horarios_tecnicos.get(horario_legajo, None)
        if dias_que_trabaja is None:
            print(f"Legajo {legajo}: no se encontró horario '{horario_legajo}'")
            continue
        
        trabaja_feriado = set(dias_que_trabaja) and set(feriados_como_fechas)

        if trabaja_feriado:

            if legajo in ausencias.keys():
                if set(dias_feriados_mes) and set(ausencias[legajo]["dias"]):
                    #tiene alguna ausencia el día de feriado
                    continue
                else:
                    #corresponde cobrar guardias feriados
                    cobra_guardias_feriados.append(legajo) 
            else:
                cobra_guardias_feriados.append(legajo)

    return cobra_guardias_feriados


    
