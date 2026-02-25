import streamlit as st
import pandas as pd
import io

from helpers import normalizar_planilla_hhee
from helpers import procesar_oficinas
from helpers import obtener_hoja_planilla
from helpers import transformar_ausencias_a_dict

import services.armado_csv as a_csv
import services.chequeo_legajos as cl
import services.extraextra as xx
import services.variacion_intermensual as v_mam

import services.chequeo_legajos as cl

st.set_page_config(page_title="Horas extras", page_icon="🕰️", layout = 'wide')

tab1, tab2, tab3, tab4= st.tabs(["🗂️ Chequeo de legajos","📝 Armado del CSV","Variación intermensual","Extra Extra"])

with tab1:
   
    st.subheader("🗂️ Chequeo de legajos por oficina")

    oficinas = None
    st.markdown("Ingresá las **oficinas** en un listado con comas, si querés indicar rangos de oficinas separalas por un guion. No uses espacios entre cada uno.")
    st.markdown("*Por ejemplo si ingresás '100-102,200,310' es que querés procesar las oficinas 100, 101, 102, 200 y 310.*")

    oficinas = st.text_area("Escribí las oficinas y presiona Ctrl + Enter")
    oficinas = procesar_oficinas(oficinas)

    st.divider()
    
    st.markdown("Subí el **archivo de los legajos con todas las oficinas**.")
    st.markdown("**Camino**: Informes > Informes de Empleados > Empleados por Oficina | **Formato**: Excel Extended (que no tenga filas en blanco).")
    archivo_legajos_oficina = st.file_uploader("", type = "xls", key = "archivo_legajos_oficina")

    st.divider()

    st.markdown("Subí la **planilla correspondiente a las horas extras** de la/s oficina/s ingresadas arriba.")
    archivo_hhee_oficina = st.file_uploader("", type = "xls", key = "archivo_hhee_oficina")

    if archivo_legajos_oficina and archivo_hhee_oficina:

        # Procesamos archivo de legajos y oficinas para crear al dataframe legajo - oficina
        df_legajos_oficina_original = cl.leer_archivo_leg_of(archivo_legajos_oficina)
        
        # Procesamos planilla de HHEE para tener los legajos de las personas
        hoja_planilla = obtener_hoja_planilla(archivo_hhee_oficina)
        planilla_hhee = pd.read_excel(archivo_hhee_oficina, sheet_name = hoja_planilla)
        df_hhee_norm = normalizar_planilla_hhee(planilla_hhee)

        # Función principal - buscamos los legajos de la planilla de HHEE 
        # para ver si coinciden con las oficinas declaradas
        cl.reportar_legajos(df_hhee_norm, df_legajos_oficina_original, oficinas)

with tab2:
    st.subheader("📝 Comparación con ausencias y armado del CSV")

    st.markdown("Subí la **planilla de horas extras** que querés procesar.")
    planilla_csv = st.file_uploader("", type = "xls", accept_multiple_files = False, key = "planilla_csv")
    
    st.divider()

    st.markdown("Subí el **listado de ausencias** (puede incluir todas las oficinas). Recordá que hay que hacer el cálculo antes de exportarlo.")
    st.markdown("**Camino**: Informes > Informes de Asistencia > Ausencias por Oficina | **Formato**: Excel Extended o Excel (no tabular).")
    ausencias = st.file_uploader("", type = 'xls', accept_multiple_files = False, key = "ausencias")

    st.divider()
    nombre_archivo = st.text_input("Escribí el nombre del archivo csv que querés generar")

    if planilla_csv and ausencias and nombre_archivo:

        # procesamos la planilla de hhe
        hoja_planilla = obtener_hoja_planilla(planilla_csv)
        planilla_hhee = pd.read_excel(planilla_csv, sheet_name = hoja_planilla, engine = "calamine")
        planilla_hhee = normalizar_planilla_hhee(planilla_hhee)
        resumen_planilla_pre = a_csv.transformar_hhee_a_csv(planilla_hhee)
        
        # procesamos las ausencias 
        ausencias_ofi = transformar_ausencias_a_dict(ausencias,es_viaticos=False)

        # comparamos la planilla de hhee con las ausencias y reportamos inconsistencias
        legajos_inconsistencias = a_csv.anular_unidad_por_ausencias(ausencias_ofi,planilla_hhee)
        a_csv.reportar_inconsistencias_hhee(legajos_inconsistencias,ausencias_ofi) #si la lista es vacía, no devolverá nada

        # reportamos diferencias
        resumen_planilla_pos = a_csv.transformar_hhee_a_csv(planilla_hhee)
        a_csv.reportar_diferencias_entre_planillas(resumen_planilla_pre,resumen_planilla_pos)

        # Eliminamos los legajos que no tengan horas extras
        resumen_planilla_final = a_csv.eliminar_legajo_sin_hhee(resumen_planilla_pos) 

        # Transformamos para descarga
        csv = resumen_planilla_final.to_csv(index=False).encode('latin1')
        st.download_button(
            label="Descargar CSV",
            data=csv,
            file_name=f"{nombre_archivo}.csv",
            mime="text/csv",
            key='download_csv_no_index'
        )

with tab3:

    st.subheader("📊 Variación intermensual")
    #st.markdown("1️⃣ Subí los archivos correspondientes al _mes anterior_. En caso de tener dos liquidaciones, subí ambos juntos.")
    #st.markdown("**Camino:** InfoSueldos > Informes Liquidación > Liquid. por Ofic. y Rango de Cpto. | **Formato:** Excel Extended (no tabular) | Sin restricciones salvo el mes de liquidación.")
    #archivos_1 = st.file_uploader("", type=["xls"], key="archivo1", accept_multiple_files=True)
    #cant_mes_anterior = len(archivos_1)
    
    st.divider()

    st.markdown("1️⃣ Subí los archivos correspondientes al _mes actual_. En caso de tener dos liquidaciones, sube ambos juntos.")
    st.markdown("**Camino:** InfoSueldos > Informes Liquidación > Liquid. por Ofic. y Rango de Cpto. | **Formato:** Excel Extended (no tabular) | Sin restricciones salvo el mes de liquidación.")

    archivos_2 = st.file_uploader("", type=["xls"], key="archivo2",accept_multiple_files=True)
    cant_mes_actual = len(archivos_2)

    # --- Cuando ambos archivos son subidos ---
    if archivos_2:
        
        st.success("Archivos cargados correctamente.")

        #dfs_1 = []
        dfs_2 = []

        #for archivo_1 in archivos_1:
        #    df_1 = pd.read_excel(archivo_1, engine='xlrd')
        #    dfs_1.append(df_1)

        for archivo_2 in archivos_2:
            df_2 = pd.read_excel(archivo_2,engine='xlrd')
            dfs_2.append(df_2)
        
        #for i,df in enumerate(dfs_1):
        #    df.columns =  ["Muni", "Legajo", "Nombre", "Liq","Base",
        #            "Cant horas","Valor por hora","Saporte",
        #            "Fecha","Valor total"]
            
        #    if i == 0:
        #        mes_anterior = v_mam.limpiar(df)
        #    else:
        #        v_mam.agregar_liquidacion_extra(mes_anterior, df)

        for j,df in enumerate(dfs_2):
            df.columns = ["Muni", "Legajo", "Nombre", "Liq","Base",
                    "Cant horas","Valor por hora","Saporte",
                    "Fecha","Valor total"]
            
            if j == 0:
                mes_actual = v_mam.limpiar(df)
            else:
                v_mam.agregar_liquidacion_extra(mes_actual, df)

        #dataSetLimpio_mes1 = v_mam.armar_data_set(mes_anterior)
        dataSetLimpio_mes2 = v_mam.armar_data_set(mes_actual)

        #v_mam.agregar_total(dataSetLimpio_mes1,dataSetLimpio_mes2)

        #df_area = v_mam.unir_oficinas(dataSetLimpio_mes1,dataSetLimpio_mes2)

        #df_personas = v_mam.unir_personas(dataSetLimpio_mes1, dataSetLimpio_mes2)

        df_area_total = v_mam.resumen_oficinas(dataSetLimpio_mes2)

        #output1 = io.BytesIO()
        #output2 = io.BytesIO()
        output3 = io.BytesIO()
        #output4 = io.BytesIO()

        #df_area.to_excel(output1, index=False)
        #df_personas.to_excel(output2, index=False)
        df_area_total.to_excel(output3, index = False)
        #dataSetLimpio_mes2.to_excel(output4, index= False)

        mes_anterior_str = v_mam.obtener_mes_anterior()
        #nombre_archivo_1 = f"Dif. horas extras por oficina_{mes_anterior_str}.xlsx"
        #nombre_archivo_2 = f"Dif. horas extras por persona_{mes_anterior_str}.xlsx"
        nombre_archivo_3 = f"Resumen horas extras mes actual_{mes_anterior_str}.xlsx"
        #nombre_archivo_4 = f"Reporte por empleado de horas extras mes actual_{mes_anterior_str}.xlsx"

        #st.download_button(
         #   label="📄 Descargar planilla de diferencias de horas extras por oficina",
         #   data=output1.getvalue(),
         #   file_name=nombre_archivo_1,
         #   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        #)

        #st.download_button(
        #    label="📄 Descargar planilla de diferencias de horas extras por persona",
        #    data=output2.getvalue(),
        #    file_name=nombre_archivo_2,
        #    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        #)

        st.download_button(
        label="📄 Descargar resumen por oficina para el mes actual",
        data=output3.getvalue(),
        file_name=nombre_archivo_3,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        #st.download_button(
        #label="📄 Descargar planilla de horas extras por empleado",
        #data=output4.getvalue(),
        #file_name=nombre_archivo_4,
        #mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        #)

with tab4:
    st.subheader('🗞️ Extra! Extra!')

    st.write('Procedimiento')
    with st.expander('Paso 1️⃣: Descargá el archivo de novedades'):
        st.markdown('''
                        - Entrar a M@JOR e ir a Informes > Informes de empleados > Empleados por novedad
                        - Elegir partición MU
                        - Seleccionar Novedades vigentes en el año y mes actual
                        - Elegir variables desde @HRSEXTR1 a @HRSEXTR3
                        - Establecer restricciones > Ejecutar
                        - ⚠️**Importante**⚠️: exportarlo en el formato "Excel 5.0 (XLS) Tabular" y confirmar "Column headings"
                        '''
                        )
            
    archivos = None
    oficinas = None
    with st.expander('Paso 2️⃣: Subí todos los archivos, tanto los csvs como el de novedades descargado del sistema'):
        archivos = st.file_uploader('Subí aca abajo los archivos arrastrando o seleccionando en \'Browse files\'',accept_multiple_files=True)
        st.write("Ingresá las oficinas en un listado con comas, si querés indicar rangos de oficinas separalas por un guion. No uses espacios entre cada uno.")
        st.write("Por ejemplo si ingresás '100-102,200,310' es que querés procesar las oficinas 100, 101, 102, 200 y 310")
        st.write("Si escribís la palabra 'TODO' vas a procesar considerando todas las oficinas (aviso: seguramente aparezcan muchas personas no reportadas pero que sí figuran en sistema)")
        oficinas = st.text_area("Escribí las oficinas o 'todo' abajo, y presioná Ctrl+Enter")
    oficinas = procesar_oficinas(oficinas)

    novedades = None
    with st.expander('Paso 3️⃣: Procesar los datos y ver los resultados'):
        if st.button("Procesar") and archivos:
            # Hallar archivo de novedades
            for archivo in archivos:
                if archivo.name.endswith('.xls'): 
                    novedades = archivo
                    break

            if novedades is None:
                st.error('No subiste el archivo de novedades, hacelo en el paso 2.', icon = '🚨')
            # Procesar
            else:
                resultados_sistema = xx.procesar_novedades_sistema(novedades)
                resultados_reporte = xx.procesar_csvs_oficinas(archivos)
                df,no_estan_en_sistema,no_reportados = xx.comparar_y_armar_df(resultados_sistema,resultados_reporte,oficinas)

                with st.expander('Ver resultados'):
                    if len(no_estan_en_sistema) > 0:
                        st.write("1) Estos legajos fueron reportados pero no cargados en el sistema.")
                        with st.expander("Ver más"):
                            xx.imprimir_lista(no_estan_en_sistema)
                    else: 
                        st.write("1) Todos los legajos reportados están cargados al sistema.")
                    
                    if len(no_reportados) > 0:
                        st.write("2) Estos legajos no fueron reportados por las oficinas pero están cargados en el sistema.")
                        with st.expander("Ver más"):
                            xx.imprimir_lista(no_reportados)
                    else:  
                        st.write("2) Todos los legajos de las oficinas dadas están reportados.")

                    buffer = io.BytesIO()
                    if df is not None:
                        st.write("3) Se encontraron las siguientes inconsistencias:")
                        st.write(df)
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df.to_excel(writer, sheet_name='inconsistencias_hrs_extra', index=True)
                        buffer.seek(0)
                        st.download_button(
                            label="Descargar resultados",
                            data=buffer,
                            file_name="inconsistencias_hrs_extra.xlsx",
                            mime="application/vnd.ms-excel",
                            icon=":material/download:",
                        )
                    else:
                        st.write("3) No se encontraron inconsistencias entre lo reportado y el sistema.")

                    nombres_no_coinciden = xx.comparar_nombres(resultados_sistema,resultados_reporte)

                    if len(nombres_no_coinciden) > 0:
                        st.write('Los siguientes nombres pueden no coincidir:')
                        xx.imprimir_nombre_no_coinciden(nombres_no_coinciden)
