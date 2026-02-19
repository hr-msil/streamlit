import streamlit as st
import pandas as pd
import io

from ExtraExtraApp import procesar_novedades_sistema
from ExtraExtraApp import procesar_csvs_oficinas
from ExtraExtraApp import comparar_y_armar_df
from ExtraExtraApp import comparar_nombres
from ExtraExtraApp import procesar_oficinas
from ExtraExtraApp import imprimir_lista
from ExtraExtraApp import imprimir_no_coinciden


# PAGINA
st.title('Extra! Extra! 🗞️')

st.header('Procedimiento')
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
            resultados_sistema = procesar_novedades_sistema(novedades)
            resultados_reporte = procesar_csvs_oficinas(archivos)
            df,no_estan_en_sistema,no_reportados = comparar_y_armar_df(resultados_sistema,resultados_reporte,oficinas)

            with st.expander('Ver resultados'):
                if len(no_estan_en_sistema) > 0:
                    st.write("1) Estos legajos fueron reportados pero no cargados en el sistema.")
                    with st.expander("Ver más"):
                        imprimir_lista(no_estan_en_sistema)
                else: 
                    st.write("1) Todos los legajos reportados están cargados al sistema.")
                
                if len(no_reportados) > 0:
                    st.write("2) Estos legajos no fueron reportados por las oficinas pero están cargados en el sistema.")
                    with st.expander("Ver más"):
                        imprimir_lista(no_reportados)
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

                nombres_no_coinciden = comparar_nombres(resultados_sistema,resultados_reporte)

                if len(nombres_no_coinciden) > 0:
                    st.write('Los siguientes nombres pueden no coincidir:')
                    imprimir_no_coinciden(nombres_no_coinciden)