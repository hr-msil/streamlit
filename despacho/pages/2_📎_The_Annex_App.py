import openpyxl
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.shared import Mm
import streamlit as st
from io import BytesIO
import xlrd

from theannexapp import armar_anexo_dosV2
from theannexapp import armar_anexoV2
from theannexapp import validar_archivo_mensualizados
from theannexapp import armar_documento
from theannexapp import validar_otro_archivo

st.title('The Annex App versión beta NO UTILIZAR📎')

st.write("Selecciona los archivos que queres procesar, validalos y selecciona cuáles queres procesar")

st.write("Que anexo querés realizar")

opciones = ["","Hacer anexo de Mensualizados", "Hacer el otro anexo"]

opcion = st.selectbox("Elegí una opción", opciones)
st.set_page_config(layout="wide")

if opcion == "":
    st.write("❗IMPORTANTE: seleccionar una acción antes de continuar")

if opcion == "Hacer anexo de Mensualizados":
    st_archivos = st.file_uploader("Clickeá donde dice 'Browse files' y subí los archivos",type = "xls",  accept_multiple_files=True)
    if st_archivos:
        st.subheader("Validación de datos")
        st.write("Recordá que para que tus archivos sean válidos, deben cumplir con los siguientes puntos:")
        st.markdown("""
        - El orden de las columnas debe ser: Nro. Oficina, Oficina, Legajo, Apellido y Nombre, Categoría, Función, Bonificación, Fecha Ingreso Cargo, Fecha Egreso Cargo
        - Debe tener 9 (nueve) columnas
        - No debe tener ningún valor nulo, la única columna que puede tener algún nulo es la de Bonificación
        """)
        archivos_a_procesar = []

        for i, archivo in enumerate(st_archivos):
            nombres_columnas, cant_columnas, tiene_nulos = validar_archivo_mensualizados(archivo)
            with st.expander(f"🗒️{archivo.name}",expanded=True):
                col1, col2,col3 = st.columns([3,2,1])
                with col1:
                    st.info(f"**Filename:** {archivo.name}")
                with col2:
                    st.write(f"**Cantidad de columnas:** {cant_columnas}")
                    st.write(f"**Nombre de las columnas:** {", ".join(nombres_columnas)}")
                    st.write(f"**¿Tiene valores nulos?** {'SI' if tiene_nulos else 'NO'}")
                with col3:
                    es_usable = st.checkbox("Marcar como archivo correco", value = True, key = f"check_{i}")
                    if es_usable:
                        archivos_a_procesar.append(archivo)
                    else:
                        st.error("No será procesado")

        st.divider()

        if st.button("🚀 Archivos seleccionados", type = "primary"):
            if not archivos_a_procesar:
                st.warning("No hay archivos seleccionados para procesar")
            else:
                st.success(f"Procesando {len(archivos_a_procesar)} archivos...")
                progress_bar = st.progress(0)
                for idx, f in enumerate(archivos_a_procesar):
                    titulo = f.name.split(".xls")[0]
                    documento = armar_documento()
                    armar_anexoV2(documento,f)
                    buffer = BytesIO()
                    documento.save(buffer)
                    buffer.seek(0)

                    st.info('Recordá revisar el documento')
                    st.download_button(
                        label="Descargar notificaciones",
                        data=buffer,
                        file_name= titulo.strip() + ".docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        icon=":material/download:",
                    )
                    progress_bar.progress((idx + 1) / len(archivos_a_procesar))
                

if opcion == "Hacer el otro anexo":
    st_archivos = st.file_uploader("Clickeá donde dice 'Browse files' y subí los archivos",type = "xls",  accept_multiple_files=True)
    if st_archivos:
        st.subheader("Validación de datos")
        st.write("Recordá que para que tus archivos sean válidos, deben cumplir con los siguientes puntos:")
        st.markdown("""
        - El orden de las columnas debe ser: Nro. Oficina, Oficina, Legajo, Apellido y Nombre, Categoría, Función, Bonificación
        - Debe tener 7 (siete) columnas
        - No debe tener ningún valor nulo, la única columna que puede tener algún nulo es la de Bonificación
        """)
        archivos_a_procesar = []

        for i, archivo in enumerate(st_archivos):
            nombres_columnas, cant_columnas, tiene_nulos = validar_otro_archivo(archivo)
            with st.expander(f"🗒️{archivo.name}",expanded=True):
                col1, col2,col3 = st.columns([3,2,1])
                with col1:
                    st.info(f"**Filename:** {archivo.name}")
                with col2:
                    st.write(f"**Cantidad de columnas:** {cant_columnas}")
                    st.write(f"**Nombre de las columnas:** {", ".join(nombres_columnas)}")
                    st.write(f"**¿Tiene valores nulos?** {'SI' if tiene_nulos else 'NO'}")
                with col3:
                    es_usable = st.checkbox("Marcar como archivo correco", value = True, key = f"check_{i}")
                    if es_usable:
                        archivos_a_procesar.append(archivo)
                    else:
                        st.error("No será procesado")

        st.divider()

        if st.button("🚀 Archivos seleccionados", type = "primary"):
            if not archivos_a_procesar:
                st.warning("No hay archivos seleccionados para procesar")
            else:
                st.success(f"Procesando {len(archivos_a_procesar)} archivos...")
                progress_bar = st.progress(0)
                for idx, f in enumerate(archivos_a_procesar):
                    titulo = f.name.split(".xls")[0]
                    documento = armar_documento()
                    armar_anexo_dosV2(documento,f)
                    buffer = BytesIO()
                    documento.save(buffer)
                    buffer.seek(0)

                    st.info('Recordá revisar el documento')
                    st.download_button(
                        label="Descargar notificaciones",
                        data=buffer,
                        file_name= titulo.strip() + ".docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        icon=":material/download:",
                    )
                    progress_bar.progress((idx + 1) / len(archivos_a_procesar))
                




        