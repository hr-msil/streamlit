import streamlit as st
from io import BytesIO
from theannexapp import armar_anexo_dosV2

st.title('The Annex App versión beta NO UTILIZAR📎')

st.write("Seleccioná los archivos que queres procesar")


st.set_page_config(layout="wide")

#Ponemos un checkbox para que la persona decida si separar o no por oficina
separar = st.checkbox("Clickeá acá si deseas hacer un archivo por oficina, caso contrario se hará un archivo con todas las oficinas separadas por hoja.")
st.write("Recordá que las primeras dos columnas deben corresponder a la oficina y nombre de la oficina de la persona.")
st_archivos = st.file_uploader("Clickeá donde dice 'Browse files' y subí los archivos", type = "xls", accept_multiple_files = True)
if st_archivos:
    
    st.success(f"Procesando {len(st_archivos)} archivos...")
    progress_bar = st.progress(0)
    st.info('Recordá revisar los documentos resultantes y eliminar las celdas con los módulos.')
    for idx, f in enumerate(st_archivos):
        titulo = f.name.split(".xls")[0]
        documentos = armar_anexo_dosV2(f, separar)

        for i,documento in enumerate(documentos):

            buffer = BytesIO()
            documento.save(buffer)
            buffer.seek(0)

            st.download_button(
                label="Descargar notificaciones",
                data=buffer,
                file_name= str(i + 1) + "_"+ titulo.strip() + ".docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                icon=":material/download:",
            )
        progress_bar.progress((idx + 1) / len(st_archivos))
                




        