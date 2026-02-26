import openpyxl
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.shared import Mm
import streamlit as st
from io import BytesIO

from theannexapp import armar_anexos


## Streamlit APP ## 

st.title('The Annex App📎')
st_archivos = st.file_uploader("Clickeá donde dice 'Browse files' y subí los archivos", accept_multiple_files=True)

if st_archivos:
    st.success(f"Subiste {len(st_archivos)} planilla(s)")
    titulo = st.text_input("Escribí el nombre del archivo y presioná Enter", "Anexo Subsecretaría ABC")

    if st.button("Procesar y armar anexos"):
        documento = armar_anexos(st_archivos)

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