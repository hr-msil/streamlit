import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pdfminer.high_level import extract_text
import re
from io import BytesIO

from notificapp import obtener_datos
from notificapp import armar_documento


## Streamlit APP ## 

st.title('NotificApp 🧢')
st_archivos = st.file_uploader("Clickeá donde dice 'Browse files' y subí los archivos", accept_multiple_files=True)
st.warning('Recordá que el formato tiene que ser \'NroDeResolucion - NroDeExpediente\'', icon="⚠️")

if st_archivos:
    st.success(f"Subiste {len(st_archivos)} expedientes(s)")
    
    if st.button("Procesar y convertir a DOCX"):
        dict_file_datos = obtener_datos(st_archivos)
        
        documento = armar_documento(dict_file_datos,st_archivos)

        buffer = BytesIO()
        documento.save(buffer)
        buffer.seek(0)

        st.info('Recordá revisar el documento, pues puede contener errores de tipeo')
        st.download_button(
            label="Descargar notificaciones",
            data=buffer,
            file_name="LISTO PARA NOTIFICAR.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            icon=":material/download:",

        )
