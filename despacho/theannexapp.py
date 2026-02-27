import openpyxl
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.shared import Mm
import streamlit as st
import pandas as pd
from io import BytesIO


def armar_anexo(documento,planilla_xls,nombre_planilla):
    # Encabezado del anexo centrado y en negrita

    
    xlsx_buffer = BytesIO()
    df = pd.read_excel(planilla_xls)

    # quitamos las columnas oficina y nombre oficina
    nombres_cols_oficinas = df.columns[:2]
    df = df.drop(nombres_cols_oficinas, axis = 1)
    
    # cambiamos el formato de fechas de columnas
    nombres_cols_fechas = df.columns[5:]
    df[nombres_cols_fechas] = df[nombres_cols_fechas].apply(
            pd.to_datetime,
            format='%d/%m/%Y',
            errors='coerce'
        )
    df.to_excel(xlsx_buffer, index=False)
    xlsx_buffer.seek(0)

    parrafo_exp = documento.add_paragraph()
    # Agregar titulo de pagina
    nombre_anexo = nombre_planilla.split('.xls')[0]
    run_exp = parrafo_exp.add_run(nombre_anexo)
    run_exp.bold = True
    run_exp.underline = True
    run_exp.font.size = Pt(16)
    parrafo_exp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Agregar tabla
    tabla = documento.add_table(rows=1, cols=8)
    tabla.style = 'Table Grid'  # Bordes visibles

    encabezado = tabla.rows[0].cells
    encabezado[0].text = "LEGAJO"
    encabezado[1].text = "APELLIDO Y NOMBRE"
    encabezado[2].text = "Cat."
    encabezado[3].text = "FUNCION"
    encabezado[4].text = "BONIF."
    encabezado[5].text = "INGRESO"
    encabezado[6].text = "EGRESO"
    encabezado[7].text = "NOTIFICACION FIRMA Y FECHA"

    wb = openpyxl.load_workbook(xlsx_buffer, read_only = True)
    ws = wb.worksheets[0]
 
    for row in ws.iter_rows(min_row = 2, max_row = ws.max_row, min_col = 1, max_col = 7):
        if not all(cell.value is None for cell in row):
            fila = tabla.add_row().cells
            idx_col = 0
            es_999 = False
            for cell in row:
                if idx_col == 2: # col categoria
                    es_999 = "999" in str(cell.value) # chequeamos si es 999 para modificar bonificacion 
                    fila[idx_col].text = str(cell.value)
                elif idx_col == 4 and es_999: # col bonificacion y es 999
                    fila[idx_col].text = "" # no ponemos nada porque los 999 no tienen bonificaciones salvo los modulos que cobran
                elif idx_col == 5 or idx_col == 6: # col fechas
                    fila[idx_col].text = str(cell.value.strftime("%d/%m/%Y"))
                elif cell.value is not None: # otra col con texto
                    fila[idx_col].text = str(cell.value)
                else: # celda sin texto
                    fila[idx_col].text = ""
                idx_col += 1
            fila[7].text = "" # espacio para firmar

    documento.add_page_break()

def armar_anexos(anexos):
    documento = Document()

    style = documento.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(13)

    section = documento.sections[0]
    section.page_height = Mm(210)
    section.page_width = Mm(297)
    section.left_margin = Mm(25.4)
    section.right_margin = Mm(25.4)
    section.top_margin = Mm(25.4)
    section.bottom_margin = Mm(25.4)
    section.header_distance = Mm(12.7)
    section.footer_distance = Mm(12.7)
    section.orientation = WD_ORIENT.LANDSCAPE

    for anexo in anexos:
        df = pd.read_excel(anexo, engine="xlrd")
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        buffer.seek(0)
        armar_anexo(documento,buffer,anexo.name)

    return documento documento


