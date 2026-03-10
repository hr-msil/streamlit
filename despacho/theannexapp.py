import openpyxl
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.shared import Mm
import streamlit as st
from io import BytesIO
import xlrd

def armar_anexo(documento,planilla):
    # Encabezado del anexo centrado y en negrita
    parrafo_exp = documento.add_paragraph()
    nombre_anexo = planilla.name.split('.xlsx')[0]
    run_exp = parrafo_exp.add_run(nombre_anexo)
    run_exp.bold = True
    run_exp.underline = True
    run_exp.font.size = Pt(16)
    parrafo_exp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Tabla
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

    wb = openpyxl.load_workbook(planilla,read_only = True)
    ws = wb.worksheets[0]

    for row in ws.iter_rows(min_row = 2, max_row = ws.max_row, min_col = 3, max_col = 9):
        if not all(cell.value is None for cell in row):
            fila = tabla.add_row().cells
            i = 0
            for cell in row:
                if i == 5 or i == 6:
                    fila[i].text = str(cell.value.strftime("%d/%m/%Y"))
                elif cell.value is not None:
                    fila[i].text = str(cell.value)
                else:
                    fila[i].text = ""
                i += 1
            fila[7].text = "" # espacio para firmar

    documento.add_page_break()

def armar_anexos(planillas):
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

    for planilla in planillas:
        armar_anexo(documento,planilla)

    return documento

def armar_anexo_dos(documento,planilla):
    # Encabezado del anexo centrado y en negrita
    parrafo_exp = documento.add_paragraph()
    nombre_anexo = planilla.name.split('.xls')[0]
    run_exp = parrafo_exp.add_run(nombre_anexo)
    run_exp.bold = True
    run_exp.underline = True
    run_exp.font.size = Pt(16)
    parrafo_exp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Tabla
    tabla = documento.add_table(rows=1, cols=6)
    tabla.style = 'Table Grid'  # Bordes visibles

    encabezado = tabla.rows[0].cells
    encabezado[0].text = "LEGAJO"
    encabezado[1].text = "APELLIDO Y NOMBRE"
    encabezado[2].text = "Cat."
    encabezado[3].text = "FUNCION"
    encabezado[4].text = "BONIF."
    encabezado[5].text = "NOTIFICACION FIRMA Y FECHA"

    wb = xlrd.open_workbook(file_contents=planilla.read())
    ws = wb.sheet_by_index(0)
    tiene_bonificacion = False

    for row_idx in range(1, ws.nrows):
        row = ws.row_values(row_idx)

        fila = tabla.add_row().cells
   
        for i, cell in enumerate(row):
            if i == 0 or i == 1:
                continue
            elif isinstance(cell, str) and cell.startswith("MODULOS"):
                fila[i - 2].text = ""
            elif i == 6 and isinstance(cell, str) and not cell.startswith("MODULOS"):
                fila[i - 2].text = str(cell) if cell else ""
                tiene_bonificacion = True
            elif i == 2 or i == 4:
                fila[i - 2].text = str(int(cell)) if cell else "" #para que nos nos aparezcan decimales
            else:
                fila[i - 2].text = str(cell) if cell else ""

        fila[5].text = "" #espacio para firmar

    if not tiene_bonificacion:
        for row in tabla.rows: #elimino la columna 5 (index base 1)
            celda = row.cells[4]
            row._tr.remove(celda._tc)


    documento.add_page_break()

def armar_anexos_dos(planillas):
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

    for planilla in planillas:
        armar_anexo_dos(documento,planilla)
    
    return documento


