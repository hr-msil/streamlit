import openpyxl
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.shared import Mm
import streamlit as st
from io import BytesIO
import xlrd
import datetime
import pandas as pd

def set_table_font_size(table, size_pt):
    """
    Cambia el tamaño de fuente de todo el texto en una tabla.
    :param table: objeto Table de python-docx
    :param size_pt: tamaño en puntos (int o float)
    """
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(size_pt)

def armar_esqueleto(documento, planilla, oficina):
    """
    Arma la hoja con determinado formato del word.
    
    :param documento: Documento que se está escribiendo
    :param planilla: Planilla actual de la cuál se sacan los datos
    :param oficina: [int, str], array con el número de la oficina y el nombre de la oicina
    """

    parrafo_exp = documento.add_paragraph()
    # El nombre del anexo correspondiente a la oficina es del tipo "oficina - nombre de la oficina"
    nombre_anexo = str(oficina[0]) + " - " + oficina[1]
    run_exp = parrafo_exp.add_run(nombre_anexo)
    run_exp.bold = True
    run_exp.underline = True
    run_exp.font.size = Pt(16)
    parrafo_exp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Tabla
    tabla = documento.add_table(rows=1, cols=8)
    tabla.style = 'Table Grid'  # Bordes visibles

    encabezado = tabla.rows[0].cells
    encabezado[0].text = "Legajo"
    encabezado[1].text = "APELLIDO Y NOMBRE"
    encabezado[2].text = "CATEGORÍA"
    encabezado[3].text = "FUNCIÓN"
    encabezado[4].text = "BONIFICACIÓN"
    encabezado[5].text = "INGRESO"
    encabezado[6].text = "EGRESO"
    encabezado[7].text = "NOTIFICACIÓN FIRMA Y FECHA"
    

    return tabla

def armar_anexoV2(documento,planilla):
    """
    Pasa los datos de la planilla a un formato tabla en un Word.
    
    :param documento: Documento que se está escribiendo.
    :param planilla: Planilla  .xlsx de la cuál se están sacando los datos.
    """
    
    wb = xlrd.open_workbook(file_contents=planilla.read())
    ws = wb.sheet_by_index(0)
    tiene_bonificacion = False
    
    oficina_anterior_num = str(int(ws.cell_value(1, 0))) #Primer número de oficina del área
    oficina_anterior_nom = str(ws.cell_value(1,1))
    numero_oficina = str(int(ws.cell_value(1, 0)))
    nombre_oficina = str(ws.cell_value(1, 1))
    tabla = armar_esqueleto(documento, planilla,[numero_oficina, nombre_oficina])

    for row_idx in range(1, ws.nrows):

        row = ws.row_values(row_idx)
        

        numero_oficina = str(int(row[0]))
        nombre_oficina = str(row[1])

        if numero_oficina != oficina_anterior_num and nombre_oficina != oficina_anterior_nom:

            documento.add_page_break()
            
            tabla = armar_esqueleto(documento, planilla,[numero_oficina,nombre_oficina])
            oficina_anterior_num = numero_oficina
            oficina_anterior_nom = nombre_oficina
            tiene_bonificacion = False

        fila = tabla.add_row().cells
    
        for i,cell in enumerate(row):
            if i == 0 and i == 1:
                continue
            
            elif i == 7 or i == 8:
                fecha = xlrd.xldate_as_datetime(cell, wb.datemode)
                texto = fecha.strftime("%d/%m/%Y")
                fila[i - 2].text = texto
            
            elif i == 6 :
    
                if cell.startswith("MODULOS"):
                    fila[i - 2].text = ""
                else:
                    fila[i - 2].text = str(cell) if cell else ""
                    tiene_bonificacion = True

            elif i == 2:
                fila[i - 2].text = str(int(cell))

            elif i == 4:
                fila[i - 2].text = str(cell) if isinstance(cell, str) else str(int(cell))

            elif cell is not None:
                fila[i - 2].text = str(cell)

            else:
                fila[i - 2].text = ""

        fila[7].text = "" # espacio para firmar

        if not tiene_bonificacion:
            for row in tabla.rows: #elimino la columna 5 (index base 1)
                celda = row.cells[4]
                row._tr.remove(celda._tc)
            
        

    



def armar_documento():
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


    return documento

#------------- Esto es para armar el otro anexo ------------------------

def armar_esqueleto_dos(documento, planilla, oficina):
    """
    Arma la hoja con determinado formato del word.
    
    :param documento: Documento que se está escribiendo
    :param planilla: Planilla actual de la cuál se sacan los datos
    :param oficina: [int, str], array con el número de la oficina y el nombre de la oicina
    """

    parrafo_exp = documento.add_paragraph()
    # El nombre del anexo correspondiente a la oficina es del tipo "oficina - nombre de la oficina"
    nombre_anexo = str(oficina[0]) + " - " + oficina[1]
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
    encabezado[2].text = "CATEGORÍA"
    encabezado[3].text = "FUNCIÓN"
    encabezado[4].text = "BONIFICACIÓN"
    encabezado[5].text = "NOTIFICACIÓN FIRMA Y FECHA"
    

    return tabla

def armar_anexo_dosV2(documento,planilla):
    """
    Pasa los datos de la planilla a un formato tabla en un Word.
    
    :param documento: Documento que se está escribiendo.
    :param planilla: Planilla  .xlsx de la cuál se están sacando los datos.
    """
    
    wb = xlrd.open_workbook(file_contents=planilla.read())
    ws = wb.sheet_by_index(0)
    tiene_bonificacion = False
    
    oficina_anterior_num = str(int(ws.cell_value(1, 0))) #Primer número de oficina del área
    oficina_anterior_nom = str(ws.cell_value(1,1))
    numero_oficina = str(int(ws.cell_value(1, 0)))
    nombre_oficina = str(ws.cell_value(1, 1))
    tabla = armar_esqueleto_dos(documento, planilla,[numero_oficina, nombre_oficina])

    for row_idx in range(1, ws.nrows):

        row = ws.row_values(row_idx)
        

        numero_oficina = str(int(row[0]))
        nombre_oficina = str(row[1])

        if numero_oficina != oficina_anterior_num and nombre_oficina != oficina_anterior_nom:
            

            documento.add_page_break()
            
            tabla = armar_esqueleto_dos(documento, planilla,[numero_oficina,nombre_oficina])
            oficina_anterior_num = numero_oficina
            oficina_anterior_nom = nombre_oficina
            tiene_bonificacion = False

        fila = tabla.add_row().cells
    
        
        for i, cell in enumerate(row):

            if i == 0 or i == 1:
                continue
            elif isinstance(cell, str) and cell.startswith("MODULOS"):
                fila[i - 2].text = ""
            elif i == 6 and isinstance(cell, str) and not cell.startswith("MODULOS"):
                fila[i - 2].text = str(cell) if cell else ""
                tiene_bonificacion = True
            elif i == 2:
                fila[i - 2].text = str(int(cell)) if cell else "" #para que nos nos aparezcan decimales
            elif i == 4:
                fila[i - 2].text = str(cell) if isinstance(cell, str) else str(int(cell))
            else:
                fila[i - 2].text = str(cell) if cell else ""

                

        fila[5].text = "" # espacio para firmar

    if not tiene_bonificacion:
            for row in tabla.rows: #elimino la columna 5 (index base 1)
                celda = row.cells[4]
                row._tr.remove(celda._tc)


    



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
        #validar_archivo_mensualizados(planilla)
        armar_anexo_dosV2(documento,planilla)
    
    return documento


def validar_archivo_mensualizados(archivo):

    df = pd.read_excel(archivo)
    cant_columnas = len(df.columns)
    valores_nulos = df.iloc[:, [0, 1, 2, 3, 4, 5, 7, 8]].isnull().any().any()

    archivo.seek(0)
    return df.columns, cant_columnas, valores_nulos

def validar_otro_archivo(archivo):

    df = pd.read_excel(archivo)
    cant_columnas = len(df.columns)
    valores_nulos = df.iloc[:, [0,1,2,3,4,5]].isnull().any().any() 
    
    archivo.seek(0)
    return df.columns, cant_columnas, valores_nulos





