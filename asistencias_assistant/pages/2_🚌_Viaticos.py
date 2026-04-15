import streamlit as st
import pandas as pd
import io

from helpers import obtener_hoja_planilla
from helpers import transformar_ausencias_a_dict

import services.viaticos as vi

st.set_page_config(page_title="Viáticos", page_icon="🚌", layout = 'wide')

st.markdown("**Versión beta: cualquier cosa rara que encuentres, no dudes en reportarla!**")

st.markdown("Subí la **planilla de viáticos**.")
archivo_planilla = st.file_uploader("Planilla de viáticos", type = 'xls', accept_multiple_files = False, key = "archivo_planilla")

st.divider()

st.markdown("Subí el **listado de legajos** (no hay que poner restricciones sobre el archivo). Recordá eliminar los totales al final de la planilla.")
st.markdown("**Camino**: Informes > Informes de Empleado > Empleados por Legajo | **Formato**: Excel (Tabular).")
archivo_legajos = st.file_uploader("Listado de legajos", type = 'xls', accept_multiple_files = False, key = "archivo_legajos")

st.divider()

st.markdown("Subí el **listado de ausencias** (puede incluir todas las oficinas). Recordá que hay que hacer el cálculo antes de exportarlo.")
st.markdown("**Camino**: Informes > Informes de Asistencia > Ausencias por Oficina | **Formato**: Excel Extended o Excel (no tabular).")
archivo_ausencias = st.file_uploader("Listado de ausencias", type = 'xls', accept_multiple_files = False, key = "archivo_ausencias")

st.divider()

nombre_archivo = st.text_input("Escribí el nombre del archivo que querés generar")

if archivo_planilla and archivo_legajos and archivo_ausencias and nombre_archivo:
    hoja_planilla = obtener_hoja_planilla(archivo_planilla)
    planilla_viaticos = pd.read_excel(archivo_planilla, sheet_name=hoja_planilla)
    planilla_viaticos = vi.normalizar_planilla_viaticos(planilla_viaticos)
    planilla_viaticos_pre = planilla_viaticos.copy()

    # chequear legajos y nombre
    datos_sistema = pd.read_excel(archivo_legajos)
    resultados = vi.validar_legajos_y_nombres(planilla_viaticos,datos_sistema)

    vi.reportar_validacion_legajos(resultados)

    # chequear ausencias
    ausencias_dict = transformar_ausencias_a_dict(archivo_ausencias, es_viaticos=True)

    planilla_viaticos_pos, inconsistencias = vi.modificar_viaticos_en_ausencia(planilla_viaticos, ausencias_dict)

    vi.reportar_inconsistencias(planilla_viaticos_pos, inconsistencias)

    st.badge("Esta es la planilla de viáticos antes de compararla con las ausencias.")
    st.write(planilla_viaticos_pre)

    st.badge("Esta es la planilla de viáticos después de compararla con las ausencias.")
    st.write(planilla_viaticos_pos)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        planilla_viaticos_pos.to_excel(writer, sheet_name=f'planilla_viaticos_{nombre_archivo}', index=True)
    buffer.seek(0)
    st.download_button(
        label="Descargar planilla final",
        data=buffer,
        file_name=f"planilla_viaticos_{nombre_archivo}.xlsx",
        mime="application/vnd.ms-excel",
        icon=":material/download:",

    )
