import streamlit as st

st.set_page_config(
    page_title="Asistencia's Assistant",
    page_icon="🤖",
    layout='wide'
)

st.title("Asistencia's Assistant 🤖")
st.write("El asistente de asistencias a tu servicio, para darte una mano en lo que necesites.")
st.badge("Frente a cualquier bug o problema, contactar al equipo de Datos.", icon='📊')

st.divider()

# Acciones rápidas
col1, col2, col3 = st.columns(3)

col1.page_link("pages/1_🕰️_Horas_extras.py", label="Para la carga y control de hh.ee.", icon="🕰️")
col2.page_link("pages/2_🚌_Viaticos.py", label="Para el control de viaticos.", icon="🚌")
col3.page_link("pages/3_✅_Descuentos.py", label="Para el proceso de descuentos.", icon="✅")
