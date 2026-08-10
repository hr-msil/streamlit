import streamlit as st

st.set_page_config(
    page_title="Hola",
    page_icon="👋",
    layout="wide"
)

st.write("# Haberes! 👋")

st.sidebar.success("Seleccionar una página arriba")

col1, col2, col3 = st.columns(3)

col1.page_link("pages/1_📊_Control_productividades.py", label="Para el control de Productividades", icon="📊")
col2.page_link("pages/2_📈_Permisos_de_cobro .py", label="Para los permisos de cobro", icon="📈")
col3.page_link("pages/3_🗞️_Extra_Extra.py", label="Extra Extra", icon="🗞️")



