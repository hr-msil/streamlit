import streamlit as st

st.set_page_config(
    page_title="Hola",
    page_icon="👋",
)

st.write("# Bienvenidos a la App de Despacho! 👋")

st.sidebar.success("Seleccionar una página arriba")

# Acciones rápidas
col1, col2 = st.columns(2)

col1.page_link("pages/1_🧢_notificapp.py", label="NotificApp.", icon="🧢")
col2.page_link("pages/2_📎_The_Annex_App.py", label="The Annex App.", icon="📎")
