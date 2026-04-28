import streamlit as st
import pandas as pd
import io
root = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(root))
from services.guardiasMedicas import armar_listados

st.subheader("🩺 Guardias médicas")
st.divider()

col1, col2 = st.columns(2)

with col1:
    st.caption("Listado de novedades")
    archivo_novedades = st.file_uploader(
        "", type="xls", accept_multiple_files=False, key="archivo_novedades"
    )
    st.caption("Listado de ausencias")
    archivo_ausencias = st.file_uploader(
        "", type="xls", accept_multiple_files=False, key="archivo_ausencias"
    )

with col2:
    st.caption("Listado de horarios")
    archivo_horarios = st.file_uploader(
        "", type="xls", accept_multiple_files=False, key="archivo_horarios"
    )
    st.caption("Empleados por oficina")
    archivo_empleados_por_ofi = st.file_uploader(
        "", type="xls", accept_multiple_files=False, key="archivo_empleados_por_ofi"
    )

if archivo_novedades and archivo_horarios and archivo_ausencias and archivo_empleados_por_ofi:
    st.divider()

    listas, legajos_a_reportar = armar_listados(
        archivo_novedades, archivo_horarios, archivo_ausencias, archivo_empleados_por_ofi
    )

    buffers = {}

    for nombre, lista in listas.items():
        if nombre == "DTO398":
            df = pd.DataFrame({"legajos": lista[0], "porcentajes": lista[1]})
        elif nombre == "LEGAJOS SIN HORARIO":
            legajos_sin_horario = lista
            continue
        else:
            df = pd.DataFrame({"legajos": lista})

        buffer = io.StringIO()
        df.to_csv(buffer, index=False)
        buffers[nombre] = buffer.getvalue()

    st.caption("Archivos generados")
    cols = st.columns(len(buffers))
    for col, (nombre, data) in zip(cols, buffers.items()):
        with col:
            st.download_button(
                label=nombre,
                data=data,
                file_name=f"{nombre}.csv",
                mime="text/csv",
                icon=":material/download:",
                use_container_width=True,
            )
    for output in legajos_a_reportar:
        st.info(output)
