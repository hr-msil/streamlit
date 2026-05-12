import streamlit as st
import zipfile
import tempfile
import os
from pathlib import Path

from streamlit_pdf_viewer import pdf_viewer
import pandas as pd

st.set_page_config(layout="wide", page_title="Visor de Facturas")

# ─────────────────────────────────────────
# HELPERS CON CACHÉ
# ─────────────────────────────────────────

@st.cache_data
def leer_excel(file_bytes):
    import io
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name="Facturas")


@st.cache_data
def extraer_pdfs(zip_bytes):
    """
    Extrae el ZIP en un directorio temporal persistente (session_state),
    retorna un dict {nombre: bytes}.
    """
    pdfs = {}
    with zipfile.ZipFile(os.devnull.__class__.__mro__[0].__new__(os.devnull.__class__), "r") if False else zipfile.ZipFile(
        __import__("io").BytesIO(zip_bytes)
    ) as zf:
        for name in zf.namelist():
            if name.lower().endswith(".pdf"):
                pdfs[Path(name).name] = zf.read(name)
    return pdfs

def next_pdf():
    if st.session_state[idx_key] < len(df_validos) - 1:
        st.session_state[idx_key] += 1

def prev_pdf():
    if st.session_state[idx_key] > 0:
        st.session_state[idx_key] -= 1

def sample_95_confidence(df, error=0.05):
    Z = 1.96
    p = 0.5

    N = len(df)

    # tamaño base
    n0 = (Z**2 * p * (1 - p)) / (error**2)

    # corrección por población finita
    n = n0 / (1 + (n0 - 1) / N)

    n = int(min(N, round(n)))

    return df.sample(n=n, random_state=42)        


# ─────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────

MESES = [
    "Enero", "Febrero", "Marzo", "Abril",
    "Mayo", "Junio", "Julio", "Agosto",
    "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

DICC_MESES = {m: i + 1 for i, m in enumerate(MESES)}

CLASIFICACION = ["Complemento", "Honorario"]
DICC_CLAS = {"Complemento": "complemento", "Honorario": "honorario"}

# ─────────────────────────────────────────
# SIDEBAR: CARGA Y FILTROS
# ─────────────────────────────────────────

with st.sidebar:
    st.header("⚙️ Configuración")

    mes = st.selectbox("Mes", MESES)
    clasificar = st.selectbox("Tipo de factura", CLASIFICACION)

    facturas_zip = st.file_uploader("ZIP con PDFs", type=["zip"])
    archivo_tabla = st.file_uploader("Tabla de facturas (.xlsx)", type="xlsx")

    cargar = st.button("Cargar / Recargar", type="primary", use_container_width=True)

# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────

st.title("🗒️ Visor de Facturas")

if not (facturas_zip and archivo_tabla):
    st.info("Subí el ZIP con PDFs y la tabla Excel para comenzar.")
    st.stop()

# ── Leer archivos (cacheados por contenido) ──────────────────────
df_raw  = leer_excel(archivo_tabla.getvalue())
#df_raw_sampleado = sample_95_confidence(df_raw) # -------> Para hacerlo COMPLETO, sacar esto
pdf_map = extraer_pdfs(facturas_zip.getvalue())   # {nombre.pdf: bytes}

# ── Filtrar por mes y clasificación ─────────────────────────────
df_filtrado = df_raw[
    (df_raw["Mes"] == DICC_MESES[mes]) &
    (df_raw["clasifcacion"] == DICC_CLAS[clasificar])
].copy()

df_filtrado_sample = sample_95_confidence(df_filtrado)



df_filtrado_sample["pdf_esperado"] = (
    str(DICC_MESES[mes]) + ". " + mes + " "
    + df_filtrado_sample["nombre"] + ".pdf"
)

df_filtrado_sample["Existe PDF"] = df_filtrado_sample["pdf_esperado"].isin(pdf_map.keys())

# ── Clave única para este filtro: recrea el estado cuando cambia ──
state_key = f"df_{mes}_{clasificar}"



if cargar or state_key not in st.session_state:
    display_df = df_filtrado_sample[[
        "nombre", "Mes", "Tipo", "Número", "Importe",
        "clasifcacion", "pdf_esperado", "Existe PDF"
    ]].copy()
    display_df["Ver"]       = False
    display_df["Procesado"] = False
    display_df["OK"]        = False
    st.session_state[state_key] = display_df

# ─────────────────────────────────────────
# LAYOUT: TABLA | PDF
# ─────────────────────────────────────────

#col_tabla, col_pdf = st.columns([2,3], gap="large")
col_tabla, col_pdf = st.tabs(["Data","Chequeo"])
with col_tabla:
    COLS_VISIBLES = ["nombre", "Mes", "Existe PDF", "Procesado", "OK"]
    st.subheader("Facturas")

    solo_no_procesados = st.checkbox("Mostrar solo no procesados")

    df_a_mostrar = st.session_state[state_key]
    if solo_no_procesados:
        df_a_mostrar = df_a_mostrar[df_a_mostrar["Procesado"] == False]

    edited_df = st.data_editor(
        df_a_mostrar[COLS_VISIBLES],
        hide_index=True,
        use_container_width=True,
        key=f"editor_{state_key}",
        column_config={
            "Ver":        st.column_config.CheckboxColumn("👁 Ver PDF"),
            "Procesado":  st.column_config.CheckboxColumn("🔄 Procesado"),
            "Existe PDF": st.column_config.CheckboxColumn("📄 Existe"),
            "OK":         st.column_config.CheckboxColumn("✅ OK"),
        },
        disabled=[
            "nombre", "Mes", "Tipo", "Número", "Importe",
            "clasifcacion", "pdf_esperado", "Existe PDF"
        ],
    )

    # Persistir cambios respetando índices originales
    # (edited_df puede ser un subconjunto si el filtro está activo)
    st.session_state[state_key].update(edited_df)

    # Para las métricas siempre usamos el df completo
    full_df    = st.session_state[state_key]

    

    if state_key not in st.session_state:
        st.session_state[state_key] = display_df
    idx_key = f"pdf_index_{state_key}"
    if idx_key not in st.session_state:
        st.session_state[idx_key] = 0
    df_validos = full_df[full_df["Existe PDF"]].reset_index()
    total      = len(full_df)
    procesados = full_df["Procesado"].sum()
    ok_count   = full_df["OK"].sum()
    sin_pdf    = (~full_df["Existe PDF"]).sum()

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total",      total)
    m2.metric("Procesados", int(procesados))
    m3.metric("OK",         int(ok_count))
    m4.metric("Sin PDF",    int(sin_pdf), delta=None if sin_pdf == 0 else f"-{sin_pdf}", delta_color="inverse")

with col_pdf:
    st.subheader("Visor de PDF")

    #selected_rows = full_df[full_df["Ver"] == True]
    #st.write(f'Las filas seleccionadas son: {selected_rows}')

    #if len(selected_rows) == 0:
    #    st.info("Marcá ✔ en la columna **👁 Ver PDF** para visualizar una factura.")

    #elif len(selected_rows) > 1:
     #   st.warning("Seleccioná **solo un** PDF a la vez.")

    #else:
    col_prev, col_info, col_next = st.columns([1,2,1])

    with col_prev:
        st.button("⬅ Anterior", on_click=prev_pdf,
              disabled=st.session_state[idx_key] == 0)
    with col_next:
        st.button("Siguiente ➡", on_click=next_pdf,
              disabled=st.session_state[idx_key] == len(df_validos) - 1)
    with col_info:
        st.write(f"{st.session_state[idx_key]+1} / {len(df_validos)}")

        if(len(df_validos) == 0):
            st.info("No hay PDFs para mostrar")
        else:
            current = df_validos.iloc[st.session_state[idx_key]]

            row         = current
            pdf_name    = row["pdf_esperado"]
            existe      = row["Existe PDF"]
            idx         = row["index"]

            st.caption(f"📄 {pdf_name}")

            if not existe:
                st.error("El archivo PDF no fue encontrado en el ZIP.")
            else:
                pdf_bytes = pdf_map[pdf_name]

                # Info de la factura seleccionada
                with st.expander("Datos de la factura", expanded=True):
                    c1, c2 = st.columns(2)
                    c1.markdown(f"**Nombre:** {row['nombre']}")
                    c1.markdown(f"**Tipo:** {row['Tipo']}")
                    c1.markdown(f"**Número:** {row['Número']}")
                    c2.markdown(f"**Mes:** {mes}")
                    c2.markdown(f"**Clasificación:** {row['clasifcacion']}")
                    c2.markdown(f"**Importe:** {row['Importe']}")

                # Botones de acción rápida
                col_a, col_b = st.columns(2)
                if col_a.button("✅ Marcar OK", use_container_width=True, type="primary"):
                    st.session_state[state_key].at[idx, "OK"]        = True
                    st.session_state[state_key].at[idx, "Procesado"] = True
                    st.session_state[state_key].at[idx, "Ver"]       = False

                    if st.session_state[idx_key] < len(df_validos) - 1:
                        st.session_state[idx_key] += 1
                    st.rerun()

                if col_b.button("❌ No OK", use_container_width=True):
                    st.session_state[state_key].at[idx, "OK"]        = False
                    st.session_state[state_key].at[idx, "Procesado"] = True
                    st.session_state[state_key].at[idx, "Ver"]       = False

                    if st.session_state[idx_key] < len(df_validos) - 1:
                        st.session_state[idx_key] += 1
                    st.rerun()


                # PDF viewer
                pdf_viewer(
                    pdf_bytes,
                    width=700,
                    height=650,
                    render_text=True,
                )