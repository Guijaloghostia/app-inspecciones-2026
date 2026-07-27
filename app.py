import os
import pandas as pd
import streamlit as st

# Configuración de página
st.set_page_config(
    page_title="Control de Inspecciones 2026",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- ESTILOS CSS CON NAVEGACIÓN FIJA EN EL BORDE INFERIOR (ESTILO APP MÓVIL) ---
st.markdown(
    """
    <style>
        /* Fija la barra de pestañas al borde inferior de la pantalla */
        .stTabs [data-baseweb="tab-list"] {
            position: fixed;
            bottom: 0;
            left: 0;
            width: 100%;
            background-color: #1e1e1e;
            z-index: 99999;
            display: flex;
            justify-content: space-around;
            padding: 8px 0;
            box-shadow: 0px -2px 10px rgba(0,0,0,0.3);
            border-top: 1px solid #333;
        }

        /* Estilo de cada botón/pestaña para facilitar el toque del pulgar */
        .stTabs [data-baseweb="tab"] {
            font-size: 16px !important;
            font-weight: bold;
            color: #ffffff !important;
            padding: 10px 15px !important;
            border: none !important;
            background-color: transparent !important;
            flex-grow: 1;
            text-align: center;
        }

        /* Color al seleccionar pestaña */
        .stTabs [aria-selected="true"] {
            color: #ff4b4b !important;
            border-top: 3px solid #ff4b4b !important;
        }

        /* Margen inferior al contenido para que los botones de abajo no tapen el texto */
        .block-container {
            padding-bottom: 90px !important;
        }

        .stTextInput input {
            font-size: 18px !important;
        }

        .metric-card {
            background-color: #262730;
            border-left: 5px solid #ff4b4b;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 15px;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# --- CARGA DE DATOS ---
@st.cache_data
def cargar_datos():
  archivo = "001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx"
  if os.path.exists(archivo):
    df = pd.read_excel(archivo)
    if "TREL" in df.columns:
      df["TREL_num"] = pd.to_numeric(df["TREL"], errors="coerce").fillna(0)
    if "TNR" in df.columns:
      df["TNR_num"] = pd.to_numeric(df["TNR"], errors="coerce").fillna(0)
    return df
  else:
    st.error(f"No se encontró el archivo: {archivo}")
    return pd.DataFrame()


df_raw = cargar_datos()

if not df_raw.empty:
  df = df_raw.copy()

  # --- PESTAÑAS (CSS LAS MUEVE AL BORDE INFERIOR) ---
  tab_busqueda, tab_resumen, tab_info = st.tabs(
      ["🔍 Buscador", "📊 Resumen", "ℹ️ Info"]
  )

  # === PESTAÑA 1: BUSCADOR Y TARJETAS ===
  with tab_busqueda:
    st.subheader("Búsqueda de Locales e Inspecciones")

    col_cuit, col_razon, col_calle = st.columns(3)

    with col_cuit:
      cuit_filtro = st.text_input("🔍 CUIT:", placeholder="Ej: 30-...")
    with col_razon:
      razon_filtro = st.text_input(
          "🏢 Razón Social:", placeholder="Ej: Nombre..."
      )
    with col_calle:
      calle_filtro = st.text_input("📍 Calle:", placeholder="Ej: San Martín...")

    # Aplicar filtros
    if cuit_filtro:
      df = df[
          df["Cuit"]
          .astype(str)
          .str.contains(cuit_filtro, case=False, na=False)
      ]
    if razon_filtro:
      df = df[
          df["RAZON SOCIAL"]
          .astype(str)
          .str.contains(razon_filtro, case=False, na=False)
      ]
    if calle_filtro:
      df = df[
          df["CALLE"].astype(str).str.contains(calle_filtro, case=False, na=False)
      ]

    st.markdown(f"**Registros encontrados:** `{len(df)}`")

    # Muestra los datos en tarjetas
    for idx, row in df.iterrows():
      cuit_val = row.get("Cuit", "-")
      razon_val = row.get("RAZON SOCIAL", "-")
      calle_val = row.get("CALLE", "")
      num_val = row.get("Núm.", "")
      trel_val = int(row.get("TREL_num", 0))
      tnr_val = int(row.get("TNR_num", 0))

      st.markdown(
          f"""
            <div class="metric-card">
                <h4><b>CUIT:</b> {cuit_val} | <b>Razón Social:</b> {razon_val}</h4>
                <p>📍 <b>Dirección:</b> {calle_val} {num_val}</p>
            </div>
            """,
          unsafe_allow_html=True,
      )

      c1, c2 = st.columns(2)
      c1.metric("Trabajadores Relevados (TREL)", trel_val)
      c2.metric("No Registrados (TNR)", tnr_val)
      st.divider()

  # === PESTAÑA 2: RESUMEN Y MÉTRICAS ===
  with tab_resumen:
    st.subheader("Métricas Generales")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Locales", len(df))
    if "TREL_num" in df.columns:
      c2.metric("Total Relevados (TREL)", int(df["TREL_num"].sum()))
    if "TNR_num" in df.columns:
      c3.metric("Total No Registrados (TNR)", int(df["TNR_num"].sum()))

  # === PESTAÑA 3: INFO ===
  with tab_info:
    st.write("### Control de Fiscalización")
    st.info(
        "Navegá entre el buscador y las estadísticas usando los botones"
        " ubicados en el borde inferior de la pantalla."
    )

else:
    st.warning("Verificá que el archivo de Excel esté cargado en GitHub.")
