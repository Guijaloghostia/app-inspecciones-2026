import os
import pandas as pd
import streamlit as st

# Configuración de página (ocultamos la sidebar por completo)
st.set_page_config(
    page_title="Control de Inspecciones 2026",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# --- ESTADO DE NAVEGACIÓN ---
if "seccion" not in st.session_state:
  st.session_state["seccion"] = "🔍 Buscador"

# --- CSS PARA BOTONES FIJOS ABAJO Y DISEÑO MÓVIL ---
st.markdown(
    """
    <style>
        /* Oculta la barra lateral de Streamlit por completo */
        [data-testid="stSidebar"] {
            display: none;
        }
        
        /* Espacio al final de la página para que el contenido no quede tapado por la barra inferior */
        .main .block-container {
            padding-bottom: 110px !important;
        }

        /* Estilo visual para las Tarjetas de Locales */
        .card-local {
            background-color: #262730;
            border-left: 6px solid #ff4b4b;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 15px;
        }
        .card-regular {
            border-left: 6px solid #4caf50 !important;
        }
        .card-irregular {
            border-left: 6px solid #f44336 !important;
        }

        /* Ajuste general de letra grande para celular */
        p, label, span, div {
            font-size: 17px !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# --- CARGA COMPLETA DE DATOS ---
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

  # ==========================================
  # SECCIÓN 1: BUSCADOR DE LOCALES
  # ==========================================
  if st.session_state["seccion"] == "🔍 Buscador":
    st.title("📋 Búsqueda de Locales e Inspecciones")

    col1, col2 = st.columns(2)
    with col1:
      cuit_filtro = st.text_input("🔍 Buscar por CUIT:", placeholder="Ej: 30-...")
      razon_filtro = st.text_input(
          "🏢 Buscar por Razón Social:", placeholder="Ej: Nombre..."
      )

    with col2:
      calle_filtro = st.text_input(
          "📍 Buscar por Calle:", placeholder="Ej: San Martín..."
      )
      if "Localidad" in df.columns:
        localidades = ["Todas"] + sorted(
            [str(x) for x in df["Localidad"].dropna().unique()]
        )
        localidad_filtro = st.selectbox("🌆 Localidad:", localidades)
      else:
        localidad_filtro = "Todas"

    # Aplicación de Filtros
    if cuit_filtro and "Cuit" in df.columns:
      df = df[
          df["Cuit"]
          .astype(str)
          .str.contains(cuit_filtro, case=False, na=False)
      ]
    if razon_filtro and "RAZON SOCIAL" in df.columns:
      df = df[
          df["RAZON SOCIAL"]
          .astype(str)
          .str.contains(razon_filtro, case=False, na=False)
      ]
    if calle_filtro and "CALLE" in df.columns:
      df = df[
          df["CALLE"].astype(str).str.contains(calle_filtro, case=False, na=False)
      ]
    if localidad_filtro != "Todas" and "Localidad" in df.columns:
      df = df[df["Localidad"].astype(str) == localidad_filtro]

    st.markdown(f"**Locales encontrados:** `{len(df)}`")
    st.markdown("---")

    # Tarjetas de Locales
    for idx, row in df.iterrows():
      razon_val = row.get("RAZON SOCIAL", "Sin Razón Social")
      cuit_val = row.get("Cuit", "-")
      calle_val = row.get("CALLE", "")
      num_val = row.get("Núm.", "")
      loc_val = row.get("Localidad", "")
      adicional_val = row.get("Adicional", "")
      tnr_val = str(row.get("TNR", "-")).upper()
      trel_val = row.get("TREL", 0)
      exp_val = row.get("Expediente", "-")
      inspec_val = row.get("Inspec.", "-")
      fecha_val = row.get("FECHA", "-")

      clase_card = "card-local"
      if "REGULAR" in tnr_val and "IRREGULAR" not in tnr_val:
        clase_card += " card-regular"
      elif "IRREGULAR" in tnr_val:
        clase_card += " card-irregular"

      st.markdown(
          f"""
            <div class="{clase_card}">
                <h3><b>🏢 {razon_val}</b></h3>
                <p><b>🔍 CUIT:</b> {cuit_val}</p>
                <p>📍 <b>Dirección:</b> {calle_val} {num_val} {f'({adicional_val})' if pd.notna(adicional_val) and adicional_val != '' else ''} - {loc_val}</p>
                <p>📌 <b>Estado (TNR):</b> {tnr_val} | <b>Trabajadores Relevados (TREL):</b> {trel_val}</p>
                <p>📂 <b>Expediente:</b> {exp_val} | 👮 <b>Inspector:</b> {inspec_val} | 📅 <b>Fecha:</b> {fecha_val}</p>
            </div>
            """,
          unsafe_allow_html=True,
      )

  # ==========================================
  # SECCIÓN 2: RESUMEN GENERAL
  # ==========================================
  elif st.session_state["seccion"] == "📊 Resumen":
    st.title("📊 Resumen General de Fiscalización")

    cuit_resumen = st.text_input(
        "🔍 Filtrar resumen por CUIT (Opcional):", placeholder="Ej: 30-..."
    )
    if cuit_resumen and "Cuit" in df.columns:
      df = df[
          df["Cuit"]
          .astype(str)
          .str.contains(cuit_resumen, case=False, na=False)
      ]

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Locales", len(df))
    if "TREL_num" in df.columns:
      c2.metric("Total Relevados (TREL)", int(df["TREL_num"].sum()))
    if "TNR_num" in df.columns:
      c3.metric("Total No Registrados (TNR)", int(df["TNR_num"].sum()))

    st.markdown("---")
    if "Localidad" in df.columns:
      st.subheader("📍 Inspecciones por Localidad")
      st.bar_chart(df["Localidad"].value_counts())

  # ==========================================
  # SECCIÓN 3: INFORMACIÓN
  # ==========================================
  elif st.session_state["seccion"] == "ℹ️ Info":
    st.title("ℹ️ Información del Sistema")
    st.info(
        "Sistema optimizado para consulta rápida desde dispositivos móviles."
    )
    st.markdown("""
        * **Navegación Táctil:** Usá la barra de botones en la parte inferior de la pantalla para cambiar de vista.
        * **Tarjetas Rápidas:** Información organizada con CUIT, estado y dirección resaltados.
        """)

  # ==========================================
  # BOTONES FIJOS EN LA PARTE INFERIOR
  # ==========================================
  st.markdown("---")
  btn_col1, btn_col2, btn_col3 = st.columns(3)

  with btn_col1:
    if st.button(
        "🔍 Buscador",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state["seccion"] == "🔍 Buscador"
            else "secondary"
        ),
    ):
      st.session_state["seccion"] = "🔍 Buscador"
      st.rerun()

  with btn_col2:
    if st.button(
        "📊 Resumen",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state["seccion"] == "📊 Resumen"
            else "secondary"
        ),
    ):
      st.session_state["seccion"] = "📊 Resumen"
      st.rerun()

  with btn_col3:
    if st.button(
        "ℹ️ Info",
        use_container_width=True,
        type=(
            "primary" if st.session_state["seccion"] == "ℹ️ Info" else "secondary"
        ),
    ):
      st.session_state["seccion"] = "ℹ️ Info"
      st.rerun()

else:
  st.warning("Verificá que el archivo de Excel esté cargado en GitHub.")
