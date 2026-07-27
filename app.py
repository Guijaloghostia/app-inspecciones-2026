import os
import pandas as pd
import streamlit as st

# Configuración de página original (con sidebar habilitada)
st.set_page_config(
    page_title="Control de Inspecciones 2026",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- CARGA DE DATOS ---
@st.cache_data
def cargar_datos():
  archivo = "001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx"
  if os.path.exists(archivo):
    df = pd.read_excel(archivo)
    # Limpieza preventiva para evitar errores en las métricas
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
  # MENÚ LATERAL DESPLEGABLE ORIGINAL (SIDEBAR)
  # ==========================================
  st.sidebar.title("📌 Menú de Opciones")
  opciones_menu = [
      "🔍 Buscador / Filtros",
      "📊 Resumen General",
      "ℹ️ Información",
  ]
  eleccion = st.sidebar.radio("Ir a:", opciones_menu)

  # ==========================================
  # OPCCIÓN 1: BUSCADOR CON FILTRO DE CUIT
  # ==========================================
  if eleccion == "🔍 Buscador / Filtros":
    st.title("Búsqueda de Locales e Inspecciones")

    # Filtros en la pantalla principal
    col_cuit, col_razon, col_calle = st.columns(3)

    with col_cuit:
      cuit_filtro = st.text_input("🔍 CUIT:", placeholder="Ej: 30-...")
    with col_razon:
      razon_filtro = st.text_input(
          "🏢 Razón Social:", placeholder="Ej: Nombre..."
      )
    with col_calle:
      calle_filtro = st.text_input("📍 Calle:", placeholder="Ej: San Martín...")

    # Aplicación de los filtros
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

    st.markdown(f"**Registros encontrados:** `{len(df)}`")

    # Función para dar color a la columna TNR
    def colorear_tnr(val):
      try:
        val_num = float(val)
        if val_num > 0:
          return (
              "background-color: #ffcdd2; color: #b71c1c; font-weight: bold"
          )
        return "background-color: #c8e6c9; color: #1b5e20; font-weight: bold"
      except:
        return ""

    columnas_mostrar = ["Cuit", "RAZON SOCIAL", "CALLE", "Núm.", "TREL", "TNR"]
    cols_existentes = [c for c in columnas_mostrar if c in df.columns]

    if cols_existentes:
      if "TNR" in cols_existentes:
        df_estilizado = df[cols_existentes].style.map(
            colorear_tnr, subset=["TNR"]
        )
        st.dataframe(df_estilizado, use_container_width=True, height=500)
      else:
        st.dataframe(df[cols_existentes], use_container_width=True, height=500)

  # ==========================================
  # OPCIÓN 2: RESUMEN GENERAL
  # ==========================================
  elif eleccion == "📊 Resumen General":
    st.title("Métricas y Estadísticas")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Locales", len(df))
    if "TREL_num" in df.columns:
      c2.metric("Total Relevados (TREL)", int(df["TREL_num"].sum()))
    if "TNR_num" in df.columns:
      c3.metric("Total No Registrados (TNR)", int(df["TNR_num"].sum()))

  # ==========================================
  # OPCIÓN 3: INFORMACIÓN
  # ==========================================
  elif eleccion == "ℹ️ Información":
    st.title("Información del Sistema")
    st.info(
        "Utilizá el menú desplegable de la izquierda (Sidebar) para navegar"
        " entre las distintas secciones de la aplicación."
    )

else:
  st.warning("Verificá que el archivo de Excel esté cargado en GitHub.")
