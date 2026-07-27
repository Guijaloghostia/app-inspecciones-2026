import os
import pandas as pd
import streamlit as st

# Configuración de página con la sidebar desplegable original
st.set_page_config(
    page_title="Control de Inspecciones 2026",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- CARGA COMPLETA DE DATOS DE LA BASE ---
@st.cache_data
def cargar_datos():
  archivo = "001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx"
  if os.path.exists(archivo):
    df = pd.read_excel(archivo)
    # Limpieza preventiva para cálculos y visualización
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

  # =========================================================
  # MENÚ LATERAL COMPLETO (SIDEBAR ORIGINAL CON MULTIFILTROS)
  # =========================================================
  st.sidebar.title("🔍 Control de Inspecciones")
  st.sidebar.markdown("---")

  opciones_menu = [
      "🔎 Buscador Avanzado y Tabla",
      "📊 Resumen General y Estadísticas",
      "ℹ️ Información del Sistema",
  ]
  eleccion = st.sidebar.radio("Navegación / Secciones:", opciones_menu)

  st.sidebar.markdown("---")

  # ==========================================
  # OPCIÓN 1: BUSCADOR COMPLETO Y MÚLTIPLES FILTROS
  # ==========================================
  if eleccion == "🔎 Buscador Avanzado y Tabla":
    st.title("📋 Base Completa de Inspecciones")

    # Contenedor de Filtros Avanzados
    with st.expander("🔻 Filtros de Búsqueda (Tocá para desplegar)", expanded=True):
      col1, col2, col3 = st.columns(3)

      with col1:
        cuit_filtro = st.text_input("🔍 CUIT:", placeholder="Ej: 30-...")
        razon_filtro = st.text_input(
            "🏢 Razón Social:", placeholder="Ej: Nombre de fantasía o firma..."
        )

      with col2:
        calle_filtro = st.text_input(
            "📍 Calle / Domicilio:", placeholder="Ej: San Martín..."
        )
        if "Localidad" in df.columns:
          localidades = ["Todas"] + sorted(
              [str(x) for x in df["Localidad"].dropna().unique()]
          )
          localidad_filtro = st.selectbox("🌆 Localidad:", localidades)
        else:
          localidad_filtro = "Todas"

      with col3:
        if "TNR" in df.columns:
          estados = ["Todos"] + sorted(
              [str(x) for x in df["TNR"].dropna().unique()]
          )
          estado_filtro = st.selectbox("📌 Estado (TNR):", estados)
        else:
          estado_filtro = "Todos"

        if "Inspec." in df.columns:
          inspectores = ["Todos"] + sorted(
              [str(x) for x in df["Inspec."].dropna().unique()]
          )
          inspector_filtro = st.selectbox("👮 Inspector:", inspectores)
        else:
          inspector_filtro = "Todos"

    # Aplicación de los Filtros al DataFrame
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
    if estado_filtro != "Todos" and "TNR" in df.columns:
      df = df[df["TNR"].astype(str) == estado_filtro]
    if inspector_filtro != "Todos" and "Inspec." in df.columns:
      df = df[df["Inspec."].astype(str) == inspector_filtro]

    # Indicador de resultados
    st.success(f"**Registros encontrados:** `{len(df)}` de `{len(df_raw)}`")

    # Función para resaltar filas según el estado
    def colorear_filas(val):
      try:
        val_str = str(val).upper()
        if "IRREGULAR" in val_str or (val_str.replace(".", "").isdigit() and float(val_str) > 0):
          return "background-color: #ffcdd2; color: #b71c1c; font-weight: bold"
        elif "REGULAR" in val_str:
          return "background-color: #c8e6c9; color: #1b5e20; font-weight: bold"
        elif "CERRADO" in val_str:
          return "background-color: #ffe0b2; color: #e65100; font-weight: bold"
      except:
        pass
      return ""

    # MOSTRAR TODAS LAS COLUMNAS ORIGINALES DE LA BASE DE DATOS
    cols_a_ocultar = ["TREL_num", "TNR_num"]
    cols_visibles = [c for c in df.columns if c not in cols_a_ocultar]

    if "TNR" in df.columns:
      df_estilizado = df[cols_visibles].style.map(
          colorear_filas, subset=["TNR"]
      )
      st.dataframe(df_estilizado, use_container_width=True, height=600)
    else:
      st.dataframe(df[cols_visibles], use_container_width=True, height=600)

  # ==========================================
  # OPCIÓN 2: RESUMEN GENERAL COMPLETO
  # ==========================================
  elif eleccion == "📊 Resumen General y Estadísticas":
    st.title("📊 Resumen General de Fiscalización")

    # Métricas Principales
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Actas / Inspecciones", len(df))

    if "TNR" in df.columns:
      regulares = len(df[df["TNR"].astype(str).str.contains("REGULAR", na=False)])
      irregulares = len(df[df["TNR"].astype(str).str.contains("IRREGULAR", na=False)])
      m2.metric("Locales Regulares", regulares)
      m3.metric("Locales Irregulares", irregulares)

    if "TREL_num" in df.columns:
      m4.metric("Total Trabaljadores Relevados", int(df["TREL_num"].sum()))

    st.markdown("---")

    # Gráficos y Tablas de Resumen Integradas
    c_left, c_right = st.columns(2)

    with c_left:
      if "Localidad" in df.columns:
        st.subheader("📍 Inspecciones por Localidad")
        st.bar_chart(df["Localidad"].value_counts())

    with c_right:
      if "Inspec." in df.columns:
        st.subheader("👮 Registros por Inspector")
        st.dataframe(
            df["Inspec."].value_counts().reset_index(),
            use_container_width=True,
        )

  # ==========================================
  # OPCIÓN 3: INFORMACIÓN
  # ==========================================
  elif eleccion == "ℹ️ Información del Sistema":
    st.title("ℹ️ Información y Ayuda")
    st.info(
        "Esta aplicación permite la consulta, filtrado y análisis integral de"
        " las actas de fiscalización."
    )
    st.markdown("""
        **Funcionalidades principales:**
        * **Filtros Combinados:** Podés filtrar por CUIT, Razón Social, Calle, Localidad, Estado y Inspector simultáneamente.
        * **Visualización de Base Completa:** La tabla incluye todas las columnas del archivo original (CIIU, Expedientes, Adicionales, etc.).
        * **Métricas en tiempo real:** Los datos del resumen se actualizan según los datos cargados.
        """)

else:
  st.warning("Verificá que el archivo de Excel esté cargado en GitHub.")
