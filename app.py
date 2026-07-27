import streamlit as st
import pandas as pd
import os

# Configuración de página
st.set_page_config(
    page_title="Control de Inspecciones 2026", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# --- ESTILOS CSS (BOTONES ABAJO) ---
st.markdown("""
    <style>
        /* Fija la barra de pestañas al borde inferior del celular */
        .stTabs [data-baseweb="tab-list"] {
            position: fixed;
            bottom: 0;
            left: 0;
            width: 100%;
            background-color: #1e1e1e; /* Fondo oscuro para los botones */
            z-index: 99999;
            display: flex;
            justify-content: space-around;
            padding: 10px 0;
            box-shadow: 0px -2px 10px rgba(0,0,0,0.5);
        }
        /* Estilos de cada botón */
        .stTabs [data-baseweb="tab"] {
            font-size: 16px !important;
            font-weight: bold;
            padding: 10px !important;
            flex-grow: 1;
            text-align: center;
            color: #ffffff !important;
        }
        /* Resalta el botón seleccionado en rojo */
        .stTabs [aria-selected="true"] {
            color: #ff4b4b !important;
            border-top: 3px solid #ff4b4b !important;
        }
        /* Margen inferior para que la tabla no quede tapada por los botones */
        .block-container {
            padding-bottom: 90px !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- CARGA DE DATOS Y CORRECCIÓN DE NÚMEROS ---
@st.cache_data
def cargar_datos():
    archivo = "001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx"
    if os.path.exists(archivo):
        df = pd.read_excel(archivo)
        # Limpiamos errores para que no falle la suma en el resumen
        if 'TREL' in df.columns:
            df['TREL_num'] = pd.to_numeric(df['TREL'], errors='coerce').fillna(0)
        if 'TNR' in df.columns:
            df['TNR_num'] = pd.to_numeric(df['TNR'], errors='coerce').fillna(0)
        return df
    else:
        st.error(f"No se encontró el archivo: {archivo}")
        return pd.DataFrame()

df_raw = cargar_datos()

if not df_raw.empty:
    df = df_raw.copy()

    # --- NAVEGACIÓN (LOS BOTONES QUE QUEDAN ABAJO) ---
    tab_busqueda, tab_resumen, tab_instrucciones = st.tabs([
        "🔍 Tabla / Filtros", 
        "📊 Resumen", 
        "ℹ️ Info"
    ])

    # === PESTAÑA 1: BUSCADOR Y TABLA COLOREADA ===
    with tab_busqueda:
        st.subheader("Búsqueda de Locales e Inspecciones")
        
        # Filtros incluyendo el CUIT
        col_cuit, col_razon, col_calle = st.columns(3)
        with col_cuit:
            cuit_filtro = st.text_input("🔍 CUIT:", placeholder="Ej: 30-...")
        with col_razon:
            razon_filtro = st.text_input("🏢 Razón Social:", placeholder="Ej: Nombre...")
        with col_calle:
            calle_filtro = st.text_input("📍 Calle:", placeholder="Ej: San Martín...")

        # Aplicamos los filtros si escribiste algo
        if cuit_filtro:
            df = df[df['Cuit'].astype(str).str.contains(cuit_filtro, case=False, na=False)]
        if razon_filtro:
            df = df[df['RAZON SOCIAL'].astype(str).str.contains(razon_filtro, case=False, na=False)]
        if calle_filtro:
            df = df[df['CALLE'].astype(str).str.contains(calle_filtro, case=False, na=False)]

        st.markdown(f"**Registros encontrados:** `{len(df)}`")

        # Función para pintar la tabla (rojo si hay no registrados, verde si está en regla)
        def colorear_tnr(val):
            try:
                val_num = float(val)
                if val_num > 0:
                    return 'background-color: #ffcdd2; color: #b71c1c; font-weight: bold'
                return 'background-color: #c8e6c9; color: #1b5e20; font-weight: bold'
            except:
                return ''

        columnas = ['Cuit', 'RAZON SOCIAL', 'CALLE', 'Núm.', 'TREL', 'TNR']
        cols_existentes = [c for c in columnas if c in df.columns]
        
        if cols_existentes:
            if 'TNR' in cols_existentes:
                df_estilizado = df[cols_existentes].style.map(colorear_tnr, subset=['TNR'])
                st.dataframe(df_estilizado, use_container_width=True, height=500)
            else:
                st.dataframe(df[cols_existentes], use_container_width=True, height=500)

    # === PESTAÑA 2: MÉTRICAS Y RESUMEN ===
    with tab_resumen:
        st.subheader("Métricas Generales")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Registros", len(df))
        if 'TREL_num' in df.columns:
            c2.metric("Trabajadores Relevados (TREL)", int(df['TREL_num'].sum()))
        if 'TNR_num' in df.columns:
            c3.metric("Trabajadores No Registrados (TNR)", int(df['TNR_num'].sum()))

    # === PESTAÑA 3: INFO ===
    with tab_instrucciones:
        st.write("### Panel de Control de Inspecciones")
        st.info("Navegá tocando los botones fijos en la parte inferior de tu pantalla.")

else:
    st.warning("Verificá que el archivo de Excel esté subido correctamente en el repositorio.")
