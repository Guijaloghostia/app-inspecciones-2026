import streamlit as st
import pandas as pd
import os

# Configuración de página con menú lateral colapsado para ver todo más limpio
st.set_page_config(
    page_title="Control de Inspecciones 2026", 
    layout="wide", 
    initial_sidebar_state="collapsed"
)

# --- ESTILOS CSS PERSONALIZADOS (Letras grandes para el celu) ---
st.markdown("""
    <style>
        .stTabs [data-baseweb="tab"] {
            font-size: 18px !important;
            font-weight: bold;
            padding: 10px 20px;
        }
        .stTextInput input {
            font-size: 18px !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- CARGA DE DATOS ---
@st.cache_data
def cargar_datos():
    archivo = "001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx"
    if os.path.exists(archivo):
        df = pd.read_excel(archivo)
        return df
    else:
        st.error(f"No se encontró el archivo: {archivo}")
        return pd.DataFrame()

df_raw = cargar_datos()

if not df_raw.empty:
    df = df_raw.copy()

    # --- NAVEGACIÓN POR PESTAÑAS (Abajo o arriba directo como botones) ---
    tab_busqueda, tab_resumen, tab_instrucciones = st.tabs([
        "🔍 Buscador y Tabla", 
        "📊 Resumen", 
        "ℹ️ Info"
    ])

    # === PESTAÑA 1: BUSCADOR Y TABLA CON COLORES ===
    with tab_busqueda:
        st.subheader("Búsqueda de Locales e Inspecciones")
        
        col_cuit, col_razon, col_calle = st.columns(3)
        
        with col_cuit:
            cuit_filtro = st.text_input("🔍 CUIT:", placeholder="Ej: 30-...")
        with col_razon:
            razon_filtro = st.text_input("🏢 Razón Social:", placeholder="Ej: Nombre...")
        with col_calle:
            calle_filtro = st.text_input("📍 Calle:", placeholder="Ej: San Martín...")

        # Filtros aplicados
        if cuit_filtro:
            df = df[df['Cuit'].astype(str).str.contains(cuit_filtro, case=False, na=False)]
        if razon_filtro:
            df = df[df['RAZON SOCIAL'].astype(str).str.contains(razon_filtro, case=False, na=False)]
        if calle_filtro:
            df = df[df['CALLE'].astype(str).str.contains(calle_filtro, case=False, na=False)]

        st.markdown(f"**Registros encontrados:** `{len(df)}`")

        # Función para pintar en rojo o verde según Trabajadores No Registrados (TNR)
        def colorear_tnr(val):
            try:
                val_num = float(val)
                if val_num > 0:
                    return 'background-color: #ffcdd2; color: #b71c1c; font-weight: bold' # Rojo si hay TNR
                return 'background-color: #c8e6c9; color: #1b5e20; font-weight: bold' # Verde si está en regla
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

    # === PESTAÑA 2: RESUMEN ===
    with tab_resumen:
        st.subheader("Métricas Generales")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Registros", len(df))
        if 'TREL' in df.columns:
            c2.metric("Trabajadores Relevados (TREL)", int(df['TREL'].sum()))
        if 'TNR' in df.columns:
            c3.metric("Trabajadores No Registrados (TNR)", int(df['TNR'].sum()))

    # === PESTAÑA 3: INFO ===
    with tab_instrucciones:
        st.write("### Panel de Control de Inspecciones")
        st.info("Usá las pestañas superiores para alternar entre el buscador con tabla coloreada y las estadísticas generales sin necesidad de abrir menús laterales.")

else:
    st.warning("Verificá que el archivo de Excel esté subido correctamente en el repositorio.")
