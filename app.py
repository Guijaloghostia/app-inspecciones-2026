import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Base de Fiscalizaciones", layout="wide", initial_sidebar_state="expanded")
st.markdown(
    """
    <style>
        /* Agranda el texto de los elementos de la barra lateral / menú */
        [data-testid="stSidebar"] * {
            font-size: 20px !important;
        }

        /* Agranda los títulos del menú */
        [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
            font-size: 24px !important;
        }

        /* Agranda los botones de la barra lateral y de la app */
        .stButton button {
            font-size: 18px !important;
        }
    </style>
""",
    unsafe_allow_html=True,
)


# --- ESTILOS CSS ---
st.markdown("""
    <style>
    .metric-card {
        background-color: #f8f9fa;
        border-left: 5px solid #d9534f;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        margin-bottom: 10px;
    }
    .stButton>button {
        width: 100%;
        background-color: #d9534f;
        color: white;
        border-radius: 6px;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

DEFAULT_FILE = "001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx"

# --- NAVEGACIÓN PRINCIPAL ---
st.sidebar.title("📌 Menú Principal")
opcion = st.sidebar.radio("Ir a:", [
    "🏠 Dashboard General",
    "DM Análisis por Calle / Cuadra",
    "🔍 Consultar Ficha por Local",
    "🔴 Tablero de Prioridades",
    "⚙️ Carga y Configuración"
])

st.sidebar.divider()
st.sidebar.subheader("📁 Archivo Activo")
uploaded_file = st.sidebar.file_uploader("Subir Excel alternativo:", type=["xlsx", "xls"], key="sidebar_uploader")

# --- FUNCIÓN DE CARGA DE DATOS PARA BASE 2026 ---
def cargar_datos(file_source):
    excel_file = pd.ExcelFile(file_source, engine='openpyxl')
    
    # Selecciona la hoja correcta o la primera disponible
    if 'BASE FISCALIZACIONES X DOMICILI' in excel_file.sheet_names:
        sheet = 'BASE FISCALIZACIONES X DOMICILI'
    elif 'TOTAL' in excel_file.sheet_names:
        sheet = 'TOTAL'
    else:
        sheet = excel_file.sheet_names[0]
        
    df = pd.read_excel(excel_file, sheet_name=sheet)
    df.columns = df.columns.astype(str).str.strip()

    # Columnas principales
    df['Calle'] = df['CALLE'].astype(str).str.strip() if 'CALLE' in df.columns else df.iloc[:, 3].astype(str).str.strip()
    col_num = 'Núm.' if 'Núm.' in df.columns else ('Num' if 'Num' in df.columns else df.columns[4])
    
    # Procesamiento de números y direcciones
    df['Num_Val'] = pd.to_numeric(df[col_num], errors='coerce').fillna(0).astype(int)
    df['Cuadra'] = (df['Num_Val'] // 100) * 100
    df['Cuadra_Texto'] = df['Calle'] + " al " + df['Cuadra'].astype(str)
    df['Núm_Clean'] = df[col_num].fillna('').astype(str).str.replace('.0', '', regex=False).str.strip()
    df['Direccion_Corta'] = df['Calle'] + " " + df['Núm_Clean']

    # Columnas complementarias
    df['TREL'] = pd.to_numeric(df['TREL'], errors='coerce').fillna(0) if 'TREL' in df.columns else 0
    df['TNR_Estado'] = df['TNR'].astype(str).str.strip().str.upper() if 'TNR' in df.columns else 'S/D'
    df['Es_Irregular'] = df['TNR_Estado'].apply(lambda x: 1 if x in ['IRREGULAR', 'NEGATIVA', 'OBSTRUCCION'] else 0)

    # Manejo de fechas
    col_fecha = 'FECHA' if 'FECHA' in df.columns else ('Fecha' if 'Fecha' in df.columns else None)
    if col_fecha:
        df['Fecha_Clean'] = pd.to_datetime(df[col_fecha], errors='coerce')
    else:
        df['Fecha_Clean'] = pd.NaT

    # Agrupación por local
    agg_dict = {
        'Calle_Nombre': ('Calle', 'first'),
        'Cuadra_Texto': ('Cuadra_Texto', 'first'),
        'Cant_Inspecciones': ('Calle', 'count'),
        'Total_TREL': ('TREL', 'sum'),
        'Cant_Irregulares': ('Es_Irregular', 'sum'),
        'Ultimo_Estado': ('TNR_Estado', 'last'),
        'Ultima_Inspeccion': ('Fecha_Clean', 'max'),
        'Razon_Social': ('RAZON SOCIAL', 'first') if 'RAZON SOCIAL' in df.columns else ('Calle', 'first'),
        'Localidad': ('Localidad', 'first') if 'Localidad' in df.columns else ('Calle', 'first')
    }

    resumen = df.groupby('Direccion_Corta').agg(**agg_dict).reset_index()
    resumen['% Irregularidad'] = ((resumen['Cant_Irregulares'] / resumen['Cant_Inspecciones']) * 100).round(1)
    
    def asignar_prioridad(row):
        if row['Cant_Inspecciones'] == 1 and row['Cant_Irregulares'] > 0:
            return "🔴 ALTA (1 sola insp. e irregular)"
        elif row['Cant_Inspecciones'] <= 2:
            return "🟡 MEDIA (1-2 inspecciones)"
        else:
            return "🟢 BAJA (3+ inspecciones)"
            
    resumen['Prioridad'] = resumen.apply(asignar_prioridad, axis=1)
    return df, resumen

df_raw, resumen = None, None

file_to_process = None
if uploaded_file is not None:
    file_to_process = uploaded_file
    st.sidebar.success("Usando archivo subido")
elif os.path.exists(DEFAULT_FILE):
    file_to_process = DEFAULT_FILE
    st.sidebar.info("Usando base predeterminada 2026")

if file_to_process is not None:
    try:
        df_raw, resumen = cargar_datos(file_to_process)
    except Exception as e:
        st.error(f"Error cargando la base de datos: {e}")

if resumen is not None:

    # --- SECCIÓN 1: DASHBOARD GENERAL ---
    if opcion == "🏠 Dashboard General":
        st.title("📊 Panel de Control e Inspecciones 2026")
        st.caption("Visión sintética de la tasa de inspecciones y estado en calle.")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total Locales / Domicilios", len(resumen))
        c2.metric("Total Inspecciones Realizadas", resumen['Cant_Inspecciones'].sum())
        c3.metric("Total Trabajadores Registrados (TREL)", int(resumen['Total_TREL'].sum()))
        c4.metric("Prioridad Alta (Refiscalizar)", len(resumen[resumen['Prioridad'].str.contains("ALTA")]))

        st.divider()
        col_left, col_right = st.columns(2)
        with col_left:
            st.subheader("Distribución por Nivel de Prioridad")
            st.bar_chart(resumen['Prioridad'].value_counts())
        with col_right:
            st.subheader("Top 10 Domicilios con Mayor Cantidad de Inspecciones")
            top_insp = resumen.sort_values(by='Cant_Inspecciones', ascending=False).head(10)
            st.dataframe(top_insp[['Direccion_Corta', 'Razon_Social', 'Cant_Inspecciones', 'Ultimo_Estado']], use_container_width=True)

    # --- SECCIÓN 2: ANÁLISIS POR CALLE / CUADRA ---
    elif opcion == "DM Análisis por Calle / Cuadra":
        st.title("DM Control por Calle y Cuadras")
        st.write("Identificá qué cuadras o calles tienen sobreinspección y cuáles faltan recorrer.")

        lista_calles = sorted([c for c in resumen['Calle_Nombre'].unique() if str(c).strip() != ''])
        calle_sel = st.selectbox("Seleccioná o buscá una Calle:", [""] + lista_calles)

        if calle_sel:
            df_calle = resumen[resumen['Calle_Nombre'].str.upper() == calle_sel.upper()]

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Calle", calle_sel)
            m2.metric("Total Domicilios", len(df_calle))
            m3.metric("Total Inspecciones", df_calle['Cant_Inspecciones'].sum())
            m4.metric("Prom. Inspecciones p/ Local", f"{(df_calle['Cant_Inspecciones'].sum() / len(df_calle)):.1f}")

            st.divider()

            col_cuadra, col_locales = st.columns([1, 1.2])

            with col_cuadra:
                st.subheader("📌 Inspecciones por Cuadra (Altura)")
                resumen_cuadra = df_calle.groupby('Cuadra_Texto').agg(
                    Locales=('Direccion_Corta', 'count'),
                    Inspecciones=('Cant_Inspecciones', 'sum'),
                    Total_TREL=('Total_TREL', 'sum')
                ).reset_index().sort_values(by='Inspecciones', ascending=False)
                
                st.dataframe(resumen_cuadra, use_container_width=True)

            with col_locales:
                st.subheader("🏪 Domicilios de la Calle")
                filtro_cant = st.radio("Mostrar:", ["Todos", "Solo 1 inspección (Prioridad)", "2 o más inspecciones"], horizontal=True)

                df_mostrar = df_calle.copy()
                if filtro_cant == "Solo 1 inspección (Prioridad)":
                    df_mostrar = df_mostrar[df_mostrar['Cant_Inspecciones'] == 1]
                elif filtro_cant == "2 o más inspecciones":
                    df_mostrar = df_mostrar[df_mostrar['Cant_Inspecciones'] >= 2]

                st.dataframe(
                    df_mostrar[['Direccion_Corta', 'Razon_Social', 'Cant_Inspecciones', 'Ultimo_Estado', 'Prioridad']].sort_values(by='Cant_Inspecciones', ascending=True),
                    use_container_width=True
                )

    # --- SECCIÓN 3: CONSULTAR FICHA POR LOCAL ---
    elif opcion == "🔍 Consultar Ficha por Local":
        st.title("🔍 Buscador Interactivo de Domicilio / Comercio")
        
        busqueda = st.selectbox("Seleccioná o buscá una dirección exacta:", [""] + list(resumen['Direccion_Corta'].unique()))
        
        if busqueda:
            local = resumen[resumen['Direccion_Corta'] == busqueda].iloc[0]
            st.success(f"📍 Ficha: **{local['Direccion_Corta']}** - {local['Razon_Social']}")
            
            f1, f2, f3, f4 = st.columns(4)
            f1.metric("Cant. Inspecciones", local['Cant_Inspecciones'])
            f2.metric("Trabajadores (TREL)", int(local['Total_TREL']))
            f3.metric("Última Inspección", str(local['Ultima_Inspeccion'])[:10] if pd.notnull(local['Ultima_Inspeccion']) else "S/D")
            f4.metric("Último Estado", local['Ultimo_Estado'])

            st.divider()
            st.subheader("📋 Historial de Inspecciones Registradas")
            historial = df_raw[df_raw['Direccion_Corta'] == busqueda]
            st.dataframe(historial[['FECHA', 'RAZON SOCIAL', 'CALLE', 'Núm.', 'Localidad', 'TREL', 'TNR', 'Inspec.', 'Expediente']], use_container_width=True)

    # --- SECCIÓN 4: TABLERO DE PRIORIDADES ---
    elif opcion == "🔴 Tablero de Prioridades":
        st.title("🔴 Tablero de Refiscalización Prioritaria")
        
        prio_filtro = st.multiselect(
            "Filtrar por Nivel de Prioridad:",
            options=list(resumen['Prioridad'].unique()),
            default=list(resumen['Prioridad'].unique())
        )
        
        res_filtrado = resumen[resumen['Prioridad'].isin(prio_filtro)]
        st.dataframe(
            res_filtrado[['Direccion_Corta', 'Razon_Social', 'Localidad', 'Cant_Inspecciones', 'Total_TREL', 'Ultimo_Estado', 'Prioridad']].sort_values(by=['Cant_Inspecciones'], ascending=[True]),
            use_container_width=True
        )

    # --- SECCIÓN 5: CARGA Y CONFIGURACIÓN ---
    elif opcion == "⚙️ Carga y Configuración":
        st.title("⚙️ Carga y Actualización de Archivos")
        st.info("💡 La aplicación está vinculada a la base oficial `001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx`.")
else:
    st.warning("Esperando carga de base de datos...")
