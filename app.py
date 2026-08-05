import json
import os
import streamlit as st
import pandas as pd
import numpy as np
import folium
from streamlit_folium import st_folium
from folium.plugins import HeatMap, Fullscreen

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema de Gestión de Inspecciones - Control de Refiscalización",
    page_icon="📋",
    layout="wide"
)

# --- SISTEMA DE AUTENTICACIÓN ---
def verificar_autenticacion():
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    if not st.session_state.autenticado:
        st.sidebar.title("🔐 Acceso al Sistema")
        usuario = st.sidebar.text_input("Usuario")
        password = st.sidebar.text_input("Contraseña", type="password")
        if st.sidebar.button("Ingresar"):
            if usuario == "admin" and password == "admin":  # Ajustar según tus credenciales reales
                st.session_state.autenticado = True
                st.rerun()
            else:
                st.sidebar.error("Usuario o contraseña incorrectos")
        return False
    return True

if not verificar_autenticacion():
    st.stop()

# --- DICCIONARIO CENTRALIZADO DE INICIALES/CÓDIGOS A NOMBRES COMPLETOS ---
MAPEO_INICIALES = {
    "A": "Aníbal",
    "ANIBAL": "Aníbal",
    "AR": "Ariel",
    "ARIEL": "Ariel",
    "C": "Cynthia",
    "CINTIA": "Cynthia",
    "CYNTHIA": "Cynthia",
    "CIMINO": "Cimino",
    "F": "Fernando",
    "FERNANDO": "Fernando",
    "G": "Guillermo",
    "GUILLERMO": "Guillermo",
    "GO": "Gonzalo",
    "GONZALO": "Gonzalo",
    "H": "Hernán",
    "HERNAN": "Hernán",
    "L": "Luciano",
    "LUCIANO": "Luciano",
    "P": "Pablo",
    "PABLO": "Pablo",
    "R": "Rubén",
    "RUBEN": "Rubén",
}

# --- CARGA DE DATOS Y CACHÉ JSON ---
@st.cache_data
def cargar_configuracion_json():
    if os.path.exists("coordenadas.json"):
        with open("coordenadas.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

config_coords = cargar_configuracion_json()

@st.cache_data
def cargar_datos():
    try:
        df = pd.read_excel("inspecciones.xlsx")
    except:
        np.random.seed(42)
        fils = 150
        df = pd.DataFrame({
            "CUIT": [f"20-{np.random.randint(10000000, 40000000)}-{np.random.randint(0, 9)}" for _ in range(fils)],
            "Razon_Social": [f"Contribuyente {i}" for i in range(fils)],
            "Direccion_Corta": [f"Calle Falsa {np.random.randint(100, 900)}" for _ in range(fils)],
            "Calle": [f"Calle {np.random.randint(1, 25)}" for _ in range(fils)],
            "Inspector_Clean": np.random.choice(["AG", "C", "F", "ANIBAL", "G", "PR", "AR"], fils),
            "TREL": np.random.choice([0, 1, 2], fils),
            "Es_Irregular": np.random.choice([0, 1], fils, p=[0.7, 0.3]),
            "Latitud": -38.005 + np.random.normal(0, 0.02, fils),
            "Longitud": -57.555 + np.random.normal(0, 0.02, fils),
        })
    return df

df_raw = cargar_datos()

# --- BARRA LATERAL DE NAVEGACIÓN ---
st.sidebar.title("Navegación")
opcion = st.sidebar.radio(
    "Seleccioná una sección:",
    [
        "📊 Tablero General",
        "📋 Ranking de Inspectores",
        "🗺️ Mapa de Calor y Control",
        "🔍 Búsqueda Avanzada (CUIT / Dirección)"
    ]
)

# ==========================================
# 1. TABLERO GENERAL
# ==========================================
if opcion == "📊 Tablero General":
    st.title("📊 Tablero General de Inspecciones")
    st.write("Vista general de las métricas principales del sistema de refiscalización.")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Inspecciones", len(df_raw))
    col2.metric("Inspecciones Irregulares", int(df_raw["Es_Irregular"].sum()))
    col3.metric("Total TREL", int(df_raw["TREL"].sum()))
    
    st.divider()
    st.subheader("Resumen de Registros")
    st.dataframe(df_raw.head(15), use_container_width=True)

# ==========================================
# 2. RANKING DE INSPECTORES
# ==========================================
elif opcion == "📋 Ranking de Inspectores":
    st.title("📋 Ranking y Desempeño Individual de Inspectores")
    st.write(
        "El sistema desglosa cada combinación o pareja/trío (ej. AG -> Aníbal y"
        " Guillermo / CG -> Cynthia y Guillermo) asignándole 1 punto completo"
        " de la inspección a cada uno de los participantes."
    )

    def desglosar_inspectores(cadena):
        if not isinstance(cadena, str) or not cadena.strip():
            return ["Otros / Sin Identificar"]

        cadena_limpia = (
            str(cadena)
            .upper()
            .replace(".", " ")
            .replace("/", " ")
            .replace("-", " ")
            .replace(",", " ")
        )
        tokens = cadena_limpia.split()
        nombres = []

        for token in tokens:
            if not token:
                continue
            
            if token in MAPEO_INICIALES:
                nombres.append(MAPEO_INICIALES[token])
            else:
                encontrado_parcial = False
                i = 0
                while i < len(token):
                    if i + 2 <= len(token) and token[i:i+2] in MAPEO_INICIALES:
                        nombres.append(MAPEO_INICIALES[token[i:i+2]])
                        i += 2
                        encontrado_parcial = True
                    elif token[i] in MAPEO_INICIALES:
                        nombres.append(MAPEO_INICIALES[token[i]])
                        i += 1
                        encontrado_parcial = True
                    else:
                        i += 1
                
                if not encontrado_parcial:
                    nombres.append("Otros / Sin Identificar")

        return list(set(nombres)) if nombres else ["Otros / Sin Identificar"]

    df_exp = df_raw.copy()
    df_exp["Inspectores_Lista"] = df_exp["Inspector_Clean"].apply(desglosar_inspectores)
    df_explotado = df_exp.explode("Inspectores_Lista").rename(
        columns={"Inspectores_Lista": "Inspector_Individual"}
    )

    ranking_df = (
        df_explotado.groupby("Inspector_Individual")
        .agg(
            Total_Inspecciones=("Direccion_Corta", "count"),
            Locales_Unicos=("Direccion_Corta", "nunique"),
            Total_TREL=("TREL", "sum"),
            Irregularidades=("Es_Irregular", "sum"),
        )
        .reset_index()
    )

    ranking_df["% Irregularidad"] = (
        (ranking_df["Irregularidades"] / ranking_df["Total_Inspecciones"]) * 100
    ).round(1)
    ranking_df = ranking_df.sort_values(
        by="Total_Inspecciones", ascending=False
    ).rename(columns={"Inspector_Individual": "Inspector"})

    # --- BUSCADOR DE INSPECTOR EN BARRA LATERAL ---
    st.sidebar.divider()
    st.sidebar.subheader("🔍 Filtrar Inspector")
    inspectores_disponibles = ["Todos"] + list(ranking_df["Inspector"].unique())
    inspector_seleccionado = st.sidebar.selectbox("Seleccioná un inspector:", inspectores_disponibles)

    if inspector_seleccionado != "Todos":
        ranking_df_tabla = ranking_df[ranking_df["Inspector"] == inspector_seleccionado]
    else:
        ranking_df_tabla = ranking_df

    i1, i2, i3, i4 = st.columns(4)
    i1.metric("Inspectores Identificados", len(ranking_df))
    i2.metric(
        "Prom. Inspecciones p/ Inspector",
        f"{(ranking_df['Total_Inspecciones'].mean()):.1f}",
    )
    i3.metric(
        "Max Inspecciones (Líder)", ranking_df["Total_Inspecciones"].max()
    )
    i4.metric("Total TREL Relevados", int(ranking_df["Total_TREL"].sum()))

    st.divider()

    col_rank_tabla, col_rank_chart = st.columns([1.2, 1])

    with col_rank_tabla:
        st.subheader("🏆 Tabla General de Rendimiento")
        st.dataframe(ranking_df_tabla, use_container_width=True)

    with col_rank_chart:
        st.subheader("📊 Gráfico de Cantidad de Inspecciones")
        chart_data = ranking_df.set_index("Inspector")[["Total_Inspecciones"]]
        st.bar_chart(chart_data)

# ==========================================
# 3. MAPA DE CALOR Y CONTROL (Con Fullscreen)
# ==========================================
elif opcion == "🗺️ Mapa de Calor y Control":
    st.title("🗺️ Mapa de Calor y Puntos de Control")
    st.write("Visualización geoespacial de la concentración de inspecciones con botón de pantalla completa.")

    df_mapa = df_raw.dropna(subset=["Latitud", "Longitud"])

    if not df_mapa.empty:
        lat_centro = df_mapa["Latitud"].mean()
        lon_centro = df_mapa["Longitud"].mean()

        m = folium.Map(location=[lat_centro, lon_centro], zoom_start=13, control_scale=True)
        
        Fullscreen(
            position="topright",
            title="Expandir a Pantalla Completa",
            title_cancel="Salir de Pantalla Completa",
            force_separate_button=True
        ).add_to(m)

        heat_data = df_mapa[["Latitud", "Longitud"]].values.tolist()
        HeatMap(heat_data, radius=15, blur=10, max_zoom=1).add_to(m)

        st.subheader("Mapa Interactivo")
        st_folium(m, width="100%", height=650)
    else:
        st.warning("No hay datos geográficos suficientes para mostrar el mapa de calor.")

# ==========================================
# 4. BÚSQUEDA AVANZADA (CUIT / Dirección / Calle)
# ==========================================
elif opcion == "🔍 Búsqueda Avanzada (CUIT / Dirección)":
    st.title("🔍 Módulo de Búsqueda y Consultas")
    st.write("Consultá registros específicos por CUIT, Razón Social, Dirección o Calle.")

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        cuit_query = st.text_input("Buscar por CUIT o Razón Social:")
    with col_s2:
        dir_query = st.text_input("Buscar por Dirección o Calle:")

    df_busqueda = df_raw.copy()

    if cuit_query:
        mask_cuit = df_busqueda["CUIT"].astype(str).str.contains(cuit_query, case=False, na=False)
        if "Razon_Social" in df_busqueda.columns:
            mask_cuit = mask_cuit | df_busqueda["Razon_Social"].astype(str).str.contains(cuit_query, case=False, na=False)
        df_busqueda = df_busqueda[mask_cuit]

    if dir_query:
        mask_dir = df_busqueda["Direccion_Corta"].astype(str).str.contains(dir_query, case=False, na=False)
        if "Calle" in df_busqueda.columns:
            mask_dir = mask_dir | df_busqueda["Calle"].astype(str).str.contains(dir_query, case=False, na=False)
        df_busqueda = df_busqueda[mask_dir]

    st.subheader(f"Resultados Encontrados ({len(df_busqueda)})")
    st.dataframe(df_busqueda, use_container_width=True)

# Cierre de la barra lateral
st.sidebar.markdown("---")
st.sidebar.caption("Aguante el rojo")
