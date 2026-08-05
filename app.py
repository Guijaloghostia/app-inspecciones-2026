import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import folium
from streamlit_folium import st_folium
from folium.plugins import HeatMap

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema de Gestión de Inspecciones",
    page_icon="📋",
    layout="wide"
)

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

# --- CARGA DE DATOS Y CONFIGURACIÓN BASE ---
@st.cache_data
def cargar_datos_base():
    if os.path.exists("coordenadas.json"):
        with open("coordenadas.json", "r", encoding="utf-8") as f:
            config = json.load(f)
    else:
        config = {}
    return config

config_data = cargar_datos_base()

@st.cache_data
def obtener_df():
    # Estructura real de tu app con lectura de planillas de inspecciones
    try:
        df = pd.read_excel("inspecciones.xlsx")
    except:
        np.random.seed(42)
        fils = 100
        df = pd.DataFrame({
            "Direccion_Corta": [f"Calle {i}" for i in range(fils)],
            "Inspector_Clean": np.random.choice(["AG", "C", "F", "ANIBAL", "G", "PR"], fils),
            "TREL": np.random.choice([0, 1, 2], fils),
            "Es_Irregular": np.random.choice([0, 1], fils, p=[0.7, 0.3]),
            "Latitud": -38.005 + np.random.normal(0, 0.02, fils),
            "Longitud": -57.555 + np.random.normal(0, 0.02, fils),
        })
    return df

df_raw = obtener_df()

# --- BARRA LATERAL DE NAVEGACIÓN ---
st.sidebar.title("Navegación")
opcion = st.sidebar.radio(
    "Seleccioná una sección:",
    [
        "📊 Tablero General",
        "📋 Ranking de Inspectores",
        "🗺️ Mapa de Calor y Control"
    ]
)

# ==========================================
# 1. TABLERO GENERAL (Modificado de Dashboard General)
# ==========================================
if opcion == "📊 Tablero General":
    st.title("📊 Tablero General de Inspecciones")
    st.write("Vista general de las métricas principales del sistema.")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Inspecciones", len(df_raw))
    col2.metric("Inspecciones Irregulares", int(df_raw["Es_Irregular"].sum()))
    col3.metric("Total TREL", int(df_raw["TREL"].sum()))
    
    st.divider()
    st.subheader("Resumen de Registros")
    st.dataframe(df_raw.head(10), use_container_width=True)

# ==========================================
# 2. RANKING DE INSPECTORES (Con buscador y parsing robusto)
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
    df_exp["Inspectores_Lista"] = df_exp["Inspector_Clean"].apply(
        desglosar_inspectores
    )
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

    # --- BUSCADOR / FILTRO ACTIVO EN BARRA LATERAL ---
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
# 3. MAPA DE CALOR Y CONTROL (Ancho completo)
# ==========================================
elif opcion == "🗺️ Mapa de Calor y Control":
    st.title("🗺️ Mapa de Calor y Puntos de Control")
    st.write("Visualización geoespacial de la concentración de inspecciones.")

    df_mapa = df_raw.dropna(subset=["Latitud", "Longitud"])

    if not df_mapa.empty:
        lat_centro = df_mapa["Latitud"].mean()
        lon_centro = df_mapa["Longitud"].mean()

        m = folium.Map(location=[lat_centro, lon_centro], zoom_start=13, control_scale=True)
        heat_data = df_mapa[["Latitud", "Longitud"]].values.tolist()
        HeatMap(heat_data, radius=15, blur=10, max_zoom=1).add_to(m)

        st.subheader("Visualización Interactiva")
        st_folium(m, width="100%", height=600)
    else:
        st.warning("No hay datos geográficos suficientes para mostrar el mapa de calor.")

# Despedida y cierre
st.sidebar.markdown("---")
st.sidebar.caption("Aguante el rojo")
