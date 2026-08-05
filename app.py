import json
import os
import streamlit as st
import pandas as pd
import numpy as np
import folium
from streamlit_folium import st_folium
from folium.plugins import HeatMap, Fullscreen

st.set_page_config(
    page_title="Sistema de Gestión de Inspecciones",
    page_icon="📋",
    layout="wide"
)

# --- CARGA DE DATOS Y CONFIGURACIÓN BASE ---
@st.cache_data
def cargar_configuracion_json():
    if os.path.exists("coordenadas.json"):
        with open("coordenadas.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

config_coords = cargar_configuracion_json()

@st.cache_data
def cargar_datos():
    # Carga directa del archivo Excel real sin datos de prueba falsos
    return pd.read_excel("inspecciones.xlsx")

df_raw = cargar_datos()

# --- MAPEO DE INICIALES A NOMBRE COMPLETO ---
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

# --- BARRA LATERAL DE NAVEGACIÓN ---
st.sidebar.title("Navegación")
opcion = st.sidebar.radio(
    "Seleccioná una sección:",
    [
        "📊 Tablero General",
        "🔍 Búsqueda Avanzada (CUIT / Dirección)",
        "🛣️ Búsqueda por Calle",
        "📋 Ranking de Inspectores",
        "🗺️ Mapa de Calor y Control"
    ]
)

# ==========================================
# 1. TABLERO GENERAL
# ==========================================
if opcion == "📊 Tablero General":
    st.title("📊 Tablero General de Inspecciones")
    st.write("Vista general de las métricas principales del sistema.")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total de Inspecciones", len(df_raw))
    col2.metric("Inspecciones Irregulares", int(df_raw["Es_Irregular"].sum()) if "Es_Irregular" in df_raw.columns else 0)
    col3.metric("Total TREL", int(df_raw["TREL"].sum()) if "TREL" in df_raw.columns else 0)
    
    st.divider()
    st.subheader("Resumen de Registros")
    st.dataframe(df_raw.head(15), use_container_width=True)

# ==========================================
# 2. BÚSQUEDA AVANZADA (CUIT / Dirección)
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
        mask_cuit = pd.Series(False, index=df_busqueda.index)
        if "CUIT" in df_busqueda.columns:
            mask_cuit = mask_cuit | df_busqueda["CUIT"].astype(str).str.contains(cuit_query, case=False, na=False)
        if "Razon_Social" in df_busqueda.columns:
            mask_cuit = mask_cuit | df_busqueda["Razon_Social"].astype(str).str.contains(cuit_query, case=False, na=False)
        df_busqueda = df_busqueda[mask_cuit]

    if dir_query:
        mask_dir = pd.Series(False, index=df_busqueda.index)
        if "Direccion_Corta" in df_busqueda.columns:
            mask_dir = mask_dir | df_busqueda["Direccion_Corta"].astype(str).str.contains(dir_query, case=False, na=False)
        if "Calle" in df_busqueda.columns:
            mask_dir = mask_dir | df_busqueda["Calle"].astype(str).str.contains(dir_query, case=False, na=False)
        df_busqueda = df_busqueda[mask_dir]

    st.subheader(f"Resultados Encontrados ({len(df_busqueda)})")
    st.dataframe(df_busqueda, use_container_width=True)

# ==========================================
# 3. BÚSQUEDA POR CALLE
# ==========================================
elif opcion == "🛣️ Búsqueda por Calle":
    st.title("🛣️ Búsqueda General por Calle")
    st.write("Filtrá y analizá las inspecciones agrupadas o filtradas directamente por nombre de calle.")

    calle_columna = "Calle" if "Calle" in df_raw.columns else ("Direccion_Corta" if "Direccion_Corta" in df_raw.columns else None)
    
    if calle_columna:
        calles_disponibles = sorted([str(x) for x in df_raw[calle_columna].dropna().unique()])
        calle_seleccionada = st.selectbox("Seleccioná o buscá una calle:", ["Todas las calles"] + calles_disponibles)

        df_calle = df_raw.copy()
        if calle_seleccionada != "Todas las calles":
            df_calle = df_calle[df_calle[calle_columna].astype(str) == calle_seleccionada]

        st.subheader(f"Registros para: {calle_seleccionada} ({len(df_calle)} encontrados)")
        st.dataframe(df_calle, use_container_width=True)
    else:
        st.warning("No se encontró una columna de calles o direcciones en el archivo Excel.")

# ==========================================
# 4. RANKING DE INSPECTORES
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

    col_inspector_src = "Inspector_Clean" if "Inspector_Clean" in df_raw.columns else (df_raw.columns[4] if len(df_raw.columns) > 4 else df_raw.columns[0])

    df_exp = df_raw.copy()
    df_exp["Inspectores_Lista"] = df_exp[col_inspector_src].apply(desglosar_inspectores)
    df_explotado = df_exp.explode("Inspectores_Lista").rename(
        columns={"Inspectores_Lista": "Inspector_Individual"}
    )

    col_conteo = "Direccion_Corta" if "Direccion_Corta" in df_raw.columns else df_raw.columns[0]
    col_trel = "TREL" if "TREL" in df_raw.columns else df_raw.columns[0]
    col_irregular = "Es_Irregular" if "Es_Irregular" in df_raw.columns else df_raw.columns[0]

    ranking_df = (
        df_explotado.groupby("Inspector_Individual")
        .agg(
            Total_Inspecciones=(col_conteo, "count"),
            Locales_Unicos=(col_conteo, "nunique"),
            Total_TREL=(col_trel, lambda x: pd.to_numeric(x, errors="coerce").sum() if col_trel in df_raw.columns else 0),
            Irregularidades=(col_irregular, lambda x: pd.to_numeric(x, errors="coerce").sum() if col_irregular in df_raw.columns else 0),
        )
        .reset_index()
    )

    ranking_df["% Irregularidad"] = np.where(
        ranking_df["Total_Inspecciones"] > 0,
        ((ranking_df["Irregularidades"] / ranking_df["Total_Inspecciones"]) * 100).round(1),
        0.0
    )
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
# 5. MAPA DE CALOR Y CONTROL
# ==========================================
elif opcion == "🗺️ Mapa de Calor y Control":
    st.title("🗺️ Mapa de Calor y Puntos de Control")
    st.write("Visualización geoespacial de la concentración de inspecciones.")

    lat_col = "Latitud" if "Latitud" in df_raw.columns else None
    lon_col = "Longitud" if "Longitud" in df_raw.columns else None

    if lat_col and lon_col:
        df_mapa = df_raw.dropna(subset=[lat_col, lon_col])
        if not df_mapa.empty:
            lat_centro = pd.to_numeric(df_mapa[lat_col], errors="coerce").mean()
            lon_centro = pd.to_numeric(df_mapa[lon_col], errors="coerce").mean()

            m = folium.Map(location=[lat_centro, lon_centro], zoom_start=13, control_scale=True)
            
            Fullscreen(
                position="topright",
                title="Expandir a Pantalla Completa",
                title_cancel="Salir de Pantalla Completa",
                force_separate_button=True
            ).add_to(m)

            heat_data = df_mapa[[lat_col, lon_col]].dropna().values.tolist()
            HeatMap(heat_data, radius=15, blur=10, max_zoom=1).add_to(m)

            st.subheader("Mapa Interactivo")
            st_folium(m, width="100%", height=650)
        else:
            st.warning("No hay datos geográficos válidos para mostrar el mapa de calor.")
    else:
        st.warning("No se encontraron las columnas de 'Latitud' y 'Longitud' en el Excel.")

st.sidebar.markdown("---")
st.sidebar.caption("Aguante el rojo")
