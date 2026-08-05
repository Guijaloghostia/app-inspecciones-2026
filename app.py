import json
import os
import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="Sistema de Gestión de Inspecciones",
    page_icon="📋",
    layout="wide"
)

# --- CARGA DE DATOS Y CONFIGURACIÓN BASE ---
@st.cache_data
def cargar_datos_locales():
    posibles_nombres = ["001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx", "inspecciones.xlsx", "Inspecciones.xlsx", "data.xlsx"]
    for archivo in posibles_nombres:
        if os.path.exists(archivo):
            xls = pd.ExcelFile(archivo)
            return pd.read_excel(archivo, sheet_name=xls.sheet_names[0])
    
    fils = 150
    return pd.DataFrame({
        "Cuit": [f"20-{np.random.randint(10000000, 40000000)}-{np.random.randint(0, 9)}" for _ in range(fils)],
        "RAZON SOCIAL": [f"Contribuyente {i}" for i in range(fils)],
        "CALLE": [f"Calle {np.random.randint(1, 25)}" for _ in range(fils)],
        "Núm.": np.random.randint(100, 900, fils),
        "Inspec.": np.random.choice(["AG", "C", "F", "ANIBAL", "G", "PR", "AR"], fils),
        "TREL": np.random.choice([0, 1, 2], fils),
        "TNR": np.random.choice(["REGULAR", "IRREGULAR"], fils, p=[0.7, 0.3]),
    })

df_raw = cargar_datos_locales()

# Normalización de columnas clave
if "CALLE" in df_raw.columns and "Núm." in df_raw.columns and "Direccion_Corta" not in df_raw.columns:
    df_raw["Direccion_Corta"] = df_raw["CALLE"].astype(str) + " " + df_raw["Núm."].astype(str)

if "TNR" in df_raw.columns and "Es_Irregular" not in df_raw.columns:
    df_raw["Es_Irregular"] = df_raw["TNR"].astype(str).str.upper().str.contains("IRREGULAR").astype(int)

if "Inspec." in df_raw.columns and "Inspector_Clean" not in df_raw.columns:
    df_raw["Inspector_Clean"] = df_raw["Inspec."]

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
        "🔍 Búsqueda Avanzada (CUIT / Razón Social)",
        "🛣️ Búsqueda por Calle",
        "📍 Búsqueda por Dirección Exacta",
        "📋 Ranking de Inspectores"
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
    col3.metric("Total TREL", int(pd.to_numeric(df_raw["TREL"], errors="coerce").sum()) if "TREL" in df_raw.columns else 0)
    
    st.divider()
    st.subheader("Resumen de Registros")
    st.dataframe(df_raw.head(15), use_container_width=True)

# ==========================================
# 2. BÚSQUEDA AVANZADA (CUIT / Razón Social)
# ==========================================
elif opcion == "🔍 Búsqueda Avanzada (CUIT / Razón Social)":
    st.title("🔍 Búsqueda por CUIT o Razón Social")
    st.write("Consultá registros específicos ingresando el CUIT o la Razón Social.")

    cuit_query = st.text_input("Ingresá CUIT o Razón Social:")
    df_busqueda = df_raw.copy()

    if cuit_query:
        mask_cuit = pd.Series(False, index=df_busqueda.index)
        for col_c in ["Cuit", "CUIT"]:
            if col_c in df_busqueda.columns:
                mask_cuit = mask_cuit | df_busqueda[col_c].astype(str).str.contains(cuit_query, case=False, na=False)
        for col_r in ["RAZON SOCIAL", "Razon_Social"]:
            if col_r in df_busqueda.columns:
                mask_cuit = mask_cuit | df_busqueda[col_r].astype(str).str.contains(cuit_query, case=False, na=False)
        df_busqueda = df_busqueda[mask_cuit]

    st.subheader(f"Resultados Encontrados ({len(df_busqueda)})")
    st.dataframe(df_busqueda, use_container_width=True)

# ==========================================
# 3. BÚSQUEDA POR CALLE
# ==========================================
elif opcion == "🛣️ Búsqueda por Calle":
    st.title("🛣️ Búsqueda General por Calle")
    st.write("Filtrá y analizá las inspecciones seleccionando o escribiendo el nombre de la calle.")

    calle_columna = "CALLE" if "CALLE" in df_raw.columns else ("Calle" if "Calle" in df_raw.columns else None)
    
    if calle_columna:
        calles_disponibles = sorted([str(x) for x in df_raw[calle_columna].dropna().unique()])
        calle_seleccionada = st.selectbox("Seleccioná o buscá una calle:", ["Todas las calles"] + calles_disponibles)

        df_calle = df_raw.copy()
        if calle_seleccionada != "Todas las calles":
            df_calle = df_calle[df_calle[calle_columna].astype(str) == calle_seleccionada]

        st.subheader(f"Registros para: {calle_seleccionada} ({len(df_calle)} encontrados)")
        st.dataframe(df_calle, use_container_width=True)
    else:
        st.warning("No se encontró una columna de calles en el archivo Excel.")

# ==========================================
# 4. BÚSQUEDA POR DIRECCIÓN EXACTA
# ==========================================
elif opcion == "📍 Búsqueda por Dirección Exacta":
    st.title("📍 Búsqueda por Dirección Exacta")
    st.write("Ingresá la dirección completa o el número para filtrar con precisión.")

    dir_query = st.text_input("Buscar por Dirección o Número:")
    df_dir = df_raw.copy()

    if dir_query:
        mask_dir = pd.Series(False, index=df_dir.index)
        for col_d in ["Direccion_Corta", "CALLE", "Núm."]:
            if col_d in df_dir.columns:
                mask_dir = mask_dir | df_dir[col_d].astype(str).str.contains(dir_query, case=False, na=False)
        df_dir = df_dir[mask_dir]

    st.subheader(f"Resultados Encontrados ({len(df_dir)})")
    st.dataframe(df_dir, use_container_width=True)

# ==========================================
# 5. RANKING DE INSPECTORES
# ==========================================
elif opcion == "📋 Ranking de Inspectores":
    st.title("📋 Ranking y Desempeño Individual de Inspectores")
    st.write(
        "El sistema desglosa cada combinación o pareja/trío (ej. AG -> Aníbal y"
        " Guillermo) asignándole 1 punto completo de la inspección a cada participante."
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

    col_inspector_src = "Inspector_Clean" if "Inspector_Clean" in df_raw.columns else ("Inspec." if "Inspec." in df_raw.columns else df_raw.columns[0])

    df_exp = df_raw.copy()
    df_exp["Inspectores_Lista"] = df_exp[col_inspector_src].apply(desglosar_inspectores)
    df_explotado = df_exp.explode("Inspectores_Lista").rename(
        columns={"Inspectores_Lista": "Inspector_Individual"}
    )

    col_conteo = "Direccion_Corta" if "Direccion_Corta" in df_raw.columns else ("CALLE" if "CALLE" in df_raw.columns else df_raw.columns[0])
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

    i1, i2, i3, i4 = st.columns(4)
    i1.metric("Inspectores Identificados", len(ranking_df))
    i2.metric("Prom. Inspecciones p/ Inspector", f"{(ranking_df['Total_Inspecciones'].mean()):.1f}")
    i3.metric("Max Inspecciones (Líder)", ranking_df["Total_Inspecciones"].max())
    i4.metric("Total TREL Relevados", int(ranking_df["Total_TREL"].sum()))

    st.divider()

    col_rank_tabla, col_rank_chart = st.columns([1.2, 1])

    with col_rank_tabla:
        st.subheader("🏆 Tabla General de Rendimiento")
        st.dataframe(ranking_df, use_container_width=True)

    with col_rank_chart:
        st.subheader("📊 Gráfico de Cantidad de Inspecciones")
        chart_data = ranking_df.set_index("Inspector")[["Total_Inspecciones"]]
        st.bar_chart(chart_data)

st.sidebar.markdown("---")
st.sidebar.caption("Aguante el rojo")
