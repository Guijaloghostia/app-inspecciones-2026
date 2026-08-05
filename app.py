import streamlit as st
import pandas as pd

# --- MAPEO DE INICIALES (Ajustalo según tus datos) ---
MAPEO_INICIALES = {
    # Agregá acá tus mapeos si los usas
}

# Configuración inicial de la página
st.set_page_config(
    page_title="Sistema de Inspecciones",
    page_icon="📋",
    layout="wide"
)

# Simulación de carga de datos (reemplazalo por tu lectura real si es necesario)
@st.cache_data
def cargar_datos():
    # Acá iría tu pd.read_csv o similar
    # Dejamos un DataFrame vacío o de estructura base por seguridad
    return pd.DataFrame()

df_raw = cargar_datos()

# Menú lateral o selector de opciones (ejemplo basado en tu estructura)
opcion = st.sidebar.selectbox(
    "Menú de Navegación",
    ["📋 Ranking de Inspectores"] # Agregá las demás opciones de tu menú original si las tenés
)


# --- SECCIÓN 5: RANKING Y DESEMPEÑO DE INSPECTORES ---
if opcion == "📋 Ranking de Inspectores":
    st.title("📋 Ranking y Desempeño Individual de Inspectores")
    st.write(
        "El sistema desglosa cada combinación o pareja/trío asignándole"
        " la intervención a cada uno de los participantes."
    )

    def desglosar_inspectores(cadena):
        if not isinstance(cadena, str) or not cadena.strip():
            return ["Otros / Sin Identificar"]

        cadena_limpia = str(cadena).upper()
        for sep in ["/", "-", "_", ".", ","]:
            cadena_limpia = cadena_limpia.replace(sep, " ")
        
        palabras = cadena_limpia.split()
        nombres = []

        for p in palabras:
            p_limpia = p.strip()
            if not p_limpia:
                continue

            if p_limpia in MAPEO_INICIALES:
                nombres.append(MAPEO_INICIALES[p_limpia])
            else:
                encontrado = False
                for clave, nombre_completo in MAPEO_INICIALES.items():
                    if clave in p_limpia or p_limpia in clave:
                        nombres.append(nombre_completo)
                        encontrado = True
                        break
                
                if not encontrado:
                    nombres.append(p_limpia.capitalize())

        return list(set(nombres)) if nombres else ["Otros / Sin Identificar"]

    df_exp = df_raw.copy()
    
    # Validamos que existan las columnas necesarias para evitar errores previos
    if "Inspector_Clean" in df_exp.columns and "Direccion_Corta" in df_exp.columns:
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
        
        # Limpiamos nulos y ordenamos
        ranking_df = ranking_df.dropna(subset=["Inspector_Individual"])
        ranking_df = ranking_df.sort_values(
            by="Total_Inspecciones", ascending=False
        ).rename(columns={"Inspector_Individual": "Inspector"})

        i1, i2, i3, i4 = st.columns(4)
        i1.metric("Inspectores Identificados", len(ranking_df))
        i2.metric(
            "Prom. Inspecciones p/ Inspector",
            f"{(ranking_df['Total_Inspecciones'].mean()):.1f}" if not ranking_df.empty else "0",
        )
        i3.metric(
            "Max Inspecciones (Líder)", int(ranking_df["Total_Inspecciones"].max()) if not ranking_df.empty else 0
        )
        i4.metric("Total TREL Relevados", int(ranking_df["Total_TREL"].sum()) if not ranking_df.empty else 0)

        st.divider()

        col_rank_tabla, col_rank_chart = st.columns([1.2, 1])

        with col_rank_tabla:
            st.subheader("🏆 Posiciones Individuales")
            st.dataframe(
                ranking_df[[
                    "Inspector",
                    "Total_Inspecciones",
                    "Locales_Unicos",
                    "Total_TREL",
                    "Irregularidades",
                    "% Irregularidad",
                ]],
                use_container_width=True,
            )

        with col_rank_chart:
            st.subheader("📊 Gráfico de Rendimiento")
            
            metrica_grafico = st.selectbox(
                "Seleccionar métrica para el gráfico:",
                [
                    "Total de Inspecciones",
                    "Locales Únicos Visitados",
                    "Total de TREL (Trabajadores)",
                    "Cantidad de Irregularidades"
                ]
            )
            
            mapa_metricas = {
                "Total de Inspecciones": "Total_Inspecciones",
                "Locales Únicos Visitados": "Locales_Unicos",
                "Total de TREL (Trabajadores)": "Total_TREL",
                "Cantidad de Irregularidades": "Irregularidades"
            }
            
            columna_elegida = mapa_metricas[metrica_grafico]
            
            # Gráfico blindado con x e y explícitas
            st.bar_chart(
                data=ranking_df,
                x="Inspector",
                y=columna_elegida,
                use_container_width=True
            )

        st.divider()

        st.subheader("🔍 Ficha y Detalle de Actas por Inspector")
        inspectores_lista = sorted(ranking_df["Inspector"].unique()) if not ranking_df.empty else []
        inspector_sel = st.selectbox(
            "Seleccioná un inspector para ver sus intervenciones:",
            [""] + inspectores_lista,
        )

        if inspector_sel:
            df_inspector = df_explotado[
                df_explotado["Inspector_Individual"] == inspector_sel
            ]

            d1, d2, d3, d4 = st.columns(4)
            d1.metric("Inspecciones Intervenidas", len(df_inspector))
            d2.metric(
                "Locales Distintos Visitó",
                df_inspector["Direccion_Corta"].nunique(),
            )
            d3.metric("Total TREL Relevado", int(df_inspector["TREL"].sum()))
            d4.metric("Irregularidades", df_inspector["Es_Irregular"].sum())

            cols_hist_insp = [
                "FECHA",
                "RAZON SOCIAL",
                "CUIT_Clean",
                "CALLE",
                "Núm.",
                "Localidad",
                "TREL",
                "TNR",
                "Inspector_Clean",
                "Expediente",
            ]
            cols_presentes_insp = [
                c for c in cols_hist_insp if c in df_inspector.columns
            ]

            st.dataframe(
                df_inspector[cols_presentes_insp].rename(
                    columns={
                        "CUIT_Clean": "CUIT",
                        "Inspector_Clean": "Equipo Inspectivo Registrado",
                    }
                ),
                use_container_width=True,
            )
    else:
        st.warning("Faltan columnas requeridas en el conjunto de datos para procesar el ranking.")
