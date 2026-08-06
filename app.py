import json
import os
import time
import pandas as pd
import streamlit as st

# --- LIBRERÍAS PARA EL MAPA DE CALOR Y GEOLOCALIZACIÓN ---
import folium
from folium.plugins import HeatMap, Fullscreen
from geopy.geocoders import Nominatim
from streamlit_folium import st_folium

st.set_page_config(
    page_title="Control de Refiscalización 2026",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================================
# --- SISTEMA DE AUTENTICACIÓN (LOGIN CON USUARIO Y CONTRASEÑA) ---
# ==============================================================================
def check_password():
    """Retorna True si el usuario ingresó las credenciales correctas."""

    def password_entered():
        """Verifica si el usuario y la clave coinciden con los secretos guardados."""
        username = st.session_state.get("username", "")
        password = st.session_state.get("password", "")

        allowed_users = st.secrets.get("passwords", {"admin": "1234"})

        if username in allowed_users and allowed_users[username] == password:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("## 🔒 Acceso Restringido - Control de Refiscalización")
        st.text_input("Usuario", key="username")
        st.text_input("Contraseña", type="password", key="password")
        st.button("Iniciar Sesión", on_click=password_entered)
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("## 🔒 Acceso Restringido - Control de Refiscalización")
        st.text_input("Usuario", key="username")
        st.text_input("Contraseña", type="password", key="password")
        st.button("Iniciar Sesión", on_click=password_entered)
        st.error("😕 Usuario o contraseña incorrectos.")
        return False
    else:
        return True

if not check_password():
    st.stop()

if st.sidebar.button("🚪 Cerrar Sesión"):
    st.session_state["password_correct"] = False
    st.rerun()

# ==============================================================================
# --- DICCIONARIO CENTRALIZADO DE INICIALES/CÓDIGOS A NOMBRES COMPLETOS ---
# ==============================================================================
MAPEO_INICIALES = {
    "A": "Aníbal",
    "ANIBAL": "Aníbal",
    "AR": "Ariel",
    "ARIEL": "Ariel",
    "C": "Cynthia",
    "CINTIA": "Cynthia",
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

# --- ESTILOS CSS INTERACTIVOS Y TIPOGRAFÍA EN SIDEBAR ---
st.markdown(
    """
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

    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
        font-size: 20px !important;
        font-weight: 600 !important;
        line-height: 1.5 !important;
    }
    
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        padding-top: 8px !important;
        padding-bottom: 8px !important;
    }

    [data-testid="stSidebar"] h1 {
        font-size: 26px !important;
    }

    [data-testid="stSidebar"] h3 {
        font-size: 22px !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

DEFAULT_FILE_2026 = "001-BASE COMPARTIDA FISCALIZACIONES 2026.xlsx"
DEFAULT_FILE_2025 = "04-BASE FISCALIZACIONES POR DOMICILIO 2025.xlsx"
CACHE_FILE = "coordenadas.json"

# --- NAVEGACIÓN PRINCIPAL ---
st.sidebar.title("📌 Menú Principal")
opcion = st.sidebar.radio(
    "Ir a:",
    [
        "🏠 Tablero General",
        "DM Análisis por Calle / Cuadra",
        "🔄 Comparativa 2025 vs 2026",
        "🔍 Consultar Ficha por Local",
        "📋 Ranking de Inspectores",
        "🔴 Tablero de Prioridades",
        "🗺️ Mapa de Control",
        "⚙️ Carga y Configuración",
    ],
)

st.sidebar.divider()
st.sidebar.subheader("📁 Base 2026 (Dinámica)")
uploaded_file_2026 = st.sidebar.file_uploader(
    "Subir Excel 2026 alternativo:",
    type=["xlsx", "xls"],
    key="uploader_2026",
)

# --- MANEJO DE CACHÉ DE COORDENADAS ---
def cargar_cache_coords():
  if os.path.exists(CACHE_FILE):
    try:
      with open(CACHE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      return {}
  return {}

def guardar_cache_coords(cache):
  try:
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
      json.dump(cache, f, ensure_ascii=False, indent=2)
  except Exception as e:
    st.error(f"Error al guardar caché de coordenadas: {e}")

def geocodificar_direcciones_seguro(df_direcciones):
  cache = cargar_cache_coords()
  geolocator = Nominatim(user_agent="app_fiscalizaciones_2026_v7", timeout=5)

  pendientes = [
      row
      for _, row in df_direcciones.iterrows()
      if row["Direccion_Corta"] not in cache
  ]
  total_pendientes = len(pendientes)

  if total_pendientes > 0:
    progress_bar = st.progress(0)
    status_text = st.empty()

    for idx, row in enumerate(pendientes):
      direccion = row["Direccion_Corta"]
      localidad = (
          row["Localidad"]
          if pd.notnull(row["Localidad"]) and str(row["Localidad"]).strip() != ""
          else "Mar del Plata"
      )

      if pd.notnull(direccion) and str(direccion).strip() != "":
        query = f"{direccion}, {localidad}, Buenos Aires, Argentina"
        lat, lon = None, None

        for intento in range(2):
          try:
            location = geolocator.geocode(query)
            if location:
              lat, lon = location.latitude, location.longitude
            break
          except Exception:
            time.sleep(1.5)

        cache[direccion] = (lat, lon)

      time.sleep(1.1)
      porcentaje = (idx + 1) / total_pendientes
      progress_bar.progress(porcentaje)
      status_text.text(
          f"Procesando nuevas direcciones: {idx + 1} de {total_pendientes}..."
      )

      if (idx + 1) % 10 == 0:
        guardar_cache_coords(cache)

    guardar_cache_coords(cache)
    status_text.empty()
    progress_bar.empty()

  return cache

# --- FUNCIÓN DE CARGA Y NORMALIZACIÓN DE BASES ---
def cargar_datos(file_source):
  excel_file = pd.ExcelFile(file_source, engine="openpyxl")

  if "BASE FISCALIZACIONES X DOMICILI" in excel_file.sheet_names:
    sheet = "BASE FISCALIZACIONES X DOMICILI"
  elif "TOTAL" in excel_file.sheet_names:
    sheet = "TOTAL"
  else:
    sheet = excel_file.sheet_names[0]

  df = pd.read_excel(excel_file, sheet_name=sheet)
  df.columns = df.columns.astype(str).str.strip()

  col_cuit = (
      "Cuit"
      if "Cuit" in df.columns
      else ("CUIT" if "CUIT" in df.columns else None)
  )
  if col_cuit:
    df["CUIT_Clean"] = df[col_cuit].fillna("-").astype(str).str.strip()
  else:
    df["CUIT_Clean"] = "-"

  df["Calle"] = (
      df["CALLE"].astype(str).str.strip()
      if "CALLE" in df.columns
      else df.iloc[:, 3].astype(str).str.strip()
  )
  col_num = (
      "Núm."
      if "Núm." in df.columns
      else ("Num" if "Num" in df.columns else df.columns[4])
  )

  df["Num_Val"] = (
      pd.to_numeric(df[col_num], errors="coerce").fillna(0).astype(int)
  )
  df["Cuadra"] = (df["Num_Val"] // 100) * 100
  df["Cuadra_Texto"] = df["Calle"] + " al " + df["Cuadra"].astype(str)
  df["Núm_Clean"] = (
      df[col_num]
      .fillna("")
      .astype(str)
      .str.replace(".0", "", regex=False)
      .str.strip()
  )
  df["Direccion_Corta"] = df["Calle"] + " " + df["Núm_Clean"]

  df["TREL"] = (
      pd.to_numeric(df["TREL"], errors="coerce").fillna(0)
      if "TREL" in df.columns
      else 0
  )
  df["TNR_Estado"] = (
      df["TNR"].astype(str).str.strip().str.upper()
      if "TNR" in df.columns
      else "S/D"
  )
  df["Es_Irregular"] = df["TNR_Estado"].apply(
      lambda x: 1 if x in ["IRREGULAR", "NEGATIVA", "OBSTRUCCION"] else 0
  )

  col_fecha = (
      "FECHA"
      if "FECHA" in df.columns
      else ("Fecha" if "Fecha" in df.columns else None)
  )
  if col_fecha:
    df["Fecha_Clean"] = pd.to_datetime(df[col_fecha], errors="coerce")
  else:
    df["Fecha_Clean"] = pd.NaT

  col_inspector = (
      "Inspec."
      if "Inspec." in df.columns
      else (
          "INSPECTOR"
          if "INSPECTOR" in df.columns
          else ("Inspector" if "Inspector" in df.columns else None)
      )
  )
  if col_inspector:
    df["Inspector_Clean"] = (
        df[col_inspector].fillna("Otros / Sin Identificar").astype(str).str.strip()
    )
  else:
    df["Inspector_Clean"] = "Otros / Sin Identificar"

  agg_dict = {
      "Calle_Nombre": ("Calle", "first"),
      "Cuadra_Texto": ("Cuadra_Texto", "first"),
      "Cant_Inspecciones": ("Calle", "count"),
      "Total_TREL": ("TREL", "sum"),
      "Cant_Irregulares": ("Es_Irregular", "sum"),
      "Ultimo_Estado": ("TNR_Estado", "last"),
      "Ultima_Inspeccion": ("Fecha_Clean", "max"),
      "Razon_Social": (
          ("RAZON SOCIAL", "first")
          if "RAZON SOCIAL" in df.columns
          else ("Calle", "first")
      ),
      "CUIT": ("CUIT_Clean", "first"),
      "Localidad": (
          ("Localidad", "first")
          if "Localidad" in df.columns
          else ("Calle", "first")
      ),
  }

  resumen = df.groupby("Direccion_Corta").agg(**agg_dict).reset_index()
  resumen["% Irregularidad"] = (
      (resumen["Cant_Irregulares"] / resumen["Cant_Inspecciones"]) * 100
  ).round(1)

  def asignar_prioridad(row):
    if row["Cant_Inspecciones"] == 1 and row["Cant_Irregulares"] > 0:
      return "🔴 ALTA (1 sola insp. e irregular)"
    elif row["Cant_Inspecciones"] <= 2:
      return "🟡 MEDIA (1-2 inspecciones)"
    else:
      return "🟢 BAJA (3+ inspecciones)"

  resumen["Prioridad"] = resumen.apply(asignar_prioridad, axis=1)
  return df, resumen

# --- CARGA BASE 2026 ---
df_raw, resumen = None, None
file_2026 = (
    uploaded_file_2026
    if uploaded_file_2026
    else (DEFAULT_FILE_2026 if os.path.exists(DEFAULT_FILE_2026) else None)
)

if file_2026:
  try:
    df_raw, resumen = cargar_datos(file_2026)
  except Exception as e:
    st.error(f"Error cargando Base 2026: {e}")

# --- CARGA BASE 2025 (FIJA EN CACHÉ) ---
@st.cache_data(show_spinner=False)
def cargar_base_2025_fija(path):
  if os.path.exists(path):
    try:
      return cargar_datos(path)
    except Exception:
      return None, None
  return None, None

df_raw_2025, resumen_2025 = cargar_base_2025_fija(DEFAULT_FILE_2025)

if resumen is not None:

  # --- SECCIÓN 1: DASHBOARD GENERAL ---
  if opcion == "🏠 Tablero General":
    st.title("📊 Panel de Control e Inspecciones 2026")
    st.caption("Visión sintética de la tasa de inspecciones y estado en calle.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Locales / Domicilios", len(resumen))
    c2.metric(
        "Total Inspecciones Realizadas", resumen["Cant_Inspecciones"].sum()
    )
    c3.metric(
        "Total Trabajadores Registrados (TREL)", int(resumen["Total_TREL"].sum())
    )
    c4.metric(
        "Prioridad Alta (Refiscalizar)",
        len(resumen[resumen["Prioridad"].str.contains("ALTA")]),
    )

    st.divider()
    col_left, col_right = st.columns(2)
    with col_left:
      st.subheader("Distribución por Nivel de Prioridad")
      st.bar_chart(resumen["Prioridad"].value_counts())
    with col_right:
      st.subheader("Top 10 Domicilios con Mayor Cantidad de Inspecciones")
      top_insp = resumen.sort_values(
          by="Cant_Inspecciones", ascending=False
      ).head(10)
      st.dataframe(
          top_insp[[
              "Direccion_Corta",
              "Razon_Social",
              "CUIT",
              "Cant_Inspecciones",
              "Ultimo_Estado",
          ]],
          use_container_width=True,
      )

  # --- SECCIÓN 2: ANÁLISIS POR CALLE / CUADRA ---
  elif opcion == "DM Análisis por Calle / Cuadra":
    st.title("DM Control por Calle y Cuadras")
    st.write("Identificá qué cuadras o calles tienen sobreinspección.")

    lista_calles = sorted(
        [c for c in resumen["Calle_Nombre"].unique() if str(c).strip() != ""]
    )
    calle_sel = st.selectbox("Seleccioná o buscá una Calle:", [""] + lista_calles)

    if calle_sel:
      df_calle = resumen[
          resumen["Calle_Nombre"].str.upper() == calle_sel.upper()
      ]

      m1, m2, m3, m4 = st.columns(4)
      m1.metric("Calle", calle_sel)
      m2.metric("Total Domicilios", len(df_calle))
      m3.metric("Total Inspecciones", df_calle["Cant_Inspecciones"].sum())
      m4.metric(
          "Prom. Inspecciones p/ Local",
          f"{(df_calle['Cant_Inspecciones'].sum() / len(df_calle)):.1f}",
      )

      st.divider()
      col_cuadra, col_locales = st.columns([1, 1.2])

      with col_cuadra:
        st.subheader("📌 Inspecciones por Cuadra (Altura)")
        resumen_cuadra = (
            df_calle.groupby("Cuadra_Texto")
            .agg(
                Locales=("Direccion_Corta", "count"),
                Inspecciones=("Cant_Inspecciones", "sum"),
                Total_TREL=("Total_TREL", "sum"),
            )
            .reset_index()
            .sort_values(by="Inspecciones", ascending=False)
        )
        st.dataframe(resumen_cuadra, use_container_width=True)

      with col_locales:
        st.subheader("🏪 Domicilios de la Calle")
        filtro_cant = st.radio(
            "Mostrar:",
            [
                "Todos",
                "Solo 1 inspección (Prioridad)",
                "2 o más inspecciones",
            ],
            horizontal=True,
        )

        df_mostrar = df_calle.copy()
        if filtro_cant == "Solo 1 inspección (Prioridad)":
          df_mostrar = df_mostrar[df_mostrar["Cant_Inspecciones"] == 1]
        elif filtro_cant == "2 o más inspecciones":
          df_mostrar = df_mostrar[df_mostrar["Cant_Inspecciones"] >= 2]

        st.dataframe(
            df_mostrar[[
                "Direccion_Corta",
                "Razon_Social",
                "CUIT",
                "Cant_Inspecciones",
                "Ultimo_Estado",
                "Prioridad",
            ]].sort_values(by="Cant_Inspecciones", ascending=True),
            use_container_width=True,
        )

  # --- SECCIÓN 3: COMPARATIVA INTERANUAL 2025 vs 2026 ---
  elif opcion == "🔄 Comparativa 2025 vs 2026":
    st.title("🔄 Cruce de Control Interanual (2025 vs 2026)")
    st.write(
        "Detectá la situación de cada local fiscalizado en 2025 respecto al"
        " control actual de 2026."
    )

    if resumen_2025 is not None:
      all_calles = sorted(
          list(
              set(resumen["Calle_Nombre"].unique()).union(
                  set(resumen_2025["Calle_Nombre"].unique())
              )
          )
      )
      calle_comp = st.selectbox(
          "Seleccioná una Calle para comparar:", [""] + all_calles
      )

      if calle_comp:
        r25 = resumen_2025[
            resumen_2025["Calle_Nombre"].str.upper() == calle_comp.upper()
        ].copy()
        r26 = resumen[
            resumen["Calle_Nombre"].str.upper() == calle_comp.upper()
        ].copy()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Locales Fiscalizados 2025", len(r25))
        k2.metric("Locales Fiscalizados 2026", len(r26))
        k3.metric("Inspecciones Totales 2025", r25["Cant_Inspecciones"].sum())
        k4.metric("Inspecciones Totales 2026", r26["Cant_Inspecciones"].sum())

        st.divider()

        dirs_2025 = set(r25["Direccion_Corta"].unique())
        dirs_2026 = set(r26["Direccion_Corta"].unique())

        faltantes = dirs_2025 - dirs_2026
        ambos = dirs_2025.intersection(dirs_2026)
        nuevos_2026 = dirs_2026 - dirs_2025

        c1, c2, c3 = st.columns(3)
        c1.error(
            f"🔴 **Faltan Inspeccionar en 2026:** {len(faltantes)} locales"
        )
        c2.success(f"🟢 **Inspeccionados en Ambos Años:** {len(ambos)} locales")
        c3.info(f"🔵 **Nuevos Detectados en 2026:** {len(nuevos_2026)} locales")

        st.divider()

        r25_sub = r25[[
            "Direccion_Corta",
            "Razon_Social",
            "CUIT",
            "Cuadra_Texto",
            "Cant_Inspecciones",
            "Ultimo_Estado",
        ]].rename(
            columns={
                "Cant_Inspecciones": "Insp_2025",
                "Ultimo_Estado": "Estado_2025",
            }
        )

        r26_sub = r26[[
            "Direccion_Corta",
            "Razon_Social",
            "CUIT",
            "Cuadra_Texto",
            "Cant_Inspecciones",
            "Ultimo_Estado",
        ]].rename(
            columns={
                "Cant_Inspecciones": "Insp_2026",
                "Ultimo_Estado": "Estado_2026",
            }
        )

        merged = pd.merge(
            r25_sub,
            r26_sub,
            on=["Direccion_Corta"],
            how="outer",
            suffixes=("_2025", "_2026"),
        )

        merged["Razon_Social"] = merged["Razon_Social_2026"].fillna(
            merged["Razon_Social_2025"]
        )
        merged["CUIT"] = merged["CUIT_2026"].fillna(merged["CUIT_2025"])
        merged["Cuadra_Texto"] = merged["Cuadra_Texto_2026"].fillna(
            merged["Cuadra_Texto_2025"]
        )

        merged["Insp_2025"] = merged["Insp_2025"].fillna(0).astype(int)
        merged["Insp_2026"] = merged["Insp_2026"].fillna(0).astype(int)
        merged["Estado_2025"] = merged["Estado_2025"].fillna("-")
        merged["Estado_2026"] = merged["Estado_2026"].fillna("-")

        def categorizar_estado(row):
          if row["Insp_2025"] > 0 and row["Insp_2026"] == 0:
            return "🔴 PENDIENTE 2026"
          elif row["Insp_2025"] > 0 and row["Insp_2026"] > 0:
            return "🟢 CONTROLADO EN AMBOS AÑOS"
          else:
            return "🔵 NUEVO EN 2026"

        merged["Estado Interanual"] = merged.apply(categorizar_estado, axis=1)

        cols_orden = [
            "Estado Interanual",
            "Direccion_Corta",
            "Razon_Social",
            "CUIT",
            "Insp_2025",
            "Insp_2026",
            "Estado_2025",
            "Estado_2026",
        ]

        tab1, tab2, tab3, tab4 = st.tabs([
            "🔴 Pendientes 2026",
            "🟢 Controlados en Ambos Años",
            "🔵 Nuevos en 2026",
            "📋 Vista Unificada (Todos)",
        ])

        with tab1:
          st.subheader(
              f"🔴 Locales de {calle_comp} inspeccionados en 2025 pero NO en"
              " 2026"
          )
          df_p = merged[
              merged["Estado Interanual"] == "🔴 PENDIENTE 2026"
          ].sort_values(by="Insp_2025", ascending=False)
          st.dataframe(df_p[cols_orden], use_container_width=True)

        with tab2:
          st.subheader(
              f"🟢 Locales de {calle_comp} con inspecciones en 2025 y 2026"
          )
          df_c = merged[
              merged["Estado Interanual"] == "🟢 CONTROLADO EN AMBOS AÑOS"
          ].sort_values(by="Insp_2026", ascending=False)
          st.dataframe(df_c[cols_orden], use_container_width=True)

        with tab3:
          st.subheader(f"🔵 Locales nuevos relevados en {calle_comp} durante 2026")
          df_n = merged[
              merged["Estado Interanual"] == "🔵 NUEVO EN 2026"
          ].sort_values(by="Insp_2026", ascending=False)
          st.dataframe(df_n[cols_orden], use_container_width=True)

        with tab4:
          st.subheader(
              f"📋 Todos los locales registrados en {calle_comp} (2025 / 2026)"
          )
          st.dataframe(
              merged[cols_orden].sort_values(
                  by="Estado Interanual", ascending=True
              ),
              use_container_width=True,
          )

    else:
      st.warning(
          f"No se encontró el archivo fijo `{DEFAULT_FILE_2025}` en la carpeta"
          " raíz del proyecto."
      )

  # --- SECCIÓN 4: CONSULTAR FICHA / CONSOLIDADO ---
  elif opcion == "🔍 Consultar Ficha por Local":
    st.title("🔍 Buscador Interactivo y Consolidado")

    col_b1, col_b2 = st.columns(2)
    with col_b1:
      busqueda_dir = st.selectbox(
          "Seleccioná o buscá una dirección exacta:",
          [""] + list(resumen["Direccion_Corta"].unique()),
      )
    with col_b2:
      busqueda_cuit_rs = st.text_input(
          "🔍 O buscá por CUIT o Razón Social (Consolidado):",
          placeholder="Ej: 30-12345678-9 o Nombre de Empresa...",
      )

    if busqueda_cuit_rs.strip():
      query_text = busqueda_cuit_rs.strip()

      res_match = resumen[
          resumen["CUIT"]
          .astype(str)
          .str.contains(query_text, case=False, na=False)
          | resumen["Razon_Social"]
          .astype(str)
          .str.contains(query_text, case=False, na=False)
      ]

      if not res_match.empty:
        razones_unicas = ", ".join(res_match["Razon_Social"].unique())
        cuits_unicos = ", ".join(res_match["CUIT"].unique())

        tot_locales = len(res_match)
        tot_inspecciones = res_match["Cant_Inspecciones"].sum()
        tot_trel = int(res_match["Total_TREL"].sum())
        ultima_fecha = res_match["Ultima_Inspeccion"].max()

        st.success(
            f"🏢 **Consolidado Contribuyente / Empresa:** {razones_unicas}  \n🆔"
            f" **CUIT:** `{cuits_unicos}`"
        )

        f1, f2, f3, f4 = st.columns(4)
        f1.metric("Total Locales / Domicilios", tot_locales)
        f2.metric("Total Inspecciones Acumuladas", tot_inspecciones)
        f3.metric("Total Trabajadores (TREL)", tot_trel)
        f4.metric(
            "Última Inspección Registrada",
            (
                str(ultima_fecha)[:10]
                if pd.notnull(ultima_fecha)
                else "S/D"
            ),
        )

        st.divider()
        st.subheader("📍 Domicilios y Locales de la Empresa")
        st.dataframe(
            res_match[[
                "Direccion_Corta",
                "Localidad",
                "Cant_Inspecciones",
                "Total_TREL",
                "Ultimo_Estado",
                "Prioridad",
            ]],
            use_container_width=True,
        )

        st.subheader("📋 Historial Completo de Inspecciones 2026")
        direcciones_grupo = res_match["Direccion_Corta"].tolist()
        historial_grupo = df_raw[
            df_raw["Direccion_Corta"].isin(direcciones_grupo)
        ]

        cols_historial = [
            "FECHA",
            "RAZON SOCIAL",
            "CUIT_Clean",
            "CALLE",
            "Núm.",
            "Localidad",
            "TREL",
            "TNR",
            "Inspec.",
            "Expediente",
        ]
        cols_presentes = [
            c for c in cols_historial if c in historial_grupo.columns
        ]

        st.dataframe(
            historial_grupo[cols_presentes].rename(
                columns={"CUIT_Clean": "CUIT", "Inspec.": "Inspector"}
            ),
            use_container_width=True,
        )
      else:
        st.warning(
            "No se encontraron coincidencias para el CUIT o Razón Social"
            " ingresado."
        )

    elif busqueda_dir:
      local = resumen[resumen["Direccion_Corta"] == busqueda_dir].iloc[0]
      st.success(
          f"📍 Ficha Local: **{local['Direccion_Corta']}** -"
          f" {local['Razon_Social']} | **CUIT:** `{local['CUIT']}`"
      )

      f1, f2, f3, f4 = st.columns(4)
      f1.metric("Cant. Inspecciones", local["Cant_Inspecciones"])
      f2.metric("Trabajadores (TREL)", int(local["Total_TREL"]))
      f3.metric(
          "Última Inspección",
          (
              str(local["Ultima_Inspeccion"])[:10]
              if pd.notnull(local["Ultima_Inspeccion"])
              else "S/D"
          ),
      )
      f4.metric("Último Estado", local["Ultimo_Estado"])

      st.divider()
      st.subheader("📋 Historial de Inspecciones Registradas")
      historial = df_raw[df_raw["Direccion_Corta"] == busqueda_dir]

      cols_historial = [
          "FECHA",
          "RAZON SOCIAL",
          "CUIT_Clean",
          "CALLE",
          "Núm.",
          "Localidad",
          "TREL",
          "TNR",
          "Inspec.",
          "Expediente",
      ]
      cols_presentes = [c for c in cols_historial if c in historial.columns]

      st.dataframe(
          historial[cols_presentes].rename(
              columns={"CUIT_Clean": "CUIT", "Inspec.": "Inspector"}
          ),
          use_container_width=True,
      )

  # --- SECCIÓN 5: RANKING Y DESEMPEÑO DE INSPECTORES ---
  elif opcion == "📋 Ranking de Inspectores":
    st.title("📋 Ranking y Desempeño Individual de Inspectores")
    st.write(
        "Comparativa de rendimiento aplicando el tradicional **Método Real** y el nuevo **Método Ramadori** con carga proporcional."
    )

    def desglosar_inspectores(cadena):
      if not isinstance(cadena, str) or not cadena.strip():
        return ["Otros / Sin Identificar"]

      partes = str(cadena).replace("/", "-").replace(" ", "-").split("-")
      nombres = []

      for parte in partes:
        p_limpia = parte.strip().upper()
        if not p_limpia:
          continue

        if p_limpia in MAPEO_INICIALES:
          nombres.append(MAPEO_INICIALES[p_limpia])
        else:
          i = 0
          while i < len(p_limpia):
            if (
                i + 2 <= len(p_limpia)
                and p_limpia[i : i + 2] in MAPEO_INICIALES
            ):
              nombres.append(MAPEO_INICIALES[p_limpia[i : i + 2]])
              i += 2
            elif p_limpia[i] in MAPEO_INICIALES:
              nombres.append(MAPEO_INICIALES[p_limpia[i]])
              i += 1
            else:
              nombres.append(f"⚠️ Sin mapear ({p_limpia[i]})")
              i += 1

      return list(set(nombres)) if nombres else ["Otros / Sin Identificar"]

    df_exp = df_raw.copy()
    
    # Cálculo de proporciones (Incluyendo la irregularidad corregida para Ramadori)
    df_exp['Cant_Inspectores_Fila'] = df_exp['Inspector_Clean'].astype(str).apply(lambda x: len(desglosar_inspectores(x)))
    df_exp['Inspecciones_Ramadori'] = 1 / df_exp['Cant_Inspectores_Fila']
    df_exp['Trabajadores_Ramadori'] = df_exp['TREL'] / df_exp['Cant_Inspectores_Fila']
    df_exp['Es_Irregular_Ramadori'] = df_exp['Es_Irregular'] / df_exp['Cant_Inspectores_Fila']
    df_exp['Inspecciones_Real'] = 1

    df_exp["Inspectores_Lista"] = df_exp["Inspector_Clean"].apply(desglosar_inspectores)
    df_explotado = df_exp.explode("Inspectores_Lista").rename(
        columns={"Inspectores_Lista": "Inspector_Individual"}
    )

    # --- TABLA MÉTODO REAL ---
    tabla_metodo_real = (
        df_explotado.groupby("Inspector_Individual")
        .agg(
            Total_Inspecciones=("Inspecciones_Real", "sum"),
            Locales_Unicos=("Direccion_Corta", "nunique"),
            Total_TREL=("TREL", "sum"),
            Irregularidades=("Es_Irregular", "sum"),
        )
        .reset_index()
    )
    tabla_metodo_real["% Irregularidad"] = ((tabla_metodo_real["Irregularidades"] / tabla_metodo_real["Total_Inspecciones"]) * 100).round(1)
    tabla_metodo_real = tabla_metodo_real.sort_values(by="Total_Inspecciones", ascending=False).rename(columns={"Inspector_Individual": "Inspector"})

    # --- TABLA MÉTODO RAMADORI (CORREGIDA) ---
    tabla_metodo_ramadori = (
        df_explotado.groupby("Inspector_Individual")
        .agg(
            Total_Inspecciones=("Inspecciones_Ramadori", "sum"),
            Locales_Unicos=("Direccion_Corta", "nunique"),
            Total_TREL=("Trabajadores_Ramadori", "sum"),
            Irregularidades=("Es_Irregular_Ramadori", "sum"), # Irregularidad fraccionada proporcionalmente
        )
        .reset_index()
    )
    tabla_metodo_ramadori["% Irregularidad"] = ((tabla_metodo_ramadori["Irregularidades"] / tabla_metodo_ramadori["Total_Inspecciones"]) * 100).round(1)
    tabla_metodo_ramadori = tabla_metodo_ramadori.sort_values(by="Total_Inspecciones", ascending=False).rename(columns={"Inspector_Individual": "Inspector"})

    st.divider()

    # --- BUSCADOR Y FILTRO POR INSPECTOR ---
    st.subheader("🔍 Búsqueda Específica por Inspector")
    lista_inspectores_unicos = sorted(tabla_metodo_real["Inspector"].unique().tolist())
    inspector_buscado = st.selectbox(
        "Seleccioná un inspector para ver sus totales bajo ambos métodos:",
        ["(Ver listado completo)"] + lista_inspectores_unicos
    )

    if inspector_buscado != "(Ver listado completo)":
      df_real_filt = tabla_metodo_real[tabla_metodo_real["Inspector"] == inspector_buscado]
      df_rama_filt = tabla_metodo_ramadori[tabla_metodo_ramadori["Inspector"] == inspector_buscado]
      
      st.success(f"📌 Resultados para: **{inspector_buscado}**")
      
      col_fi1, col_fi2 = st.columns(2)
      with col_fi1:
          st.markdown("##### 🏢 Método Real (Totales)")
          st.dataframe(df_real_filt, use_container_width=True)
      with col_fi2:
          st.markdown("##### 🧠 Método Ramadori (Totales)")
          st.dataframe(df_rama_filt, use_container_width=True)
      
      st.divider()

    # --- RANKINGS COMPLETOS ---
    st.subheader("📊 Rankings Completos")
    col_real, col_ramadori = st.columns(2)

    with col_real:
        st.markdown("### 🏢 Método Real")
        st.caption("Inspección entera y total de trabajadores por cada inspector.")
        st.dataframe(
            tabla_metodo_real[[
                "Inspector",
                "Total_Inspecciones",
                "Locales_Unicos",
                "Total_TREL",
                "Irregularidades",
                "% Irregularidad",
            ]],
            use_container_width=True,
        )

    with col_ramadori:
        st.markdown("### 🧠 Método Ramadori")
        st.caption("Carga de trabajo y trabajadores fraccionados.")
        st.dataframe(
            tabla_metodo_ramadori[[
                "Inspector",
                "Total_Inspecciones",
                "Locales_Unicos",
                "Total_TREL",
                "Irregularidades",
                "% Irregularidad",
            ]],
            use_container_width=True,
        )

    st.divider()
    st.subheader("📊 Gráfico de Inspecciones (Método Real)")
    st.bar_chart(data=tabla_metodo_real.set_index("Inspector")["Total_Inspecciones"])

  # --- SECCIÓN 6: TABLERO DE PRIORIDADES ---
  elif opcion == "🔴 Tablero de Prioridades":
    st.title("🔴 Tablero de Refiscalización Prioritaria")
    st.write(
        "Listado de locales que requieren una nueva visita (Generalmente 1 sola "
        "inspección con resultado irregular)."
    )

    prioridades_altas = resumen[resumen["Prioridad"].str.contains("ALTA")]
    
    st.metric("Total Locales Prioridad Alta", len(prioridades_altas))

    st.dataframe(
        prioridades_altas[[
            "Direccion_Corta",
            "Razon_Social",
            "CUIT",
            "Cant_Inspecciones",
            "Total_TREL",
            "Ultimo_Estado",
            "Prioridad",
        ]].sort_values(by="Cant_Inspecciones", ascending=True),
        use_container_width=True,
    )

  # --- SECCIÓN 7: MAPA DE CONTROL ---
  elif opcion == "🗺️ Mapa de Control":
    st.title("🗺️ Mapa de Georreferenciación y Calor")
    st.write("Visualización espacial de las fiscalizaciones.")

    with st.spinner("Geocodificando direcciones faltantes (esto puede demorar un momento la primera vez)..."):
        cache_coords = geocodificar_direcciones_seguro(resumen)

    mapa = folium.Map(location=[-38.0055, -57.5426], zoom_start=13)

    heat_data = []
    for _, row in resumen.iterrows():
        direccion = row["Direccion_Corta"]
        coords = cache_coords.get(direccion)
        if coords and coords[0] is not None:
            color = "green"
            if "ALTA" in row["Prioridad"]:
                color = "red"
            elif "MEDIA" in row["Prioridad"]:
                color = "orange"

            folium.CircleMarker(
                location=coords,
                radius=6,
                popup=f"<b>{direccion}</b><br>{row['Razon_Social']}<br>Inspecciones: {row['Cant_Inspecciones']}<br>Estado: {row['Ultimo_Estado']}",
                color=color,
                fill=True,
                fill_color=color,
                fill_opacity=0.7,
            ).add_to(mapa)
            heat_data.append(coords)

    if heat_data:
        HeatMap(heat_data).add_to(mapa)

    Fullscreen().add_to(mapa)
    st_folium(mapa, width="100%", height=600)

  # --- SECCIÓN 8: CARGA Y CONFIGURACIÓN ---
  elif opcion == "⚙️ Carga y Configuración":
    st.title("⚙️ Carga y Configuración")
    st.write("Vista de los datos procesados en memoria y gestión de la herramienta.")

    st.subheader("Base de Datos Procesada (Memoria 2026)")
    st.dataframe(df_raw, use_container_width=True)

    st.divider()
    st.subheader("Mantenimiento")
    if st.button("🗑️ Borrar Caché de Coordenadas"):
        if os.path.exists(CACHE_FILE):
            os.remove(CACHE_FILE)
            st.success("Caché de coordenadas eliminado exitosamente. Se regenerará en la próxima carga del mapa.")
            st.rerun()
        else:
            st.info("No hay caché de coordenadas local para borrar.")

else:
    st.warning(
        "⚠️ No se han cargado datos válidos. Por favor, subí un archivo Excel "
        "desde el menú lateral para comenzar o verificá que el archivo por defecto "
        "exista en la misma carpeta."
    )
