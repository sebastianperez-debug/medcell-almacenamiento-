"""
Dashboard de Almacenamiento — réplica en Streamlit de un reporte de Power BI.
Versión con interfaz mejorada: tema oscuro, KPIs con semáforo, tabs,
gráficos de dona/treemap y descarga del reporte filtrado.

Fuente de datos: hoja "UBICACIONES" del archivo Excel exportado desde
Power BI / el sistema de origen. Columnas esperadas:
SUB INVENTARIO, LOCALIZADOR, RESUMEN, PASILLO, NIVEL, POSICION,
ALMACENAMIENTO, OBSERVACIONES, Bodega, VACIAS, STOCK, Par
"""

import io
import os
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.colors import sample_colorscale
import streamlit as st
import streamlit.components.v1 as components
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# Configuración de página y paleta de colores
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Informe de Almacenamiento",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

COLOR_OCUPADA = "#EAB308"       # amarillo dorado (ocupado, cálido)
COLOR_DISPONIBLE = "#22C55E"    # verde vívido (disponible, libre)

COLOR_ROJO = "#FB7185"          # crítico (>90%)
COLOR_AMARILLO = "#FBBF24"      # atención (70-90%)
COLOR_VERDE = "#34D399"         # saludable (<70%)

COLOR_ACENTO_1 = "#818CF8"      # índigo claro
COLOR_ACENTO_2 = "#22C55E"      # verde
COLOR_NEUTRO = "#334155"        # slate oscuro de fondo para escalas

COLOR_CARD_BG = "#161B2C"
COLOR_CARD_BORDER = "#2A2F45"
COLOR_TEXT_MUTED = "#94A3B8"
COLOR_GRID = "rgba(148, 163, 184, 0.12)"   # grilla sutil
PLOTLY_FONT_FAMILY = "Inter, -apple-system, Segoe UI, sans-serif"

DATA_PATH = "data/Almacenamiento_2026.xlsx"

CSS = f"""
<style>
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {{
        background-color: #0B0F17 !important;
        color: #E2E8F0 !important;
    }}
    [data-testid="stHeader"] {{ background-color: #0B0F17 !important; }}
    /* Refuerzo para computadores/navegadores en modo claro: la app tiene
       tema oscuro fijado en .streamlit/config.toml, pero estos selectores
       cubren widgets nativos como respaldo adicional. */
    [data-testid="stSidebar"], [data-testid="stSidebarContent"] {{
        background-color: #0B0F17 !important;
        color: #E2E8F0 !important;
    }}
    [data-testid="stSidebar"] * {{ color: #E2E8F0 !important; }}
    [data-testid="stExpander"] {{
        background-color: {COLOR_CARD_BG} !important;
        border: 1px solid {COLOR_CARD_BORDER} !important;
        border-radius: 8px !important;
    }}
    [data-testid="stFileUploaderDropzone"] {{
        background-color: {COLOR_CARD_BG} !important;
        border: 1px dashed {COLOR_CARD_BORDER} !important;
    }}
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    [data-testid="stTextInput"] input {{
        background-color: {COLOR_CARD_BG} !important;
        color: #E2E8F0 !important;
        border-color: {COLOR_CARD_BORDER} !important;
    }}
    div[data-baseweb="popover"] ul {{
        background-color: {COLOR_CARD_BG} !important;
    }}
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] li {{ color: #E2E8F0; }}
    /* Refuerzo adicional: botones de descarga/acción y el botón "Browse
       files" del uploader, que en modo claro del navegador se ven blancos. */
    [data-testid="stDownloadButton"] button,
    [data-testid="stBaseButton-secondary"],
    [data-testid="stFileUploader"] section button,
    div[data-testid="stButton"] button {{
        background-color: {COLOR_CARD_BG} !important;
        color: #E2E8F0 !important;
        border: 1px solid {COLOR_CARD_BORDER} !important;
    }}
    [data-testid="stFileUploader"] section,
    [data-testid="stFileUploader"] small {{
        color: #E2E8F0 !important;
    }}
    /* La tabla (st.dataframe) se dibuja en canvas y toma sus colores
       directamente del tema activo de Streamlit, no de este CSS: por eso
       el arreglo de fondo para ella vive en .streamlit/config.toml. Aquí
       solo cubrimos el borde/toolbar exterior como respaldo visual. */
    [data-testid="stDataFrame"] {{
        border: 1px solid {COLOR_CARD_BORDER} !important;
        border-radius: 6px;
    }}
    /* Tarjetas KPI */
    .metric-card {{
        background-color:{COLOR_CARD_BG};
        border:1px solid {COLOR_CARD_BORDER};
        border-radius:10px;
        padding:16px 18px;
        text-align:center;
        height:100%;
    }}
    .metric-card h1 {{
        font-size:30px;
        margin:0;
        color:#F2F2F2;
        font-weight:700;
    }}
    .metric-card p {{
        margin-top:4px;
        margin-bottom:0;
        color:{COLOR_TEXT_MUTED};
        font-size:13px;
        letter-spacing:.2px;
    }}
    .metric-card .badge {{
        display:inline-block;
        margin-top:6px;
        padding:2px 10px;
        border-radius:12px;
        font-size:11px;
        font-weight:600;
    }}

    /* Header superior */
    .app-header {{
        display:flex;
        align-items:center;
        justify-content:space-between;
        padding:6px 0 18px 0;
        border-bottom:1px solid {COLOR_CARD_BORDER};
        margin-bottom:18px;
    }}
    .app-header h1 {{
        font-size:26px;
        margin:0;
        color:#F2F2F2;
    }}
    .app-header p {{
        margin:0;
        color:{COLOR_TEXT_MUTED};
        font-size:13px;
    }}

    /* Sección títulos */
    .section-title {{
        font-size:16px;
        font-weight:600;
        color:#F2F2F2;
        margin:6px 0 10px 0;
    }}

    div[data-testid="stMetricValue"] {{ font-size: 26px; }}

    /* Tarjetas de detalle por pasillo (vacías / ocupadas) */
    .pasillo-card {{
        background-color:{COLOR_CARD_BG};
        border:1px solid {COLOR_CARD_BORDER};
        border-radius:10px;
        padding:10px 8px 12px 8px;
        text-align:center;
        margin-bottom:12px;
    }}
    .pasillo-card .pasillo-nombre {{
        font-size:13px;
        font-weight:700;
        color:#F2F2F2;
        margin:0 0 5px 0;
    }}
    .pasillo-card.pasillo-card-detalle {{
        min-height:0;
        padding-left:4px;
        padding-right:4px;
    }}
    .pasillo-cat-head,
    .pasillo-cat-row {{
        display:grid;
        grid-template-columns: minmax(0, 1.7fr) minmax(0, 1fr) minmax(0, 1fr);
        gap:2px;
        align-items:center;
        justify-content:start;
        width:100%;
    }}
    .pasillo-cat-head {{
        margin-top:7px;
        padding:0 0 4px;
        color:#8EA7C7;
        font-size:7px;
        font-weight:700;
        border-bottom:1px solid #2B344A;
    }}
    .pasillo-cat-head span:nth-child(2),
    .pasillo-cat-head span:nth-child(3) {{
        text-align:center;
    }}
    .pasillo-cat-row {{
        padding:5px 0;
        font-size:10px;
        line-height:1.2;
        border-bottom:1px solid rgba(43,52,74,.55);
    }}
    .pasillo-cat-name {{
        color:#E2E8F0;
        text-align:left;
        white-space:normal;
        overflow:hidden;
        text-overflow:ellipsis;
        overflow-wrap:break-word;
        word-break:normal;
        hyphens:none;
        min-width:0;
    }}
    .pasillo-cat-empty {{
        color:#20D46B;
        font-weight:800;
        font-size:13px;
        line-height:1;
        text-align:center;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
        min-width:0;
    }}
    .pasillo-cat-occupied {{
        color:#FFC400;
        font-weight:800;
        font-size:13px;
        line-height:1;
        text-align:center;
        white-space:nowrap;
        overflow:hidden;
        text-overflow:ellipsis;
        min-width:0;
    }}
    .pasillo-stats {{
        display:flex;
        justify-content:space-around;
    }}
    .pasillo-card .pasillo-stat h1 {{
        font-size:20px;
        margin:0;
        font-weight:700;
    }}
    .pasillo-card .pasillo-stat p {{
        margin:2px 0 0 0;
        font-size:10px;
        color:{COLOR_TEXT_MUTED};
        letter-spacing:.2px;
    }}

    /* Tarjetas de bodega con barra de capacidad (semáforo) */
    .bodega-card {{
        background-color:{COLOR_CARD_BG};
        border:1px solid {COLOR_CARD_BORDER};
        border-radius:10px;
        padding:14px 16px 12px 16px;
        height:100%;
    }}
    .bodega-card .bodega-top {{
        display:flex;
        justify-content:space-between;
        align-items:baseline;
        margin-bottom:8px;
    }}
    .bodega-card .bodega-nombre {{
        font-size:13px;
        font-weight:700;
        color:#F2F2F2;
        letter-spacing:.2px;
    }}
    .bodega-card .bodega-pct {{
        font-size:18px;
        font-weight:800;
    }}
    .bodega-card .bodega-track {{
        width:100%;
        height:8px;
        border-radius:6px;
        background-color:rgba(148,163,184,.15);
        overflow:hidden;
        margin-bottom:10px;
    }}
    .bodega-card .bodega-fill {{
        height:100%;
        border-radius:6px;
    }}
    .bodega-card .bodega-stats {{
        display:flex;
        justify-content:space-between;
    }}
    .bodega-card .bodega-stat {{
        text-align:left;
    }}
    .bodega-card .bodega-stat:last-child {{
        text-align:right;
    }}
    .bodega-card .bodega-stat .valor {{
        font-size:15px;
        font-weight:700;
        color:#F2F2F2;
        display:block;
    }}
    .bodega-card .bodega-stat .etiqueta {{
        font-size:10px;
        color:{COLOR_TEXT_MUTED};
        letter-spacing:.2px;
    }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def semaforo_color(pct: float) -> str:
    """Devuelve color según nivel de ocupación."""
    if pct >= 90:
        return COLOR_ROJO
    elif pct >= 70:
        return COLOR_AMARILLO
    return COLOR_VERDE


def semaforo_label(pct: float) -> str:
    if pct >= 90:
        return "Crítico"
    elif pct >= 70:
        return "Atención"
    return "Saludable"


# ----------------------------------------------------------------------
# Carga de datos
# ----------------------------------------------------------------------
def _cache_bust_key(file) -> str:
    """Genera una 'llave' que cambia si el archivo cambia, para invalidar
    el caché de Streamlit automáticamente (aunque la ruta/nombre sea igual).
    """
    if isinstance(file, str):
        try:
            stat = os.stat(file)
            return f"{file}:{stat.st_mtime_ns}:{stat.st_size}"
        except OSError:
            return file
    # Archivo subido por file_uploader (UploadedFile): su tamaño y nombre
    # ya identifican su contenido de forma única para este propósito.
    return f"{getattr(file, 'name', '')}:{getattr(file, 'size', '')}"


@st.cache_data(show_spinner="Cargando datos...")
def load_data(file, _bust: str) -> pd.DataFrame:
    df = pd.read_excel(file, sheet_name="UBICACIONES")
    df["NIVEL"] = df["NIVEL"].astype(str)
    df["PASILLO"] = df["PASILLO"].astype(str)
    df["OCUPADA"] = df["VACIAS"] == 0
    df["ALMACENAMIENTO_FLAG"] = df["ALMACENAMIENTO"].apply(
        lambda x: "Sí" if pd.isna(x) else "No"
    )
    df["ES_RECETARIO"] = df["OBSERVACIONES"].apply(
        lambda x: "Sí" if x == "RECETARIO" else "No"
    )
    return df


@st.cache_data(show_spinner="Cargando datos de stock...")
def load_stock_data(file, _bust: str) -> pd.DataFrame:
    posibles_nombres = ["STOCK EN POSICION", "STOCK"]
    xls = pd.ExcelFile(file)
    nombre_encontrado = next(
        (n for n in posibles_nombres if n in xls.sheet_names), None
    )
    if nombre_encontrado is None:
        nombre_encontrado = next(
            (n for n in xls.sheet_names if "stock" in n.strip().lower()), None
        )
    if nombre_encontrado is None:
        raise ValueError(
            "No encontré ninguna hoja de STOCK. "
            f"Hojas disponibles en el archivo: {xls.sheet_names}"
        )

    df = pd.read_excel(xls, sheet_name=nombre_encontrado, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def fmt_code(val):
    if pd.isna(val) or val == "" or val is None or str(val).lower() == "nan":
        return "S/N"
    val_str = str(val).strip()
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str


def limpiar_numero(val):
    if pd.isna(val) or val == "" or val is None or str(val).lower() == "nan":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace("$", "").replace(" ", "")
    if not val_str:
        return 0.0

    if "." in val_str and "," in val_str:
        if val_str.rfind(",") > val_str.rfind("."):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "," in val_str:
        if val_str.count(",") == 1:
            val_str = val_str.replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "." in val_str:
        partes = val_str.split(".")
        if len(partes) > 2:
            val_str = val_str.replace(".", "")
        elif len(partes) == 2:
            if len(partes[1]) == 3 and len(partes[0]) <= 3:
                val_str = val_str.replace(".", "")

    try:
        return float(val_str)
    except (ValueError, TypeError):
        return 0.0


def formato_unidades(valor):
    try:
        val_int = int(round(valor))
        return f"{val_int:,}".replace(",", ".")
    except (ValueError, TypeError):
        return "0"


def _dedent_html(html: str) -> str:
    return "\n".join(line.strip() for line in html.strip("\n").split("\n"))


def _estilizar_hoja_resumen(ws):
    """Da formato de tabla al resumen tipo tabla dinámica (hoja 'Resumen')."""
    COLOR_HEADER = "1F2A44"     # azul oscuro (encabezado)
    COLOR_TOTAL = "E8EDF7"      # celeste muy claro (fila/columna de totales)
    COLOR_BORDE = "B9C2D0"

    max_row = ws.max_row
    max_col = ws.max_column

    borde_fino = Border(
        left=Side(style="thin", color=COLOR_BORDE),
        right=Side(style="thin", color=COLOR_BORDE),
        top=Side(style="thin", color=COLOR_BORDE),
        bottom=Side(style="thin", color=COLOR_BORDE),
    )

    # Encabezado (fila 1: "Bodega" + nombres de pasillo + "Total general")
    for col in range(1, max_col + 1):
        celda = ws.cell(row=1, column=col)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=COLOR_HEADER)
        celda.alignment = Alignment(horizontal="center", vertical="center")
        celda.border = borde_fino

    # Cuerpo de la tabla
    for fila in range(2, max_row + 1):
        es_fila_total = ws.cell(row=fila, column=1).value == "Total general"
        for col in range(1, max_col + 1):
            celda = ws.cell(row=fila, column=col)
            es_col_total = col == max_col
            celda.border = borde_fino
            if col == 1:
                celda.alignment = Alignment(horizontal="left", vertical="center")
            else:
                celda.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(celda.value, (int, float)):
                    celda.number_format = "#,##0"
            if es_fila_total or es_col_total:
                celda.font = Font(bold=True)
                celda.fill = PatternFill("solid", fgColor=COLOR_TOTAL)

    # Ancho de columnas ajustado al contenido
    for col in range(1, max_col + 1):
        letra = get_column_letter(col)
        largo = max(
            (len(str(ws.cell(row=r, column=col).value)) for r in range(1, max_row + 1)),
            default=8,
        )
        ws.column_dimensions[letra].width = max(largo + 3, 10 if col > 1 else 16)

    ws.row_dimensions[1].height = 20
    ws.freeze_panes = "B2"


with st.sidebar:
    st.header("Fuente de datos")
    uploaded = st.file_uploader(
        "Reemplazar con un Excel más nuevo (opcional)", type=["xlsx"]
    )
    st.caption(
        "Si no subes nada, se usa el archivo incluido en el repositorio "
        f"(`{DATA_PATH}`). Debe tener las hojas **UBICACIONES** y **STOCK**."
    )
    if uploaded is None and os.path.exists(DATA_PATH):
        _mtime = date.fromtimestamp(os.path.getmtime(DATA_PATH))
        st.caption(f"🕒 Última modificación del archivo: **{_mtime.strftime('%d-%m-%Y')}**")

_fuente = uploaded if uploaded is not None else DATA_PATH
_bust = _cache_bust_key(_fuente)

df_raw = None
error_ubicaciones = None
try:
    df_raw = load_data(_fuente, _bust)
except Exception as e:
    error_ubicaciones = str(e)

df_stock_raw = None
error_stock = None
try:
    df_stock_raw = load_stock_data(_fuente, _bust)
except Exception as e:
    error_stock = str(e)

if df_raw is None and df_stock_raw is None:
    st.error(
        f"No pude leer `{DATA_PATH}` ni encontrar las hojas `UBICACIONES` / "
        "`STOCK`. Sube un Excel válido desde la barra lateral."
    )
    st.stop()

# ----------------------------------------------------------------------
# Filtros (barra lateral)
# ----------------------------------------------------------------------
st.sidebar.header("Filtros — Almacenamiento")

pasillos = sorted(df_raw["PASILLO"].unique()) if df_raw is not None else []
bodegas = sorted(df_raw["Bodega"].unique()) if df_raw is not None else []

localizador_q = st.sidebar.text_input("🔎 Buscar localizador")

with st.sidebar.expander("📍 Pasillo y Bodega", expanded=False):
    bp1, bp2 = st.columns(2)
    if bp1.button("Todo", key="pasillo_todo", use_container_width=True):
        st.session_state["sel_pasillos"] = pasillos
    if bp2.button("Limpiar", key="pasillo_limpiar", use_container_width=True):
        st.session_state["sel_pasillos"] = []
    sel_pasillos = st.multiselect(
        "Pasillo", pasillos, default=[], key="sel_pasillos",
        placeholder="Todos (sin selección = todos)",
    )

    bb1, bb2 = st.columns(2)
    if bb1.button("Todo", key="bodega_todo", use_container_width=True):
        st.session_state["sel_bodegas"] = bodegas
    if bb2.button("Limpiar", key="bodega_limpiar", use_container_width=True):
        st.session_state["sel_bodegas"] = []
    sel_bodegas = st.multiselect(
        "Tipo bodega", bodegas, default=[], key="sel_bodegas",
        placeholder="Todas (sin selección = todas)",
    )

pasillos_activos = sel_pasillos if sel_pasillos else pasillos
bodegas_activas = sel_bodegas if sel_bodegas else bodegas

resumen_pasillo = "Todos" if not sel_pasillos else f"{len(sel_pasillos)} seleccionados"
resumen_bodega = "Todas" if not sel_bodegas else f"{len(sel_bodegas)} seleccionadas"
st.sidebar.caption(f"Pasillo: **{resumen_pasillo}** · Bodega: **{resumen_bodega}**")

st.sidebar.caption(
    "💡 El filtro **Código artículo / SKU** ahora vive en la pestaña "
    "**🗓️ Stock y Caducidad**, usando la hoja STOCK del mismo Excel."
)

almacenamiento_sel = st.sidebar.radio(
    "Almacenamiento", ["Sí", "No", "Todos"], index=0, horizontal=True
)
recetario_sel = st.sidebar.radio(
    "Es Recetario", ["No", "Sí", "Todos"], index=0, horizontal=True
)


def render_almacenamiento(
    df_raw, pasillos_activos, bodegas_activas, localizador_q,
    almacenamiento_sel, recetario_sel,
):
    """Renderiza el dashboard de Almacenamiento (pestaña 1)."""
    df = df_raw[
        df_raw["PASILLO"].isin(pasillos_activos) & df_raw["Bodega"].isin(bodegas_activas)
    ]
    if localizador_q:
        df = df[df["LOCALIZADOR"].str.contains(localizador_q, case=False, na=False)]
    if almacenamiento_sel != "Todos":
        df = df[df["ALMACENAMIENTO_FLAG"] == almacenamiento_sel]
    if recetario_sel != "Todos":
        df = df[df["ES_RECETARIO"] == recetario_sel]

    if df.empty:
        st.warning("No hay datos para los filtros seleccionados.")
        return

    # ----------------------------------------------------------------------
    # Cálculos base / KPIs
    # ----------------------------------------------------------------------
    total_localizadores = df["LOCALIZADOR"].nunique()
    ocupadas = int(df["OCUPADA"].sum())
    disponibles = int((~df["OCUPADA"]).sum())
    pct_ocupacion = round(ocupadas / total_localizadores * 100, 1)

    df_espacio_real = df_raw[
        df_raw["PASILLO"].isin(pasillos_activos)
        & df_raw["Bodega"].isin(bodegas_activas)
    ]
    if localizador_q:
        df_espacio_real = df_espacio_real[
            df_espacio_real["LOCALIZADOR"].str.contains(localizador_q, case=False, na=False)
        ]
    df_espacio_real = df_espacio_real.copy()
    df_espacio_real["OCUPADA_REAL"] = (
        df_espacio_real["OCUPADA"] | (df_espacio_real["ES_RECETARIO"] == "Sí")
    )

    por_pasillo_pct = (
        df_espacio_real.groupby("PASILLO")["OCUPADA_REAL"]
        .mean().mul(100).round(1).sort_values()
    )
    pasillo_mas_libre = por_pasillo_pct.index[0]
    pasillo_mas_saturado = por_pasillo_pct.index[-1]

    # ----------------------------------------------------------------------
    # Header superior
    # ----------------------------------------------------------------------
    st.markdown(
        f"""
        <div class="app-header">
            <div>
                <h1>📦 Informe de Almacenamiento</h1>
                <p>Ocupación de ubicaciones por bodega, pasillo y nivel</p>
            </div>
            <div style="text-align:right;">
                <p>Última actualización</p>
                <p style="color:#F2F2F2; font-size:15px; font-weight:600;">
                    {date.today().strftime("%d-%m-%Y")}
                </p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    def kpi_card(col, value, label, badge_text=None, badge_color=None):
        badge_html = ""
        if badge_text:
            badge_html = (
                f'<span class="badge" style="background-color:{badge_color}22;'
                f'color:{badge_color};">{badge_text}</span>'
            )
        col.markdown(
            f'<div class="metric-card"><h1>{value}</h1>'
            f"<p>{label}</p>{badge_html}</div>",
            unsafe_allow_html=True,
        )

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    kpi_card(k1, f"{total_localizadores:,}".replace(",", "."), "Localizadores")
    kpi_card(k2, f"{ocupadas:,}".replace(",", "."), "Ubicaciones ocupadas")
    kpi_card(k3, f"{disponibles:,}".replace(",", "."), "Ubicaciones disponibles")
    kpi_card(
        k4,
        f"{pct_ocupacion}%",
        "% Ocupación general",
        semaforo_label(pct_ocupacion),
        semaforo_color(pct_ocupacion),
    )
    kpi_card(
        k5,
        pasillo_mas_saturado,
        "Pasillo más saturado",
        f"{por_pasillo_pct.iloc[-1]}%",
        semaforo_color(por_pasillo_pct.iloc[-1]),
    )
    kpi_card(
        k6,
        pasillo_mas_libre,
        "Pasillo con más espacio",
        f"{por_pasillo_pct.iloc[0]}%",
        semaforo_color(por_pasillo_pct.iloc[0]),
    )

    st.caption(
        "ℹ️ *Pasillo más saturado/con más espacio* cuentan las posiciones "
        "reservadas para **RECETARIO** como no disponibles, aunque figuren vacías."
    )

    st.write("")

    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, sheet_name="UBICACIONES_FILTRADO")

    # El reporte de "solo vacías" debe usar exactamente el mismo DataFrame
    # filtrado que alimenta los letreros/KPIs. Así, la cantidad descargada
    # coincide con "Ubicaciones disponibles/vacías" que se muestra arriba.
    df_vacias = df[df["VACIAS"] > 0].copy()

    # Columnas auxiliares/internas que no queremos exportar.
    df_vacias = df_vacias.drop(
        columns=["Par", "OCUPADA", "ALMACENAMIENTO_FLAG", "ES_RECETARIO"],
        errors="ignore",
    )

    def _nombre_hoja_excel(nombre: str, usados: set) -> str:
        """Sanitiza y garantiza nombres de hoja únicos y válidos (máx 31 car.)."""
        limpio = str(nombre).strip() or "SinPasillo"
        for ch in r"[]:*?/\\":
            limpio = limpio.replace(ch, "-")
        limpio = limpio[:31]
        base = limpio
        i = 2
        while limpio in usados:
            sufijo = f"_{i}"
            limpio = f"{base[: 31 - len(sufijo)]}{sufijo}"
            i += 1
        usados.add(limpio)
        return limpio

    # Hoja 1: resumen estilo tabla dinámica (Suma de VACIAS por Bodega x Pasillo)
    pivot_vacias = pd.pivot_table(
        df_vacias,
        values="VACIAS",
        index="Bodega",
        columns="PASILLO",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="Total general",
    )

    buffer_vacias = io.BytesIO()
    with pd.ExcelWriter(buffer_vacias, engine="openpyxl") as writer:
        pivot_vacias.to_excel(writer, sheet_name="Resumen")
        _estilizar_hoja_resumen(writer.sheets["Resumen"])

        # Hojas siguientes: una por pasillo, con el detalle de ubicaciones vacías.
        hojas_usadas = {"Resumen"}
        pasillos_orden = sorted(
            df_vacias["PASILLO"].dropna().astype(str).unique(),
            key=lambda x: (0, int(float(x))) if x.replace(".", "", 1).isdigit() else (1, x),
        )
        for pasillo in pasillos_orden:
            grupo = df_vacias[df_vacias["PASILLO"].astype(str) == pasillo]
            nombre_hoja = _nombre_hoja_excel(f"Pasillo {pasillo}", hojas_usadas)
            grupo.to_excel(writer, sheet_name=nombre_hoja, index=False)

    col_desc_principal, col_desc_vacias = st.columns([3, 2])
    with col_desc_principal:
        st.download_button(
            label="⬇️ Descargar reporte filtrado (Excel)",
            data=buffer.getvalue(),
            file_name=f"almacenamiento_filtrado_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col_desc_vacias:
        st.download_button(
            label=f"⬇️ Descargar solo vacías (Excel) · {len(df_vacias):,} ubicaciones"
            .replace(",", "."),
            data=buffer_vacias.getvalue(),
            file_name=f"ubicaciones_vacias_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.write("")

    # ---- Ocupación por tipo de bodega (tarjetas con barra semáforo) --------
    st.markdown('<p class="section-title">Ocupación por tipo de bodega</p>', unsafe_allow_html=True)

    g = (
        df.groupby("Bodega")
        .agg(
            total=("LOCALIZADOR", "count"),
            ocupadas=("OCUPADA", "sum"),
            disponibles=("OCUPADA", lambda s: (~s).sum()),
        )
        .assign(pct=lambda d: (d["ocupadas"] / d["total"] * 100).round(0))
        .sort_values("pct", ascending=False)
    )

    def bodega_card(col, nombre, pct, ocupadas, disponibles):
        color = semaforo_color(pct)
        ancho = min(max(pct, 0), 100)
        col.markdown(
            f'<div class="bodega-card">'
            f'<div class="bodega-top">'
            f'<span class="bodega-nombre">{nombre}</span>'
            f'<span class="bodega-pct" style="color:{color};">{pct:.0f}%</span>'
            f'</div>'
            f'<div class="bodega-track">'
            f'<div class="bodega-fill" style="width:{ancho:.0f}%; background-color:{color};"></div>'
            f'</div>'
            f'<div class="bodega-stats">'
            f'<div class="bodega-stat">'
            f'<span class="valor" style="color:{COLOR_OCUPADA};">{ocupadas:,}</span>'
            f'<span class="etiqueta">OCUPADAS</span>'
            f'</div>'
            f'<div class="bodega-stat">'
            f'<span class="valor" style="color:{COLOR_DISPONIBLE};">{disponibles:,}</span>'
            f'<span class="etiqueta">DISPONIBLES</span>'
            f'</div>'
            f'</div>'
            f'</div>'.replace(",", "."),
            unsafe_allow_html=True,
        )

    bodegas_lista = list(g.index)
    cols_bodega = st.columns(len(bodegas_lista))
    for col, nombre in zip(cols_bodega, bodegas_lista):
        bodega_card(
            col, nombre,
            float(g.loc[nombre, "pct"]),
            int(g.loc[nombre, "ocupadas"]),
            int(g.loc[nombre, "disponibles"]),
        )

    st.write("")

    BASE_LAYOUT = dict(
        paper_bgcolor="#0B0F17",
        plot_bgcolor="#0B0F17",
        font=dict(family=PLOTLY_FONT_FAMILY, color="#E2E8F0", size=12),
        hoverlabel=dict(
            bgcolor=COLOR_CARD_BG, bordercolor=COLOR_CARD_BORDER,
            font=dict(family=PLOTLY_FONT_FAMILY, color="#F1F5F9"),
        ),
    )

    st.markdown('<p class="section-title">% Ocupación por Nivel y Pasillo</p>', unsafe_allow_html=True)

    def nivel_numero(v):
        if pd.isna(v):
            return None
        txt = str(v).strip().lower().replace("nivel", "").strip()
        try:
            n = float(txt.replace(",", "."))
            return int(n) if n.is_integer() else n
        except Exception:
            return None

    df_heat = df.copy()
    df_heat["NIVEL_NUM"] = df_heat["NIVEL"].apply(nivel_numero)
    df_heat = df_heat[df_heat["NIVEL_NUM"].isin(range(1, 8))].copy()

    niveles = list(range(1, 8))
    pasillos_heat = sorted(
        df_heat["PASILLO"].dropna().astype(str).unique(),
        key=lambda x: (0, int(float(x))) if str(x).replace('.', '', 1).isdigit() else (1, str(x))
    )

    piv_base = (
        df_heat.pivot_table(
            index="NIVEL_NUM", columns="PASILLO",
            values="OCUPADA", aggfunc="mean"
        ) * 100
    ).reindex(index=niveles, columns=pasillos_heat)

    piv = piv_base.copy()

    total_base = (
        df_heat.pivot_table(
            index="NIVEL_NUM", columns="PASILLO",
            values="OCUPADA", aggfunc="size"
        ).reindex(index=niveles, columns=pasillos_heat).fillna(0)
    )
    ocupadas_base = (
        df_heat.pivot_table(
            index="NIVEL_NUM", columns="PASILLO",
            values="OCUPADA", aggfunc="sum"
        ).reindex(index=niveles, columns=pasillos_heat).fillna(0)
    )
    vacias_base = total_base - ocupadas_base

    total_m = total_base.copy()
    ocupadas_m = ocupadas_base.copy()
    vacias_m = vacias_base.copy()

    text_vals = piv.map(lambda x: "" if pd.isna(x) else f"{x:.1f}%")

    customdata = []
    for r in piv.index:
        row = []
        for c in piv.columns:
            row.append([
                int(ocupadas_m.loc[r, c]),
                int(vacias_m.loc[r, c]),
                int(total_m.loc[r, c]),
            ])
        customdata.append(row)

    x_positions = list(range(len(piv.columns)))
    y_positions = list(range(1, 8))

    fig_heat = go.Figure(
        data=go.Heatmap(
            z=piv.values,
            x=x_positions,
            y=y_positions,
            customdata=customdata,
            colorscale=[
                [0.00, COLOR_VERDE],
                [0.50, COLOR_AMARILLO],
                [1.00, COLOR_ROJO],
            ],
            zmin=0,
            zmax=100,
            text=text_vals.values,
            texttemplate="%{text}",
            textfont=dict(
                family=PLOTLY_FONT_FAMILY,
                size=11,
                color="#0B1220",
            ),
            hoverongaps=False,
            showscale=True,
            xgap=3,
            ygap=3,
            colorbar=dict(
                title=dict(
                    text="%",
                    font=dict(family=PLOTLY_FONT_FAMILY, color="#E2E8F0"),
                ),
                tickfont=dict(
                    family=PLOTLY_FONT_FAMILY,
                    color="#E2E8F0",
                ),
                outlinewidth=0,
                len=1.0,
                thickness=14,
                y=0.5,
            ),
            hovertemplate=(
                "<b>Nivel %{y} · Pasillo %{customdata[3]}</b><br>"
                "Ocupación: %{z:.1f}%<br>"
                "<b>Posiciones vacías: %{customdata[1]:,.0f}</b><br>"
                "Posiciones ocupadas: %{customdata[0]:,.0f}<br>"
                "Posiciones totales: %{customdata[2]:,.0f}"
                "<extra></extra>"
            ),
        )
    )

    for i in range(len(customdata)):
        for j in range(len(customdata[i])):
            customdata[i][j].append(str(piv.columns[j]))
    fig_heat.data[0].customdata = customdata

    fig_heat.update_layout(
        **BASE_LAYOUT,
        height=520,
        margin=dict(l=95, r=100, t=75, b=45),
        xaxis=dict(
            title=dict(
                text="Pasillo",
                font=dict(size=12, color="#E2E8F0"),
            ),
            side="top",
            tickmode="array",
            tickvals=x_positions,
            ticktext=[str(c).strip() for c in piv.columns],
            tickangle=0,
            tickfont=dict(size=12, color="#E2E8F0"),
            showgrid=False,
            zeroline=False,
            range=[-0.5, len(piv.columns)-0.5],
        ),
        yaxis=dict(
            title=dict(
                text="Nivel",
                font=dict(size=12, color="#E2E8F0"),
            ),
            tickmode="array",
            tickvals=y_positions,
            ticktext=[f"Nivel {r}" for r in niveles],
            tickfont=dict(size=11, color="#E2E8F0"),
            range=[7.5, 0.5],
            showgrid=False,
            zeroline=False,
        ),
    )

    st.plotly_chart(
        fig_heat,
        use_container_width=True,
        config={
            "displayModeBar": False,
            "responsive": True,
            "scrollZoom": False,
        },
        theme=None,
    )

    st.caption(
        "🟢 Saludable (<70%) · 🟡 Atención (70-90%) · 🔴 Crítico (>90%)."
    )

    st.write("")



    # Datos para las tarjetas por pasillo (independiente del gráfico posterior).
    gp_cards = (
        df.groupby("PASILLO")
        .agg(ocupadas=("OCUPADA", "sum"), disponibles=("OCUPADA", lambda s: (~s).sum()))
        .reindex(sorted(df["PASILLO"].astype(str).str.strip().unique()))
    )
    gp_cards.index = gp_cards.index.astype(str).str.strip()

    st.markdown('<p class="section-title">Detalle por pasillo: vacías y ocupadas</p>', unsafe_allow_html=True)

    def pasillo_card(col, pasillo, vacias, ocupadas):
        if str(pasillo) in ("A", "B"):
            sub = df[df["PASILLO"].astype(str).str.strip() == str(pasillo)].copy()
            if "Bodega" in sub.columns:
                detalle = (
                    sub.groupby("Bodega", dropna=False)
                    .agg(
                        vacias=("VACIAS", "sum"),
                        ocupadas=("OCUPADA", "sum"),
                    )
                    .reset_index()
                )
                detalle["Bodega"] = detalle["Bodega"].fillna("SIN CATEGORÍA").astype(str)
                detalle = detalle.sort_values("Bodega")
            else:
                detalle = pd.DataFrame(columns=["Bodega", "vacias", "ocupadas"])

            nombres = {
                "COSMETICO": "COSMÉTICO",
                "DISP.MEDICOS": "DISP. MÉDICOS",
                "INFLAMABLE": "INFLAMABLE",
                "ALIMENTO": "ALIMENTO",
                "ALTILLO": "ALTILLO",
            }

            filas = ""
            for _, r in detalle.iterrows():
                categoria = nombres.get(str(r["Bodega"]).strip(), str(r["Bodega"]).strip())
                filas += (
                    f'<div class="pasillo-cat-row">'
                    f'<span class="pasillo-cat-name">{categoria}</span>'
                    f'<span class="pasillo-cat-empty">{int(r["vacias"]):,}</span>'
                    f'<span class="pasillo-cat-occupied">{int(r["ocupadas"]):,}</span>'
                    f'</div>'
                ).replace(",", ".")

            col.markdown(
                f'<div class="pasillo-card pasillo-card-detalle">'
                f'<p class="pasillo-nombre">Pasillo {pasillo}</p>'
                f'<div class="pasillo-cat-head">'
                f'<span></span><span>VACÍAS</span><span>OCUPADAS</span>'
                f'</div>'
                f'{filas}'
                f'</div>',
                unsafe_allow_html=True,
            )
        else:
            col.markdown(
                f'<div class="pasillo-card">'
                f'<p class="pasillo-nombre">Pasillo {pasillo}</p>'
                f'<div class="pasillo-stats">'
                f'<div class="pasillo-stat"><h1 style="color:{COLOR_DISPONIBLE};">{vacias:,}</h1>'
                f'<p>VACÍAS</p></div>'
                f'<div class="pasillo-stat"><h1 style="color:{COLOR_OCUPADA};">{ocupadas:,}</h1>'
                f'<p>OCUPADAS</p></div>'
                f'</div></div>'.replace(",", "."),
                unsafe_allow_html=True,
            )

    TARJETAS_POR_FILA = 7
    pasillos_lista = list(gp_cards.index)
    for i in range(0, len(pasillos_lista), TARJETAS_POR_FILA):
        fila = pasillos_lista[i:i + TARJETAS_POR_FILA]
        cols_fila = st.columns(TARJETAS_POR_FILA)
        for col, pasillo in zip(cols_fila, fila):
            pasillo_card(
                col, pasillo,
                int(gp_cards.loc[pasillo, "disponibles"]),
                int(gp_cards.loc[pasillo, "ocupadas"]),
            )

    st.write("")


    # =========================================================================
    # SECCIÓN DE LA IMAGEN 2 (AHORA ARRIBA): BODEGAS, TREEMAP Y HEATMAP
    # =========================================================================
    st.markdown('<p class="section-title">Ocupadas y disponibles por pasillo</p>', unsafe_allow_html=True)
    gp = (
        df.groupby("PASILLO")
        .agg(ocupadas=("OCUPADA", "sum"), disponibles=("OCUPADA", lambda s: (~s).sum()))
        .reindex(sorted(df["PASILLO"].unique()))
    )
    gp["total"] = gp["ocupadas"] + gp["disponibles"]

    fig_pas = go.Figure()
    fig_pas.add_trace(
        go.Bar(x=gp.index, y=gp["ocupadas"], name="Ocupadas",
               marker=dict(color=COLOR_OCUPADA, line=dict(width=0)),
               text=gp["ocupadas"], textposition="inside",
               textfont=dict(color="#0B1220", family=PLOTLY_FONT_FAMILY),
               hovertemplate="Pasillo %{x}<br>Ocupadas: %{y:,.0f}<extra></extra>")
    )
    fig_pas.add_trace(
        go.Bar(x=gp.index, y=gp["disponibles"], name="Disponibles",
               marker=dict(color=COLOR_DISPONIBLE, line=dict(width=0)),
               text=gp["disponibles"].replace(0, ""), textposition="inside",
               textfont=dict(color="#0B1220", family=PLOTLY_FONT_FAMILY),
               hovertemplate="Pasillo %{x}<br>Disponibles: %{y:,.0f}<extra></extra>")
    )
    for pasillo, row in gp.iterrows():
        fig_pas.add_annotation(
            x=pasillo, y=row["total"], text=f"{int(row['total'])}",
            showarrow=False, yshift=14,
            font=dict(size=11, color=COLOR_TEXT_MUTED, family=PLOTLY_FONT_FAMILY),
        )
    fig_pas.update_layout(
        **BASE_LAYOUT,
        barmode="stack", height=380, bargap=0.32,
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis=dict(title="Pasillo", showgrid=False, linecolor=COLOR_GRID),
        yaxis=dict(title="Ubicaciones", showgrid=True, gridcolor=COLOR_GRID,
                    zeroline=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0,
                     bgcolor="rgba(0,0,0,0)"),
    )
    st.plotly_chart(fig_pas, use_container_width=True)

    st.markdown('<p class="section-title">Distribución de ubicaciones por bodega (treemap)</p>', unsafe_allow_html=True)
    g3 = df.groupby("Bodega").size().reset_index(name="cantidad")
    fig_tree = px.treemap(
        g3, path=["Bodega"], values="cantidad",
        color="cantidad", color_continuous_scale=[COLOR_NEUTRO, COLOR_ACENTO_1, COLOR_ACENTO_2],
    )
    fig_tree.update_traces(
        marker=dict(line=dict(color=COLOR_CARD_BG, width=2)),
        textfont=dict(family=PLOTLY_FONT_FAMILY, size=14, color="#F8FAFC"),
        hovertemplate="%{label}<br>%{value:,.0f} ubicaciones<extra></extra>",
        root_color="rgba(0,0,0,0)",
    )
    fig_tree.update_layout(
        height=320, margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(family=PLOTLY_FONT_FAMILY, color="#E2E8F0"),
        coloraxis_showscale=False,
    )
    st.plotly_chart(fig_tree, use_container_width=True)

    st.markdown('<p class="section-title">Detalle de datos filtrados</p>', unsafe_allow_html=True)

    col_filtro_pas, col_filtro_estado = st.columns([1, 1])
    with col_filtro_pas:
        pasillos_tabla = st.multiselect(
            "Pasillo",
            sorted(df["PASILLO"].unique()),
            default=[],
            placeholder="Todos los pasillos",
            key="detalle_pasillo_filtro",
        )
    with col_filtro_estado:
        st.markdown(
            '<p style="margin-bottom:6px; font-size:14px;">Ver ubicaciones:</p>',
            unsafe_allow_html=True,
        )
        estado_tabla = st.radio(
            "Ver ubicaciones",
            ["Todas", "Ocupadas", "Disponibles"],
            index=0, horizontal=True, key="detalle_estado_filtro",
            label_visibility="collapsed",
        )

    df_detalle = df.copy()
    if pasillos_tabla:
        df_detalle = df_detalle[df_detalle["PASILLO"].isin(pasillos_tabla)]
    if estado_tabla == "Ocupadas":
        df_detalle = df_detalle[df_detalle["OCUPADA"]]
    elif estado_tabla == "Disponibles":
        df_detalle = df_detalle[~df_detalle["OCUPADA"]]

    df_detalle = df_detalle.drop(columns=["ALMACENAMIENTO", "OBSERVACIONES"], errors="ignore")
    if "VACIAS" in df_detalle.columns:
        df_detalle["VACIAS"] = df_detalle["VACIAS"].apply(
            lambda x: "Vacía" if x else "Con stock"
        )
    if "OCUPADA" in df_detalle.columns:
        df_detalle["OCUPADA"] = df_detalle["OCUPADA"].apply(
            lambda x: "Ocupada" if x else "Disponible"
        )
    if "Par" in df_detalle.columns:
        df_detalle["Par"] = df_detalle["Par"].apply(lambda x: "Sí" if x else "No")

    st.dataframe(df_detalle, use_container_width=True, height=320)
    st.caption(
        f"{len(df_detalle):,} ubicaciones mostradas de {len(df):,} totales."
        .replace(",", ".")
    )


def render_stock(df_stock_raw):
    """Renderiza el dashboard de Stock / Fecha de Caducidad (pestaña 2)."""
    df = df_stock_raw.copy()

    st.markdown("### 📦 Dashboard de Fecha de Caducidad")

    col_cod = next(
        (
            c
            for c in df.columns
            if c.strip().lower()
            in ["codigo_articulo", "id_producto", "sku", "codigo"]
        ),
        None,
    )
    col_estado_sub = next(
        (c for c in df.columns if c.strip().lower() == "estado_subin"), None
    ) or next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["sub_inventario", "estado sub inventario"]
        ),
        None,
    )
    col_estado_lote = next(
        (
            c
            for c in df.columns
            if c.strip().lower()
            in ["estado_lote", "estado lote", "estado_lote_prov"]
        ),
        None,
    )
    col_lote = next(
        (
            c
            for c in df.columns
            if c.strip().lower() == "lote_proveedor"
        ),
        None,
    ) or next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["lote", "lote_prov"]
        ),
        None,
    )
    col_loc = next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["localizador", "ubicacion"]
        ),
        None,
    )
    col_desc_stock = next(
        (c for c in df.columns if "descripcion" in c.lower()), None
    )
    if not col_desc_stock and len(df.columns) > 3:
      col_desc_stock = df.columns[3]
    col_fecha = next(
        (
            c
            for c in df.columns
            if c.strip().lower()
            in [
                "fecha_expiracion_lote",
                "vencimiento",
                "fecha expiracion",
                "fecha_expiracion",
            ]
        ),
        None,
    )
    col_cant = next(
        (
            c
            for c in df.columns
            if c.strip().lower() in ["cantidad", "stock", "unidades"]
        ),
        None,
    )

    if col_cod and col_cod in df.columns:
      df[col_cod] = df[col_cod].apply(fmt_code)

    col_sku_sb = next(
        (c for c in df.columns if c.strip().lower() == "codigo_sb"), None
    )
    col_sku_pu = next(
        (c for c in df.columns if c.strip().lower() == "codigo_pu"), None
    )
    if not col_sku_sb and len(df.columns) > 1:
      col_sku_sb = df.columns[1]
    if not col_sku_pu and len(df.columns) > 2:
      col_sku_pu = df.columns[2]

    if col_sku_sb and col_sku_sb in df.columns:
      df[col_sku_sb] = df[col_sku_sb].apply(fmt_code)
    if col_sku_pu and col_sku_pu in df.columns:
      df[col_sku_pu] = df[col_sku_pu].apply(fmt_code)

    if col_cant:
      df[col_cant] = df[col_cant].apply(limpiar_numero)

    hoy = pd.Timestamp.today()
    limite_6m = hoy + pd.DateOffset(months=6)
    limite_13m = hoy + pd.DateOffset(months=13)

    if col_fecha:
      df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")

      def calcular_alerta(fecha):
        if pd.isna(fecha):
          return "Sin Fecha"
        if fecha < hoy:
          return "Vencido"
        if fecha < limite_6m:
          return "Menos de 6 meses"
        elif fecha <= limite_13m:
          return "Pronto vence (6-13m)"
        else:
          return "Vigente (> 13m)"

      df["Alerta_Caducidad"] = df[col_fecha].apply(calcular_alerta)
    else:
      df["Alerta_Caducidad"] = "Sin Fecha"
      df[col_fecha] = "N/A"

    col_dash1, col_dash2 = st.columns([1, 2.3])

    key_codigo = f"sel_codigo_stock"
    key_sku_sb = f"sel_sku_sb_stock"
    key_sku_pu = f"sel_sku_pu_stock"

    def _limpiar_otros_filtros(keys_a_limpiar):
      for k in keys_a_limpiar:
        if k in st.session_state:
          st.session_state[k] = "Todos"

    with col_dash2:
      _pad_izq, filtro_codigo_col, filtro_sku_sb_col, filtro_sku_pu_col, _pad_der = (
          st.columns([0.3, 1, 1, 1, 0.3])
      )

      with filtro_codigo_col:
        if col_cod:
          lista_codigos = sorted(
              [str(x) for x in df[col_cod].dropna().unique() if str(x).strip() != ""]
          )
          codigo_sel = st.selectbox(
              "Código:",
              ["Todos"] + lista_codigos,
              key=key_codigo,
              on_change=_limpiar_otros_filtros,
              args=([key_sku_sb, key_sku_pu],),
          )
        else:
          codigo_sel = "Todos"

      with filtro_sku_sb_col:
        if col_sku_sb and col_sku_sb in df.columns:
          lista_sku_sb = sorted(
              [str(x) for x in df[col_sku_sb].dropna().unique() if str(x).strip() != "" and str(x) != "S/N"]
          )
          sku_sb_sel = st.selectbox(
              "SKU SB:",
              ["Todos"] + lista_sku_sb,
              key=key_sku_sb,
              on_change=_limpiar_otros_filtros,
              args=([key_codigo, key_sku_pu],),
          )
        else:
          sku_sb_sel = "Todos"

      with filtro_sku_pu_col:
        if col_sku_pu and col_sku_pu in df.columns:
          lista_sku_pu = sorted(
              [str(x) for x in df[col_sku_pu].dropna().unique() if str(x).strip() != "" and str(x) != "S/N"]
          )
          sku_pu_sel = st.selectbox(
              "SKU PU:",
              ["Todos"] + lista_sku_pu,
              key=key_sku_pu,
              on_change=_limpiar_otros_filtros,
              args=([key_codigo, key_sku_sb],),
          )
        else:
          sku_pu_sel = "Todos"

    df_dash = df.copy()
    if codigo_sel != "Todos" and col_cod:
      df_dash = df_dash[df_dash[col_cod].astype(str) == codigo_sel].copy()

    if sku_sb_sel != "Todos" and col_sku_sb and col_sku_sb in df_dash.columns:
      df_dash = df_dash[df_dash[col_sku_sb].astype(str) == sku_sb_sel].copy()

    if sku_pu_sel != "Todos" and col_sku_pu and col_sku_pu in df_dash.columns:
      df_dash = df_dash[df_dash[col_sku_pu].astype(str) == sku_pu_sel].copy()

    if codigo_sel != "Todos":
      prod_sel = codigo_sel
    elif sku_sb_sel != "Todos":
      prod_sel = sku_sb_sel
    elif sku_pu_sel != "Todos":
      prod_sel = sku_pu_sel
    else:
      prod_sel = "Seleccione..."

    if col_cant:
      total_unidades = df_dash[col_cant].sum()
      total_vencido = df_dash[
          df_dash["Alerta_Caducidad"] == "Vencido"
      ][col_cant].sum()
      total_menos_6m = df_dash[
          df_dash["Alerta_Caducidad"] == "Menos de 6 meses"
      ][col_cant].sum()
      total_pronto = df_dash[
          df_dash["Alerta_Caducidad"] == "Pronto vence (6-13m)"
      ][col_cant].sum()
      total_vigentes = df_dash[
          df_dash["Alerta_Caducidad"] == "Vigente (> 13m)"
      ][col_cant].sum()
    else:
      total_unidades = len(df_dash)
      total_vencido = len(df_dash[df_dash["Alerta_Caducidad"] == "Vencido"])
      total_menos_6m = len(
          df_dash[df_dash["Alerta_Caducidad"] == "Menos de 6 meses"]
      )
      total_pronto = len(
          df_dash[df_dash["Alerta_Caducidad"] == "Pronto vence (6-13m)"]
      )
      total_vigentes = len(
          df_dash[df_dash["Alerta_Caducidad"] == "Vigente (> 13m)"]
      )

    total_critico = total_vencido + total_menos_6m
    pct_critico = (
        (total_critico / total_unidades * 100) if total_unidades > 0 else 0.0
    )

    st.markdown(
        """
            <style>
            .critico-card { border-radius: 8px; padding: 14px 18px; margin-bottom: 15px;
              border: 1px solid #333; display: flex; justify-content: space-between; align-items: center; }
            </style>
            """,
        unsafe_allow_html=True,
    )
    color_pct_critico = (
        "#e74c3c" if pct_critico >= 15
        else "#f1c40f" if pct_critico >= 5
        else "#2ecc71"
    )
    st.markdown(
        '<div class="critico-card" style="background-color: #141414;">'
        '<span style="color:#aaaaaa; font-weight:600; text-transform:uppercase; font-size:13px;">'
        '⚠️ % de Stock Crítico (vencido + vence en &lt; 6 meses)</span>'
        f'<span style="color:{color_pct_critico}; font-size:26px; font-weight:bold;">{pct_critico:.2f}%</span>'
        "</div>",
        unsafe_allow_html=True,
    )

    label_map_alerta = {
        "Todos": "Todos",
        "Vencido": "Vencido",
        "Vence en < 6 meses": "Menos de 6 meses",
        "Pronto vence (6-13m)": "Pronto vence (6-13m)",
        "Vigente (> 13m)": "Vigente (> 13m)",
    }
    key_alerta = f"radio_alerta_stock"
    etiqueta_sel = st.radio(
        "🔍 Filtrar por categoría de caducidad:",
        list(label_map_alerta.keys()),
        horizontal=True,
        key=key_alerta,
    )
    filtro_actual = label_map_alerta[etiqueta_sel]

    if filtro_actual != "Todos":
      df_dash_alerta = df_dash[df_dash["Alerta_Caducidad"] == filtro_actual].copy()
    else:
      df_dash_alerta = df_dash.copy()

    with col_dash1:
      st.markdown(
          """
              <style>
              .stock-card { border-radius: 5px; padding: 15px; margin-bottom: 10px; text-align: center; color: white; font-weight: bold; }
              </style>
              """,
          unsafe_allow_html=True,
      )

      def _borde(valor):
        return "border: 2px solid #ffffff;" if filtro_actual == valor else ""

      st.markdown(
          '<div class="stock-card" style="background-color: #333; color:'
          f' white; {_borde("Todos")}">Unidades Registradas<br><span'
          f' style="font-size:24px;">{formato_unidades(total_unidades)}</span></div>',
          unsafe_allow_html=True,
      )
      st.markdown(
          '<div class="stock-card" style="background-color:'
          f' #8b0000; {_borde("Vencido")}">Vencido<br><span'
          f' style="font-size:24px;">{formato_unidades(total_vencido)}</span></div>',
          unsafe_allow_html=True,
      )
      st.markdown(
          '<div class="stock-card" style="background-color:'
          f' #e74c3c; {_borde("Menos de 6 meses")}">Vence en &lt; 6 meses<br><span'
          f' style="font-size:24px;">{formato_unidades(total_menos_6m)}</span></div>',
          unsafe_allow_html=True,
      )
      st.markdown(
          '<div class="stock-card" style="background-color: #f1c40f; color:'
          f' black; {_borde("Pronto vence (6-13m)")}">Pronto vence (6 a 13'
          ' meses)<br><span'
          f' style="font-size:24px;">{formato_unidades(total_pronto)}</span></div>',
          unsafe_allow_html=True,
      )
      st.markdown(
          '<div class="stock-card" style="background-color:'
          f' #2ecc71; {_borde("Vigente (> 13m)")}">Vigentes (> 13 meses)<br><span'
          f' style="font-size:24px;">{formato_unidades(total_vigentes)}</span></div>',
          unsafe_allow_html=True,
      )

    with col_dash2:
      st.markdown("#### Estado de caducidad")
      labels = ["Vencido", "< 6 meses", "6 a 13 meses", "Vigente (> 13m)"]
      values = [total_vencido, total_menos_6m, total_pronto, total_vigentes]
      colors = ["#8b0000", "#e74c3c", "#f1c40f", "#2ecc71"]

      total_donut = sum(values)
      if total_donut > 0:
        textos_pct = [
            f"{lbl}<br>{(v / total_donut * 100):.2f}%"
            for lbl, v in zip(labels, values)
        ]
        fig_pie = go.Figure(
            data=[
                go.Pie(
                    labels=labels,
                    values=values,
                    hole=0.55,
                    marker=dict(colors=colors, line=dict(color="#0e1117", width=2)),
                    text=textos_pct,
                    texttemplate="%{text}",
                    textposition="outside",
                    textfont=dict(size=12, color="#ffffff"),
                )
            ]
        )
        fig_pie.update_layout(
            height=380,
            margin=dict(t=20, b=60, l=60, r=60),
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#ffffff"),
            showlegend=True,
            legend=dict(
                orientation="h",
                y=-0.15,
                x=0.5,
                xanchor="center",
                yanchor="top",
            ),
        )
        _pad_chart_izq, col_chart, _pad_chart_der = st.columns([0.3, 2, 0.3])
        with col_chart:
          st.plotly_chart(
              fig_pie, use_container_width=True, key="pie_stock"
          )
      else:
        st.info("Sin registros para mostrar.")

      if prod_sel != "Seleccione...":
        stock_actual = df_dash[col_cant].sum() if col_cant else 0
        prox_vencer = (
            df_dash[df_dash[col_fecha].notna()][col_fecha].min()
            if col_fecha
            else None
        )
        dias_vencer = (
            (prox_vencer - hoy).days if pd.notna(prox_vencer) else "N/A"
        )

        st.markdown(
            '<div class="stock-card" style="background-color: #7f8c8d;">Stock'
            ' actual<br><span'
            f' style="font-size:24px;">{formato_unidades(stock_actual)}</span></div>',
            unsafe_allow_html=True,
        )

        if isinstance(dias_vencer, int):
          if prox_vencer < limite_6m:
            texto_vence = (
                f"Vence en {dias_vencer} días"
                if dias_vencer >= 0
                else f"Venció hace {abs(dias_vencer)} días"
            )
            color_vence = "#e74c3c"
            color_texto = "color: white;"
          elif prox_vencer <= limite_13m:
            texto_vence = f"Vence en {dias_vencer} días"
            color_vence = "#f1c40f"
            color_texto = "color: black;"
          else:
            texto_vence = f"Vence en {dias_vencer} días"
            color_vence = "#2ecc71"
            color_texto = "color: white;"
        else:
          texto_vence = "Sin fecha registrada"
          color_vence = "#333333"
          color_texto = "color: white;"

        st.markdown(
            '<div class="stock-card" style="background-color:'
            f" {color_vence}; {color_texto} border: 1px solid #555;\">Plazo de"
            f' vencimiento<br><span style="font-size:20px;">{texto_vence}</span></div>',
            unsafe_allow_html=True,
        )

    st.divider()

    if col_estado_lote and col_estado_lote in df_dash_alerta.columns:
      st.markdown("##### 🏷️ Cantidad de Unidades por Estado de Lote")
      df_est_grp = (
          df_dash_alerta.groupby(col_estado_lote, dropna=False)[col_cant]
          .sum()
          .reset_index()
          if col_cant
          else df_dash_alerta[col_estado_lote].value_counts().reset_index()
      )

      if not df_est_grp.empty:
        c_e, c_q = df_est_grp.columns[0], df_est_grp.columns[1]
        num_items = len(df_est_grp)
        cols_est = st.columns(min(num_items, 6))
        for idx_e, row_e in df_est_grp.iterrows():
          nombre_est = (
              str(row_e[c_e]) if pd.notna(row_e[c_e]) else "Sin Estado"
          )
          cant_est = row_e[c_q]

          with cols_est[idx_e % min(num_items, 6)]:
            st.markdown(
                _dedent_html(f"""<div style="background-color: #141414; border: 1px solid #0070f3; border-radius: 8px; padding: 10px; text-align: center; margin-bottom: 15px;">
                                  <div style="font-size: 12px; color: #aaaaaa; font-weight: 600; text-transform: uppercase;">{nombre_est}</div>
                                  <div style="font-size: 20px; font-weight: bold; color: #ffffff; margin-top: 3px;">{formato_unidades(cant_est)}</div>
                              </div>"""),
                unsafe_allow_html=True,
            )

      st.divider()

    detalle_filtro = "(General)"
    partes_filtro = []
    if codigo_sel != "Todos":
      partes_filtro.append(f"Código: {codigo_sel}")
    if sku_sb_sel != "Todos":
      partes_filtro.append(f"SKU SB: {sku_sb_sel}")
    if sku_pu_sel != "Todos":
      partes_filtro.append(f"SKU PU: {sku_pu_sel}")
    if filtro_actual != "Todos":
      partes_filtro.append(f"Caducidad: {filtro_actual}")
    if partes_filtro:
      detalle_filtro = f"({' | '.join(partes_filtro)})"

    st.subheader(f"📋 Detalle de Stock y Lotes {detalle_filtro}")

    cols_mostrar = []
    nombres_amigables = {}
    if col_cod:
      cols_mostrar.append(col_cod)
      nombres_amigables[col_cod] = "Código Artículo"
    if col_sku_sb and col_sku_sb in df_dash_alerta.columns:
      cols_mostrar.append(col_sku_sb)
      nombres_amigables[col_sku_sb] = "SKU SB"
    if col_sku_pu and col_sku_pu in df_dash_alerta.columns:
      cols_mostrar.append(col_sku_pu)
      nombres_amigables[col_sku_pu] = "SKU PU"
    if col_estado_sub:
      cols_mostrar.append(col_estado_sub)
      nombres_amigables[col_estado_sub] = "Estado Sub-Inv"
    if col_estado_lote:
      cols_mostrar.append(col_estado_lote)
      nombres_amigables[col_estado_lote] = "Estado Lote"
    if col_lote:
      cols_mostrar.append(col_lote)
      nombres_amigables[col_lote] = "Lote Proveedor"
    if col_loc:
      cols_mostrar.append(col_loc)
      nombres_amigables[col_loc] = "Localizador"
    if col_cant:
      cols_mostrar.append(col_cant)
      nombres_amigables[col_cant] = "Cantidad"
    if col_fecha:
      cols_mostrar.append(col_fecha)
      nombres_amigables[col_fecha] = "Fecha Expiración"
    cols_mostrar.append("Alerta_Caducidad")
    nombres_amigables["Alerta_Caducidad"] = "Rango Caducidad"

    df_vista_stock = df_dash_alerta[cols_mostrar].copy()
    df_vista_stock = df_vista_stock.rename(columns=nombres_amigables)

    if "Fecha Expiración" in df_vista_stock.columns:
      df_vista_stock["Fecha Expiración"] = pd.to_datetime(
          df_vista_stock["Fecha Expiración"], errors="coerce"
      ).dt.strftime("%d-%m-%Y")

    st.dataframe(df_vista_stock, hide_index=True, use_container_width=True)

    st.divider()

    if col_loc and col_loc in df_dash.columns:
      if filtro_actual != "Todos":
        titulo_loc = f"##### 📍 Top Localizadores — {filtro_actual}"
        df_critico = df_dash_alerta.copy()
      else:
        titulo_loc = "##### 📍 Top Localizadores con más Stock por Vencer"
        df_critico = df_dash[
            df_dash["Alerta_Caducidad"].isin(["Vencido", "Menos de 6 meses"])
        ].copy()

      st.markdown(titulo_loc)

      df_critico = df_critico[
          df_critico[col_loc].notna()
          & (df_critico[col_loc].astype(str).str.strip() != "")
      ].copy()

      if not df_critico.empty:
        cols_group = [col_loc]
        if col_desc_stock and col_desc_stock in df_critico.columns:
          cols_group.append(col_desc_stock)

        if col_cant:
          grp_loc = (
              df_critico.groupby(cols_group, dropna=False)[col_cant]
              .sum()
              .reset_index()
              .rename(columns={col_cant: "Cantidad"})
          )
        else:
          grp_loc = (
              df_critico.groupby(cols_group, dropna=False)
              .size()
              .reset_index(name="Cantidad")
          )

        grp_loc = grp_loc.sort_values(by="Cantidad", ascending=False).head(10)

        etiqueta_barra = (
            grp_loc[col_loc].astype(str)
            + (
                " — " + grp_loc[col_desc_stock].astype(str)
                if col_desc_stock and col_desc_stock in grp_loc.columns
                else ""
            )
        )
        grp_loc_sorted = grp_loc.assign(_etiqueta=etiqueta_barra).sort_values(
            by="Cantidad", ascending=True
        )
        fig_loc = px.bar(
            grp_loc_sorted,
            x="Cantidad",
            y="_etiqueta",
            orientation="h",
            text_auto=",.0f",
            color_discrete_sequence=["#e74c3c"],
        )
        fig_loc.update_traces(
            textfont_size=11, textposition="outside", cliponaxis=False
        )
        fig_loc.update_layout(
            template="plotly_dark",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            margin=dict(t=10, b=10, l=10, r=10),
            height=320,
            xaxis_title="",
            yaxis_title="",
        )
        st.plotly_chart(
            fig_loc, use_container_width=True, key="top_loc_stock"
        )

        rename_cols = {col_loc: "Localizador"}
        if col_desc_stock and col_desc_stock in grp_loc.columns:
          rename_cols[col_desc_stock] = "Descripción Producto"
        grp_loc_disp = grp_loc.rename(columns=rename_cols)
        st.dataframe(
            grp_loc_disp,
            column_config={
                "Cantidad": st.column_config.NumberColumn(
                    "Cantidad", format="%,d"
                ),
            },
            hide_index=True,
            use_container_width=True,
        )
      else:
        st.info(
            "No hay stock (con localizador registrado) para la categoría seleccionada."
        )


def render_escanear(df_stock_raw, hojas_disponibles):
    """Renderiza la pestaña de escaneo de Localizador (pestaña 3)."""
    st.markdown("### 📷 Escanear Localizador")
    st.caption("Apunta la cámara al texto MCD de la posición.")

    video_scan_html = """
        <div style="position:relative; width:100%; max-height:320px; overflow:hidden;
                    border-radius:8px; background:#000;">
          <video id="video" style="width:100%; max-height:320px; object-fit:cover;
                 display:block;" muted playsinline autoplay></video>
          <div style="position:absolute; top:50%; left:50%; transform:translate(-50%,-50%);
                      width:82%; height:105px; border:3px solid #00e676; border-radius:6px;
                      box-shadow:0 0 0 2000px rgba(0,0,0,0.35); pointer-events:none;"></div>
        </div>
        <div style="text-align:center; margin-top:10px; display:flex; gap:8px; justify-content:center;">
          <button id="btn-torch" style="background:#0070f3; color:#fff; border:none;
                  border-radius:8px; padding:8px 16px; font-weight:600; cursor:pointer;">
            💡 Linterna
          </button>
        </div>
        <p id="estado-scan" style="color:#888; font-size:13px; text-align:center; margin-top:6px;">
          🎥 Activando cámara...
        </p>
        <p id="detalle-scan" style="color:#666; font-size:12px; text-align:center; margin:0 8px;">
          Primero intentará reconocer el Localizador MCD directamente.
        </p>

        <script src="https://unpkg.com/@zxing/library@0.21.3/umd/index.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js"></script>
        <script>
          const estado = document.getElementById("estado-scan");
          const detalle = document.getElementById("detalle-scan");
          const video = document.getElementById("video");
          const torchBtn = document.getElementById("btn-torch");
          let yaEnvio = false;
          let streamActual = null;
          let ocrWorker = null;
          let ocrActivo = false;

          function mostrarError(msg) {
            estado.textContent = msg;
          }

          function normalizarLocalizador(texto) {
            if (!texto) return null;
            let s = String(texto).toUpperCase();
            s = s.replace(/[\n\r\t]/g, " ");
            s = s.replace(/[|]/g, "I");
            s = s.replace(/\s+/g, " ");

            const m = s.match(/MCD\s*[.\-\s]\s*\d+\s*[.\-\s]\s*\d+\s*[.\-\s]\s*[A-Z0-9]+\s*[.\-\s]\s*\d+\s*[.\-\s]\s*\d+/);
            if (!m) return null;

            let loc = m[0]
              .replace(/\s+/g, "")
              .replace(/-/g, ".");

            loc = loc.replace(/\.\.+/g, ".");
            loc = loc.replace(/^MCD/i, "MCD");

            if (/^MCD\.\d+\.\d+\.[A-Z0-9]+\.\d+\.\d+$/.test(loc)) {
              return loc;
            }
            return null;
          }

          function enviarValor(valor, origen) {
            if (yaEnvio || !valor) return;
            const loc = normalizarLocalizador(valor);
            if (!loc) return;

            yaEnvio = true;
            estado.textContent = "✅ Localizador detectado: " + loc;
            detalle.textContent = origen === "ocr"
              ? "🔎 Reconocido desde el texto de la etiqueta. Buscando productos..."
              : "📦 Obtenido desde el código de barras. Buscando productos...";

            try {
              const url = new URL(window.parent.location.href);
              url.searchParams.set("loc", loc);
              window.parent.location.href = url.href;
            } catch (e) {
              window.location.href = "?loc=" + encodeURIComponent(loc);
            }
          }

          function enviarCodigoBarras(codigo) {
            if (yaEnvio || !codigo) return;
            const valor = String(codigo).trim();

            if (!/^\d{8,14}$/.test(valor)) return;

            const loc = normalizarLocalizador(valor);
            if (loc) {
              enviarValor(loc, "barcode");
              return;
            }

            yaEnvio = true;
            estado.textContent = "✅ Código detectado: " + valor;
            detalle.textContent = "🔎 Buscando la relación código → Localizador...";
            try {
              const url = new URL(window.parent.location.href);
              url.searchParams.set("loc", valor);
              window.parent.location.href = url.href;
            } catch (e) {
              window.location.href = "?loc=" + encodeURIComponent(valor);
            }
          }

          async function iniciarOCR() {
            if (ocrActivo || typeof Tesseract === "undefined" || yaEnvio) return;
            ocrActivo = true;
            try {
              detalle.textContent = "🔎 OCR activo: busca el texto MCD.0.3.G.x.xxx...";
              ocrWorker = await Tesseract.createWorker("eng", 1, {
                logger: function(m) {
                  if (m.status === "recognizing text") {
                    const pct = Math.round((m.progress || 0) * 100);
                    estado.textContent = "🔎 Reconociendo Localizador... " + pct + "%";
                  }
                }
              });
              await ocrWorker.setParameters({
                tessedit_char_whitelist: "MCD.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-",
                preserve_interword_spaces: "0"
              });

              const canvas = document.createElement("canvas");
              const ctx = canvas.getContext("2d", {willReadFrequently:true});

              while (!yaEnvio) {
                if (!video.videoWidth || !video.videoHeight) {
                  await new Promise(r => setTimeout(r, 700));
                  continue;
                }

                const vw = video.videoWidth;
                const vh = video.videoHeight;
                const cropW = Math.floor(vw * 0.82);
                const cropH = Math.floor(vh * 0.34);
                const sx = Math.floor((vw - cropW) / 2);
                const sy = Math.floor((vh - cropH) / 2);
                canvas.width = cropW;
                canvas.height = cropH;
                ctx.drawImage(video, sx, sy, cropW, cropH, 0, 0, cropW, cropH);

                try {
                  const resultado = await ocrWorker.recognize(canvas);
                  const texto = resultado?.data?.text || "";
                  const loc = normalizarLocalizador(texto);
                  if (loc) {
                    enviarValor(loc, "ocr");
                    break;
                  }
                } catch (e) {}

                if (!yaEnvio) {
                  estado.textContent = "📷 Buscando Localizador MCD...";
                  await new Promise(r => setTimeout(r, 400));
                }
              }
            } catch (e) {
              detalle.textContent = "⚠️ OCR no disponible; se mantiene el lector de barras.";
            } finally {
              ocrActivo = false;
            }
          }

          async function iniciarCamara() {
            try {
              streamActual = await navigator.mediaDevices.getUserMedia({
                video: {
                  facingMode: {ideal: "environment"},
                  width: {ideal: 1920},
                  height: {ideal: 1080},
                  focusMode: {ideal: "continuous"}
                },
                audio: false
              });
              video.srcObject = streamActual;
              await video.play();
              estado.textContent = "📷 Buscando Localizador MCD...";

              try {
                const track = streamActual.getVideoTracks()[0];
                const caps = track.getCapabilities ? track.getCapabilities() : {};
                if (caps.focusMode && caps.focusMode.includes("continuous")) {
                  await track.applyConstraints({advanced:[{focusMode:"continuous"}]});
                }
              } catch (e) {}

              iniciarOCR();
            } catch (e) {
              mostrarError("❌ No se pudo acceder a la cámara: " + (e.message || e));
            }
          }

          function iniciarBarras() {
            if (typeof ZXing === "undefined") return;
            try {
              const hints = new Map();
              hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [
                ZXing.BarcodeFormat.CODE_128,
                ZXing.BarcodeFormat.CODE_39,
                ZXing.BarcodeFormat.EAN_13,
                ZXing.BarcodeFormat.EAN_8,
                ZXing.BarcodeFormat.ITF,
                ZXing.BarcodeFormat.UPC_A
              ]);
              hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
              const reader = new ZXing.BrowserMultiFormatReader(hints);
              reader.decodeFromVideoDevice(null, video, (result, err) => {
                if (!result || yaEnvio) return;
                enviarCodigoBarras(result.getText());
              });
            } catch (e) {}
          }

          torchBtn.onclick = function() {
            try {
              const track = streamActual && streamActual.getVideoTracks()[0];
              if (!track) throw new Error("No hay cámara activa");
              const settings = track.getSettings();
              track.applyConstraints({advanced:[{torch:!settings.torch}]})
                .catch(() => alert("Este dispositivo no permite linterna desde el navegador."));
            } catch (e) {
              alert("No se pudo acceder a la linterna.");
            }
          };

          let intentos = 0;
          const esperarLibrerias = setInterval(() => {
            intentos++;
            if (typeof ZXing !== "undefined" && typeof Tesseract !== "undefined") {
              clearInterval(esperarLibrerias);
              iniciarCamara();
              setTimeout(iniciarBarras, 1200);
            } else if (intentos >= 100) {
              clearInterval(esperarLibrerias);
              if (typeof Tesseract !== "undefined") {
                iniciarCamara();
              } else {
                mostrarError("❌ No se pudieron cargar los lectores. Recarga la página.");
              }
            }
          }, 100);
        </script>
        """

    if "mc_scan_activo" not in st.session_state:
      st.session_state["mc_scan_activo"] = False

    if not st.session_state["mc_scan_activo"]:
      if st.button("📷 Activar cámara y escanear", key="btn_activar_scan", use_container_width=True):
        st.session_state["mc_scan_activo"] = True
        st.rerun()
      st.info("Presiona el botón superior para activar la cámara y el escáner.")
    else:
      components.html(video_scan_html, height=430)

    st.caption(
        "💡 Recomendado: centra el texto MCD.0.3.G.x.xxx dentro del recuadro verde."
    )

    with st.expander("⌨️ ¿No lee el código? Ingresa el Localizador manualmente", expanded=True):
      loc_manual = st.text_input(
          "Localizador (ej: MCD.0.3.C.2.013):", key="loc_manual_input"
      )
      buscar_click = st.button("Buscar", key="btn_buscar_manual")

    loc_query = st.query_params.get("loc", None)
    loc_escaneado = None
    if buscar_click and loc_manual.strip():
      loc_escaneado = loc_manual.strip()
      st.query_params["loc"] = loc_escaneado
    elif loc_query:
      loc_escaneado = loc_query

    if st.button("🔄 Limpiar escaneo", key="btn_limpiar_scan"):
      st.query_params.clear()
      st.rerun()

    if loc_escaneado:
      df_stock_scan = df_stock_raw
      if df_stock_scan is None:
        st.error("No se encontró la hoja 'STOCK' en el Excel.")
      else:
        df_stock_scan = df_stock_scan.copy()

        col_loc_scan = next(
            (c for c in df_stock_scan.columns
             if c.strip().lower() in ["localizador", "ubicacion"]),
            None,
        )
        col_cod_scan = next(
            (c for c in df_stock_scan.columns
             if c.strip().lower() in ["codigo_articulo", "id_producto", "sku", "codigo"]),
            None,
        )
        col_desc_scan = next(
            (c for c in df_stock_scan.columns if "descripcion" in c.lower()), None
        )
        if not col_desc_scan and len(df_stock_scan.columns) > 3:
          col_desc_scan = df_stock_scan.columns[3]
        col_lote_scan = next(
            (c for c in df_stock_scan.columns if c.strip().lower() == "lote_proveedor"),
            None,
        )
        col_cant_scan = next(
            (c for c in df_stock_scan.columns
             if c.strip().lower() in ["cantidad", "stock", "unidades"]),
            None,
        )
        col_fecha_scan = next(
            (c for c in df_stock_scan.columns
             if c.strip().lower() in ["fecha_expiracion_lote", "vencimiento", "fecha_expiracion"]),
            None,
        )

        if not col_loc_scan:
          st.error("La hoja STOCK no tiene columna de Localizador reconocible.")
        else:
          def _norm_scan_value(v):
              if v is None or pd.isna(v):
                  return ""
              s = str(v).strip().upper()
              if s.endswith(".0"):
                  s = s[:-2]
              return s

          def _parece_localizador(v):
              s = _norm_scan_value(v)
              if not s:
                  return False
              partes = s.split(".")
              return len(partes) >= 5 and partes[0] == "MCD" and all(part.strip() for part in partes)

          scan_norm = _norm_scan_value(loc_escaneado)
          localizadores_encontrados = []
          hoja_mapeo = None

          MAPEO_PRUEBA = {
              "9631187073887": "MCD.0.3.G.2.120",
              "11111283": "MCD.0.3.G.4.120",
          }

          resultado = df_stock_scan[
              df_stock_scan[col_loc_scan].apply(_norm_scan_value) == scan_norm
          ].copy()

          if not resultado.empty:
              localizadores_encontrados = [str(loc_escaneado).strip()]
              hoja_mapeo = "STOCK"

          if resultado.empty and scan_norm and scan_norm in MAPEO_PRUEBA:
              localizadores_encontrados = [MAPEO_PRUEBA[scan_norm]]
              hoja_mapeo = "MAPEO_PRUEBA"
              loc_norms = {_norm_scan_value(x) for x in localizadores_encontrados}
              resultado = df_stock_scan[
                  df_stock_scan[col_loc_scan].apply(_norm_scan_value).isin(loc_norms)
              ].copy()

          if resultado.empty and scan_norm:
              for nombre_hoja, df_mapeo in hojas_disponibles.items():
                  if df_mapeo is None or not hasattr(df_mapeo, "columns"):
                      continue
                  try:
                      df_mapeo = df_mapeo.copy()
                  except Exception:
                      continue

                  for col in df_mapeo.columns:
                      try:
                          mask_codigo = df_mapeo[col].apply(_norm_scan_value) == scan_norm
                      except Exception:
                          continue
                      if not mask_codigo.any():
                          continue

                      filas_match = df_mapeo.loc[mask_codigo]
                      columnas_loc = [
                          c for c in df_mapeo.columns
                          if any(palabra in str(c).strip().lower()
                                 for palabra in ["localizador", "ubicacion", "ubicación", "loc"])
                      ]
                      candidatos = []
                      for c_loc in columnas_loc:
                          try:
                              candidatos.extend(filas_match[c_loc].dropna().astype(str).str.strip().tolist())
                          except Exception:
                              pass
                      if not candidatos:
                          for _, fila_match in filas_match.iterrows():
                              for valor in fila_match.tolist():
                                  if _parece_localizador(valor):
                                      candidatos.append(str(valor).strip())

                      candidatos = [x for x in candidatos if _parece_localizador(x)]
                      candidatos = list(dict.fromkeys(candidatos))
                      if candidatos:
                          localizadores_encontrados.extend(candidatos)
                          hoja_mapeo = nombre_hoja
                          break
                  if localizadores_encontrados:
                      break

              localizadores_encontrados = list(dict.fromkeys(localizadores_encontrados))
              if localizadores_encontrados:
                  loc_norms = {_norm_scan_value(x) for x in localizadores_encontrados}
                  resultado = df_stock_scan[
                      df_stock_scan[col_loc_scan].apply(_norm_scan_value).isin(loc_norms)
                  ].copy()

          if localizadores_encontrados and not resultado.empty:
              loc_mostrado = ", ".join(localizadores_encontrados)
              st.success(f"📍 Localizador: **{loc_mostrado}**")
          elif scan_norm:
              st.warning(
                  f"⚠️ Se detectó **{loc_escaneado}**, pero no encontré ese Localizador ni una relación código → Localizador en el Excel."
              )

          if resultado.empty:
            st.warning("No se encontró ningún producto registrado en esa posición.")
          else:
            if col_cant_scan:
              resultado[col_cant_scan] = resultado[col_cant_scan].apply(limpiar_numero)
            if col_cod_scan:
              resultado[col_cod_scan] = resultado[col_cod_scan].apply(fmt_code)
            if col_fecha_scan:
              resultado[col_fecha_scan] = pd.to_datetime(
                  resultado[col_fecha_scan], errors="coerce"
              ).dt.strftime("%d-%m-%Y")

            for _, fila in resultado.iterrows():
              desc_txt = fila[col_desc_scan] if col_desc_scan else "Sin descripción"
              cod_txt = fila[col_cod_scan] if col_cod_scan else "S/N"
              cant_txt = (
                  formato_unidades(fila[col_cant_scan]) if col_cant_scan else "N/A"
              )
              lote_txt = fila[col_lote_scan] if col_lote_scan else "N/A"
              fecha_txt = fila[col_fecha_scan] if col_fecha_scan else "N/A"

              st.markdown(
                  _dedent_html(f"""
                  <div style="background-color:#141414; border:1px solid #0070f3;
                              border-radius:10px; padding:16px; margin-bottom:12px;">
                      <div style="color:#aaaaaa; font-size:12px; text-transform:uppercase;">Producto</div>
                      <div style="color:#ffffff; font-size:20px; font-weight:bold;">{desc_txt}</div>
                      <div style="margin-top:8px; color:#cccccc; font-size:14px;">
                          Código: <b>{cod_txt}</b> · Lote: <b>{lote_txt}</b> · Vence: <b>{fecha_txt}</b>
                      </div>
                      <div style="margin-top:8px; color:#2ecc71; font-size:22px; font-weight:bold;">
                          Stock: {cant_txt} unidades
                      </div>
                  </div>
                  """),
                  unsafe_allow_html=True,
              )
    else:
      st.info("Aún no se ha escaneado ningún código.")


# ----------------------------------------------------------------------
# Pestañas
# ----------------------------------------------------------------------
tab_almacen, tab_stock, tab_escanear = st.tabs(
    ["📦 Almacenamiento", "🗓️ Stock y Caducidad", "📷 Escanear"]
)

with tab_almacen:
    if df_raw is None:
        st.error(
            f"No pude cargar la hoja `UBICACIONES` del Excel: {error_ubicaciones}"
        )
    else:
        render_almacenamiento(
            df_raw, pasillos_activos, bodegas_activas, localizador_q,
            almacenamiento_sel, recetario_sel,
        )

with tab_stock:
    if df_stock_raw is None:
        st.error(
            f"No pude cargar la hoja `STOCK` del Excel: {error_stock}"
        )
    else:
        render_stock(df_stock_raw)

with tab_escanear:
    if df_stock_raw is None:
        st.error(
            f"No pude cargar la hoja `STOCK` del Excel: {error_stock}"
        )
    else:
        hojas_disponibles = {}
        if df_stock_raw is not None:
            hojas_disponibles["STOCK"] = df_stock_raw
        if df_raw is not None:
            hojas_disponibles["UBICACIONES"] = df_raw
        render_escanear(df_stock_raw, hojas_disponibles)
