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
from datetime import date

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# ----------------------------------------------------------------------
# Configuración de página y paleta de colores
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Informe de Almacenamiento",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Paleta base (misma esencia que el PBI original) + semáforo para KPIs
COLOR_OCUPADA = "#C0755F"       # terracota / rojo ladrillo
COLOR_DISPONIBLE = "#7FA87A"    # verde salvia
COLOR_ROJO = "#D9534F"          # crítico (>90%)
COLOR_AMARILLO = "#E8B84B"      # atención (70-90%)
COLOR_VERDE = "#5CB868"         # saludable (<70%)
COLOR_CARD_BG = "#1A1D24"
COLOR_CARD_BORDER = "#2E323C"
COLOR_TEXT_MUTED = "#9AA0A8"

DATA_PATH = "data/Almacenamiento_2026.xlsx"

CSS = f"""
<style>
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
@st.cache_data(show_spinner="Cargando datos...")
def load_data(file) -> pd.DataFrame:
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


with st.sidebar:
    st.header("Fuente de datos")
    uploaded = st.file_uploader(
        "Reemplazar con un Excel más nuevo (opcional)", type=["xlsx"]
    )
    st.caption(
        "Si no subes nada, se usa el archivo incluido en el repositorio "
        f"(`{DATA_PATH}`)."
    )

try:
    df_raw = load_data(uploaded if uploaded is not None else DATA_PATH)
except FileNotFoundError:
    st.error(
        f"No encontré `{DATA_PATH}`. Sube un archivo Excel con la hoja "
        "`UBICACIONES` desde la barra lateral."
    )
    st.stop()

# ----------------------------------------------------------------------
# Filtros (barra lateral)
# ----------------------------------------------------------------------
st.sidebar.header("Filtros")

pasillos = sorted(df_raw["PASILLO"].unique())
sel_pasillos = st.sidebar.multiselect("Pasillo", pasillos, default=pasillos)

localizador_q = st.sidebar.text_input("Buscar localizador")

bodegas = sorted(df_raw["Bodega"].unique())
sel_bodegas = st.sidebar.multiselect("Tipo bodega", bodegas, default=bodegas)

st.sidebar.caption(
    "💡 El filtro **Código artículo** del reporte original vive en otra "
    "tabla (maestro de artículos) que no venía en este Excel. Si me "
    "compartes esa tabla, la agrego."
)

almacenamiento_sel = st.sidebar.radio(
    "Almacenamiento", ["Sí", "No", "Todos"], index=0, horizontal=True
)
recetario_sel = st.sidebar.radio(
    "Es Recetario", ["No", "Sí", "Todos"], index=0, horizontal=True
)

# Aplicar filtros
df = df_raw[
    df_raw["PASILLO"].isin(sel_pasillos) & df_raw["Bodega"].isin(sel_bodegas)
]
if localizador_q:
    df = df[df["LOCALIZADOR"].str.contains(localizador_q, case=False, na=False)]
if almacenamiento_sel != "Todos":
    df = df[df["ALMACENAMIENTO_FLAG"] == almacenamiento_sel]
if recetario_sel != "Todos":
    df = df[df["ES_RECETARIO"] == recetario_sel]

if df.empty:
    st.warning("No hay datos para los filtros seleccionados.")
    st.stop()

# ----------------------------------------------------------------------
# Cálculos base / KPIs
# ----------------------------------------------------------------------
total_localizadores = df["LOCALIZADOR"].nunique()
ocupadas = int(df["OCUPADA"].sum())
disponibles = int((~df["OCUPADA"]).sum())
pct_ocupacion = round(ocupadas / total_localizadores * 100, 1)

# Pasillo más saturado / con más espacio
por_pasillo_pct = (
    df.groupby("PASILLO")["OCUPADA"].mean().mul(100).round(1).sort_values()
)
pasillo_mas_libre = por_pasillo_pct.index[0]
pasillo_mas_saturado = por_pasillo_pct.index[-1]

# Ubicaciones "otras" (observaciones distintas de vacío/recetario)
otras_obs = df["OBSERVACIONES"].apply(
    lambda x: pd.notna(x) and x != "RECETARIO"
).sum()

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


# Fila de KPIs principales (6 tarjetas)
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

st.write("")

# Botón de descarga del reporte filtrado
buffer = io.BytesIO()
df.to_excel(buffer, index=False, sheet_name="UBICACIONES_FILTRADO")
st.download_button(
    label="⬇️ Descargar reporte filtrado (Excel)",
    data=buffer.getvalue(),
    file_name=f"almacenamiento_filtrado_{date.today().strftime('%Y%m%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.write("")

# ----------------------------------------------------------------------
# Tabs para organizar el contenido
# ----------------------------------------------------------------------
tab_resumen, tab_bodega, tab_pasillo_nivel, tab_composicion = st.tabs(
    ["📊 Resumen general", "🏬 Por bodega", "🧭 Pasillo y nivel", "🧩 Composición"]
)

# ---------------- TAB 1: Resumen general ----------------
with tab_resumen:
    st.markdown('<p class="section-title">Ubicaciones ocupadas vs. disponibles</p>', unsafe_allow_html=True)
    fig_total = go.Figure()
    fig_total.add_trace(
        go.Bar(
            x=[ocupadas], y=["Ubicaciones"], orientation="h",
            name="Ocupadas", marker_color=COLOR_OCUPADA,
            text=[f"{ocupadas:,}".replace(",", ".")], textposition="inside",
        )
    )
    fig_total.add_trace(
        go.Bar(
            x=[disponibles], y=["Ubicaciones"], orientation="h",
            name="Disponibles", marker_color=COLOR_DISPONIBLE,
            text=[f"{disponibles:,}".replace(",", ".")], textposition="inside",
        )
    )
    fig_total.update_layout(
        barmode="stack", height=110,
        margin=dict(l=10, r=10, t=10, b=10),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.05, x=0),
        xaxis=dict(visible=False), yaxis=dict(visible=False),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E6E6E6"),
    )
    st.plotly_chart(fig_total, use_container_width=True)

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
               marker_color=COLOR_OCUPADA, text=gp["ocupadas"], textposition="inside")
    )
    fig_pas.add_trace(
        go.Bar(x=gp.index, y=gp["disponibles"], name="Disponibles",
               marker_color=COLOR_DISPONIBLE,
               text=gp["disponibles"].replace(0, ""), textposition="inside")
    )
    for pasillo, row in gp.iterrows():
        fig_pas.add_annotation(
            x=pasillo, y=row["total"], text=f"{int(row['total'])}",
            showarrow=False, yshift=12, font=dict(size=11, color="#E6E6E6"),
        )
    fig_pas.update_layout(
        barmode="stack", height=380,
        margin=dict(l=10, r=10, t=30, b=10),
        xaxis_title="Pasillo", yaxis_title="Ubicaciones",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E6E6E6"),
    )
    st.plotly_chart(fig_pas, use_container_width=True)

# ---------------- TAB 2: Por bodega ----------------
with tab_bodega:
    c1, c2 = st.columns(2)

    with c1:
        st.markdown('<p class="section-title">% Ocupación por tipo de bodega</p>', unsafe_allow_html=True)
        g = (
            df.groupby("Bodega")
            .agg(total=("LOCALIZADOR", "count"), ocupadas=("OCUPADA", "sum"))
            .assign(pct=lambda d: (d["ocupadas"] / d["total"] * 100).round(0))
            .sort_values("pct")
        )
        colores_bodega = [semaforo_color(v) for v in g["pct"]]
        fig_pct = go.Figure(
            go.Bar(
                x=g["pct"], y=g.index, orientation="h",
                marker_color=colores_bodega,
                text=[f"{v:.0f}%" for v in g["pct"]], textposition="outside",
            )
        )
        fig_pct.update_layout(
            height=280, margin=dict(l=10, r=40, t=10, b=10),
            xaxis=dict(visible=False),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#E6E6E6"),
        )
        st.plotly_chart(fig_pct, use_container_width=True)

    with c2:
        st.markdown('<p class="section-title">Ocupadas vs. disponibles por bodega</p>', unsafe_allow_html=True)
        g2 = (
            df.groupby("Bodega")
            .agg(ocupadas=("OCUPADA", "sum"), disponibles=("OCUPADA", lambda s: (~s).sum()))
            .sort_values("ocupadas")
        )
        fig_bd = go.Figure()
        fig_bd.add_trace(
            go.Bar(x=g2["ocupadas"], y=g2.index, orientation="h",
                   name="Ocupadas", marker_color=COLOR_OCUPADA,
                   text=g2["ocupadas"], textposition="outside")
        )
        fig_bd.add_trace(
            go.Bar(x=g2["disponibles"], y=g2.index, orientation="h",
                   name="Disponibles", marker_color=COLOR_DISPONIBLE,
                   text=g2["disponibles"], textposition="outside")
        )
        fig_bd.update_layout(
            barmode="group", height=280,
            margin=dict(l=10, r=10, t=10, b=10),
            xaxis=dict(visible=False),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#E6E6E6"),
        )
        st.plotly_chart(fig_bd, use_container_width=True)

    st.markdown('<p class="section-title">Distribución de ubicaciones por bodega (treemap)</p>', unsafe_allow_html=True)
    g3 = df.groupby("Bodega").size().reset_index(name="cantidad")
    fig_tree = px.treemap(
        g3, path=["Bodega"], values="cantidad",
        color="cantidad", color_continuous_scale=["#3A3F4B", COLOR_OCUPADA],
    )
    fig_tree.update_layout(
        height=320, margin=dict(l=10, r=10, t=10, b=10),
        paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#E6E6E6"),
        coloraxis_showscale=False,
    )
    st.plotly_chart(fig_tree, use_container_width=True)

# ---------------- TAB 3: Pasillo y nivel ----------------
with tab_pasillo_nivel:
    st.markdown('<p class="section-title">% Ocupación por Nivel y Pasillo</p>', unsafe_allow_html=True)
    piv = df.pivot_table(
        index="NIVEL", columns="PASILLO", values="OCUPADA", aggfunc="mean"
    ) * 100
    piv = piv.reindex(sorted(piv.index, key=lambda x: int(x)))
    piv = piv[sorted(piv.columns)]

    nivel_totales = df.groupby("NIVEL")["OCUPADA"].mean() * 100
    piv["Total"] = nivel_totales.reindex(piv.index)

    pasillo_totales = df.groupby("PASILLO")["OCUPADA"].mean() * 100
    fila_total = pasillo_totales.reindex(piv.columns[:-1])
    fila_total["Total"] = df["OCUPADA"].mean() * 100
    piv.loc["Total"] = fila_total

    text_vals = piv.round(1).astype(str) + "%"

    fig_heat = go.Figure(
        data=go.Heatmap(
            z=piv.values, x=piv.columns, y=piv.index,
            colorscale=[[0, COLOR_VERDE], [0.5, COLOR_AMARILLO], [1, COLOR_ROJO]],
            text=text_vals.values, texttemplate="%{text}",
            showscale=True, xgap=2, ygap=2,
            colorbar=dict(title="%", tickfont=dict(color="#E6E6E6")),
        )
    )
    fig_heat.update_layout(
        height=380, margin=dict(l=10, r=10, t=10, b=10),
        yaxis=dict(autorange="reversed", title="Nivel"),
        xaxis=dict(title="Pasillo", side="top"),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#E6E6E6"),
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    st.caption(
        "🟢 Saludable (<70%) · 🟡 Atención (70-90%) · 🔴 Crítico (>90%). "
        "La columna/fila **Total** se calcula como Ubicaciones ocupadas ÷ "
        "Ubicaciones totales, según los filtros activos."
    )

# ---------------- TAB 4: Composición ----------------
with tab_composicion:
    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown('<p class="section-title">Almacenamiento</p>', unsafe_allow_html=True)
        vc = df["ALMACENAMIENTO_FLAG"].value_counts()
        fig_d1 = go.Figure(
            go.Pie(labels=vc.index, values=vc.values, hole=0.55,
                   marker_colors=[COLOR_OCUPADA, COLOR_DISPONIBLE])
        )
        fig_d1.update_layout(
            height=260, margin=dict(l=10, r=10, t=10, b=10),
            paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#E6E6E6"),
            showlegend=True, legend=dict(orientation="h", y=-0.1),
        )
        st.plotly_chart(fig_d1, use_container_width=True)

    with c2:
        st.markdown('<p class="section-title">Es Recetario</p>', unsafe_allow_html=True)
        vc2 = df["ES_RECETARIO"].value_counts()
        fig_d2 = go.Figure(
            go.Pie(labels=vc2.index, values=vc2.values, hole=0.55,
                   marker_colors=[COLOR_DISPONIBLE, COLOR_OCUPADA])
        )
        fig_d2.update_layout(
            height=260, margin=dict(l=10, r=10, t=10, b=10),
            paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#E6E6E6"),
            showlegend=True, legend=dict(orientation="h", y=-0.1),
        )
        st.plotly_chart(fig_d2, use_container_width=True)

    with c3:
        st.markdown('<p class="section-title">Otras observaciones</p>', unsafe_allow_html=True)
        kpi_card(
            st, int(otras_obs), "Ubicaciones con observación distinta de recetario/vacío"
        )

    st.write("")
    st.markdown('<p class="section-title">Detalle de datos filtrados</p>', unsafe_allow_html=True)
    st.dataframe(df, use_container_width=True, height=320)
