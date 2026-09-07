"""
Dashboard de Almacenamiento — réplica en Streamlit de un reporte de Power BI.

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
# Configuración de página y paleta de colores (misma paleta que el PBI)
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Informe de Almacenamiento",
    page_icon="📦",
    layout="wide",
)

COLOR_OCUPADA = "#C0755F"      # terracota / rojo ladrillo
COLOR_DISPONIBLE = "#7FA87A"   # verde salvia
COLOR_BG_CARD = "#FFFFFF"

DATA_PATH = "data/Almacenamiento_2026.xlsx"

CSS = f"""
<style>
    .metric-card {{
        background-color:{COLOR_BG_CARD};
        border:1px solid #E4E4E4;
        border-radius:8px;
        padding:14px 18px;
        text-align:center;
    }}
    .metric-card h1 {{
        font-size:28px;
        margin-bottom:0px;
        color:#222;
    }}
    .metric-card p {{
        margin-top:2px;
        color:#666;
        font-size:13px;
    }}
    div[data-testid="stMetricValue"] {{ font-size: 26px; }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


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
# Filtros (barra lateral) — replican los slicers del reporte de Power BI
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
# KPIs
# ----------------------------------------------------------------------
total_localizadores = df["LOCALIZADOR"].nunique()
ocupadas = int(df["OCUPADA"].sum())
disponibles = int((~df["OCUPADA"]).sum())
pct_ocupacion = round(ocupadas / total_localizadores * 100, 1)

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(
        f'<div class="metric-card"><h1>{total_localizadores:,}</h1>'
        f"<p>Recuento de Localizador</p></div>".replace(",", "."),
        unsafe_allow_html=True,
    )
with k2:
    st.markdown(
        f'<div class="metric-card"><h1>{date.today().strftime("%d-%m-%Y")}</h1>'
        f"<p>Última fecha: fecha proceso</p></div>",
        unsafe_allow_html=True,
    )
with k3:
    st.markdown(
        f'<div class="metric-card"><h1>{ocupadas:,}</h1>'
        f"<p>Ubicaciones ocupadas</p></div>".replace(",", "."),
        unsafe_allow_html=True,
    )
with k4:
    st.markdown(
        f'<div class="metric-card"><h1>{pct_ocupacion}%</h1>'
        f"<p>% Ocupación general</p></div>",
        unsafe_allow_html=True,
    )

st.write("")

# ----------------------------------------------------------------------
# Barra total: ocupadas vs disponibles
# ----------------------------------------------------------------------
st.subheader("Ubicaciones")
fig_total = go.Figure()
fig_total.add_trace(
    go.Bar(
        x=[ocupadas],
        y=["Ubicaciones"],
        orientation="h",
        name="Ubicaciones Ocupadas",
        marker_color=COLOR_OCUPADA,
        text=[f"{ocupadas:,}".replace(",", ".")],
        textposition="inside",
    )
)
fig_total.add_trace(
    go.Bar(
        x=[disponibles],
        y=["Ubicaciones"],
        orientation="h",
        name="Ubicaciones Disponibles",
        marker_color=COLOR_DISPONIBLE,
        text=[f"{disponibles:,}".replace(",", ".")],
        textposition="inside",
    )
)
fig_total.update_layout(
    barmode="stack",
    height=110,
    margin=dict(l=10, r=10, t=10, b=10),
    showlegend=True,
    legend=dict(orientation="h", yanchor="bottom", y=1.05, x=0),
    xaxis=dict(visible=False),
    yaxis=dict(visible=False),
)
st.plotly_chart(fig_total, use_container_width=True)

# ----------------------------------------------------------------------
# % Ocupación por Bodega (horizontal, ordenado desc)
# ----------------------------------------------------------------------
c1, c2 = st.columns([1, 1])

with c1:
    st.markdown("**Ubicaciones ocupadas (%) por tipo de bodega**")
    g = (
        df.groupby("Bodega")
        .agg(total=("LOCALIZADOR", "count"), ocupadas=("OCUPADA", "sum"))
        .assign(pct=lambda d: (d["ocupadas"] / d["total"] * 100).round(0))
        .sort_values("pct")
    )
    fig_pct = go.Figure(
        go.Bar(
            x=g["pct"],
            y=g.index,
            orientation="h",
            marker_color=COLOR_OCUPADA,
            text=[f"{v:.0f}%" for v in g["pct"]],
            textposition="outside",
        )
    )
    fig_pct.update_layout(
        height=260, margin=dict(l=10, r=40, t=10, b=10), xaxis=dict(visible=False)
    )
    st.plotly_chart(fig_pct, use_container_width=True)

with c2:
    st.markdown("**Ocupadas vs. disponibles por tipo de bodega**")
    g2 = (
        df.groupby("Bodega")
        .agg(ocupadas=("OCUPADA", "sum"), disponibles=("OCUPADA", lambda s: (~s).sum()))
        .sort_values("ocupadas")
    )
    fig_bd = go.Figure()
    fig_bd.add_trace(
        go.Bar(
            x=g2["ocupadas"],
            y=g2.index,
            orientation="h",
            name="Ocupadas",
            marker_color=COLOR_OCUPADA,
            text=g2["ocupadas"],
            textposition="outside",
        )
    )
    fig_bd.add_trace(
        go.Bar(
            x=g2["disponibles"],
            y=g2.index,
            orientation="h",
            name="Disponibles",
            marker_color=COLOR_DISPONIBLE,
            text=g2["disponibles"],
            textposition="outside",
        )
    )
    fig_bd.update_layout(
        barmode="group",
        height=260,
        margin=dict(l=10, r=10, t=10, b=10),
        xaxis=dict(visible=False),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
    )
    st.plotly_chart(fig_bd, use_container_width=True)

# ----------------------------------------------------------------------
# Ocupadas y disponibles por pasillo (stacked bar)
# ----------------------------------------------------------------------
st.markdown("**Ubicaciones ocupadas y disponibles por pasillo**")
gp = (
    df.groupby("PASILLO")
    .agg(ocupadas=("OCUPADA", "sum"), disponibles=("OCUPADA", lambda s: (~s).sum()))
    .reindex(sorted(df["PASILLO"].unique()))
)
gp["total"] = gp["ocupadas"] + gp["disponibles"]

fig_pas = go.Figure()
fig_pas.add_trace(
    go.Bar(
        x=gp.index,
        y=gp["ocupadas"],
        name="Ubicaciones Ocupadas",
        marker_color=COLOR_OCUPADA,
        text=gp["ocupadas"],
        textposition="inside",
    )
)
fig_pas.add_trace(
    go.Bar(
        x=gp.index,
        y=gp["disponibles"],
        name="Ubicaciones Disponibles",
        marker_color=COLOR_DISPONIBLE,
        text=gp["disponibles"].replace(0, ""),
        textposition="inside",
    )
)
# etiqueta de total sobre cada barra
for pasillo, row in gp.iterrows():
    fig_pas.add_annotation(
        x=pasillo,
        y=row["total"],
        text=f"{int(row['total'])}",
        showarrow=False,
        yshift=12,
        font=dict(size=11),
    )
fig_pas.update_layout(
    barmode="stack",
    height=380,
    margin=dict(l=10, r=10, t=30, b=10),
    xaxis_title="Pasillo",
    yaxis_title="Ubicaciones Ocupadas / Disp.",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
)
st.plotly_chart(fig_pas, use_container_width=True)

# ----------------------------------------------------------------------
# Tabla % ocupación por Nivel x Pasillo (heatmap)
# ----------------------------------------------------------------------
st.markdown("**% Ocupación por Nivel y Pasillo**")
piv = df.pivot_table(
    index="NIVEL", columns="PASILLO", values="OCUPADA", aggfunc="mean"
) * 100
piv = piv.reindex(sorted(piv.index, key=lambda x: int(x)))
piv = piv[sorted(piv.columns)]

# columna total por nivel (antes de agregar la fila Total)
nivel_totales = df.groupby("NIVEL")["OCUPADA"].mean() * 100
piv["Total"] = nivel_totales.reindex(piv.index)

# fila total por pasillo (incluye la columna Total general)
pasillo_totales = df.groupby("PASILLO")["OCUPADA"].mean() * 100
fila_total = pasillo_totales.reindex(piv.columns[:-1])
fila_total["Total"] = df["OCUPADA"].mean() * 100
piv.loc["Total"] = fila_total

text_vals = piv.round(1).astype(str) + "%"

fig_heat = go.Figure(
    data=go.Heatmap(
        z=piv.values,
        x=piv.columns,
        y=piv.index,
        colorscale=[[0, COLOR_OCUPADA], [1, COLOR_DISPONIBLE]],
        text=text_vals.values,
        texttemplate="%{text}",
        showscale=False,
        xgap=2,
        ygap=2,
    )
)
fig_heat.update_layout(
    height=380,
    margin=dict(l=10, r=10, t=10, b=10),
    yaxis=dict(autorange="reversed", title="Nivel"),
    xaxis=dict(title="Pasillo", side="top"),
)
st.plotly_chart(fig_heat, use_container_width=True)

st.caption(
    "Nota: la columna **% Ocupación** se calcula como Ubicaciones "
    "ocupadas ÷ Ubicaciones totales, según los filtros activos en la "
    "barra lateral."
)
