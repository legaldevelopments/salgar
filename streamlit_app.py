"""
Tablero Predial Salgar 2026 - Ley 44/1990
Análisis de formación catastral y liquidación del impuesto predial
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import numpy as np

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Predial Salgar 2026",
    page_icon="🏘️",
    layout="wide",
    initial_sidebar_state="expanded",
)

import os
RUTA_EXCEL = os.path.join(os.path.dirname(__file__), "Predial_Salgar_2026_Ley44.xlsx")

AZUL_OSC  = "#1F4E79"
AZUL_MED  = "#2E75B6"
AZUL_CLAR = "#BDD7EE"
VERDE     = "#70AD47"
AMBAR     = "#FFD966"
ROJO      = "#FF7C80"
GRIS      = "#F2F2F2"

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    /* Fondo general */
    .stApp { background-color: #F0F4F8; }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
    }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stMultiSelect label { color: #BDD7EE !important; font-weight: 600; }

    /* Tarjetas KPI */
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 20px 16px;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.10);
        border-top: 4px solid #2E75B6;
        height: 130px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    .kpi-card.verde  { border-top-color: #70AD47; }
    .kpi-card.ambar  { border-top-color: #FFD966; }
    .kpi-card.rojo   { border-top-color: #FF7C80; }
    .kpi-card.oscuro { border-top-color: #1F4E79; }

    .kpi-valor {
        font-size: 1.7rem;
        font-weight: 800;
        color: #1F4E79;
        line-height: 1.1;
    }
    .kpi-label {
        font-size: 0.75rem;
        color: #666;
        margin-top: 6px;
        font-weight: 500;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-sub {
        font-size: 0.8rem;
        color: #2E75B6;
        font-weight: 600;
        margin-top: 4px;
    }

    /* Títulos de sección */
    .seccion-titulo {
        background: linear-gradient(90deg, #1F4E79, #2E75B6);
        color: white;
        padding: 8px 16px;
        border-radius: 8px;
        font-size: 0.95rem;
        font-weight: 700;
        letter-spacing: 0.5px;
        margin: 18px 0 12px 0;
    }

    /* Tabla */
    .dataframe thead tr th {
        background-color: #1F4E79 !important;
        color: white !important;
        font-weight: 700;
    }

    /* Header principal */
    .header-main {
        background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 60%, #41A0D6 100%);
        border-radius: 14px;
        padding: 28px 32px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 4px 15px rgba(31,78,121,0.3);
    }
    .header-main h1 { margin: 0; font-size: 1.8rem; font-weight: 800; }
    .header-main p  { margin: 4px 0 0; opacity: 0.85; font-size: 0.92rem; }

    /* Nota legal */
    .nota-legal {
        background: #FFF9E6;
        border-left: 4px solid #FFD966;
        border-radius: 6px;
        padding: 10px 14px;
        font-size: 0.82rem;
        color: #555;
        margin-top: 8px;
    }
</style>
""", unsafe_allow_html=True)


# ─── CARGA DE DATOS ───────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando datos del análisis predial...")
def cargar_datos():
    df_urb = pd.read_excel(RUTA_EXCEL, sheet_name="Predios_Urbanos_2026")
    df_tabla = pd.read_excel(RUTA_EXCEL, sheet_name="Tabla_Salgar")

    # Renombrar columnas "MAURO" → "2025"
    rename_map = {
        "TERCERO_MAURO":   "TERCERO_2025",
        "AVALÚO_2025_MAURO": "AVALÚO_2025",
        "TARIFA_ACT":      "TARIFA_2025",
        "DEST_MAURO":      "DEST_2025",
        "EN_MAURO":        "EN_ANÁLISIS_2025",
        "LIQ_2025":        "LIQ_2025",
    }
    df_tabla.rename(columns={k: v for k, v in rename_map.items() if k in df_tabla.columns}, inplace=True)

    # Columna auxiliar
    df_urb["_TIENE_IGAC"] = df_urb["EN_SALGAR_IGAC"].astype(str).str.upper().str.strip() == "SÍ"
    df_urb["_APLICA_LIM"] = df_urb["APLICA_LÍMITE"].astype(str).str.upper().str.strip() == "SÍ"
    df_urb["_EXENTO"]     = df_urb["ES_EXENTO"].astype(str).str.upper().str.strip() == "S"

    return df_urb, df_tabla


df_urb, df_tabla = cargar_datos()


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏘️ Predial Salgar 2026")
    st.markdown("**Análisis Catastral · Ley 44/1990**")
    st.divider()

    st.markdown("#### Filtros")

    # Filtro: estado IGAC
    igac_opts = ["Todos", "Con datos IGAC 2026", "Sin datos IGAC 2026"]
    sel_igac = st.selectbox("Datos IGAC:", igac_opts)

    # Filtro: aplica límite
    lim_opts = ["Todos", "Aplica límite Ley 44", "No aplica límite"]
    sel_limite = st.selectbox("Límite Ley 44:", lim_opts)

    # Filtro: destino
    destinos = sorted(df_urb["DEST_2025"].dropna().unique().tolist())
    sel_dest = st.multiselect("Destino económico:", destinos, placeholder="Todos")

    # Filtro: exento
    sel_exento = st.checkbox("Excluir exentos", value=False)

    st.divider()
    st.markdown("""
    <div style='font-size:0.75rem; opacity:0.8;'>
    📂 Fuentes:<br>
    · <i>archivo salgar convertido.xlsx</i><br>
    · <i>Base de datos predial 2024-2025</i><br><br>
    📋 Normativa:<br>
    · Ley 44/1990 Art. 6<br>
    · Ley 1955/2019<br>
    · Ley 1819/2016 Art. 354
    </div>
    """, unsafe_allow_html=True)


# ─── APLICAR FILTROS ─────────────────────────────────────────────────────────
df = df_urb.copy()
if sel_igac == "Con datos IGAC 2026":
    df = df[df["_TIENE_IGAC"]]
elif sel_igac == "Sin datos IGAC 2026":
    df = df[~df["_TIENE_IGAC"]]
if sel_limite == "Aplica límite Ley 44":
    df = df[df["_APLICA_LIM"]]
elif sel_limite == "No aplica límite":
    df = df[~df["_APLICA_LIM"]]
if sel_dest:
    df = df[df["DEST_2025"].isin(sel_dest)]
if sel_exento:
    df = df[~df["_EXENTO"]]


# ─── HEADER ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-main">
  <h1>🏘️ Análisis Predial — Municipio de Salgar</h1>
  <p>Formación catastral 2026 · Aplicación Ley 44 de 1990 · Liquidación impuesto predial urbano</p>
</div>
""", unsafe_allow_html=True)

n_filtrado = len(df)
n_total    = len(df_urb)
if n_filtrado < n_total:
    st.info(f"Mostrando **{n_filtrado:,}** de **{n_total:,}** predios urbanos según los filtros aplicados.")


# ─── KPIs PRINCIPALES ────────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">📊 Indicadores Clave</div>', unsafe_allow_html=True)

liq25  = df["LIQ_2025"].fillna(0).sum()
liq26s = df["LIQ_2026_SIN_LÍMITE"].fillna(0).sum()
liq26c = df["LIQ_2026_CON_LÍMITE"].fillna(0).sum()
ahorro = df["AHORRO_CONTRIBUYENTE"].fillna(0).sum()
n_lim  = df["_APLICA_LIM"].sum()
n_igac = df["_TIENE_IGAC"].sum()
var_rec = (liq26c - liq25) / liq25 * 100 if liq25 > 0 else 0

col1, col2, col3, col4, col5, col6 = st.columns(6)

def fmt_cop(v):
    if v >= 1e9:
        return f"${v/1e9:.2f}B"
    elif v >= 1e6:
        return f"${v/1e6:.1f}M"
    else:
        return f"${v:,.0f}"

with col1:
    st.markdown(f"""
    <div class="kpi-card oscuro">
      <div class="kpi-valor">{n_filtrado:,}</div>
      <div class="kpi-label">Predios Urbanos</div>
      <div class="kpi-sub">{n_igac:,} con datos IGAC</div>
    </div>""", unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
      <div class="kpi-valor">{fmt_cop(liq25)}</div>
      <div class="kpi-label">Recaudo 2025</div>
      <div class="kpi-sub">Análisis anterior</div>
    </div>""", unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="kpi-card verde">
      <div class="kpi-valor">{fmt_cop(liq26c)}</div>
      <div class="kpi-label">Recaudo 2026 Proyectado</div>
      <div class="kpi-sub">Con límite Ley 44</div>
    </div>""", unsafe_allow_html=True)

with col4:
    color_var = "verde" if var_rec >= 0 else "rojo"
    st.markdown(f"""
    <div class="kpi-card {color_var}">
      <div class="kpi-valor">{var_rec:+.1f}%</div>
      <div class="kpi-label">Variación Recaudo</div>
      <div class="kpi-sub">2025 → 2026</div>
    </div>""", unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="kpi-card ambar">
      <div class="kpi-valor">{n_lim:,}</div>
      <div class="kpi-label">Predios Limitados</div>
      <div class="kpi-sub">Ley 44 aplica tope 2×</div>
    </div>""", unsafe_allow_html=True)

with col6:
    st.markdown(f"""
    <div class="kpi-card rojo">
      <div class="kpi-valor">{fmt_cop(ahorro)}</div>
      <div class="kpi-label">Ahorro Contribuyentes</div>
      <div class="kpi-sub">Por tope Ley 44</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─── FILA 1 DE GRÁFICAS ──────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">📈 Comparativo Fiscal 2025 vs 2026</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns([2, 2, 1.5])

# Barras: recaudo comparativo
with c1:
    fig = go.Figure()
    categorias = ["Recaudo 2025", "2026 Sin límite Ley 44", "2026 Con límite Ley 44"]
    valores    = [liq25, liq26s, liq26c]
    colores    = [AZUL_MED, ROJO, VERDE]
    fig.add_trace(go.Bar(
        x=categorias, y=valores, marker_color=colores,
        text=[fmt_cop(v) for v in valores],
        textposition="outside", textfont=dict(size=11, color=AZUL_OSC),
    ))
    fig.update_layout(
        title=dict(text="Comparativo Recaudo ($)", font=dict(size=13, color=AZUL_OSC)),
        yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee"),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=20, l=10, r=10), height=320,
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)

# Dona: predios con/sin IGAC
with c2:
    igac_no  = n_filtrado - n_igac
    labels   = ["Con datos IGAC 2026", "Sin datos IGAC", "Predios exentos"]
    values   = [n_igac - df["_EXENTO"].sum(), igac_no, df["_EXENTO"].sum()]
    colors   = [VERDE, AMBAR, AZUL_CLAR]
    fig2 = go.Figure(go.Pie(
        labels=labels, values=values,
        hole=0.55, marker_colors=colors,
        textinfo="label+percent", textfont=dict(size=11),
        hovertemplate="%{label}: %{value:,}<extra></extra>",
    ))
    fig2.update_layout(
        title=dict(text="Estado Catastral 2026", font=dict(size=13, color=AZUL_OSC)),
        showlegend=False, margin=dict(t=50, b=0, l=0, r=0), height=320,
        paper_bgcolor="white",
        annotations=[dict(text=f"<b>{n_filtrado:,}</b><br>predios", x=0.5, y=0.5,
                          font_size=13, showarrow=False, font_color=AZUL_OSC)],
    )
    st.plotly_chart(fig2, use_container_width=True)

# Semáforo Ley 44
with c3:
    n_no_lim = (~df["_APLICA_LIM"] & df["_TIENE_IGAC"]).sum()
    n_sin    = (~df["_TIENE_IGAC"]).sum()
    n_exen   = df["_EXENTO"].sum()

    fig3 = go.Figure(go.Pie(
        labels=["Aplica tope 2× Ley 44", "No supera tope", "Sin datos IGAC", "Exentos"],
        values=[n_lim, n_no_lim, n_sin, n_exen],
        hole=0.5,
        marker_colors=[ROJO, VERDE, AMBAR, AZUL_CLAR],
        textinfo="percent", textfont=dict(size=10),
        hovertemplate="%{label}: %{value:,}<extra></extra>",
    ))
    fig3.update_layout(
        title=dict(text="Aplicación Ley 44", font=dict(size=13, color=AZUL_OSC)),
        legend=dict(orientation="v", font=dict(size=9)),
        margin=dict(t=50, b=0, l=0, r=10), height=320,
        paper_bgcolor="white",
    )
    st.plotly_chart(fig3, use_container_width=True)


# ─── FILA 2 DE GRÁFICAS ──────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">📉 Distribución de Avalúos y Variaciones</div>', unsafe_allow_html=True)
c4, c5 = st.columns(2)

# Histograma variación avalúo
with c4:
    df_var = df[df["VARIACIÓN_AVALÚO_%"].notna() & (df["VARIACIÓN_AVALÚO_%"].abs() < 2000)]
    fig4 = px.histogram(
        df_var, x="VARIACIÓN_AVALÚO_%", nbins=50,
        color_discrete_sequence=[AZUL_MED],
        labels={"VARIACIÓN_AVALÚO_%": "Variación Avalúo (%)"},
        title="Distribución Variación de Avalúos 2025→2026",
    )
    fig4.add_vline(x=0,   line_dash="dash", line_color=AZUL_OSC, annotation_text="0%")
    fig4.add_vline(x=100, line_dash="dot",  line_color=ROJO,     annotation_text="100%")
    fig4.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=20, l=10, r=10), height=310,
        yaxis_title="N° predios", xaxis_title="Variación %",
        bargap=0.05,
    )
    st.plotly_chart(fig4, use_container_width=True)

# Top destinos por recaudo
with c5:
    if "DEST_2025" in df.columns:
        df_dest = (df.groupby("DEST_2025")
                   .agg(predios=("FICHA", "count"),
                        liq26=("LIQ_2026_CON_LÍMITE", "sum"))
                   .reset_index()
                   .sort_values("liq26", ascending=True)
                   .tail(10))
        fig5 = go.Figure(go.Bar(
            x=df_dest["liq26"], y=df_dest["DEST_2025"],
            orientation="h",
            marker_color=AZUL_MED,
            text=[fmt_cop(v) for v in df_dest["liq26"]],
            textposition="outside",
        ))
        fig5.update_layout(
            title=dict(text="Recaudo 2026 por Destino Económico", font=dict(size=13, color=AZUL_OSC)),
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee"),
            margin=dict(t=50, b=20, l=10, r=80), height=310,
            yaxis_title="", xaxis_title="Liquidación 2026 ($)",
        )
        st.plotly_chart(fig5, use_container_width=True)


# ─── FILA 3: SCATTER AVALUO vs LIQ ───────────────────────────────────────────
st.markdown('<div class="seccion-titulo">🔍 Relación Avalúo 2026 vs Liquidación</div>', unsafe_allow_html=True)

df_scatter = df[df["AVALÚO_TOTAL_2026"].notna() & df["LIQ_2026_CON_LÍMITE"].notna()].copy()
df_scatter["color_cat"] = df_scatter["_APLICA_LIM"].map(
    {True: "Aplica límite Ley 44", False: "No aplica límite"}
)

fig6 = px.scatter(
    df_scatter.sample(min(len(df_scatter), 1500)),  # muestra para rendimiento
    x="AVALÚO_TOTAL_2026", y="LIQ_2026_CON_LÍMITE",
    color="color_cat",
    color_discrete_map={"Aplica límite Ley 44": ROJO, "No aplica límite": VERDE},
    hover_data={"FICHA": True, "TERCERO": True, "VARIACIÓN_AVALÚO_%": True,
                "color_cat": False},
    labels={
        "AVALÚO_TOTAL_2026":   "Avalúo Total 2026 ($)",
        "LIQ_2026_CON_LÍMITE": "Liquidación 2026 con Ley 44 ($)",
        "color_cat":           "Estado",
    },
    title="Avalúo 2026 vs Liquidación 2026 (muestra de hasta 1,500 predios)",
    opacity=0.65,
)
fig6.update_layout(
    plot_bgcolor="white", paper_bgcolor="white",
    height=380, margin=dict(t=50, b=20, l=10, r=10),
    legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
)
fig6.update_traces(marker=dict(size=6))
st.plotly_chart(fig6, use_container_width=True)


# ─── TABLA DETALLE ────────────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">📋 Detalle de Predios Urbanos</div>', unsafe_allow_html=True)

# Columnas a mostrar
cols_tabla = [
    "FICHA", "TERCERO", "DIRECCIÓN_IGAC", "DEST_2025",
    "AVALÚO_2025", "LIQ_2025",
    "AVALÚO_TOTAL_2026", "VARIACIÓN_AVALÚO_%",
    "LIQ_2026_SIN_LÍMITE", "LIQ_2026_CON_LÍMITE",
    "APLICA_LÍMITE", "AHORRO_CONTRIBUYENTE",
    "EN_SALGAR_IGAC", "NOVEDADES",
]
cols_disp = [c for c in cols_tabla if c in df.columns]

df_display = df[cols_disp].copy()

# Renombrar para la vista
rename_vista = {
    "TERCERO":            "Propietario",
    "DIRECCIÓN_IGAC":     "Dirección",
    "DEST_2025":          "Destino 2025",
    "AVALÚO_2025":        "Avalúo 2025",
    "LIQ_2025":           "Liq. 2025",
    "AVALÚO_TOTAL_2026":  "Avalúo 2026",
    "VARIACIÓN_AVALÚO_%": "Var. Avalúo %",
    "LIQ_2026_SIN_LÍMITE":"Liq. 2026 bruta",
    "LIQ_2026_CON_LÍMITE":"Liq. 2026 Ley44",
    "APLICA_LÍMITE":      "Límite Ley44",
    "AHORRO_CONTRIBUYENTE":"Ahorro",
    "EN_SALGAR_IGAC":     "IGAC 2026",
    "NOVEDADES":          "Novedades",
}
df_display.rename(columns={k: v for k, v in rename_vista.items() if k in df_display.columns}, inplace=True)

# Buscador
txt_busq = st.text_input("🔎 Buscar por nombre, ficha o dirección:", placeholder="Escriba para filtrar...")
if txt_busq:
    mask = df_display.apply(lambda col: col.astype(str).str.contains(txt_busq, case=False, na=False)).any(axis=1)
    df_display = df_display[mask]

st.markdown(f"**{len(df_display):,}** registros · use los filtros del panel izquierdo para refinar")

# Formateo de columnas numéricas
fmt_cols_peso = ["Avalúo 2025", "Liq. 2025", "Avalúo 2026", "Liq. 2026 bruta", "Liq. 2026 Ley44", "Ahorro"]
fmt_cols_pct  = ["Var. Avalúo %"]

col_cfg = {}
for c in fmt_cols_peso:
    if c in df_display.columns:
        col_cfg[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
for c in fmt_cols_pct:
    if c in df_display.columns:
        col_cfg[c] = st.column_config.NumberColumn(c, format="%.2f %%")

st.dataframe(
    df_display.reset_index(drop=True),
    use_container_width=True,
    height=400,
    column_config=col_cfg,
)

# Descarga
csv_bytes = df_display.to_csv(index=False).encode("utf-8-sig")
st.download_button(
    "⬇️  Descargar tabla filtrada (.csv)",
    data=csv_bytes,
    file_name="predios_urbanos_salgar_2026_filtrado.csv",
    mime="text/csv",
)


# ─── NOTA LEGAL ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="nota-legal">
⚖️ <strong>Nota legal:</strong> El límite aplicado corresponde a LIQ_2025 × 2, según el Art. 6 de la Ley 44 de 1990
(el impuesto predial de un año no puede exceder el doble del año anterior). Las tarifas aplicadas provienen del
análisis 2025. Se recomienda verificar con el Estatuto Tributario Municipal vigente antes de liquidar oficialmente.
</div>
""", unsafe_allow_html=True)

st.markdown("<br><center style='color:#aaa; font-size:0.78rem;'>Municipio de Salgar · Análisis Predial 2026 · Ley 44/1990</center>", unsafe_allow_html=True)

