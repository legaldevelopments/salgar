"""
Tablero Predial Girardota 2026 — Streamlit
Análisis de avalúos, tarifas y límite Ley 44/1990 Art.6 (no más del doble del impuesto anterior)
Base: Predial_Girardota_2026_Ley44.xlsx (generado por procesar_predial_girardota.py)

Ejecutar: streamlit run tablero_girardota.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import os

# ── CONFIG ────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Predial Girardota 2026",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded",
)

RUTA_EXCEL = os.path.join(os.path.dirname(__file__), "Predial_Girardota_2026_Ley44.xlsx")

AZUL_OSC  = "#1F4E79"
AZUL_MED  = "#2E75B6"
AZUL_CLAR = "#BDD7EE"
VERDE     = "#70AD47"
AMBAR     = "#FFD966"
ROJO      = "#FF7C80"
NARANJA   = "#F4B183"
GRIS      = "#D9D9D9"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .stApp { background-color: #F0F4F8; }
  [data-testid="stSidebar"] {
      background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
  }
  [data-testid="stSidebar"] * { color: white !important; }
  [data-testid="stSidebar"] .stSelectbox label,
  [data-testid="stSidebar"] .stMultiSelect label,
  [data-testid="stSidebar"] .stSlider label { color: #BDD7EE !important; font-weight: 600; }

  .kpi-card {
      background: white; border-radius: 12px; padding: 18px 14px;
      text-align: center; box-shadow: 0 2px 8px rgba(0,0,0,0.10);
      border-top: 4px solid #2E75B6; min-height: 120px;
      display: flex; flex-direction: column; justify-content: center;
  }
  .kpi-card.verde  { border-top-color: #70AD47; }
  .kpi-card.ambar  { border-top-color: #FFD966; }
  .kpi-card.rojo   { border-top-color: #FF7C80; }
  .kpi-card.oscuro { border-top-color: #1F4E79; }
  .kpi-card.naranja{ border-top-color: #F4B183; }
  .kpi-valor { font-size: 1.65rem; font-weight: 800; color: #1F4E79; line-height:1.1; }
  .kpi-label { font-size: 0.73rem; color: #666; margin-top:5px; font-weight:500;
               text-transform:uppercase; letter-spacing:0.5px; }
  .kpi-sub   { font-size: 0.78rem; color: #2E75B6; font-weight:600; margin-top:3px; }

  .sec-tit {
      background: linear-gradient(90deg, #1F4E79, #2E75B6);
      color: white; padding: 7px 16px; border-radius: 8px;
      font-size: 0.92rem; font-weight: 700; letter-spacing: 0.4px;
      margin: 20px 0 10px 0;
  }
  .header-main {
      background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 55%, #3A9AD9 100%);
      border-radius: 14px; padding: 26px 30px; color: white;
      margin-bottom: 22px; box-shadow: 0 4px 15px rgba(31,78,121,0.3);
  }
  .header-main h1 { margin:0; font-size:1.75rem; font-weight:800; }
  .header-main p  { margin:4px 0 0; opacity:0.85; font-size:0.9rem; }
  .alerta-rojo {
      background:#FFF0F0; border-left:4px solid #FF7C80;
      border-radius:6px; padding:10px 14px; font-size:0.82rem; color:#555; margin-top:6px;
  }
  .alerta-verde {
      background:#F0FFF0; border-left:4px solid #70AD47;
      border-radius:6px; padding:10px 14px; font-size:0.82rem; color:#555; margin-top:6px;
  }
  .nota-legal {
      background:#FFF9E6; border-left:4px solid #FFD966;
      border-radius:6px; padding:10px 14px; font-size:0.80rem; color:#555; margin-top:8px;
  }
</style>
""", unsafe_allow_html=True)


# ── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Cargando datos Girardota 2026...")
def cargar():
    df_u = pd.read_excel(RUTA_EXCEL, sheet_name="Predios_Urbanos", header=1)
    df_r = pd.read_excel(RUTA_EXCEL, sheet_name="Predios_Rurales", header=1)
    df_u["ZONA"] = "URBANO"
    df_r["ZONA"] = "RURAL"
    df = pd.concat([df_u, df_r], ignore_index=True)

    # Fichas únicas: una fila por predio (primera ocurrencia)
    df = df.drop_duplicates(subset=["FICHA"], keep="first").reset_index(drop=True)

    # Columnas booleanas — nombres nuevos
    def yn(col):
        return df[col].astype(str).str.strip().str.upper() == "SÍ"

    df["_aplica"]     = yn("APLICA_LIMITE_LEY44")
    df["_violo"]      = yn("LIQUIDACION_EN_EXCESO")
    df["_tiene_hist"] = df["IMPTO_TOTAL_2025"].fillna(0) > 0

    # Columnas numéricas
    for c in [
        "AVALUO_TOTAL_2025", "AVALUO_TOTAL_2026", "AVALUO_DER_2026",
        "IMPTO_TOTAL_2025", "IMPTO_INICIAL_2026", "IMPUESTO_PLENO_2026",
        "LIMITE_LEY44", "IMPTO_CORRECTO_2026",
        "EXCESO_LIQ_INICIAL", "IMPACTO_CONTRIBUYENTE",
        "LIMITE_LOCAL", "INC_LIMITE_LOCAL", "EXCESO_LOCAL", "AHORRO_LOCAL",
        "VAR_AVALUO_%", "VAR_IMPTO_PCT", "TARIFA_2025_MIL", "TARIFA_2026_MIL",
    ]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    if "APLICA_LIMITE_LOCAL" in df.columns:
        df["_aplica_local"] = df["APLICA_LIMITE_LOCAL"].astype(str).str.strip().str.upper() == "SÍ"
    else:
        df["_aplica_local"] = False

    df["DEST_NOM_2026"]   = df["DEST_NOM_2026"].fillna("Sin destino")
    df["RANGO_AVALUO_2026"] = df["RANGO_AVALUO_2026"].fillna("Sin dato")

    orden_rng = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M","Sin dato"]
    df["_rango_ord"] = pd.Categorical(df["RANGO_AVALUO_2026"], categories=orden_rng, ordered=True)

    return df


if not os.path.exists(RUTA_EXCEL):
    st.error(
        f"No se encontró **{RUTA_EXCEL}**. "
        "Ejecute primero `procesar_predial_girardota.py` para generar el archivo."
    )
    st.stop()

df_all = cargar()


# ── SIDEBAR ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏛️ Predial Girardota 2026")
    st.markdown("**Análisis Catastral · Ley 44**")
    st.divider()
    st.markdown("#### Filtros globales")

    sel_zona = st.selectbox("Zona:", ["Todos","Solo Urbano","Solo Rural"])
    sel_lim  = st.selectbox("Límite Ley 44:",
                            ["Todos","Liq. inicial excedió límite","Dentro del límite","No aplica límite"])

    destinos_disp = sorted(df_all["DEST_NOM_2026"].dropna().unique())
    sel_dest = st.multiselect("Destino 2026:", destinos_disp, placeholder="Todos")

    rangos_disp = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M"]
    sel_rng  = st.multiselect("Rango de avalúo 2026:", rangos_disp, placeholder="Todos")

    st.divider()
    umbral_av = st.slider("Umbral 'cambio extremo' avalúo (%):", 20, 500, 100, 10)

    st.divider()
    st.markdown("""
    <div style='font-size:0.73rem; opacity:0.8;'>
    📂 Fuente: <i>Estructura_02.xlsx</i><br>
    Sistema de Información + Gestor Catastral 2026<br><br>
    📋 Marco legal:<br>
    · Ley 44/1990 Art. 6<br>
    · Límite: doble del impuesto 2025<br>
    · UVT 2026: $52.374
    </div>""", unsafe_allow_html=True)


# ── FILTRAR ───────────────────────────────────────────────────────────────────
df = df_all.copy()
if sel_zona == "Solo Urbano":
    df = df[df["ZONA"] == "URBANO"]
elif sel_zona == "Solo Rural":
    df = df[df["ZONA"] == "RURAL"]

if sel_lim == "Liq. inicial excedió límite":
    df = df[df["_violo"]]
elif sel_lim == "Dentro del límite":
    df = df[df["_aplica"] & ~df["_violo"] & df["_tiene_hist"]]
elif sel_lim == "No aplica límite":
    df = df[~df["_aplica"]]

if sel_dest:
    df = df[df["DEST_NOM_2026"].isin(sel_dest)]
if sel_rng:
    df = df[df["RANGO_AVALUO_2026"].isin(sel_rng)]


# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-main">
  <h1>🏛️ Análisis Predial — Municipio de Girardota</h1>
  <p>Actualización catastral 2026 · Sistema de Información vs Impuesto Inicial · Ley 44/1990 Art.6 · Límite: doble del impuesto 2025 · UVT 2026: $52.374</p>
</div>""", unsafe_allow_html=True)

n_f = df["FICHA"].nunique()
n_t = df_all["FICHA"].nunique()
n_u = df[df["ZONA"]=="URBANO"]["FICHA"].nunique()
n_r = df[df["ZONA"]=="RURAL"]["FICHA"].nunique()
if n_f < n_t:
    st.info(f"Mostrando **{n_f:,}** de **{n_t:,}** fichas únicas según filtros — **{n_u:,} urbanas · {n_r:,} rurales**")
else:
    st.info(f"**{n_t:,}** fichas únicas totales — **{n_u:,} urbanas · {n_r:,} rurales**")


# ── HELPERS ───────────────────────────────────────────────────────────────────
def fmt_cop(v):
    if pd.isna(v): return "$0"
    if abs(v) >= 1e6: return f"${v/1e6:,.0f} M"
    return f"${v:,.0f}"

def kpi(col, val, lab, sub="", color=""):
    with col:
        st.markdown(f"""
        <div class="kpi-card {color}">
          <div class="kpi-valor">{val}</div>
          <div class="kpi-label">{lab}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)


# ── KPIs ─────────────────────────────────────────────────────────────────────
st.markdown('<div class="sec-tit">📊 Indicadores Clave</div>', unsafe_allow_html=True)

i25    = df["IMPTO_TOTAL_2025"].fillna(0).sum()
ig26   = df["IMPTO_INICIAL_2026"].fillna(0).sum()
ic26   = df["IMPTO_CORRECTO_2026"].fillna(0).sum()
exceso = df["EXCESO_LIQ_INICIAL"].fillna(0).sum()
n_viol = df[df["_violo"]]["FICHA"].nunique()
var_g  = (ig26 - i25) / i25 * 100 if i25 > 0 else 0

cols_kpi = st.columns(7)
kpi(cols_kpi[0], f"{n_f:,}",        "Total Predios",            sel_zona or "Todos",       "oscuro")
kpi(cols_kpi[1], fmt_cop(i25),      "Recaudo 2025",             "Sistema información",     "")
kpi(cols_kpi[2], fmt_cop(ig26),     "Recaudo Inicial 2026",     "Impuesto inicial en txt", "ambar")
kpi(cols_kpi[3], fmt_cop(ic26),     "Recaudo Correcto 2026",    "Con límite Ley 44",       "verde")
kpi(cols_kpi[4], f"{var_g:+.1f}%", "Var. Inicial vs 2025",     "Sin aplicar Ley 44",      "rojo" if var_g > 100 else "verde")
kpi(cols_kpi[5], f"{n_viol:,}",    "Liq. Inicial Excedió",     "Excedió límite Ley 44",   "rojo")
kpi(cols_kpi[6], fmt_cop(exceso),   "Exceso Liq. Inicial",      "vs límite Ley 44",        "naranja")
st.markdown("<br>", unsafe_allow_html=True)

# KPIs Acuerdo 49
n_loc    = df[df["_aplica_local"]]["FICHA"].nunique()
exc_loc  = df["EXCESO_LOCAL"].fillna(0).sum()       if "EXCESO_LOCAL"       in df.columns else 0
ahor_loc = df["AHORRO_LOCAL"].fillna(0).sum()       if "AHORRO_LOCAL"       in df.columns else 0
n_loc25  = df[df["TIPO_LIMITE_LOCAL"]=="25%"]["FICHA"].nunique() if "TIPO_LIMITE_LOCAL" in df.columns else 0
n_loc50  = df[df["TIPO_LIMITE_LOCAL"]=="50%"]["FICHA"].nunique() if "TIPO_LIMITE_LOCAL" in df.columns else 0

st.markdown('<div class="sec-tit">🏛️ Límites Acuerdo 49</div>', unsafe_allow_html=True)
cols_loc = st.columns(5)
kpi(cols_loc[0], f"{n_loc:,}",     "Predios Acuerdo 49",      "Sin cambio dest/área",         "oscuro")
kpi(cols_loc[1], f"{n_loc25:,}",   "Dest 01/24 — Límite 25%", "Habitacional/Agrícola",         "ambar")
kpi(cols_loc[2], f"{n_loc50:,}",   "Otros — Límite 50%",      "Dentro de Ley 44",              "ambar")
kpi(cols_loc[3], fmt_cop(exc_loc), "Exceso s/ Acuerdo 49",    "Inicial vs límite Acuerdo 49",  "naranja")
kpi(cols_loc[4], fmt_cop(ahor_loc),"Impacto Contribuyente",   "Con Acuerdo 49",                "verde")
st.markdown("<br>", unsafe_allow_html=True)


# ── SECCIÓN 1: COMPARATIVO FISCAL ────────────────────────────────────────────
st.markdown('<div class="sec-tit">📈 Comparativo Fiscal 2025 vs 2026</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns([2.2, 1.8, 1.8])

with c1:
    fig_bar = go.Figure(go.Bar(
        x=["Recaudo 2025", "Inicial 2026\n(tal como está)", "Correcto 2026\n(con Ley 44)"],
        y=[i25, ig26, ic26],
        marker_color=[AZUL_MED, ROJO if ig26 > ic26 else AMBAR, VERDE],
        text=[fmt_cop(v) for v in [i25, ig26, ic26]],
        textposition="outside", textfont=dict(size=11, color="black"),
    ))
    fig_bar.update_layout(
        title=dict(text="Comparativo Recaudo ($)", font=dict(size=13, color=AZUL_OSC)),
        xaxis=dict(tickfont=dict(color="black")),
        yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee", tickfont=dict(color="black")),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50,b=20,l=10,r=20), height=320, showlegend=False,
    )
    st.plotly_chart(fig_bar, use_container_width=True)

with c2:
    _fs = df.groupby("FICHA").agg(
        violo=("_violo","any"), aplica=("_aplica","any"), tiene_hist=("_tiene_hist","any")
    ).reset_index()
    n_violo_v  = int(_fs["violo"].sum())
    n_ok_v     = int((~_fs["violo"] & _fs["aplica"] & _fs["tiene_hist"]).sum())
    n_noapl    = int((~_fs["aplica"]).sum())
    n_pred2026 = int((_fs["aplica"] & ~_fs["tiene_hist"] & ~_fs["violo"]).sum())
    fig_pie = go.Figure(go.Pie(
        labels=["Liq. Inicial Excedió","Dentro del límite","No aplica límite","Predios 2026"],
        values=[n_violo_v, n_ok_v, n_noapl, n_pred2026],
        hole=0.52,
        marker_colors=[ROJO, VERDE, AMBAR, GRIS],
        textinfo="percent+label", textfont=dict(size=9),
        hovertemplate="%{label}: %{value:,}<extra></extra>",
    ))
    fig_pie.update_layout(
        title=dict(text="Aplicación Límite Ley 44", font=dict(size=13, color=AZUL_OSC)),
        showlegend=False, margin=dict(t=50,b=0,l=0,r=0), height=320,
        paper_bgcolor="white",
        annotations=[dict(text=f"<b>{n_f:,}</b><br>predios",
                          x=0.5,y=0.5,font_size=12,showarrow=False,font_color=AZUL_OSC)],
    )
    st.plotly_chart(fig_pie, use_container_width=True)

with c3:
    liq_u = float(df[df["ZONA"]=="URBANO"]["IMPTO_CORRECTO_2026"].fillna(0).sum())
    liq_r = float(df[df["ZONA"]=="RURAL"]["IMPTO_CORRECTO_2026"].fillna(0).sum())
    fig_urr = go.Figure(go.Pie(
        labels=["Urbano","Rural"],
        values=[liq_u, liq_r],
        hole=0.55,
        marker_colors=[AZUL_MED, VERDE],
        textinfo="label+percent", textfont=dict(size=10),
        hovertemplate="%{label}: %{value:$,.0f}<extra></extra>",
    ))
    fig_urr.update_layout(
        title=dict(text="Recaudo Correcto por Zona", font=dict(size=13, color=AZUL_OSC)),
        showlegend=False, margin=dict(t=50,b=0,l=0,r=0), height=320,
        paper_bgcolor="white",
    )
    st.plotly_chart(fig_urr, use_container_width=True)


# ── SECCIÓN 2: DISTRIBUCIÓN AVALÚOS ──────────────────────────────────────────
st.markdown('<div class="sec-tit">📉 Distribución de Avalúos y Variación</div>', unsafe_allow_html=True)
c4, c5 = st.columns(2)

with c4:
    col_var = "VAR_AVALUO_%"
    df_var = df[df[col_var].notna() & (df[col_var].abs() < 2000)] if col_var in df.columns else pd.DataFrame()
    if len(df_var) > 0:
        fig_hist = px.histogram(
            df_var, x=col_var, nbins=60,
            color_discrete_sequence=[AZUL_MED],
            labels={col_var: "Variación Avalúo (%)"},
            title="Distribución Variación de Avalúos (Sistema 2025 → Inicial 2026)",
        )
        fig_hist.add_vline(x=0,   line_dash="dash", line_color=AZUL_OSC, annotation_text="0%")
        fig_hist.add_vline(x=100, line_dash="dot",  line_color=ROJO,
                           annotation_text="Límite Ley 44 = 100% (doble)")
        fig_hist.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=50,b=20,l=10,r=10), height=320,
            yaxis_title="N° predios", bargap=0.05,
        )
        st.plotly_chart(fig_hist, use_container_width=True)

with c5:
    df_dest_agg = (
        df.groupby("DEST_NOM_2026")
        .agg(predios=("FICHA","count"),
             recaudo_correcto=("IMPTO_CORRECTO_2026","sum"),
             exceso_total=("EXCESO_LIQ_INICIAL","sum"),
             n_violo=("_violo","sum"))
        .reset_index()
        .sort_values("recaudo_correcto", ascending=True)
        .tail(14)
    )
    fig_dest = go.Figure(go.Bar(
        x=df_dest_agg["recaudo_correcto"],
        y=df_dest_agg["DEST_NOM_2026"],
        orientation="h", marker_color=AZUL_MED,
        text=[fmt_cop(v) for v in df_dest_agg["recaudo_correcto"]],
        textposition="outside",
    ))
    fig_dest.update_layout(
        title=dict(text="Recaudo Correcto 2026 por Destino", font=dict(size=13, color=AZUL_OSC)),
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee"),
        margin=dict(t=50,b=20,l=10,r=80), height=320,
    )
    st.plotly_chart(fig_dest, use_container_width=True)


# ── SECCIÓN 3: ANÁLISIS LEY 44 ────────────────────────────────────────────────
st.markdown('<div class="sec-tit">⚖️ Análisis de Cumplimiento Ley 44/1990 Art.6 — Límite: doble del impuesto 2025</div>',
            unsafe_allow_html=True)
c6, c7 = st.columns(2)

with c6:
    df_viol_dest = (
        df[df["_violo"]]
        .groupby("DEST_NOM_2026")
        .agg(n_violo=("FICHA","count"), exceso=("EXCESO_LIQ_INICIAL","sum"))
        .reset_index()
        .sort_values("exceso", ascending=True)
    )
    if len(df_viol_dest) > 0:
        fig_vd = go.Figure()
        fig_vd.add_trace(go.Bar(
            x=df_viol_dest["exceso"], y=df_viol_dest["DEST_NOM_2026"],
            orientation="h", marker_color=ROJO,
            name="Exceso liq. inicial ($)",
            text=[fmt_cop(v) for v in df_viol_dest["exceso"]], textposition="outside",
        ))
        fig_vd.update_layout(
            title=dict(text="Exceso Liq. Inicial por Destino (vs límite Ley 44)", font=dict(size=13, color=AZUL_OSC)),
            plot_bgcolor="white", paper_bgcolor="white",
            xaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee"),
            margin=dict(t=50,b=20,l=10,r=80), height=340,
        )
        st.plotly_chart(fig_vd, use_container_width=True)
    else:
        st.success("No hay predios donde la liq. inicial haya excedido el límite Ley 44 en la selección actual.")

with c7:
    df_rng_agg = (
        df.groupby("RANGO_AVALUO_2026")
        .agg(total=("FICHA","count"),
             violo=("_violo","sum"),
             exceso=("EXCESO_LIQ_INICIAL","sum"),
             recaudo_25=("IMPTO_TOTAL_2025","sum"),
             recaudo_g=("IMPTO_INICIAL_2026","sum"),
             recaudo_c=("IMPTO_CORRECTO_2026","sum"))
        .reset_index()
    )
    orden = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M"]
    df_rng_agg["_ord"] = df_rng_agg["RANGO_AVALUO_2026"].apply(
        lambda x: orden.index(x) if x in orden else 99)
    df_rng_agg = df_rng_agg.sort_values("_ord")

    fig_rng = go.Figure()
    fig_rng.add_trace(go.Bar(
        name="Recaudo 2025", x=df_rng_agg["RANGO_AVALUO_2026"], y=df_rng_agg["recaudo_25"],
        marker_color=AZUL_CLAR,
    ))
    fig_rng.add_trace(go.Bar(
        name="Inicial 2026", x=df_rng_agg["RANGO_AVALUO_2026"], y=df_rng_agg["recaudo_g"],
        marker_color=ROJO,
    ))
    fig_rng.add_trace(go.Bar(
        name="Correcto 2026 (Ley 44)", x=df_rng_agg["RANGO_AVALUO_2026"], y=df_rng_agg["recaudo_c"],
        marker_color=VERDE,
    ))
    fig_rng.update_layout(
        title=dict(text="Recaudo por Rango de Avalúo", font=dict(size=13, color=AZUL_OSC)),
        barmode="group", plot_bgcolor="white", paper_bgcolor="white",
        yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee"),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
        margin=dict(t=60,b=20,l=10,r=10), height=340,
    )
    st.plotly_chart(fig_rng, use_container_width=True)


# ── SECCIÓN 4: SCATTER AVALÚO vs IMPUESTO ────────────────────────────────────
st.markdown('<div class="sec-tit">🔍 Relación Avalúo 2026 vs Impuesto Inicial</div>',
            unsafe_allow_html=True)

df_sc = df[df["AVALUO_TOTAL_2026"].notna() & df["IMPTO_INICIAL_2026"].notna()].copy()
if len(df_sc) > 0:
    df_sc["_cat"] = df_sc["_violo"].map(
        {True: "Liq. Inicial Excedió", False: "Dentro del límite"})
    muestra = df_sc.sample(min(len(df_sc), 2000), random_state=42)
    fig_sc = px.scatter(
        muestra,
        x="AVALUO_TOTAL_2026", y="IMPTO_INICIAL_2026",
        color="_cat",
        color_discrete_map={"Liq. Inicial Excedió": ROJO, "Dentro del límite": VERDE},
        hover_data={"FICHA": True, "DEST_NOM_2026": True,
                    "VAR_AVALUO_%": ":.1f", "_cat": False},
        labels={"AVALUO_TOTAL_2026": "Avalúo Total 2026 ($)",
                "IMPTO_INICIAL_2026": "Impuesto Inicial 2026 ($)",
                "_cat": "Estado Ley 44"},
        title="Avalúo 2026 vs Impuesto Inicial — muestra hasta 2.000 predios",
        opacity=0.6,
    )
    fig_sc.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, margin=dict(t=50,b=20,l=10,r=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.01),
        xaxis=dict(tickformat="$,.0f"), yaxis=dict(tickformat="$,.0f"),
    )
    fig_sc.update_traces(marker=dict(size=5))
    st.plotly_chart(fig_sc, use_container_width=True)


# ── SECCIÓN 5: TABLA DINÁMICA — MATRIZ ───────────────────────────────────────
st.markdown('<div class="sec-tit">📊 Matriz Dinámica: Destino × Rango de Avalúo</div>',
            unsafe_allow_html=True)

tab1, tab2, tab3 = st.tabs(["N° Predios", "Impto Inicial 2026 ($)", "Exceso Liq. Inicial ($)"])

orden_rng_list = ["0–5 M","5–15 M","15–50 M","50–150 M","150–500 M",">500 M"]
df_mat = df[df["RANGO_AVALUO_2026"].isin(orden_rng_list)].copy()

def pivote(val_col, aggfunc="sum", fmt_fn=None):
    if df_mat.empty or val_col not in df_mat.columns:
        return pd.DataFrame()
    piv = pd.pivot_table(
        df_mat, values=val_col, index="DEST_NOM_2026",
        columns="RANGO_AVALUO_2026", aggfunc=aggfunc, fill_value=0,
        observed=True,
    )
    cols_ord = [c for c in orden_rng_list if c in piv.columns]
    piv = piv[cols_ord]
    piv["TOTAL"] = piv.sum(axis=1)
    piv = piv.sort_values("TOTAL", ascending=False)
    fila_total = piv.sum()
    fila_total.name = "▶ TOTAL"
    piv = pd.concat([piv, fila_total.to_frame().T])
    if fmt_fn:
        return piv.applymap(fmt_fn)
    return piv

with tab1:
    piv_n = pivote("FICHA", aggfunc="count")
    if not piv_n.empty:
        st.dataframe(piv_n.astype(int).style.background_gradient(
            cmap="Blues", axis=None, subset=orden_rng_list,
        ), use_container_width=True, height=380)

with tab2:
    piv_g = pivote("IMPTO_INICIAL_2026")
    if not piv_g.empty:
        st.dataframe(
            piv_g.style.format("${:,.0f}").background_gradient(cmap="Oranges", axis=None, subset=orden_rng_list),
            use_container_width=True, height=380,
        )

with tab3:
    piv_e = pivote("EXCESO_LIQ_INICIAL")
    if not piv_e.empty:
        st.dataframe(
            piv_e.style.format("${:,.0f}").background_gradient(cmap="Reds", axis=None, subset=orden_rng_list),
            use_container_width=True, height=380,
        )


# ── SECCIÓN 6: CAMBIOS EXTREMOS ──────────────────────────────────────────────
st.markdown(f'<div class="sec-tit">🚨 Predios con Cambios Extremos (Avalúo ≥ {umbral_av}% o Liq. inicial excedió)</div>',
            unsafe_allow_html=True)

col_var = "VAR_AVALUO_%"
df_ext = df[
    (df[col_var].abs().fillna(0) >= umbral_av) | df["_violo"]
].copy().sort_values(col_var, ascending=False) if col_var in df.columns else df[df["_violo"]].copy()

n_ext      = len(df_ext)
n_ext_viol = int(df_ext["_violo"].sum())
st.markdown(
    f"**{n_ext:,}** predios con cambio extremo · "
    f"de estos **{n_ext_viol:,}** con liq. inicial excedió el límite."
)

if len(df_ext) > 0:
    top20 = df_ext.head(20)
    fig_ext = go.Figure(go.Bar(
        x=top20[col_var],
        y=top20["FICHA"].astype(str),
        orientation="h",
        marker_color=[ROJO if v else NARANJA for v in top20["_violo"]],
        text=[f"{v:.1f}%" if pd.notna(v) else "N/A" for v in top20[col_var]],
        textposition="outside",
    ))
    fig_ext.update_layout(
        title=dict(text="Top 20 mayores variaciones de avalúo (%)", font=dict(size=13, color=AZUL_OSC)),
        plot_bgcolor="white", paper_bgcolor="white",
        xaxis_title="Variación %", yaxis_title="Ficha",
        margin=dict(t=50,b=20,l=10,r=60), height=420,
        showlegend=False,
    )
    st.plotly_chart(fig_ext, use_container_width=True)

# Tabla detalle de extremos
cols_ext = [c for c in [
    "FICHA","ZONA","DEST_NOM_2026","AVALUO_TOTAL_2025","AVALUO_TOTAL_2026",
    "VAR_AVALUO_%","IMPTO_TOTAL_2025","IMPTO_INICIAL_2026",
    "LIMITE_LEY44","LIQUIDACION_EN_EXCESO","EXCESO_LIQ_INICIAL","IMPTO_CORRECTO_2026",
] if c in df_ext.columns]

col_cfg_ext = {}
for c in ["AVALUO_TOTAL_2025","AVALUO_TOTAL_2026","IMPTO_TOTAL_2025",
          "IMPTO_INICIAL_2026","LIMITE_LEY44","EXCESO_LIQ_INICIAL","IMPTO_CORRECTO_2026"]:
    if c in cols_ext:
        col_cfg_ext[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
if "VAR_AVALUO_%" in cols_ext:
    col_cfg_ext["VAR_AVALUO_%"] = st.column_config.NumberColumn("VAR_AVALUO_%", format="%.2f %%")

st.dataframe(df_ext[cols_ext].reset_index(drop=True),
             use_container_width=True, height=360, column_config=col_cfg_ext)

csv_ext = df_ext[cols_ext].to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Descargar predios extremos (.csv)", data=csv_ext,
                   file_name="girardota_cambios_extremos.csv", mime="text/csv")


# ── SECCIÓN 7: RESUMEN POR DESTINO (TABLA COMPLETA) ─────────────────────────
st.markdown('<div class="sec-tit">📋 Resumen por Destino Económico</div>', unsafe_allow_html=True)

df_res_dest = (
    df.groupby(["ZONA","DEST_NOM_2026"])
    .agg(
        N_Predios=("FICHA","count"),
        Aplica_Limite_Ley44=("_aplica","sum"),
        Liq_Inicial_Excedio=("_violo","sum"),
        Predios_2026=("_tiene_hist", lambda x: (~x).sum()),
        Avaluo_2025=("AVALUO_TOTAL_2025","sum"),
        Avaluo_2026=("AVALUO_TOTAL_2026","sum"),
        Impto_2025=("IMPTO_TOTAL_2025","sum"),
        Impto_Inicial_2026=("IMPTO_INICIAL_2026","sum"),
        Impto_Correcto_2026=("IMPTO_CORRECTO_2026","sum"),
        Exceso_Liq_Inicial=("EXCESO_LIQ_INICIAL","sum"),
        Impacto_Contribuyente=("IMPACTO_CONTRIBUYENTE","sum"),
    )
    .reset_index()
    .sort_values(["ZONA","Liq_Inicial_Excedio"], ascending=[True, False])
)

df_res_dest["Var_Impto_%"] = (
    (df_res_dest["Impto_Correcto_2026"] - df_res_dest["Impto_2025"])
    / df_res_dest["Impto_2025"].replace(0, np.nan) * 100
).round(2)

df_res_dest["Pct_Liq_Inicial"] = (
    df_res_dest["Liq_Inicial_Excedio"] / df_res_dest["Aplica_Limite_Ley44"].replace(0, np.nan) * 100
).round(1)

col_cfg_res = {}
for c in ["Avaluo_2025","Avaluo_2026","Impto_2025","Impto_Inicial_2026",
          "Impto_Correcto_2026","Exceso_Liq_Inicial","Impacto_Contribuyente"]:
    col_cfg_res[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
col_cfg_res["Var_Impto_%"]    = st.column_config.NumberColumn("Var_Impto_%",   format="%.2f %%")
col_cfg_res["Pct_Liq_Inicial"]= st.column_config.NumberColumn("% liq. inicial", format="%.1f %%")

st.dataframe(df_res_dest.reset_index(drop=True), use_container_width=True,
             height=380, column_config=col_cfg_res)

csv_res = df_res_dest.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Descargar resumen por destino (.csv)", data=csv_res,
                   file_name="girardota_resumen_destino.csv", mime="text/csv")


# ── SECCIÓN 8: DETALLE PREDIOS ────────────────────────────────────────────────
st.markdown('<div class="sec-tit">📋 Detalle de Predios</div>', unsafe_allow_html=True)

cols_vis = [c for c in [
    "FICHA","N_PROPIETARIOS","ZONA","DEST_NOM_2026","DIRECCION",
    "AVALUO_TOTAL_2025","AVALUO_TOTAL_2026","VAR_AVALUO_%","AVALUO_DER_2026",
    "TARIFA_2025_MIL","TARIFA_2026_MIL",
    "IMPTO_TOTAL_2025","IMPTO_INICIAL_2026","IMPUESTO_PLENO_2026",
    "APLICA_LIMITE_LEY44","LIMITE_LEY44",
    "LIQUIDACION_EN_EXCESO","IMPTO_CORRECTO_2026",
    "EXCESO_LIQ_INICIAL","IMPACTO_CONTRIBUYENTE","VAR_IMPTO_PCT",
    "APLICA_LIMITE_LOCAL","TIPO_LIMITE_LOCAL","LIMITE_LOCAL",
    "INC_LIMITE_LOCAL","EXCESO_LOCAL","AHORRO_LOCAL",
    "RANGO_AVALUO_2026","NOVEDADES",
] if c in df.columns]

busq = st.text_input("🔎 Buscar por ficha, propietario o dirección:", placeholder="Escriba para filtrar…")
df_vis = df[cols_vis].copy()
if busq:
    mask = df_vis.apply(
        lambda col: col.astype(str).str.contains(busq, case=False, na=False)
    ).any(axis=1)
    df_vis = df_vis[mask]

st.markdown(f"**{len(df_vis):,}** registros")

col_cfg_vis = {}
for c in ["AVALUO_TOTAL_2025","AVALUO_TOTAL_2026","AVALUO_DER_2026",
          "IMPTO_TOTAL_2025","IMPTO_INICIAL_2026","IMPUESTO_PLENO_2026",
          "LIMITE_LEY44","IMPTO_CORRECTO_2026",
          "EXCESO_LIQ_INICIAL","IMPACTO_CONTRIBUYENTE",
          "LIMITE_LOCAL","INC_LIMITE_LOCAL","EXCESO_LOCAL","AHORRO_LOCAL"]:
    if c in cols_vis:
        col_cfg_vis[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
for c in ["VAR_AVALUO_%","VAR_IMPTO_PCT"]:
    if c in cols_vis:
        col_cfg_vis[c] = st.column_config.NumberColumn(c, format="%.2f %%")
for c in ["TARIFA_2025_MIL","TARIFA_2026_MIL"]:
    if c in cols_vis:
        col_cfg_vis[c] = st.column_config.NumberColumn(c, format="%.2f ‰")

st.dataframe(df_vis.reset_index(drop=True), use_container_width=True,
             height=450, column_config=col_cfg_vis)

csv_all_bytes = df_vis.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️ Descargar tabla filtrada (.csv)", data=csv_all_bytes,
                   file_name="girardota_predios_2026.csv", mime="text/csv")


# ── NOTA LEGAL ────────────────────────────────────────────────────────────────
st.markdown("""
<div class="nota-legal">
⚖️ <strong>Nota legal:</strong> Límite Ley 44/1990 Art.6 = IMP_2025 × 2 (el impuesto no puede ser mayor al doble del año anterior).
No aplica para lotes sin construir (destinos 12,13) ni uso público/exento (19).
Dest. 14 (Lote No Urbanizable) y 31 (Lote Rural) aplican Ley 44. Parcelas rurales (21,22,23) van a tarifa plena.
Los predios sin impuesto 2025 (Predios 2026) no tienen límite calculable. Los datos están agregados por ficha (total de todos los propietarios).
UVT 2026: $52.374. Verificar con el Acuerdo Municipal vigente de Girardota.
</div>""", unsafe_allow_html=True)

st.markdown(
    "<br><center style='color:#aaa; font-size:0.76rem;'>"
    "Municipio de Girardota · Predial 2026 · Ley 44/1990 Art.6 · UVT 2026: $52.374 · Límite: doble del impuesto 2025"
    "</center>", unsafe_allow_html=True,
)



