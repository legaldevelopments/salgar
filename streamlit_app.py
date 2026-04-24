"""
Tablero Predial Salgar 2026
Análisis de formación catastral y liquidación del impuesto predial
Base: Gestor Catastral (urbano) + Base anterior Municipio (urbano y rural)
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import os

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Predial Salgar 2026",
    page_icon="🏘️",
    layout="wide",
    initial_sidebar_state="expanded",
)

RUTA_EXCEL = os.path.join(os.path.dirname(__file__), "Predial_Salgar_2026_Ley44.xlsx")

AZUL_OSC  = "#1F4E79"
AZUL_MED  = "#2E75B6"
AZUL_CLAR = "#BDD7EE"
VERDE     = "#70AD47"
AMBAR     = "#FFD966"
ROJO      = "#FF7C80"
NARANJA   = "#F4B183"

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F0F4F8; }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
    }
    [data-testid="stSidebar"] * { color: white !important; }
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stMultiSelect label { color: #BDD7EE !important; font-weight: 600; }

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
    .kpi-valor { font-size: 1.7rem; font-weight: 800; color: #1F4E79; line-height: 1.1; }
    .kpi-label { font-size: 0.75rem; color: #666; margin-top: 6px; font-weight: 500;
                 text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-sub   { font-size: 0.8rem; color: #2E75B6; font-weight: 600; margin-top: 4px; }

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
    df_urb = pd.read_excel(RUTA_EXCEL, sheet_name="Predios_Urbanos_2026", header=1)
    df_rur = pd.read_excel(RUTA_EXCEL, sheet_name="Predios_Rurales_2026", header=1)
    df_nev = pd.read_excel(RUTA_EXCEL, sheet_name="Predios_Nuevos_Gestor")
    df_tab = pd.read_excel(RUTA_EXCEL, sheet_name="Tabla_Salgar")

    df_urb["SECTOR"] = "URBANO"
    df_rur["SECTOR"] = "RURAL"
    df_all = pd.concat([df_urb, df_rur], ignore_index=True)

    df_all["_TIENE_GESTOR"] = df_all["EN_BASE_GESTOR"].astype(str).str.upper().str.strip() == "SÍ"
    df_all["_APLICA_LIM"]   = df_all["APLICA_LIMITE"].astype(str).str.upper().str.strip() == "SÍ"
    df_all["_EXENTO"]       = df_all["EXENTO"].astype(str).str.upper().str.strip() == "S"

    return df_all, df_nev, df_tab


df_all, df_nev, df_tab = cargar_datos()


# ─── SIDEBAR ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏘️ Predial Salgar 2026")
    st.markdown("**Análisis Catastral · Acuerdo 022/2024**")
    st.divider()

    st.markdown("#### Filtros")

    # Filtro: sector
    sector_opts = ["Todos (Urbano + Rural)", "Solo Urbano", "Solo Rural"]
    sel_sector = st.selectbox("Sector:", sector_opts)

    # Filtro: estado Gestor Catastral (solo relevante para urbano)
    gest_opts = ["Todos", "Con datos Gestor 2026", "Sin datos Gestor 2026"]
    sel_gest = st.selectbox(
        "Base del Gestor Catastral:", gest_opts,
        disabled=(sel_sector == "Solo Rural"),
        help="Los predios rurales no están en el Gestor Catastral 2026",
    )

    # Filtro: aplica límite
    lim_opts = ["Todos", "Aplica límite Ley 44", "No aplica límite"]
    sel_limite = st.selectbox("Límite Ley 44:", lim_opts)

    # Filtro: destino
    destinos = sorted(df_all["DEST_2025"].dropna().unique().tolist())
    sel_dest = st.multiselect("Destino económico:", destinos, placeholder="Todos")

    # Filtro: exento
    sel_exento = st.checkbox("Excluir exentos / Uso público", value=False)

    # Filtro: solo lotes (aplica a urbano)
    sel_lote = st.checkbox("Solo lotes (tarifa 33‰)", value=False)

    st.divider()
    st.markdown("""
    <div style='font-size:0.75rem; opacity:0.8;'>
    📂 Fuentes:<br>
    · <i>Gestor Catastral 2026 (urbano)</i><br>
    · <i>Base anterior Municipio 2025</i><br><br>
    📋 Normativa:<br>
    · Ley 44/1990 Art. 6<br>
    · Ley 1995/2019<br>
    · Acuerdo Municipal 022/2024
    </div>
    """, unsafe_allow_html=True)


# ─── APLICAR FILTROS ─────────────────────────────────────────────────────────
df = df_all.copy()
if sel_sector == "Solo Urbano":
    df = df[df["SECTOR"] == "URBANO"]
elif sel_sector == "Solo Rural":
    df = df[df["SECTOR"] == "RURAL"]
if sel_sector != "Solo Rural":
    if sel_gest == "Con datos Gestor 2026":
        df = df[df["_TIENE_GESTOR"]]
    elif sel_gest == "Sin datos Gestor 2026":
        df = df[~df["_TIENE_GESTOR"]]
if sel_limite == "Aplica límite Ley 44":
    df = df[df["_APLICA_LIM"]]
elif sel_limite == "No aplica límite":
    df = df[~df["_APLICA_LIM"]]
if sel_dest:
    df = df[df["DEST_2025"].isin(sel_dest)]
if sel_exento:
    df = df[~df["_EXENTO"]]
if sel_lote:
    df = df[df["TARIFA_2026"].fillna(0) == 33]


# ─── HEADER ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="header-main">
  <h1>🏘️ Análisis Predial — Municipio de Salgar</h1>
  <p>Actualización catastral 2026 · Gestor Catastral (urbano) · Acuerdo Municipal 022/2024 · Ley 44 de 1990</p>
</div>
""", unsafe_allow_html=True)

n_filtrado = len(df)
n_total    = len(df_all)
n_urb      = int((df["SECTOR"] == "URBANO").sum())
n_rur      = int((df["SECTOR"] == "RURAL").sum())
if n_filtrado < n_total:
    st.info(
        f"Mostrando **{n_filtrado:,}** de **{n_total:,}** predios según los filtros aplicados "
        f"(**{n_urb:,} urbanos** · **{n_rur:,} rurales**)."
    )
else:
    st.info(f"**{n_total:,}** predios en total: **{n_urb:,} urbanos** + **{n_rur:,} rurales**")


# ─── KPIs ────────────────────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">📊 Indicadores Clave</div>', unsafe_allow_html=True)

liq25   = df["LIQ_2025"].fillna(0).sum()
liq26b  = df["LIQ_2026_BRUTA"].fillna(0).sum()
liq26f  = df["LIQ_2026_FINAL"].fillna(0).sum()
ahorro  = df["AHORRO_CONTRIB"].fillna(0).sum()
n_lim   = int(df["_APLICA_LIM"].sum())
n_gest  = int(df["_TIENE_GESTOR"].sum())
n_nev   = len(df_nev)
var_rec = (liq26f - liq25) / liq25 * 100 if liq25 > 0 else 0

sector_label = {
    "Todos (Urbano + Rural)": "Urbano + Rural",
    "Solo Urbano": "Solo Urbano",
    "Solo Rural": "Solo Rural",
}[sel_sector]


def fmt_cop(v):
    if abs(v) >= 1e9:  return f"${v/1e9:.2f}B"
    if abs(v) >= 1e6:  return f"${v/1e6:.1f}M"
    return f"${v:,.0f}"


cols_kpi = st.columns(6)
kpis = [
    (n_filtrado,          "Total Predios",        sector_label,              "oscuro"),
    (fmt_cop(liq25),      "Recaudo 2025",         "Base anterior municipio",  ""),
    (fmt_cop(liq26f),     "Recaudo 2026 Proy.",   "Con límite Ley 44",       "verde"),
    (f"{var_rec:+.1f}%",  "Variación Recaudo",    "2025 → 2026",             "verde" if var_rec >= 0 else "rojo"),
    (n_lim,               "Predios Limitados",    "Aplica tope Ley 44",      "ambar"),
    (fmt_cop(ahorro),     "Ahorro Contrib.",      "Por tope Ley 44",         "rojo"),
]
for col, (val, lab, sub, color) in zip(cols_kpi, kpis):
    with col:
        st.markdown(f"""
        <div class="kpi-card {color}">
          <div class="kpi-valor">{val}</div>
          <div class="kpi-label">{lab}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─── FILA 1: COMPARATIVO FISCAL ──────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">📈 Comparativo Fiscal 2025 vs 2026</div>',
            unsafe_allow_html=True)
c1, c2, c3 = st.columns([2, 2, 1.5])

with c1:
    fig = go.Figure(go.Bar(
        x=["Recaudo 2025", "2026 Bruto (sin tope)", "2026 Final (con tope)"],
        y=[liq25, liq26b, liq26f],
        marker_color=[AZUL_MED, ROJO, VERDE],
        text=[fmt_cop(v) for v in [liq25, liq26b, liq26f]],
        textposition="outside", textfont=dict(size=11, color=AZUL_OSC),
    ))
    fig.update_layout(
        title=dict(text="Comparativo Recaudo ($)", font=dict(size=13, color=AZUL_OSC)),
        yaxis=dict(tickformat="$,.0f", showgrid=True, gridcolor="#eee"),
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=50, b=20, l=10, r=10), height=320, showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)

with c2:
    liq_urb = float(df[df["SECTOR"] == "URBANO"]["LIQ_2026_FINAL"].fillna(0).sum())
    liq_rur = float(df[df["SECTOR"] == "RURAL"]["LIQ_2026_FINAL"].fillna(0).sum())
    fig2 = go.Figure(go.Pie(
        labels=["Urbano", "Rural"],
        values=[liq_urb, liq_rur],
        hole=0.55,
        marker_colors=[AZUL_MED, VERDE],
        textinfo="label+percent", textfont=dict(size=10),
        hovertemplate="%{label}: %{value:$,.0f}<extra></extra>",
    ))
    fig2.update_layout(
        title=dict(text="Recaudo 2026 por Sector", font=dict(size=13, color=AZUL_OSC)),
        showlegend=False, margin=dict(t=50, b=0, l=0, r=0), height=320,
        paper_bgcolor="white",
        annotations=[dict(text=f"<b>{n_filtrado:,}</b><br>predios",
                          x=0.5, y=0.5, font_size=13, showarrow=False,
                          font_color=AZUL_OSC)],
    )
    st.plotly_chart(fig2, use_container_width=True)

with c3:
    n_exen   = int(df["_EXENTO"].sum())
    n_no_lim = int((~df["_APLICA_LIM"] & ~df["_EXENTO"]).sum())
    fig3 = go.Figure(go.Pie(
        labels=["Aplica tope Ley 44", "No supera tope", "Exentos"],
        values=[n_lim, n_no_lim, n_exen],
        hole=0.5,
        marker_colors=[ROJO, VERDE, AZUL_CLAR],
        textinfo="percent", textfont=dict(size=10),
        hovertemplate="%{label}: %{value:,}<extra></extra>",
    ))
    fig3.update_layout(
        title=dict(text="Aplicación Límite Ley 44", font=dict(size=13, color=AZUL_OSC)),
        legend=dict(orientation="v", font=dict(size=9)),
        margin=dict(t=50, b=0, l=0, r=10), height=320,
        paper_bgcolor="white",
    )
    st.plotly_chart(fig3, use_container_width=True)


# ─── FILA 2: AVALÚOS Y RECAUDO POR DESTINO ───────────────────────────────────
st.markdown('<div class="seccion-titulo">📉 Distribución de Avalúos y Recaudo por Destino</div>',
            unsafe_allow_html=True)
c4, c5 = st.columns(2)

with c4:
    df_var_urb = df[
        (df["SECTOR"] == "URBANO") &
        df["VAR_AVALUO_PCT"].notna() &
        (df["VAR_AVALUO_PCT"].abs() < 2000)
    ]
    if len(df_var_urb) > 0:
        fig4 = px.histogram(
            df_var_urb, x="VAR_AVALUO_PCT", nbins=50,
            color_discrete_sequence=[AZUL_MED],
            labels={"VAR_AVALUO_PCT": "Variación Avalúo (%)"},
            title="Variación de Avalúos Urbanos (Gestor 2026 vs Base 2025)",
        )
        fig4.add_vline(x=0,   line_dash="dash", line_color=AZUL_OSC, annotation_text="0%")
        fig4.add_vline(x=100, line_dash="dot",  line_color=ROJO,     annotation_text="100%")
        fig4.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            margin=dict(t=50, b=20, l=10, r=10), height=310,
            yaxis_title="N° predios", bargap=0.05,
        )
        st.plotly_chart(fig4, use_container_width=True)
    else:
        st.info(
            "No hay datos de variación de avalúo para la selección actual. "
            "Los predios rurales conservan el avalúo 2025 (sin actualización del Gestor Catastral)."
        )

with c5:
    df_dest = (
        df.groupby("DEST_2025")
        .agg(predios=("FICHA", "count"), liq26=("LIQ_2026_FINAL", "sum"))
        .reset_index()
        .sort_values("liq26", ascending=True)
        .tail(12)
    )
    fig5 = go.Figure(go.Bar(
        x=df_dest["liq26"], y=df_dest["DEST_2025"],
        orientation="h", marker_color=AZUL_MED,
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


# ─── SCATTER AVALÚO vs LIQUIDACIÓN (predios urbanos con datos Gestor) ────────
st.markdown('<div class="seccion-titulo">🔍 Relación Avalúo Gestor 2026 vs Liquidación (Urbano)</div>',
            unsafe_allow_html=True)

df_sc = df[
    (df["SECTOR"] == "URBANO") &
    df["AVALUO_TOTAL_2026"].notna() &
    df["LIQ_2026_FINAL"].notna()
].copy()

if len(df_sc) > 0:
    df_sc["_cat"] = df_sc["_APLICA_LIM"].map(
        {True: "Aplica tope Ley 44", False: "No aplica tope"})
    fig6 = px.scatter(
        df_sc.sample(min(len(df_sc), 1500)),
        x="AVALUO_TOTAL_2026", y="LIQ_2026_FINAL",
        color="_cat",
        color_discrete_map={"Aplica tope Ley 44": ROJO, "No aplica tope": VERDE},
        hover_data={"FICHA": True, "TERCERO": True, "VAR_AVALUO_PCT": True, "_cat": False},
        labels={
            "AVALUO_TOTAL_2026": "Avalúo Total 2026 - Gestor ($)",
            "LIQ_2026_FINAL":    "Liquidación 2026 Final ($)",
            "_cat": "Estado",
        },
        title="Avalúo Gestor 2026 vs Liquidación Final — Predios Urbanos (muestra hasta 1,500)",
        opacity=0.65,
    )
    fig6.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        height=380, margin=dict(t=50, b=20, l=10, r=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
    )
    fig6.update_traces(marker=dict(size=6))
    st.plotly_chart(fig6, use_container_width=True)
else:
    st.info("No hay predios urbanos con avalúo Gestor 2026 en la selección actual.")


# ─── PREDIOS NUEVOS DEL GESTOR ────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">🆕 Predios Nuevos del Gestor Catastral 2026</div>',
            unsafe_allow_html=True)
st.markdown(
    f"**{n_nev:,}** predios que aparecen en la base del Gestor pero no estaban en la base anterior del municipio."
)

cols_nev = [c for c in ["FICHA IGAC", "DOCUMENTO", "NOMBRE PROPIETARIO", "DIRECCIÓN",
                         "DEST GESTOR", "USO/DESTINO (Sec5)", "AVALÚO TOTAL 2026",
                         "TARIFA 2026 ‰", "LIQ 2026 ESTIMADA $", "OBSERVACIÓN"]
            if c in df_nev.columns]
if cols_nev:
    cfg_nev = {}
    for c in ["AVALÚO TOTAL 2026", "LIQ 2026 ESTIMADA $"]:
        if c in df_nev.columns:
            cfg_nev[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
    st.dataframe(df_nev[cols_nev].reset_index(drop=True),
                 use_container_width=True, height=280, column_config=cfg_nev)


# ─── TABLA DETALLE ────────────────────────────────────────────────────────────
st.markdown('<div class="seccion-titulo">📋 Detalle de Predios (Urbano + Rural)</div>',
            unsafe_allow_html=True)

cols_vis = [
    "SECTOR", "FICHA", "TERCERO", "DIRECCION", "DEST_2025",
    "AVALUO_2025", "LIQ_2025",
    "AVALUO_TOTAL_2026", "VAR_AVALUO_PCT",
    "TARIFA_2026", "DESC_TARIFA_2026",
    "LIQ_2026_BRUTA", "LIQ_2026_FINAL",
    "APLICA_LIMITE", "AHORRO_CONTRIB",
    "EN_BASE_GESTOR", "NOVEDADES",
]
cols_dis = [c for c in cols_vis if c in df.columns]
df_dis = df[cols_dis].copy()

rename_vista = {
    "SECTOR":           "Sector",
    "TERCERO":          "Propietario",
    "DIRECCION":        "Dirección",
    "DEST_2025":        "Destino 2025",
    "AVALUO_2025":      "Avalúo 2025",
    "LIQ_2025":         "Liq. 2025",
    "AVALUO_TOTAL_2026": "Avalúo 2026",
    "VAR_AVALUO_PCT":   "Var. %",
    "TARIFA_2026":      "Tarifa 2026 ‰",
    "DESC_TARIFA_2026": "Tipo Tarifa",
    "LIQ_2026_BRUTA":   "Liq. 2026 bruta",
    "LIQ_2026_FINAL":   "Liq. 2026 final",
    "APLICA_LIMITE":    "Límite Ley44",
    "AHORRO_CONTRIB":   "Ahorro",
    "EN_BASE_GESTOR":   "En Gestor 2026",
    "NOVEDADES":        "Novedades",
}
df_dis.rename(columns={k: v for k, v in rename_vista.items() if k in df_dis.columns}, inplace=True)

txt_busq = st.text_input("🔎 Buscar por nombre, ficha o dirección:", placeholder="Escriba para filtrar...")
if txt_busq:
    mask = df_dis.apply(
        lambda col: col.astype(str).str.contains(txt_busq, case=False, na=False)
    ).any(axis=1)
    df_dis = df_dis[mask]

st.markdown(f"**{len(df_dis):,}** registros · use los filtros del panel izquierdo para refinar")

col_cfg = {}
for c in ["Avalúo 2025", "Liq. 2025", "Avalúo 2026", "Liq. 2026 bruta", "Liq. 2026 final", "Ahorro"]:
    if c in df_dis.columns:
        col_cfg[c] = st.column_config.NumberColumn(c, format="$ %,.0f")
if "Var. %" in df_dis.columns:
    col_cfg["Var. %"] = st.column_config.NumberColumn("Var. %", format="%.2f %%")
if "Tarifa 2026 ‰" in df_dis.columns:
    col_cfg["Tarifa 2026 ‰"] = st.column_config.NumberColumn("Tarifa 2026 ‰", format="%.0f ‰")

st.dataframe(df_dis.reset_index(drop=True), use_container_width=True, height=420,
             column_config=col_cfg)

csv_bytes = df_dis.to_csv(index=False).encode("utf-8-sig")
st.download_button("⬇️  Descargar tabla filtrada (.csv)", data=csv_bytes,
                   file_name="predios_salgar_2026.csv", mime="text/csv")


# ─── NOTA LEGAL ───────────────────────────────────────────────────────────────
st.markdown("""
<div class="nota-legal">
⚖️ <strong>Nota legal:</strong> Límite aplicado según Art. 18 Acuerdo Municipal 022/2024:
LIQ_2025 × (1 + IPC_2025 + 8pp). No aplica el límite para lotes ni predios tipo Anexo/Industrial.
Tarifas según Acuerdo 022/2024 Art. 16. Base del Gestor Catastral 2026 (solo predios urbanos).
Los predios rurales conservan el avalúo 2025 de la base anterior del municipio (Gestor Catastral
no incluye predios rurales en esta actualización).
Verificar con el Estatuto Tributario Municipal vigente antes de liquidar oficialmente.
</div>
""", unsafe_allow_html=True)

st.markdown(
    "<br><center style='color:#aaa; font-size:0.78rem;'>"
    "Municipio de Salgar · Predial 2026 · Acuerdo 022/2024 · Ley 44/1990"
    "</center>",
    unsafe_allow_html=True,
)


