"""
Dashboard: Segmento EMPRESAS — Projeto Metaday 2025
Recriação fiel do Power BI em Streamlit + Plotly
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, date

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Segmento: EMPRESAS",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# PALETA DE CORES (fiel ao Power BI)
# ─────────────────────────────────────────────
COR_GOLD   = "#B8972E"
COR_NAVY   = "#1F3864"
COR_CINZA  = "#A0A0A0"
COR_BG     = "#F0F2F5"
COR_TEXTO  = "#1A1A2E"
TREEMAP_COLORS = px.colors.qualitative.Bold

# ─────────────────────────────────────────────
# CSS GLOBAL
# ─────────────────────────────────────────────
st.markdown(f"""
<style>
/* Fundo geral */
.stApp {{ background-color: {COR_BG}; }}

/* Título principal */
h1.titulo-principal {{
    font-family: 'Georgia', serif;
    font-size: 2.6rem;
    font-weight: 800;
    color: {COR_TEXTO};
    text-align: center;
    letter-spacing: 1px;
    margin-bottom: 0.2rem;
}}

/* Card KPI */
.kpi-card {{
    background: white;
    border-radius: 12px;
    padding: 18px 24px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-top: 4px solid {COR_GOLD};
}}
.kpi-label {{
    font-size: 0.85rem;
    font-style: italic;
    color: #666;
    font-weight: 600;
    letter-spacing: 0.5px;
}}
.kpi-value {{
    font-size: 2.4rem;
    font-weight: 900;
    color: {COR_TEXTO};
    line-height: 1.1;
}}

/* Título de seção */
.section-title {{
    font-family: 'Georgia', serif;
    font-size: 1.1rem;
    font-style: italic;
    font-weight: 700;
    color: {COR_TEXTO};
    margin: 12px 0 4px 0;
}}

/* Sidebar */
section[data-testid="stSidebar"] {{
    background-color: #1A2744;
}}
section[data-testid="stSidebar"] * {{
    color: white !important;
}}
section[data-testid="stSidebar"] .stMultiSelect > div > div {{
    background-color: #2a3a5c !important;
}}

/* Separador */
hr.divisor {{
    border: none;
    border-top: 1px solid #ddd;
    margin: 8px 0 16px 0;
}}

/* Tabela IBGE */
.ibge-table {{
    font-size: 0.82rem;
}}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CARREGAMENTO DE DADOS
# ─────────────────────────────────────────────
@st.cache_data
def carregar_dados(arquivo):
    """Carrega e une todas as abas do Excel, aplicando as transformações DAX."""
    xls = pd.ExcelFile(arquivo)

    # --- Tabelas principais ---
    empresas  = pd.read_excel(xls, sheet_name="Empresas")
    ramo      = pd.read_excel(xls, sheet_name="Ramo")        if "Ramo"         in xls.sheet_names else pd.DataFrame()
    sub_ramos = pd.read_excel(xls, sheet_name="Sub-Ramos")   if "Sub-Ramos"    in xls.sheet_names else pd.DataFrame()
    ram_princ = pd.read_excel(xls, sheet_name="Ramo Principal") if "Ramo Principal" in xls.sheet_names else pd.DataFrame()
    bairros   = pd.read_excel(xls, sheet_name="Bairros")     if "Bairros"      in xls.sheet_names else pd.DataFrame()
    distrito  = pd.read_excel(xls, sheet_name="Distrito")    if "Distrito"     in xls.sheet_names else pd.DataFrame()
    porte_tb  = pd.read_excel(xls, sheet_name="Porte")       if "Porte"        in xls.sheet_names else pd.DataFrame()
    horario   = pd.read_excel(xls, sheet_name="Horario")     if "Horario"      in xls.sheet_names else pd.DataFrame()

    # --- Normaliza nomes de colunas ---
    empresas.columns  = empresas.columns.str.strip()

    # ── Coluna de presença digital (site/rede social) ──
    # Aceita várias possíveis nomenclaturas
    for col in ["Site", "Presença Digital", "Presenca Digital", "PresencaDigital", "site"]:
        if col in empresas.columns:
            empresas["Presenca_Digital"] = empresas[col].apply(
                lambda x: "SIM" if str(x).strip().upper() not in ["", "NAN", "NAO", "NÃO", "N", "NO", "FALSE", "0"] else "NÃO"
            )
            break
    if "Presenca_Digital" not in empresas.columns:
        empresas["Presenca_Digital"] = "NÃO"

    # ── Faixa etária da empresa (DAX equivalente) ──
    if "Data Abertura" in empresas.columns:
        empresas["Data Abertura"] = pd.to_datetime(empresas["Data Abertura"], errors="coerce")
        hoje = pd.Timestamp.today()
        empresas["Anos_Existencia"] = ((hoje - empresas["Data Abertura"]).dt.days / 365.25)
        def faixa_idade(anos):
            if pd.isna(anos):    return "Sem dados"
            elif anos <= 2:      return "0 a 2 anos"
            elif anos <= 5:      return "3 a 5 anos"
            elif anos <= 10:     return "6 a 10 anos"
            elif anos <= 20:     return "11 a 20 anos"
            else:                return "Mais de 20 anos"
        empresas["Faixa_Idade"] = empresas["Anos_Existencia"].apply(faixa_idade)
    else:
        empresas["Faixa_Idade"] = "Sem dados"

    # ── Faixa de distância (DAX equivalente) ──
    dist_col = next((c for c in ["Distancia (km)", "Distancia", "Distância (km)", "distancia_km"] if c in empresas.columns), None)
    if dist_col:
        empresas["Dist_km"] = pd.to_numeric(empresas[dist_col], errors="coerce")
        def faixa_dist(d):
            if pd.isna(d):    return "Sem dados"
            elif d <= 2:      return "0 a 2km"
            elif d <= 4:      return "3 a 4km"
            elif d <= 7:      return "5 a 7km"
            else:             return "Mais de 7km"
        empresas["Faixa_Dist"] = empresas["Dist_km"].apply(faixa_dist)
    else:
        empresas["Faixa_Dist"] = "Sem dados"

    # ── Faixa de tempo (DAX equivalente) ──
    tempo_col = next((c for c in ["Distância desde a FATEC", "Tempo", "tempo_min", "Tempo (min)"] if c in empresas.columns), None)
    if tempo_col:
        empresas["Tempo_min"] = pd.to_numeric(empresas[tempo_col], errors="coerce")
        def faixa_tempo(t):
            if pd.isna(t):    return "Sem dados"
            elif t <= 5:      return "0 a 5 minutos"
            elif t <= 10:     return "6 a 10 minutos"
            elif t <= 20:     return "11 a 20 minutos"
            else:             return "Mais de 20 minutos"
        empresas["Faixa_Tempo"] = empresas["Tempo_min"].apply(faixa_tempo)
    else:
        empresas["Faixa_Tempo"] = "Sem dados"

    # ── Tem CNPJ? ──
    if "CNPJ" in empresas.columns:
        empresas["Tem_CNPJ"] = empresas["CNPJ"].apply(
            lambda x: "SIM" if str(x).strip() not in ["", "nan", "NaN"] else "NÃO"
        )

    # ── Garante colunas essenciais com fallback ──
    for col, default in [
        ("Porte", "Não informado"),
        ("Distrito", "Não informado"),
        ("Bairro", "Não informado"),
        ("Expediente", "HC"),
    ]:
        if col not in empresas.columns:
            empresas[col] = default

    # ── Joins com tabelas de ramo ──
    ramo_col_emp = next((c for c in ["RAMO_ID", "Ramo_ID", "ramo_id"] if c in empresas.columns), None)
    if not ramo.empty and ramo_col_emp:
        ramo_id_col = next((c for c in ["RAMO_ID", "Ramo_ID"] if c in ramo.columns), ramo.columns[0])
        ramo_nome   = next((c for c in ["ESPECIALIZAÇÃO / TIPO DE NEGÓ...", "Especializacao", "Ramo"] if c in ramo.columns), ramo.columns[-1])
        ramo_map = ramo.set_index(ramo_id_col)[ramo_nome].to_dict()
        empresas["Ramo_Nome"] = empresas[ramo_col_emp].map(ramo_map)

        ramo_princ_id = next((c for c in ["RAMO_PRINCIPAL_ID"] if c in ramo.columns), None)
        if not ram_princ.empty and ramo_princ_id:
            rp_id_col   = next((c for c in ["RAMO_PRINCIPAL_ID"] if c in ram_princ.columns), ram_princ.columns[0])
            rp_nome_col = next((c for c in ["RAMO PRINCIPAL", "RamoPrincipal"] if c in ram_princ.columns), ram_princ.columns[1])
            rp_map = ram_princ.set_index(rp_id_col)[rp_nome_col].to_dict()
            empresas["Ramo_Principal"] = ramo[ramo_princ_id].map(rp_map).reindex(
                empresas[ramo_col_emp].map(
                    ramo.set_index(ramo_id_col)[ramo_princ_id].to_dict()
                )
            ).values

        sub_id_col_ramo = next((c for c in ["SUB-RAMO_ID", "SubRamo_ID"] if c in ramo.columns), None)
        if not sub_ramos.empty and sub_id_col_ramo:
            sr_id  = next((c for c in ["SUB-RAMO_ID", "SubRamo_ID"] if c in sub_ramos.columns), sub_ramos.columns[0])
            sr_nom = next((c for c in ["SUB-RAMO", "SubRamo"] if c in sub_ramos.columns), sub_ramos.columns[-1])
            sr_map = sub_ramos.set_index(sr_id)[sr_nom].to_dict()
            empresas["Sub_Ramo"] = ramo[sub_id_col_ramo].map(sr_map).reindex(
                empresas[ramo_col_emp].map(
                    ramo.set_index(ramo_id_col)[sub_id_col_ramo].to_dict()
                )
            ).values

    # Fallbacks para colunas de ramo
    for col, default in [("Ramo_Nome", "Não informado"), ("Ramo_Principal", "Não informado"), ("Sub_Ramo", "Não informado")]:
        if col not in empresas.columns:
            empresas[col] = default

    # ── Latitude / Longitude ──
    lat_col = next((c for c in ["Latitude", "LAT", "lat"] if c in empresas.columns), None)
    lon_col = next((c for c in ["Longitude", "LON", "lon", "LNG", "lng"] if c in empresas.columns), None)
    if lat_col: empresas["Lat"] = pd.to_numeric(empresas[lat_col], errors="coerce")
    if lon_col: empresas["Lon"] = pd.to_numeric(empresas[lon_col], errors="coerce")

    return empresas, distrito


def kpi_card(label, value):
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
    </div>"""


# ─────────────────────────────────────────────
# HEADER PADRÃO
# ─────────────────────────────────────────────
def render_header():
    st.markdown('<h1 class="titulo-principal">Segmento: EMPRESAS</h1>', unsafe_allow_html=True)
    st.markdown('<hr class="divisor">', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 1 — VISÃO GERAL (KPIs + Mapa de calor por distrito)
# ═══════════════════════════════════════════════════════════════
def pagina_visao_geral(df, df_distrito):
    render_header()

    total_emp   = len(df)
    tot_dist    = df["Distrito"].nunique()
    tot_espec   = df["Ramo_Nome"].nunique() if "Ramo_Nome" in df.columns else df["Sub_Ramo"].nunique()

    # KPIs
    c1, c2, c3 = st.columns(3)
    with c1: st.markdown(kpi_card("Empresas Catalogadas", f"{total_emp:,}".replace(",", ".")), unsafe_allow_html=True)
    with c2: st.markdown(kpi_card("Distritos Considerados", tot_dist), unsafe_allow_html=True)
    with c3: st.markdown(kpi_card("Especializações Identificadas", tot_espec), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    col_esq, col_dir = st.columns([1, 2.5])

    with col_esq:
        # Presença Digital — Pizza
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        pres = df["Presenca_Digital"].value_counts().reset_index()
        pres.columns = ["Status", "Qtde"]
        fig_pres = px.pie(pres, names="Status", values="Qtde",
                          color="Status",
                          color_discrete_map={"SIM": COR_GOLD, "NÃO": "#D0D0D0"},
                          hole=0.35)
        fig_pres.update_traces(textposition="outside", textinfo="label+value+percent",
                                textfont_size=11)
        fig_pres.update_layout(showlegend=False, margin=dict(t=10, b=10, l=10, r=10),
                               height=230, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_pres, use_container_width=True)

        # Porte — Barras horizontais
        st.markdown('<div class="section-title">Porte das Empresas</div>', unsafe_allow_html=True)
        porte_order = ["Pequeno", "Médio", "Grande"]
        porte = df["Porte"].value_counts().reindex(porte_order).dropna().reset_index()
        porte.columns = ["Porte", "Qtde"]
        colors_porte = {
            "Pequeno": COR_GOLD,
            "Médio":   COR_NAVY,
            "Grande":  COR_CINZA,
        }
        fig_porte = px.bar(porte, y="Porte", x="Qtde", orientation="h",
                           color="Porte",
                           color_discrete_map=colors_porte,
                           text="Qtde")
        fig_porte.update_traces(textposition="inside", textfont_size=13, textfont_color="white")
        fig_porte.update_layout(
            showlegend=False, height=200,
            margin=dict(t=5, b=5, l=10, r=10),
            xaxis=dict(visible=False), yaxis=dict(title=""),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig_porte, use_container_width=True)

    with col_dir:
        # Mapa de bolhas por distrito
        st.markdown('<div class="section-title">Distribuição Geográfica por Distrito</div>', unsafe_allow_html=True)
        if "Lat" in df.columns and "Lon" in df.columns and df["Lat"].notna().sum() > 10:
            df_map = df.dropna(subset=["Lat", "Lon"])
            fig_map = px.scatter_mapbox(
                df_map, lat="Lat", lon="Lon",
                color="Distrito",
                size_max=12,
                hover_name=df_map.columns[0] if df_map.shape[1] > 0 else None,
                hover_data={"Lat": False, "Lon": False, "Porte": True, "Distrito": True},
                zoom=12, height=430,
                mapbox_style="open-street-map",
                color_discrete_sequence=px.colors.qualitative.Bold,
            )
            fig_map.update_layout(margin=dict(t=0, b=0, l=0, r=0), showlegend=False,
                                  paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_map, use_container_width=True)
        else:
            # Fallback: bolhas por quantidade de empresas por distrito
            dist_count = df.groupby("Distrito").size().reset_index(name="Qtde")
            dist_count = dist_count.sort_values("Qtde", ascending=False)
            fig_bar = px.bar(dist_count, x="Distrito", y="Qtde",
                             text="Qtde",
                             color="Qtde",
                             color_continuous_scale=[[0, "#D4B44A"], [1, COR_NAVY]])
            fig_bar.update_traces(textposition="outside")
            fig_bar.update_layout(
                height=430, showlegend=False,
                xaxis_title="", yaxis_title="Empresas",
                coloraxis_showscale=False,
                margin=dict(t=10, b=10, l=10, r=10),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_bar, use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 2 — DISTRITOS (Treemap + Tempo de Existência + Ramo)
# ═══════════════════════════════════════════════════════════════
def pagina_distritos(df):
    render_header()

    col_esq, col_dir = st.columns([1, 2.5])

    with col_esq:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        pres = df["Presenca_Digital"].value_counts().reset_index()
        pres.columns = ["Status", "Qtde"]
        fig_pres = px.pie(pres, names="Status", values="Qtde",
                          color="Status",
                          color_discrete_map={"SIM": COR_GOLD, "NÃO": "#D0D0D0"},
                          hole=0.35)
        fig_pres.update_traces(textposition="outside", textinfo="label+value+percent", textfont_size=10)
        fig_pres.update_layout(showlegend=False, margin=dict(t=5, b=5, l=5, r=5), height=200,
                               paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_pres, use_container_width=True)

    with col_dir:
        # Treemap por distrito
        st.markdown('<div class="section-title">Empresas por Distrito</div>', unsafe_allow_html=True)
        dist_count = df.groupby("Distrito").size().reset_index(name="Qtde")
        fig_tree = px.treemap(dist_count, path=["Distrito"], values="Qtde",
                              color="Qtde",
                              color_continuous_scale=[
                                  [0, "#1a237e"], [0.15, "#283593"], [0.3, "#5c6bc0"],
                                  [0.45, "#00897b"], [0.6, "#43a047"], [0.75, "#e53935"],
                                  [0.9, "#8e24aa"], [1.0, "#00acc1"]
                              ])
        fig_tree.update_traces(
            textinfo="label+value",
            textfont_size=13,
            textposition="bottom left",
        )
        fig_tree.update_layout(
            height=290, margin=dict(t=5, b=5, l=5, r=5),
            coloraxis_showscale=False,
            paper_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig_tree, use_container_width=True)

    # Linha inferior: Tempo Existência + Quantidade por Ramo
    c1, c2 = st.columns(2)

    with c1:
        st.markdown('<div class="section-title">Tempo de Existência</div>', unsafe_allow_html=True)
        ordem = ["Mais de 20 anos", "0 a 2 anos", "11 a 20 anos", "6 a 10 anos", "3 a 5 anos", "Sem dados"]
        faixa = df["Faixa_Idade"].value_counts().reindex(ordem).dropna().reset_index()
        faixa.columns = ["Faixa", "Qtde"]
        fig_faixa = px.bar(faixa, y="Faixa", x="Qtde", orientation="h",
                           text="Qtde",
                           color="Qtde",
                           color_continuous_scale=[[0, "#90CAF9"], [0.5, "#1565C0"], [1, COR_NAVY]])
        fig_faixa.update_traces(textposition="outside", textfont_size=12)
        fig_faixa.update_layout(
            height=280, showlegend=False,
            xaxis=dict(visible=False), yaxis=dict(title=""),
            coloraxis_showscale=False,
            margin=dict(t=5, b=5, l=5, r=5),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig_faixa, use_container_width=True)

    with c2:
        st.markdown('<div class="section-title">Quantidade por Ramo</div>', unsafe_allow_html=True)
        ramo_count = df["Ramo_Principal"].value_counts().reset_index()
        ramo_count.columns = ["Ramo", "Qtde"]
        ramo_count = ramo_count[ramo_count["Ramo"] != "Não informado"].head(10)
        fig_ramo = px.bar(ramo_count, x="Ramo", y="Qtde",
                          text="Qtde",
                          color_discrete_sequence=[COR_GOLD])
        fig_ramo.update_traces(textposition="outside", textfont_size=11)
        fig_ramo.update_layout(
            height=280, showlegend=False,
            xaxis=dict(title="", tickangle=-30, tickfont_size=10),
            yaxis=dict(visible=False),
            margin=dict(t=20, b=5, l=5, r=5),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig_ramo, use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 3 — FATEC (Tempo/Distância + Sub-Ramos)
# ═══════════════════════════════════════════════════════════════
def pagina_fatec(df):
    render_header()

    # Linha superior: Tempo e Distância
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">Tempo da FATEC até a Empresa</div>', unsafe_allow_html=True)
        ordem_tempo = ["11 a 20 minutos", "6 a 10 minutos", "0 a 5 minutos", "Mais de 20 minutos", "Sem dados"]
        tempo_c = df["Faixa_Tempo"].value_counts().reindex(ordem_tempo).dropna().reset_index()
        tempo_c.columns = ["Faixa", "Qtde"]
        fig_tempo = px.bar(tempo_c, y="Faixa", x="Qtde", orientation="h",
                           text="Qtde",
                           color_discrete_sequence=[COR_GOLD])
        fig_tempo.update_traces(textposition="outside", textfont_size=12)
        fig_tempo.update_layout(
            height=230, showlegend=False,
            xaxis=dict(visible=False), yaxis=dict(title=""),
            margin=dict(t=5, b=5, l=5, r=5),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig_tempo, use_container_width=True)

    with c2:
        st.markdown('<div class="section-title">Distância da FATEC até a Empresa</div>', unsafe_allow_html=True)
        ordem_dist = ["3 a 4km", "0 a 2km", "5 a 7km", "Mais de 7km", "Sem dados"]
        dist_c = df["Faixa_Dist"].value_counts().reindex(ordem_dist).dropna().reset_index()
        dist_c.columns = ["Faixa", "Qtde"]
        fig_dist = px.bar(dist_c, x="Faixa", y="Qtde",
                          text="Qtde",
                          color="Qtde",
                          color_continuous_scale=[[0, "#E0D0A0"], [1, COR_GOLD]])
        fig_dist.update_traces(textposition="outside", textfont_size=12)
        fig_dist.update_layout(
            height=230, showlegend=False,
            xaxis=dict(title=""), yaxis=dict(visible=False),
            coloraxis_showscale=False,
            margin=dict(t=20, b=5, l=5, r=5),
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
        )
        st.plotly_chart(fig_dist, use_container_width=True)

    # Linha inferior: Filtro Ramo + Treemap Sub-Ramos
    col_filtro, col_tree = st.columns([1, 3])

    with col_filtro:
        st.markdown('<div class="section-title">Ramo de Atividade</div>', unsafe_allow_html=True)
        ramos_disp = sorted(df["Ramo_Principal"].dropna().unique().tolist())
        ramos_disp = [r for r in ramos_disp if r != "Não informado"]
        ramos_sel  = st.multiselect("", ramos_disp, default=[], label_visibility="collapsed")

    with col_tree:
        st.markdown('<div class="section-title">Sub-Divisão dos Ramos de Atividade</div>', unsafe_allow_html=True)
        df_f = df[df["Ramo_Principal"].isin(ramos_sel)] if ramos_sel else df
        sub_count = df_f.groupby(["Ramo_Principal", "Sub_Ramo"]).size().reset_index(name="Qtde")
        sub_count  = sub_count[(sub_count["Sub_Ramo"] != "Não informado") & (sub_count["Qtde"] > 0)]
        if not sub_count.empty:
            fig_sub = px.treemap(sub_count, path=["Ramo_Principal", "Sub_Ramo"],
                                 values="Qtde",
                                 color="Ramo_Principal",
                                 color_discrete_sequence=TREEMAP_COLORS)
            fig_sub.update_traces(textinfo="label+value", textfont_size=11, textposition="bottom left")
            fig_sub.update_layout(
                height=380, margin=dict(t=5, b=5, l=5, r=5),
                paper_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_sub, use_container_width=True)
        else:
            st.info("Selecione um ou mais ramos para ver a sub-divisão.")


# ═══════════════════════════════════════════════════════════════
# PÁGINA 4 — MAPA SEGMENTADO POR RAMO
# ═══════════════════════════════════════════════════════════════
def pagina_mapa_ramos(df):
    render_header()

    col_filtro, col_mapa = st.columns([1, 3])

    with col_filtro:
        st.markdown('<div class="section-title">Segmentação dos Ramos</div>', unsafe_allow_html=True)
        sub_ramos_disp = sorted(df["Sub_Ramo"].dropna().unique().tolist())
        sub_ramos_disp = [s for s in sub_ramos_disp if s != "Não informado"]
        sel_sub = st.multiselect("", sub_ramos_disp, default=[], label_visibility="collapsed",
                                 key="sel_sub_mapa")

    with col_mapa:
        df_f = df[df["Sub_Ramo"].isin(sel_sub)] if sel_sub else df
        if "Lat" in df_f.columns and "Lon" in df_f.columns and df_f["Lat"].notna().sum() > 5:
            df_map = df_f.dropna(subset=["Lat", "Lon"])
            fig_map = px.scatter_mapbox(
                df_map, lat="Lat", lon="Lon",
                color="Sub_Ramo",
                hover_data={"Lat": False, "Lon": False, "Porte": True, "Distrito": True, "Sub_Ramo": True},
                zoom=12, height=550,
                mapbox_style="open-street-map",
                color_discrete_sequence=px.colors.qualitative.Alphabet,
            )
            fig_map.update_layout(
                margin=dict(t=0, b=0, l=0, r=0),
                legend=dict(title="Sub-Ramo", font_size=10, x=1.0),
                paper_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_map, use_container_width=True)
        else:
            # Fallback: barras de sub-ramo
            sub_count = df_f["Sub_Ramo"].value_counts().head(20).reset_index()
            sub_count.columns = ["Sub_Ramo", "Qtde"]
            sub_count = sub_count[sub_count["Sub_Ramo"] != "Não informado"]
            fig_sub = px.bar(sub_count, x="Qtde", y="Sub_Ramo", orientation="h",
                             text="Qtde",
                             color="Sub_Ramo",
                             color_discrete_sequence=px.colors.qualitative.Bold)
            fig_sub.update_traces(textposition="outside", textfont_size=11)
            fig_sub.update_layout(
                height=550, showlegend=False,
                xaxis=dict(visible=False), yaxis=dict(title=""),
                margin=dict(t=5, b=5, l=5, r=5),
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_sub, use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 5 — DISTRITOS POR SEGMENTO (Treemap filtrado)
# ═══════════════════════════════════════════════════════════════
def pagina_distritos_segmento(df):
    render_header()

    col_filtro, col_tree = st.columns([1, 3])

    with col_filtro:
        st.markdown('<div class="section-title">Segmentação dos Ramos</div>', unsafe_allow_html=True)
        sub_ramos_disp = sorted(df["Sub_Ramo"].dropna().unique().tolist())
        sub_ramos_disp = [s for s in sub_ramos_disp if s != "Não informado"]
        sel_sub = st.multiselect("", sub_ramos_disp, default=[], label_visibility="collapsed",
                                 key="sel_sub_dist")

    with col_tree:
        df_f = df[df["Sub_Ramo"].isin(sel_sub)] if sel_sub else df
        st.markdown('<div class="section-title">Distritos que têm os Segmentos selecionados</div>', unsafe_allow_html=True)
        dist_count = df_f.groupby("Distrito").size().reset_index(name="Qtde")
        if not dist_count.empty:
            fig_tree = px.treemap(dist_count, path=["Distrito"], values="Qtde",
                                  color="Qtde",
                                  color_continuous_scale=[
                                      [0, "#1a237e"], [0.15, "#283593"], [0.3, "#5c6bc0"],
                                      [0.45, "#00897b"], [0.6, "#43a047"], [0.75, "#e53935"],
                                      [0.9, "#8e24aa"], [1.0, "#00acc1"]
                                  ])
            fig_tree.update_traces(
                textinfo="label+value", textfont_size=14, textposition="bottom left"
            )
            fig_tree.update_layout(
                height=560, margin=dict(t=5, b=5, l=5, r=5),
                coloraxis_showscale=False,
                paper_bgcolor="rgba(0,0,0,0)"
            )
            st.plotly_chart(fig_tree, use_container_width=True)
        else:
            st.info("Nenhuma empresa encontrada com o filtro atual.")


# ═══════════════════════════════════════════════════════════════
# PÁGINA 6 — DADOS IBGE / CENSO + GRÁFICOS DE PERFIL
# ═══════════════════════════════════════════════════════════════
def pagina_ibge(df, df_distrito):
    render_header()

    # Gráficos de perfil no topo
    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown('<div class="section-title">Porte das Empresas</div>', unsafe_allow_html=True)
        porte_order = ["Pequeno", "Médio", "Grande"]
        porte = df["Porte"].value_counts().reindex(porte_order).dropna().reset_index()
        porte.columns = ["Porte", "Qtde"]
        colors_porte = {"Pequeno": COR_GOLD, "Médio": COR_NAVY, "Grande": COR_CINZA}
        fig_porte = px.bar(porte, x="Porte", y="Qtde", text="Qtde",
                           color="Porte", color_discrete_map=colors_porte)
        fig_porte.update_traces(textposition="outside", textfont_size=12)
        fig_porte.update_layout(height=220, showlegend=False,
                                xaxis_title="Porte", yaxis=dict(visible=False),
                                margin=dict(t=10, b=5, l=5, r=5),
                                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_porte, use_container_width=True)

    with c2:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        pres = df["Presenca_Digital"].value_counts().reset_index()
        pres.columns = ["Status", "Qtde"]
        fig_pres = px.pie(pres, names="Status", values="Qtde",
                          color="Status",
                          color_discrete_map={"SIM": COR_GOLD, "NÃO": "#D0D0D0"},
                          hole=0.3)
        fig_pres.update_traces(textposition="outside", textinfo="label+value+percent", textfont_size=10)
        fig_pres.update_layout(showlegend=False, margin=dict(t=5, b=5, l=5, r=5), height=220,
                               paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_pres, use_container_width=True)

    with c3:
        st.markdown('<div class="section-title">Expediente</div>', unsafe_allow_html=True)
        exp_count = df["Expediente"].value_counts().reset_index()
        exp_count.columns = ["Tipo", "Qtde"]
        fig_exp = px.pie(exp_count, names="Tipo", values="Qtde",
                         color_discrete_sequence=[COR_NAVY, COR_GOLD, COR_CINZA],
                         hole=0.3)
        fig_exp.update_traces(textposition="outside", textinfo="label+percent", textfont_size=10)
        fig_exp.update_layout(showlegend=False, margin=dict(t=5, b=5, l=5, r=5), height=220,
                               paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_exp, use_container_width=True)

    # Tabela IBGE
    st.markdown("---")
    st.markdown(f'<div style="font-size:1.1rem; font-weight:800; color:{COR_TEXTO}; margin-bottom:8px;">📊 Dados IBGE/Censo 2022</div>', unsafe_allow_html=True)

    if not df_distrito.empty:
        # Tenta montar a tabela diretamente da aba Distrito
        df_distrito.columns = df_distrito.columns.str.strip()
        st.dataframe(df_distrito, use_container_width=True, height=340)
    else:
        # Fallback: monta a tabela a partir de Empresas
        emp_por_dist = df.groupby("Distrito").size().reset_index(name="Na Amostra")
        tabela_ibge = emp_por_dist.sort_values("Distrito")
        total_row   = pd.DataFrame([{"Distrito": "Total", "Na Amostra": len(df)}])
        tabela_ibge = pd.concat([tabela_ibge, total_row], ignore_index=True)
        st.dataframe(tabela_ibge, use_container_width=True, height=340)

    st.markdown("""
    <div style="font-size: 0.78rem; color: #666; margin-top: 8px;">
    Os dados acima, exceto pela coluna NA AMOSTRA, foram obtidos cruzando dados da
    <b>PREFEITURA DE SP, SEADE e OBSERVASAMPA</b>.<br>
    Os dados de População Projetada para 2030 e 2050 usam como base o Censo 2010.
    </div>
    """, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SIDEBAR + NAVEGAÇÃO
# ─────────────────────────────────────────────
def render_sidebar(df):
    with st.sidebar:
        st.markdown("""
        <div style="text-align:center; padding: 10px 0 20px 0;">
            <div style="font-size:2.5rem;">🏢</div>
            <div style="font-size:1.1rem; font-weight:800; letter-spacing:1px;">METADAY 2025</div>
            <div style="font-size:0.8rem; opacity:0.7;">Fatec Sebrae · Ciência de Dados</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("---")

        pagina = st.radio(
            "📄 Página",
            options=[
                "1 · Visão Geral",
                "2 · Distritos & Ramos",
                "3 · FATEC & Sub-Ramos",
                "4 · Mapa por Segmento",
                "5 · Distritos por Segmento",
                "6 · Dados IBGE / Perfil",
            ],
            label_visibility="collapsed"
        )

        st.markdown("---")
        st.markdown("**🔍 Filtros Globais**")

        # Filtro de Distrito
        distritos = sorted(df["Distrito"].dropna().unique().tolist())
        distritos = [d for d in distritos if d not in ["Não informado", ""]]
        sel_dist = st.multiselect("Distrito", distritos, default=distritos, key="filtro_distrito")

        # Filtro de Porte
        portes = sorted(df["Porte"].dropna().unique().tolist())
        portes = [p for p in portes if p not in ["Não informado", ""]]
        sel_porte = st.multiselect("Porte", portes, default=portes, key="filtro_porte")

        # Filtro de Presença Digital
        sel_pres = st.multiselect(
            "Presença Digital",
            ["SIM", "NÃO"],
            default=["SIM", "NÃO"],
            key="filtro_pres"
        )

        st.markdown("---")
        st.caption(f"Total filtrado: **{len(df):,}** empresas")

    return pagina, sel_dist, sel_porte, sel_pres


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    # Upload do Excel
    with st.sidebar:
        st.markdown("**📂 Dados**")
        arquivo = st.file_uploader("Carregar planilha Excel", type=["xlsx", "xls"],
                                   label_visibility="collapsed")

    if arquivo is None:
        st.markdown('<h1 class="titulo-principal">Segmento: EMPRESAS</h1>', unsafe_allow_html=True)
        st.markdown('<hr class="divisor">', unsafe_allow_html=True)
        st.info("👈 Faça upload da planilha Excel com os dados para iniciar o dashboard.")
        st.markdown("""
        **Abas esperadas no Excel:**
        - `Empresas` — tabela principal
        - `Ramo`, `Sub-Ramos`, `Ramo Principal` — hierarquia de atividades
        - `Bairros`, `Distrito` — dados geográficos e IBGE
        - `Porte`, `Horario` — tabelas auxiliares
        """)
        return

    df_raw, df_distrito = carregar_dados(arquivo)

    # Aplica filtros globais
    pagina, sel_dist, sel_porte, sel_pres = render_sidebar(df_raw)

    df = df_raw.copy()
    if sel_dist:
        df = df[df["Distrito"].isin(sel_dist)]
    if sel_porte:
        df = df[df["Porte"].isin(sel_porte)]
    if sel_pres:
        df = df[df["Presenca_Digital"].isin(sel_pres)]

    # Roteamento de páginas
    if   pagina.startswith("1"): pagina_visao_geral(df, df_distrito)
    elif pagina.startswith("2"): pagina_distritos(df)
    elif pagina.startswith("3"): pagina_fatec(df)
    elif pagina.startswith("4"): pagina_mapa_ramos(df)
    elif pagina.startswith("5"): pagina_distritos_segmento(df)
    elif pagina.startswith("6"): pagina_ibge(df, df_distrito)


if __name__ == "__main__":
    main()
