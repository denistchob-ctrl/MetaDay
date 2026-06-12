"""
Dashboard: Segmento EMPRESAS — Projeto Metaday 2025
Recriação fiel do Power BI em Streamlit + Plotly
v2 — auto-detecção de abas + logos + leitura automática da pasta
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import os
import base64
import glob
from pathlib import Path

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
# PALETA DE CORES
# ─────────────────────────────────────────────
COR_GOLD  = "#B8972E"
COR_NAVY  = "#1F3864"
COR_CINZA = "#A0A0A0"
COR_BG    = "#F0F2F5"
COR_TEXTO = "#1A1A2E"
TREEMAP_COLORS = px.colors.qualitative.Bold


# ─────────────────────────────────────────────
# UTILITÁRIOS DE IMAGEM
# ─────────────────────────────────────────────
def img_to_base64(path: str) -> str:
    """Converte imagem local para base64 para embed no HTML."""
    try:
        with open(path, "rb") as f:
            data = f.read()
        ext = Path(path).suffix.lower().lstrip(".")
        mime = "png" if ext == "png" else "jpeg"
        return f"data:image/{mime};base64,{base64.b64encode(data).decode()}"
    except Exception:
        return ""


def _logo_path(filename: str) -> str:
    """Procura o logo na pasta do script ou na raiz do projeto."""
    base = Path(__file__).parent
    for candidate in [base / filename, Path(filename), Path("assets") / filename]:
        if candidate.exists():
            return str(candidate)
    return ""


# ─────────────────────────────────────────────
# CSS GLOBAL
# ─────────────────────────────────────────────
st.markdown(f"""
<style>
.stApp {{ background-color: {COR_BG}; }}

h1.titulo-principal {{
    font-family: 'Georgia', serif;
    font-size: 2.4rem;
    font-weight: 800;
    color: {COR_TEXTO};
    text-align: center;
    letter-spacing: 1px;
    margin: 0;
}}

.kpi-card {{
    background: white;
    border-radius: 12px;
    padding: 16px 20px;
    text-align: center;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    border-top: 4px solid {COR_GOLD};
}}
.kpi-label {{
    font-size: 0.82rem;
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

.section-title {{
    font-family: 'Georgia', serif;
    font-size: 1.05rem;
    font-style: italic;
    font-weight: 700;
    color: {COR_TEXTO};
    margin: 10px 0 2px 0;
}}

/* Header row com logos */
.header-row {{
    display: flex;
    align-items: center;
    justify-content: space-between;
    padding: 4px 0 8px 0;
    border-bottom: 2px solid #ddd;
    margin-bottom: 12px;
}}
.header-logo-left  {{ width: 80px; height: 80px; object-fit: contain; }}
.header-logo-right {{ width: 60px; height: 60px; object-fit: contain; border-radius: 50%; }}

section[data-testid="stSidebar"] {{ background-color: #1A2744; }}
section[data-testid="stSidebar"] * {{ color: white !important; }}
section[data-testid="stSidebar"] .stMultiSelect > div > div {{ background-color: #2a3a5c !important; }}

div[data-testid="stFileUploader"] {{ display: none !important; }}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# DETECÇÃO AUTOMÁTICA DO EXCEL
# ─────────────────────────────────────────────
def encontrar_excel() -> str | None:
    """Procura o primeiro .xlsx na pasta do script."""
    base = Path(__file__).parent
    arquivos = sorted(
        list(base.glob("*.xlsx")) + list(base.glob("*.xls")),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return str(arquivos[0]) if arquivos else None


# ─────────────────────────────────────────────
# DETECÇÃO INTELIGENTE DE ABAS
# ─────────────────────────────────────────────

# Colunas-chave para identificar cada aba
_ABA_CHAVES = {
    "empresas":      ["cnpj", "bairro", "cep", "endereço", "endereco", "porte", "latitude", "lat"],
    "ramo":          ["cnae", "ramo_id", "especialização", "especializacao", "sub-ramo_id", "subramo"],
    "sub_ramos":     ["sub-ramo", "subramo", "sub_ramo"],
    "ramo_principal":["ramo principal", "ramo_principal", "ramo_principal_id"],
    "bairros":       ["bairro", "distrito"],
    "distrito":      ["imóveis comerciais", "imoveis comerciais", "pop projetada", "amostra considerada"],
    "porte":         ["porte"],
    "horario":       ["expediente", "descrição", "descricao"],
}


def _cols_lower(df: pd.DataFrame) -> list[str]:
    return [str(c).strip().lower() for c in df.columns]


def detectar_abas(xls: pd.ExcelFile) -> dict[str, pd.DataFrame]:
    """
    Lê cada aba e tenta classificá-la por conteúdo,
    independentemente do nome.
    """
    resultado: dict[str, pd.DataFrame] = {k: pd.DataFrame() for k in _ABA_CHAVES}
    nomes = xls.sheet_names

    for nome in nomes:
        try:
            df = pd.read_excel(xls, sheet_name=nome, nrows=3)
            cols = _cols_lower(df)
        except Exception:
            continue

        melhor_cat  = None
        melhor_score = 0
        for cat, chaves in _ABA_CHAVES.items():
            score = sum(any(chave in col for col in cols) for chave in chaves)
            if score > melhor_score:
                melhor_score = score
                melhor_cat = cat

        if melhor_cat and melhor_score >= 1:
            if resultado[melhor_cat].empty:
                resultado[melhor_cat] = pd.read_excel(xls, sheet_name=nome)

    # Log das abas detectadas (visível só em dev)
    detectadas = {k: v.shape for k, v in resultado.items() if not v.empty}
    return resultado


# ─────────────────────────────────────────────
# TRANSFORMAÇÕES (equivalentes às medidas DAX)
# ─────────────────────────────────────────────

def _detectar_col(df: pd.DataFrame, candidatos: list[str]) -> str | None:
    cols_lower = {c.strip().lower(): c for c in df.columns}
    for c in candidatos:
        if c.lower() in cols_lower:
            return cols_lower[c.lower()]
    return None


def aplicar_transformacoes(empresas: pd.DataFrame,
                            ramo: pd.DataFrame,
                            sub_ramos: pd.DataFrame,
                            ramo_principal: pd.DataFrame) -> pd.DataFrame:
    df = empresas.copy()
    df.columns = df.columns.str.strip()

    # ── Presença Digital ──────────────────────────────────────────────
    pres_col = _detectar_col(df, [
        "site", "presença digital", "presenca digital", "presencadigital",
        "rede social", "instagram", "facebook", "online"
    ])
    if pres_col:
        df["Presenca_Digital"] = df[pres_col].apply(
            lambda x: "NÃO" if str(x).strip().upper() in
                      ["", "NAN", "NÃO", "NAO", "N", "NO", "FALSE", "0", "NONE", "-"] else "SIM"
        )
    else:
        df["Presenca_Digital"] = "NÃO"

    # ── Faixa etária ──────────────────────────────────────────────────
    data_col = _detectar_col(df, ["data abertura", "data_abertura", "abertura", "fundacao", "fundação"])
    if data_col:
        df[data_col] = pd.to_datetime(df[data_col], errors="coerce", dayfirst=True)
        hoje = pd.Timestamp.today()
        df["Anos_Existencia"] = (hoje - df[data_col]).dt.days / 365.25

        def faixa_idade(a):
            if pd.isna(a):   return "Sem dados"
            elif a <= 2:     return "0 a 2 anos"
            elif a <= 5:     return "3 a 5 anos"
            elif a <= 10:    return "6 a 10 anos"
            elif a <= 20:    return "11 a 20 anos"
            else:            return "Mais de 20 anos"
        df["Faixa_Idade"] = df["Anos_Existencia"].apply(faixa_idade)
    else:
        df["Faixa_Idade"] = "Sem dados"

    # ── Faixa de distância ────────────────────────────────────────────
    dist_col = _detectar_col(df, [
        "distancia (km)", "distância (km)", "distancia_km", "distancia",
        "distância", "dist_km", "km"
    ])
    if dist_col:
        df["Dist_km"] = pd.to_numeric(df[dist_col], errors="coerce")
        def faixa_dist(d):
            if pd.isna(d):   return "Sem dados"
            elif d <= 2:     return "0 a 2km"
            elif d <= 4:     return "3 a 4km"
            elif d <= 7:     return "5 a 7km"
            else:            return "Mais de 7km"
        df["Faixa_Dist"] = df["Dist_km"].apply(faixa_dist)
    else:
        df["Faixa_Dist"] = "Sem dados"

    # ── Faixa de tempo ────────────────────────────────────────────────
    tempo_col = _detectar_col(df, [
        "distância desde a fatec", "distancia desde a fatec",
        "tempo", "tempo (min)", "tempo_min", "minutos", "min"
    ])
    if tempo_col:
        df["Tempo_min"] = pd.to_numeric(df[tempo_col], errors="coerce")
        def faixa_tempo(t):
            if pd.isna(t):   return "Sem dados"
            elif t <= 5:     return "0 a 5 minutos"
            elif t <= 10:    return "6 a 10 minutos"
            elif t <= 20:    return "11 a 20 minutos"
            else:            return "Mais de 20 minutos"
        df["Faixa_Tempo"] = df["Tempo_min"].apply(faixa_tempo)
    else:
        df["Faixa_Tempo"] = "Sem dados"

    # ── Tem CNPJ ─────────────────────────────────────────────────────
    cnpj_col = _detectar_col(df, ["cnpj"])
    if cnpj_col:
        df["Tem_CNPJ"] = df[cnpj_col].apply(
            lambda x: "SIM" if str(x).strip() not in ["", "nan", "NaN", "-", "None"] else "NÃO"
        )

    # ── Fallbacks para colunas essenciais ────────────────────────────
    for col, default in [("Porte","Não informado"),("Distrito","Não informado"),
                          ("Bairro","Não informado"),("Expediente","HC")]:
        if col not in df.columns:
            # tenta detectar pelo nome similar
            found = _detectar_col(df, [col.lower()])
            if found:
                df[col] = df[found]
            else:
                df[col] = default

    # ── Hierarquia de Ramos ───────────────────────────────────────────
    df["Ramo_Nome"]      = "Não informado"
    df["Sub_Ramo"]       = "Não informado"
    df["Ramo_Principal"] = "Não informado"

    ramo_id_emp = _detectar_col(df, ["ramo_id", "ramo id"])

    if not ramo.empty and ramo_id_emp:
        ramo.columns = ramo.columns.str.strip()
        r_id  = _detectar_col(ramo, ["ramo_id", "ramo id"]) or ramo.columns[0]
        r_nom = _detectar_col(ramo, ["especialização / tipo de negó", "especializacao",
                                      "especializacao / tipo", "ramo", "nome"]) or ramo.columns[-1]
        ramo_map = ramo.set_index(r_id)[r_nom].to_dict()
        df["Ramo_Nome"] = df[ramo_id_emp].map(ramo_map).fillna("Não informado")

        # Sub-Ramo
        if not sub_ramos.empty:
            sub_ramos.columns = sub_ramos.columns.str.strip()
            sr_id  = _detectar_col(sub_ramos, ["sub-ramo_id","subramo_id","id"]) or sub_ramos.columns[0]
            sr_nom = _detectar_col(sub_ramos, ["sub-ramo","subramo","nome"]) or sub_ramos.columns[-1]
            r_sr_id = _detectar_col(ramo, ["sub-ramo_id","subramo_id"])
            if r_sr_id:
                sr_map = sub_ramos.set_index(sr_id)[sr_nom].to_dict()
                df["Sub_Ramo"] = df[ramo_id_emp].map(
                    ramo.set_index(r_id)[r_sr_id].to_dict()
                ).map(sr_map).fillna("Não informado")

        # Ramo Principal
        if not ramo_principal.empty:
            ramo_principal.columns = ramo_principal.columns.str.strip()
            rp_id  = _detectar_col(ramo_principal, ["ramo_principal_id","id"]) or ramo_principal.columns[0]
            rp_nom = _detectar_col(ramo_principal, ["ramo principal","ramo_principal","nome"]) or ramo_principal.columns[1]
            r_rp_id = _detectar_col(ramo, ["ramo_principal_id"])
            if r_rp_id:
                rp_map = ramo_principal.set_index(rp_id)[rp_nom].to_dict()
                df["Ramo_Principal"] = df[ramo_id_emp].map(
                    ramo.set_index(r_id)[r_rp_id].to_dict()
                ).map(rp_map).fillna("Não informado")

    # ── Lat / Lon ─────────────────────────────────────────────────────
    lat_col = _detectar_col(df, ["latitude","lat"])
    lon_col = _detectar_col(df, ["longitude","lon","lng"])
    if lat_col: df["Lat"] = pd.to_numeric(df[lat_col], errors="coerce")
    if lon_col: df["Lon"] = pd.to_numeric(df[lon_col], errors="coerce")

    return df


# ─────────────────────────────────────────────
# CARREGAMENTO PRINCIPAL
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Carregando dados...")
def carregar_dados(caminho: str):
    xls   = pd.ExcelFile(caminho)
    abas  = detectar_abas(xls)
    empresas_raw = abas["empresas"]
    if empresas_raw.empty:
        st.error(f"Não foi possível detectar a aba de Empresas.\n"
                 f"Abas encontradas: {xls.sheet_names}")
        st.stop()

    df = aplicar_transformacoes(
        empresas_raw,
        abas["ramo"],
        abas["sub_ramos"],
        abas["ramo_principal"],
    )
    return df, abas["distrito"]


# ─────────────────────────────────────────────
# COMPONENTES REUTILIZÁVEIS
# ─────────────────────────────────────────────
LOGO_EMP  = _logo_path("Empresa_Simbolo.png")
LOGO_META = _logo_path("metaday_logo_sem_fundo.png")

def render_header(titulo="Segmento: EMPRESAS"):
    logo_emp_b64  = img_to_base64(LOGO_EMP)  if LOGO_EMP  else ""
    logo_meta_b64 = img_to_base64(LOGO_META) if LOGO_META else ""

    img_esq  = f'<img src="{logo_emp_b64}"  class="header-logo-left">'  if logo_emp_b64  else "🏢"
    img_dir  = f'<img src="{logo_meta_b64}" class="header-logo-right">' if logo_meta_b64 else "⬤"

    st.markdown(f"""
    <div class="header-row">
        <div>{img_esq}</div>
        <h1 class="titulo-principal">{titulo}</h1>
        <div>{img_dir}</div>
    </div>
    """, unsafe_allow_html=True)


def kpi_card(label: str, value) -> str:
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
    </div>"""


def _chart_defaults(fig, height=260):
    fig.update_layout(
        height=height,
        margin=dict(t=10, b=10, l=10, r=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def pizza_presenca(df, height=210):
    pres = df["Presenca_Digital"].value_counts().reset_index()
    pres.columns = ["Status", "Qtde"]
    fig = px.pie(pres, names="Status", values="Qtde",
                 color="Status",
                 color_discrete_map={"SIM": COR_GOLD, "NÃO": "#D0D0D0"},
                 hole=0.35)
    fig.update_traces(textposition="outside", textinfo="label+value+percent", textfont_size=10)
    return _chart_defaults(fig, height)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 1 — VISÃO GERAL
# ═══════════════════════════════════════════════════════════════
def pagina_visao_geral(df, df_distrito):
    render_header()

    total_emp  = len(df)
    tot_dist   = df["Distrito"].nunique()
    tot_espec  = df["Ramo_Nome"].nunique()

    c1, c2, c3 = st.columns(3)
    for col, label, val in [
        (c1, "Empresas Catalogadas",       f"{total_emp:,}".replace(",",".")),
        (c2, "Distritos Considerados",     str(tot_dist)),
        (c3, "Especializações Identificadas", str(tot_espec)),
    ]:
        with col:
            st.markdown(kpi_card(label, val), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_esq, col_dir = st.columns([1, 2.5])

    with col_esq:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        st.plotly_chart(pizza_presenca(df, 220), use_container_width=True)

        st.markdown('<div class="section-title">Porte das Empresas</div>', unsafe_allow_html=True)
        porte = (df["Porte"].value_counts()
                   .reindex(["Pequeno","Médio","Grande"])
                   .dropna().reset_index())
        porte.columns = ["Porte", "Qtde"]
        cor_porte = {"Pequeno": COR_GOLD, "Médio": COR_NAVY, "Grande": COR_CINZA}
        fig_p = px.bar(porte, y="Porte", x="Qtde", orientation="h",
                       color="Porte", color_discrete_map=cor_porte, text="Qtde")
        fig_p.update_traces(textposition="inside", textfont_size=13, textfont_color="white")
        fig_p.update_layout(showlegend=False, xaxis=dict(visible=False), yaxis_title="")
        st.plotly_chart(_chart_defaults(fig_p, 180), use_container_width=True)

    with col_dir:
        st.markdown('<div class="section-title">Distribuição Geográfica por Distrito</div>',
                    unsafe_allow_html=True)
        tem_coords = ("Lat" in df.columns and "Lon" in df.columns
                      and df["Lat"].notna().sum() > 10)
        if tem_coords:
            df_m = df.dropna(subset=["Lat","Lon"])
            fig_m = px.scatter_mapbox(
                df_m, lat="Lat", lon="Lon",
                color="Distrito", size_max=12, zoom=12, height=440,
                mapbox_style="open-street-map",
                color_discrete_sequence=px.colors.qualitative.Bold,
                hover_data={"Lat":False,"Lon":False,"Porte":True,"Distrito":True},
            )
            fig_m.update_layout(margin=dict(t=0,b=0,l=0,r=0), showlegend=False,
                                 paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_m, use_container_width=True)
        else:
            dist_c = df.groupby("Distrito").size().reset_index(name="Qtde").sort_values("Qtde",ascending=False)
            fig_b = px.bar(dist_c, x="Distrito", y="Qtde", text="Qtde",
                           color="Qtde",
                           color_continuous_scale=[[0,COR_GOLD],[1,COR_NAVY]])
            fig_b.update_traces(textposition="outside")
            fig_b.update_layout(coloraxis_showscale=False,
                                 xaxis_title="", yaxis_title="Empresas")
            st.plotly_chart(_chart_defaults(fig_b, 440), use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 2 — DISTRITOS & RAMOS
# ═══════════════════════════════════════════════════════════════
def pagina_distritos(df):
    render_header()
    col_esq, col_dir = st.columns([1, 2.5])

    with col_esq:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        st.plotly_chart(pizza_presenca(df, 200), use_container_width=True)

    with col_dir:
        st.markdown('<div class="section-title">Empresas por Distrito</div>', unsafe_allow_html=True)
        dist_c = df.groupby("Distrito").size().reset_index(name="Qtde")
        fig_t = px.treemap(dist_c, path=["Distrito"], values="Qtde",
                           color="Qtde",
                           color_continuous_scale=[
                               [0,"#1a237e"],[.15,"#283593"],[.3,"#5c6bc0"],
                               [.45,"#00897b"],[.6,"#43a047"],[.75,"#e53935"],
                               [.9,"#8e24aa"],[1,"#00acc1"],
                           ])
        fig_t.update_traces(textinfo="label+value", textfont_size=13, textposition="bottom left")
        fig_t.update_layout(coloraxis_showscale=False, margin=dict(t=5,b=5,l=5,r=5),
                             height=280, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_t, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">Tempo de Existência</div>', unsafe_allow_html=True)
        ordem = ["Mais de 20 anos","0 a 2 anos","11 a 20 anos","6 a 10 anos","3 a 5 anos","Sem dados"]
        fi = df["Faixa_Idade"].value_counts().reindex(ordem).dropna().reset_index()
        fi.columns = ["Faixa","Qtde"]
        fig_fi = px.bar(fi, y="Faixa", x="Qtde", orientation="h", text="Qtde",
                        color="Qtde",
                        color_continuous_scale=[[0,"#90CAF9"],[.5,"#1565C0"],[1,COR_NAVY]])
        fig_fi.update_traces(textposition="outside", textfont_size=12)
        fig_fi.update_layout(coloraxis_showscale=False,
                              xaxis=dict(visible=False), yaxis_title="", showlegend=False)
        st.plotly_chart(_chart_defaults(fig_fi, 270), use_container_width=True)

    with c2:
        st.markdown('<div class="section-title">Quantidade por Ramo</div>', unsafe_allow_html=True)
        ramo_c = (df["Ramo_Principal"].value_counts()
                    .reset_index()
                    .rename(columns={"index":"Ramo","Ramo_Principal":"Qtde",
                                     "count":"Qtde","Ramo_Principal":"Ramo"})
                  )
        # garante colunas corretas independente da versão do pandas
        ramo_c.columns = ["Ramo","Qtde"]
        ramo_c = ramo_c[ramo_c["Ramo"] != "Não informado"].head(10)
        fig_r = px.bar(ramo_c, x="Ramo", y="Qtde", text="Qtde",
                       color_discrete_sequence=[COR_GOLD])
        fig_r.update_traces(textposition="outside", textfont_size=11)
        fig_r.update_layout(xaxis=dict(title="", tickangle=-30, tickfont_size=10),
                             yaxis=dict(visible=False), showlegend=False)
        st.plotly_chart(_chart_defaults(fig_r, 270), use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 3 — FATEC & SUB-RAMOS
# ═══════════════════════════════════════════════════════════════
def pagina_fatec(df):
    render_header()
    c1, c2 = st.columns(2)

    with c1:
        st.markdown('<div class="section-title">Tempo da FATEC até a Empresa</div>', unsafe_allow_html=True)
        ordem_t = ["11 a 20 minutos","6 a 10 minutos","0 a 5 minutos","Mais de 20 minutos","Sem dados"]
        tc = df["Faixa_Tempo"].value_counts().reindex(ordem_t).dropna().reset_index()
        tc.columns = ["Faixa","Qtde"]
        fig_tc = px.bar(tc, y="Faixa", x="Qtde", orientation="h",
                        text="Qtde", color_discrete_sequence=[COR_GOLD])
        fig_tc.update_traces(textposition="outside", textfont_size=12)
        fig_tc.update_layout(xaxis=dict(visible=False), yaxis_title="", showlegend=False)
        st.plotly_chart(_chart_defaults(fig_tc, 220), use_container_width=True)

    with c2:
        st.markdown('<div class="section-title">Distância da FATEC até a Empresa</div>', unsafe_allow_html=True)
        ordem_d = ["3 a 4km","0 a 2km","5 a 7km","Mais de 7km","Sem dados"]
        dc = df["Faixa_Dist"].value_counts().reindex(ordem_d).dropna().reset_index()
        dc.columns = ["Faixa","Qtde"]
        fig_dc = px.bar(dc, x="Faixa", y="Qtde", text="Qtde",
                        color="Qtde",
                        color_continuous_scale=[[0,"#E0D0A0"],[1,COR_GOLD]])
        fig_dc.update_traces(textposition="outside", textfont_size=12)
        fig_dc.update_layout(coloraxis_showscale=False,
                              xaxis_title="", yaxis=dict(visible=False), showlegend=False)
        st.plotly_chart(_chart_defaults(fig_dc, 220), use_container_width=True)

    col_f, col_t = st.columns([1, 3])
    with col_f:
        st.markdown('<div class="section-title">Ramo de Atividade</div>', unsafe_allow_html=True)
        ramos_disp = sorted([r for r in df["Ramo_Principal"].dropna().unique()
                             if r != "Não informado"])
        ramos_sel = st.multiselect("", ramos_disp, default=[], label_visibility="collapsed")

    with col_t:
        st.markdown('<div class="section-title">Sub-Divisão dos Ramos de Atividade</div>',
                    unsafe_allow_html=True)
        df_f = df[df["Ramo_Principal"].isin(ramos_sel)] if ramos_sel else df
        sub_c = (df_f.groupby(["Ramo_Principal","Sub_Ramo"])
                      .size().reset_index(name="Qtde")
                      .query("Sub_Ramo != 'Não informado' and Qtde > 0"))
        if not sub_c.empty:
            fig_s = px.treemap(sub_c, path=["Ramo_Principal","Sub_Ramo"],
                               values="Qtde", color="Ramo_Principal",
                               color_discrete_sequence=TREEMAP_COLORS)
            fig_s.update_traces(textinfo="label+value", textfont_size=11,
                                 textposition="bottom left")
            fig_s.update_layout(height=370, margin=dict(t=5,b=5,l=5,r=5),
                                 paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_s, use_container_width=True)
        else:
            st.info("Selecione um ou mais ramos para ver a sub-divisão.")


# ═══════════════════════════════════════════════════════════════
# PÁGINA 4 — MAPA POR SEGMENTO
# ═══════════════════════════════════════════════════════════════
def pagina_mapa_ramos(df):
    render_header()
    col_f, col_m = st.columns([1, 3])

    with col_f:
        st.markdown('<div class="section-title">Segmentação dos Ramos</div>', unsafe_allow_html=True)
        subs = sorted([s for s in df["Sub_Ramo"].dropna().unique() if s != "Não informado"])
        sel  = st.multiselect("", subs, default=[], label_visibility="collapsed", key="mapa_sub")

    with col_m:
        df_f = df[df["Sub_Ramo"].isin(sel)] if sel else df
        tem_coords = ("Lat" in df_f.columns and "Lon" in df_f.columns
                      and df_f["Lat"].notna().sum() > 5)
        if tem_coords:
            df_m = df_f.dropna(subset=["Lat","Lon"])
            fig_m = px.scatter_mapbox(
                df_m, lat="Lat", lon="Lon", color="Sub_Ramo",
                hover_data={"Lat":False,"Lon":False,"Porte":True,"Distrito":True,"Sub_Ramo":True},
                zoom=12, height=560, mapbox_style="open-street-map",
                color_discrete_sequence=px.colors.qualitative.Alphabet,
            )
            fig_m.update_layout(margin=dict(t=0,b=0,l=0,r=0),
                                 legend=dict(title="Sub-Ramo",font_size=10),
                                 paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_m, use_container_width=True)
        else:
            sub_c = (df_f["Sub_Ramo"].value_counts().head(20).reset_index()
                       .rename(columns={"index":"Sub_Ramo","Sub_Ramo":"Qtde",
                                        "count":"Qtde"}))
            sub_c.columns = ["Sub_Ramo","Qtde"]
            sub_c = sub_c[sub_c["Sub_Ramo"] != "Não informado"]
            fig_b = px.bar(sub_c, x="Qtde", y="Sub_Ramo", orientation="h",
                           text="Qtde", color="Sub_Ramo",
                           color_discrete_sequence=px.colors.qualitative.Bold)
            fig_b.update_traces(textposition="outside")
            fig_b.update_layout(showlegend=False,
                                 xaxis=dict(visible=False), yaxis_title="")
            st.plotly_chart(_chart_defaults(fig_b, 560), use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# PÁGINA 5 — DISTRITOS POR SEGMENTO
# ═══════════════════════════════════════════════════════════════
def pagina_distritos_segmento(df):
    render_header()
    col_f, col_t = st.columns([1, 3])

    with col_f:
        st.markdown('<div class="section-title">Segmentação dos Ramos</div>', unsafe_allow_html=True)
        subs = sorted([s for s in df["Sub_Ramo"].dropna().unique() if s != "Não informado"])
        sel  = st.multiselect("", subs, default=[], label_visibility="collapsed", key="dist_sub")

    with col_t:
        df_f  = df[df["Sub_Ramo"].isin(sel)] if sel else df
        st.markdown('<div class="section-title">Distritos que têm os Segmentos selecionados</div>',
                    unsafe_allow_html=True)
        dist_c = df_f.groupby("Distrito").size().reset_index(name="Qtde")
        if not dist_c.empty:
            fig_t = px.treemap(dist_c, path=["Distrito"], values="Qtde",
                               color="Qtde",
                               color_continuous_scale=[
                                   [0,"#1a237e"],[.15,"#283593"],[.3,"#5c6bc0"],
                                   [.45,"#00897b"],[.6,"#43a047"],[.75,"#e53935"],
                                   [.9,"#8e24aa"],[1,"#00acc1"],
                               ])
            fig_t.update_traces(textinfo="label+value", textfont_size=14,
                                 textposition="bottom left")
            fig_t.update_layout(height=560, margin=dict(t=5,b=5,l=5,r=5),
                                 coloraxis_showscale=False,
                                 paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_t, use_container_width=True)
        else:
            st.info("Nenhuma empresa encontrada com o filtro atual.")


# ═══════════════════════════════════════════════════════════════
# PÁGINA 6 — IBGE / PERFIL
# ═══════════════════════════════════════════════════════════════
def pagina_ibge(df, df_distrito):
    render_header()
    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown('<div class="section-title">Porte das Empresas</div>', unsafe_allow_html=True)
        porte = (df["Porte"].value_counts()
                   .reindex(["Pequeno","Médio","Grande"]).dropna().reset_index())
        porte.columns = ["Porte","Qtde"]
        cor_p = {"Pequeno":COR_GOLD,"Médio":COR_NAVY,"Grande":COR_CINZA}
        fig_p = px.bar(porte, x="Porte", y="Qtde", text="Qtde",
                       color="Porte", color_discrete_map=cor_p)
        fig_p.update_traces(textposition="outside")
        fig_p.update_layout(showlegend=False,
                             xaxis_title="Porte", yaxis=dict(visible=False))
        st.plotly_chart(_chart_defaults(fig_p, 220), use_container_width=True)

    with c2:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        st.plotly_chart(pizza_presenca(df, 220), use_container_width=True)

    with c3:
        st.markdown('<div class="section-title">Expediente</div>', unsafe_allow_html=True)
        exp_c = df["Expediente"].value_counts().reset_index()
        exp_c.columns = ["Tipo","Qtde"]
        fig_e = px.pie(exp_c, names="Tipo", values="Qtde",
                       color_discrete_sequence=[COR_NAVY, COR_GOLD, COR_CINZA], hole=0.3)
        fig_e.update_traces(textposition="outside", textinfo="label+percent", textfont_size=10)
        fig_e.update_layout(showlegend=False, margin=dict(t=5,b=5,l=5,r=5),
                             height=220, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_e, use_container_width=True)

    st.markdown("---")
    st.markdown(f'<div style="font-size:1.05rem;font-weight:800;color:{COR_TEXTO};margin-bottom:6px;">'
                f'📊 Dados IBGE / Censo 2022</div>', unsafe_allow_html=True)

    if not df_distrito.empty:
        df_distrito.columns = df_distrito.columns.str.strip()
        st.dataframe(df_distrito, use_container_width=True, height=340)
    else:
        tab = (df.groupby("Distrito").size().reset_index(name="Na Amostra")
                 .sort_values("Distrito"))
        total = pd.DataFrame([{"Distrito":"Total","Na Amostra":len(df)}])
        st.dataframe(pd.concat([tab, total], ignore_index=True),
                     use_container_width=True, height=340)

    st.caption("Os dados, exceto NA AMOSTRA, foram obtidos cruzando Prefeitura SP, SEADE e OBSERVASAMPA. "
               "Populações projetadas usam o Censo 2010 como base.")


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
def render_sidebar(df):
    logo_meta_b64 = img_to_base64(LOGO_META) if LOGO_META else ""
    with st.sidebar:
        if logo_meta_b64:
            st.markdown(
                f'<div style="text-align:center;padding:10px 0 4px 0;">'
                f'<img src="{logo_meta_b64}" style="width:90px;height:90px;'
                f'object-fit:contain;border-radius:50%;"></div>',
                unsafe_allow_html=True,
            )
        st.markdown("""
        <div style="text-align:center;padding:0 0 14px 0;">
            <div style="font-size:1.05rem;font-weight:800;letter-spacing:1px;">METADAY 2025</div>
            <div style="font-size:0.75rem;opacity:0.7;">Fatec Sebrae · Ciência de Dados</div>
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
            label_visibility="collapsed",
        )
        st.markdown("---")
        st.markdown("**🔍 Filtros Globais**")

        distritos = sorted([d for d in df["Distrito"].dropna().unique()
                            if d not in ("Não informado","")])
        sel_dist  = st.multiselect("Distrito", distritos, default=distritos)

        portes    = sorted([p for p in df["Porte"].dropna().unique()
                            if p not in ("Não informado","")])
        sel_porte = st.multiselect("Porte", portes, default=portes)

        sel_pres  = st.multiselect("Presença Digital", ["SIM","NÃO"],
                                   default=["SIM","NÃO"])
        st.markdown("---")
        n_filt = len(df[df["Distrito"].isin(sel_dist) &
                        df["Porte"].isin(sel_porte) &
                        df["Presenca_Digital"].isin(sel_pres)])
        st.caption(f"Total filtrado: **{n_filt:,}** empresas")

    return pagina, sel_dist, sel_porte, sel_pres


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    # 1) Tenta encontrar Excel automaticamente na pasta
    caminho_auto = encontrar_excel()

    if caminho_auto:
        df_raw, df_distrito = carregar_dados(caminho_auto)
    else:
        # 2) Fallback: upload manual (oculto por padrão via CSS, aparece só se necessário)
        st.markdown('<h1 class="titulo-principal">Segmento: EMPRESAS</h1>', unsafe_allow_html=True)
        st.info("📂 Nenhum arquivo Excel encontrado na pasta do app. "
                "Coloque o `.xlsx` na mesma pasta que o `app.py` **ou** faça upload abaixo.")
        # Remove o display:none do uploader para esta situação
        st.markdown("<style>div[data-testid='stFileUploader']{display:block!important;}</style>",
                    unsafe_allow_html=True)
        arquivo = st.file_uploader("Carregar planilha Excel", type=["xlsx","xls"])
        if arquivo is None:
            st.stop()
        df_raw, df_distrito = carregar_dados(arquivo)

    # 3) Sidebar + filtros
    pagina, sel_dist, sel_porte, sel_pres = render_sidebar(df_raw)

    df = df_raw.copy()
    if sel_dist:  df = df[df["Distrito"].isin(sel_dist)]
    if sel_porte: df = df[df["Porte"].isin(sel_porte)]
    if sel_pres:  df = df[df["Presenca_Digital"].isin(sel_pres)]

    # 4) Roteamento
    rotas = {
        "1": pagina_visao_geral,
        "2": pagina_distritos,
        "3": pagina_fatec,
        "4": pagina_mapa_ramos,
        "5": pagina_distritos_segmento,
        "6": pagina_ibge,
    }
    n = pagina[0]
    if n in ("1","6"):
        rotas[n](df, df_distrito)
    else:
        rotas[n](df)


if __name__ == "__main__":
    main()
