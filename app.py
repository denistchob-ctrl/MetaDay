"""
Dashboard: Segmento EMPRESAS — Projeto Metaday 2025
v4 — hierarquia correta + filtros em árvore expansível nas páginas 3, 4, 5
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import base64
from pathlib import Path

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Segmento: EMPRESAS",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

COR_GOLD  = "#B8972E"
COR_NAVY  = "#1F3864"
COR_CINZA = "#BDBDBD"
COR_AZUL  = "#5B9BD5"
COR_BG    = "#F0F2F5"
COR_TEXTO = "#1A1A2E"
TREEMAP_COLORS = px.colors.qualitative.Bold

# ──────────────────────────────────────────────────────────────
# UTILITÁRIOS DE IMAGEM
# ──────────────────────────────────────────────────────────────
def img_to_base64(path: str) -> str:
    try:
        with open(path, "rb") as f:
            data = f.read()
        mime = "png" if Path(path).suffix.lower() == ".png" else "jpeg"
        return f"data:image/{mime};base64,{base64.b64encode(data).decode()}"
    except Exception:
        return ""

def _logo_path(filename: str) -> str:
    base = Path(__file__).parent
    for c in [base / filename, Path(filename), Path("assets") / filename]:
        if c.exists():
            return str(c)
    return ""

# ──────────────────────────────────────────────────────────────
# CSS
# ──────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
.stApp {{ background-color:{COR_BG}; }}
h1.titulo-principal {{
    font-family:'Georgia',serif; font-size:2.4rem; font-weight:800;
    color:{COR_TEXTO}; text-align:center; letter-spacing:1px; margin:0;
}}
.kpi-card {{
    background:white; border-radius:12px; padding:16px 20px;
    text-align:center; box-shadow:0 2px 8px rgba(0,0,0,0.08);
    border-top:4px solid {COR_GOLD};
}}
.kpi-label {{ font-size:0.82rem; font-style:italic; color:#666; font-weight:600; }}
.kpi-value {{ font-size:2.4rem; font-weight:900; color:{COR_TEXTO}; line-height:1.1; }}
.section-title {{
    font-family:'Georgia',serif; font-size:1.05rem; font-style:italic;
    font-weight:700; color:{COR_TEXTO}; margin:10px 0 2px 0;
}}
.header-row {{
    display:flex; align-items:center; justify-content:space-between;
    padding:4px 0 8px 0; border-bottom:2px solid #ddd; margin-bottom:12px;
}}
.header-logo-left  {{ width:80px; height:80px; object-fit:contain; }}
.header-logo-right {{ width:60px; height:60px; object-fit:contain; border-radius:50%; }}

/* Árvore de filtro */
.tree-sub {{
    font-weight:700; font-size:0.88rem;
    color:{COR_NAVY}; padding:4px 0 2px 0; margin-top:6px;
}}
.tree-ramo {{
    font-size:0.82rem; color:#333; padding:1px 0 1px 12px;
}}

section[data-testid="stSidebar"] {{ background-color:#1A2744; }}
section[data-testid="stSidebar"] * {{ color:white !important; }}
section[data-testid="stSidebar"] .stMultiSelect>div>div {{ background-color:#2a3a5c !important; }}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# LOCALIZAR EXCEL
# ──────────────────────────────────────────────────────────────
def encontrar_excel() -> str | None:
    base = Path(__file__).parent
    arquivos = sorted(
        list(base.glob("*.xlsx")) + list(base.glob("*.xls")),
        key=lambda p: p.stat().st_mtime, reverse=True,
    )
    return str(arquivos[0]) if arquivos else None


# ──────────────────────────────────────────────────────────────
# DETECÇÃO DE ABAS (por conteúdo, não por nome)
# ──────────────────────────────────────────────────────────────
_ABA_CHAVES = {
    "empresas":       ["cnpj","tipo de porte","site/redes sociais",
                       "distancia de carro","tempo para chegar","latitude"],
    "ramo":           ["cnae","ramo_id","especialização","sub-ramo_id","ramo_principal_id"],
    "sub_ramos":      ["sub-ramo_id","sub-ramo","subramo"],
    "ramo_principal": ["ramo_principal_id","ramo principal"],
    "bairros":        ["bairro","distrito"],
    "distrito":       ["imóveis comerciais","pop projetada","amostra considerada"],
    "porte":          ["porte"],
    "horario":        ["expediente","descrição"],
}

def detectar_abas(xls: pd.ExcelFile) -> dict:
    resultado = {k: pd.DataFrame() for k in _ABA_CHAVES}
    for nome in xls.sheet_names:
        try:
            df_p = pd.read_excel(xls, sheet_name=nome, nrows=3)
            cols = [str(c).strip().lower() for c in df_p.columns]
        except Exception:
            continue
        melhor_cat, melhor_score = None, 0
        for cat, chaves in _ABA_CHAVES.items():
            score = sum(any(ch in col for col in cols) for ch in chaves)
            if score > melhor_score:
                melhor_score, melhor_cat = score, cat
        if melhor_cat and melhor_score >= 1 and resultado[melhor_cat].empty:
            resultado[melhor_cat] = pd.read_excel(xls, sheet_name=nome)
    return resultado


# ──────────────────────────────────────────────────────────────
# DETECÇÃO DE COLUNA (parcial, case-insensitive)
# ──────────────────────────────────────────────────────────────
def _col(df: pd.DataFrame, candidatos: list) -> str | None:
    lower = {c.strip().lower(): c for c in df.columns}
    for c in candidatos:
        if c.lower() in lower:
            return lower[c.lower()]
    for c in candidatos:
        for kl, kr in lower.items():
            if c.lower() in kl:
                return kr
    return None


# ──────────────────────────────────────────────────────────────
# TRANSFORMAÇÕES  ←  equivalentes DAX
# ──────────────────────────────────────────────────────────────
def aplicar_transformacoes(empresas, ramo_df, sub_ramos_df, ramo_principal_df):
    df = empresas.copy()
    df.columns = df.columns.str.strip()

    # 1) Presença Digital  →  "Site/Redes Sociais"
    pc = _col(df, ["site/redes sociais","site redes sociais","site","redes sociais",
                   "presença digital","presenca digital"])
    _vazios = {"","nan","none","-","não","nao","n","no","false","0"}
    df["Presenca_Digital"] = (
        df[pc].apply(lambda x: "NÃO" if str(x).strip().lower() in _vazios else "SIM")
        if pc else "NÃO"
    )

    # 2) Porte  →  "Tipo de Porte"
    ptc = _col(df, ["tipo de porte","porte"])
    if ptc:
        df["Porte"] = df[ptc].astype(str).str.strip()
    elif "Porte" not in df.columns:
        df["Porte"] = "Não informado"

    # 3) Faixa etária  →  "Data de Abertura da Empresa"
    dc = _col(df, ["data de abertura da empresa","data abertura","data_abertura","abertura"])
    if dc:
        df[dc] = pd.to_datetime(df[dc], errors="coerce", dayfirst=True)
        hoje = pd.Timestamp.today()
        df["Anos_Existencia"] = (hoje - df[dc]).dt.days / 365.25
        def faixa_idade(a):
            if pd.isna(a):  return "Sem dados"
            elif a <= 2:    return "0 a 2 anos"
            elif a <= 5:    return "3 a 5 anos"
            elif a <= 10:   return "6 a 10 anos"
            elif a <= 20:   return "11 a 20 anos"
            else:           return "Mais de 20 anos"
        df["Faixa_Idade"] = df["Anos_Existencia"].apply(faixa_idade)
    else:
        df["Faixa_Idade"] = "Sem dados"

    # 4) Faixa Tempo  →  "Tempo para chegar de carro da FATEC ... (minutos)"
    tc = _col(df, ["tempo para chegar de carro da fatec","tempo para chegar",
                   "tempo (min)","tempo_min","minutos"])
    if tc:
        df["Tempo_min"] = pd.to_numeric(df[tc], errors="coerce")
        def faixa_tempo(t):
            if pd.isna(t):  return "Sem dados"
            elif t <= 5:    return "0 a 5 minutos"
            elif t <= 10:   return "6 a 10 minutos"
            elif t <= 20:   return "11 a 20 minutos"
            else:           return "Mais de 20 minutos"
        df["Faixa_Tempo"] = df["Tempo_min"].apply(faixa_tempo)
    else:
        df["Faixa_Tempo"] = "Sem dados"

    # 5) Faixa Distância  →  "Distancia de Carro da FATEC ... (km)"
    distc = _col(df, ["distancia de carro da fatec","distância de carro da fatec",
                      "distancia (km)","distância (km)","distancia_km"])
    if distc:
        df["Dist_km"] = pd.to_numeric(df[distc], errors="coerce")
        def faixa_dist(d):
            if pd.isna(d):  return "Sem dados"
            elif d <= 2:    return "0 a 2km"
            elif d <= 4:    return "3 a 4km"
            elif d <= 7:    return "5 a 7km"
            else:           return "Mais de 7km"
        df["Faixa_Dist"] = df["Dist_km"].apply(faixa_dist)
    else:
        df["Faixa_Dist"] = "Sem dados"

    # Fallbacks coluna Distrito / Bairro / Expediente
    for col, alts, default in [
        ("Distrito",   ["distrito"],                       "Não informado"),
        ("Bairro",     ["bairro"],                         "Não informado"),
        ("Expediente", ["expediente","horário","horario"], "HC"),
    ]:
        if col not in df.columns:
            found = _col(df, alts)
            df[col] = df[found] if found else default

    # ── HIERARQUIA CORRETA ────────────────────────────────────────────
    #
    #  Empresas  ──[RAMO_ID]──►  Ramo  ──[SUB-RAMO_ID]──►  Sub-Ramos
    #                                   ──[RAMO_PRINCIPAL_ID]──► Ramo Principal
    #
    df["Ramo_Nome"]      = "Não informado"
    df["Sub_Ramo"]       = "Não informado"
    df["Sub_Ramo_ID"]    = pd.NA
    df["Ramo_Principal"] = "Não informado"

    # chave que liga Empresa → Ramo
    emp_ramo_id = _col(df, ["ramo_id","ramo id"])

    if not ramo_df.empty and emp_ramo_id:
        ramo_df = ramo_df.copy()
        ramo_df.columns = ramo_df.columns.str.strip()

        # colunas da tabela Ramo
        r_id   = _col(ramo_df, ["ramo_id","ramo id"])            or ramo_df.columns[0]
        r_nom  = _col(ramo_df, ["especialização / tipo de negó",
                                 "especializacao","cnae","nome"])  or ramo_df.columns[-1]
        r_sr   = _col(ramo_df, ["sub-ramo_id","subramo_id"])      # FK → Sub-Ramos
        r_rp   = _col(ramo_df, ["ramo_principal_id"])             # FK → Ramo Principal

        # Ramo_Nome: nome do ramo/especialização da empresa
        ramo_nome_map = ramo_df.set_index(r_id)[r_nom].to_dict()
        df["Ramo_Nome"] = df[emp_ramo_id].map(ramo_nome_map).fillna("Não informado")

        # Sub_Ramo: Empresa → Ramo → Sub-Ramo
        if not sub_ramos_df.empty and r_sr:
            sub_ramos_df = sub_ramos_df.copy()
            sub_ramos_df.columns = sub_ramos_df.columns.str.strip()

            # colunas da tabela Sub-Ramos
            # Sub-Ramos tem: SUB-RAMO_ID (PK), RAMO_ID (FK → Ramo Principal), SUB-RAMO (nome)
            sr_id  = _col(sub_ramos_df, ["sub-ramo_id","subramo_id"]) or sub_ramos_df.columns[0]
            sr_nom = _col(sub_ramos_df, ["sub-ramo","subramo","nome"]) or sub_ramos_df.columns[-1]

            # mapeia: ramo_id da empresa → sub-ramo_id (via tabela Ramo)
            ramo_to_sr = ramo_df.set_index(r_id)[r_sr].to_dict()
            # mapeia: sub-ramo_id → nome do sub-ramo
            sr_nome_map = sub_ramos_df.set_index(sr_id)[sr_nom].to_dict()

            df["Sub_Ramo_ID"] = df[emp_ramo_id].map(ramo_to_sr)
            df["Sub_Ramo"]    = df["Sub_Ramo_ID"].map(sr_nome_map).fillna("Não informado")

        # Ramo_Principal: Sub-Ramo → Ramo Principal
        # A tabela Sub-Ramos tem RAMO_ID que aponta para Ramo Principal
        if not ramo_principal_df.empty and not sub_ramos_df.empty:
            ramo_principal_df = ramo_principal_df.copy()
            ramo_principal_df.columns = ramo_principal_df.columns.str.strip()

            rp_id  = _col(ramo_principal_df, ["ramo_principal_id","id"]) or ramo_principal_df.columns[0]
            rp_nom = _col(ramo_principal_df, ["ramo principal","ramo_principal","nome"]) or ramo_principal_df.columns[1]

            # Sub-Ramos.RAMO_ID aponta para Ramo Principal
            sr_rp_col = _col(sub_ramos_df, ["ramo_id","ramo id"])
            if sr_rp_col:
                rp_nome_map  = ramo_principal_df.set_index(rp_id)[rp_nom].to_dict()
                sr_to_rp     = sub_ramos_df.set_index(sr_id)[sr_rp_col].to_dict()
                df["Ramo_Principal"] = (
                    df["Sub_Ramo_ID"].map(sr_to_rp).map(rp_nome_map).fillna("Não informado")
                )

    # Lat / Lon
    latc = _col(df, ["latitude","lat"])
    lonc = _col(df, ["longitude","lon","lng"])
    if latc: df["Lat"] = pd.to_numeric(df[latc], errors="coerce")
    if lonc: df["Lon"] = pd.to_numeric(df[lonc], errors="coerce")

    return df


# ──────────────────────────────────────────────────────────────
# CARREGAMENTO
# ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Carregando dados…")
def carregar_dados(caminho):
    xls  = pd.ExcelFile(caminho)
    abas = detectar_abas(xls)
    if abas["empresas"].empty:
        st.error(f"Aba de Empresas não detectada. Abas: {xls.sheet_names}")
        st.stop()
    df = aplicar_transformacoes(
        abas["empresas"], abas["ramo"], abas["sub_ramos"], abas["ramo_principal"]
    )
    return df, abas["distrito"]


# ──────────────────────────────────────────────────────────────
# FILTRO EM ÁRVORE  — replica o Power BI exatamente:
#
#  ⊞ □ Sub-Ramo A          ← expander com checkbox "selecionar todos"
#       □ Especialização 1
#       □ Especialização 2
#  ⊟ □ Sub-Ramo B          ← expandido mostra filhos
#       □ Especialização 3
#
# Retorna (ramos_nome_sel, sub_ramos_sel) onde:
#   ramos_nome_sel = Ramo_Nome (especializações) individualmente marcadas
#   sub_ramos_sel  = Sub_Ramos cujo checkbox-pai foi marcado (todos os filhos)
# ──────────────────────────────────────────────────────────────
def filtro_arvore(df: pd.DataFrame, key_prefix: str) -> list:
    """
    Árvore 2 níveis: Sub-Ramo (expansível) → Especialização/Tipo de Negócio (checkbox).
    Retorna lista de Ramo_Nome selecionados para filtrar o dataframe.
    """
    # Monta dicionário  Sub-Ramo → [Especialização, ...]
    tree: dict[str, list[str]] = {}
    validos = df[["Sub_Ramo","Ramo_Nome"]].drop_duplicates()
    validos = validos[
        (validos["Sub_Ramo"]  != "Não informado") &
        (validos["Ramo_Nome"] != "Não informado")
    ].sort_values(["Sub_Ramo","Ramo_Nome"])

    for _, row in validos.iterrows():
        sr = str(row["Sub_Ramo"])
        rn = str(row["Ramo_Nome"])
        tree.setdefault(sr, [])
        if rn not in tree[sr]:
            tree[sr].append(rn)

    selecionados: list[str] = []

    # CSS extra para simular o visual PBI (símbolo ⊞/⊟ + indentação)
    st.markdown("""
    <style>
    /* Remove padding padrão dos expanders para aproximar do visual PBI */
    div[data-testid="stExpander"] > details > summary {
        font-size: 0.88rem !important;
        font-weight: 600 !important;
        padding: 4px 6px !important;
        color: #1A1A2E !important;
    }
    div[data-testid="stExpander"] > details {
        border: none !important;
        border-bottom: 1px solid #e0e0e0 !important;
        box-shadow: none !important;
    }
    /* checkbox de especialização — indentado */
    div[data-testid="stExpander"] .stCheckbox label {
        font-size: 0.82rem !important;
        padding-left: 12px !important;
        color: #333 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    for sr in sorted(tree.keys()):
        especializacoes = tree[sr]
        # expander recolhido por padrão (⊞), expandido mostra filhos (⊟)
        with st.expander(f"□  {sr}", expanded=False):
            # checkbox "pai" — seleciona todos os filhos deste sub-ramo
            tudo = st.checkbox(
                f"Selecionar todos",
                key=f"{key_prefix}_ALL_{sr}",
                value=False,
            )
            for espec in especializacoes:
                # chave sanitizada (sem caracteres problemáticos)
                safe_key = f"{key_prefix}_{sr}_{espec}".replace(" ","_")[:120]
                marcado = st.checkbox(
                    espec,
                    key=safe_key,
                    value=tudo,
                )
                if marcado or tudo:
                    selecionados.append(espec)

    return list(set(selecionados))


# ──────────────────────────────────────────────────────────────
# COMPONENTES VISUAIS
# ──────────────────────────────────────────────────────────────
LOGO_EMP  = _logo_path("Empresa_Simbolo.png")
LOGO_META = _logo_path("metaday_logo_sem_fundo.png")

def render_header(titulo="Segmento: EMPRESAS"):
    l_emp  = img_to_base64(LOGO_EMP)  if LOGO_EMP  else ""
    l_meta = img_to_base64(LOGO_META) if LOGO_META else ""
    img_e  = f'<img src="{l_emp}"  class="header-logo-left">'  if l_emp  else "🏢"
    img_d  = f'<img src="{l_meta}" class="header-logo-right">' if l_meta else "⬤"
    st.markdown(f"""
    <div class="header-row">
        <div>{img_e}</div>
        <h1 class="titulo-principal">{titulo}</h1>
        <div>{img_d}</div>
    </div>""", unsafe_allow_html=True)

def kpi_card(label, value):
    return (f'<div class="kpi-card">'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div></div>')

def _bl(fig, h=260):
    fig.update_layout(height=h, margin=dict(t=10,b=10,l=10,r=10),
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    return fig

def grafico_porte(df, height=230, orientacao="v"):
    porte = (df["Porte"].value_counts()
               .reindex(["Pequeno","Médio","Grande"]).dropna().reset_index())
    porte.columns = ["Porte","Qtde"]
    cor_p = {"Pequeno":COR_GOLD,"Médio":COR_NAVY,"Grande":COR_CINZA}
    if orientacao == "v":
        fig = px.bar(porte, x="Porte", y="Qtde", text="Qtde",
                     color="Porte", color_discrete_map=cor_p)
        fig.update_traces(textposition="outside", textfont_size=13)
        fig.update_layout(xaxis=dict(title="Porte",tickfont_size=12),
                          yaxis=dict(visible=False), showlegend=False)
    else:
        fig = px.bar(porte, y="Porte", x="Qtde", orientation="h",
                     text="Qtde", color="Porte", color_discrete_map=cor_p)
        fig.update_traces(textposition="inside", textfont_size=13, textfont_color="white")
        fig.update_layout(xaxis=dict(visible=False), yaxis_title="", showlegend=False)
    return _bl(fig, height)

def grafico_presenca(df, height=230):
    pres = df["Presenca_Digital"].value_counts().reset_index()
    pres.columns = ["Status","Qtde"]
    fig = px.pie(pres, names="Status", values="Qtde", color="Status",
                 color_discrete_map={"SIM":COR_GOLD,"NÃO":"#D8D8D8"}, hole=0.0)
    fig.update_traces(textposition="outside", textinfo="label+value+percent",
                      textfont_size=11, pull=[0,0.04])
    fig.update_layout(showlegend=False, margin=dict(t=20,b=20,l=10,r=10),
                      height=height, paper_bgcolor="rgba(0,0,0,0)")
    return fig

def grafico_expediente(df, height=230):
    exp_c = df["Expediente"].value_counts().reset_index()
    exp_c.columns = ["Tipo","Qtde"]
    cores = {t: (COR_NAVY if "24" in str(t).upper() else
                 COR_AZUL if "HC" in str(t).upper() else COR_CINZA)
             for t in exp_c["Tipo"]}
    fig = px.pie(exp_c, names="Tipo", values="Qtde",
                 color="Tipo", color_discrete_map=cores, hole=0.0)
    fig.update_traces(textposition="outside", textinfo="label+percent",
                      textfont_size=11)
    fig.update_layout(showlegend=False, margin=dict(t=20,b=20,l=10,r=10),
                      height=height, paper_bgcolor="rgba(0,0,0,0)")
    return fig

def grafico_tempo_fatec(df, height=240):
    ordem = ["11 a 20 minutos","6 a 10 minutos","0 a 5 minutos","Mais de 20 minutos"]
    tc = df["Faixa_Tempo"].value_counts().reindex(ordem).fillna(0).reset_index()
    tc.columns = ["Faixa","Qtde"]
    max_v = tc["Qtde"].max() or 1
    fig = go.Figure()
    fig.add_trace(go.Bar(y=tc["Faixa"], x=[max_v*1.05]*len(tc), orientation="h",
                         marker_color="#E0E0E0", showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Bar(y=tc["Faixa"], x=tc["Qtde"], orientation="h",
                         marker_color=COR_GOLD,
                         text=tc["Qtde"].astype(int), textposition="inside",
                         textfont=dict(size=13,color="white"),
                         showlegend=False,
                         hovertemplate="%{y}: %{x}<extra></extra>"))
    fig.update_layout(barmode="overlay",
                      xaxis=dict(visible=False, range=[0,max_v*1.15]),
                      yaxis=dict(title="", tickfont_size=11, autorange="reversed"),
                      margin=dict(t=5,b=5,l=0,r=60), height=height,
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    return fig

def grafico_dist_fatec(df, height=240):
    ordem = ["3 a 4km","0 a 2km","5 a 7km","Mais de 7km"]
    dc = df["Faixa_Dist"].value_counts().reindex(ordem).fillna(0).reset_index()
    dc.columns = ["Faixa","Qtde"]
    max_v = dc["Qtde"].max() or 1
    fig = go.Figure()
    fig.add_trace(go.Bar(x=dc["Faixa"], y=[max_v*1.08]*len(dc),
                         marker_color="#E0E0E0", showlegend=False, hoverinfo="skip"))
    fig.add_trace(go.Bar(x=dc["Faixa"], y=dc["Qtde"], marker_color=COR_GOLD,
                         text=dc["Qtde"].astype(int), textposition="inside",
                         textfont=dict(size=13,color="white"),
                         showlegend=False,
                         hovertemplate="%{x}: %{y}<extra></extra>"))
    fig.update_layout(barmode="overlay",
                      xaxis=dict(title="",tickfont_size=11),
                      yaxis=dict(visible=False, range=[0,max_v*1.18]),
                      margin=dict(t=30,b=5,l=5,r=5), height=height,
                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
    return fig


# ══════════════════════════════════════════════════════════════
# PÁGINA 1 — VISÃO GERAL
# ══════════════════════════════════════════════════════════════
def pagina_visao_geral(df, df_distrito):
    render_header()
    c1,c2,c3 = st.columns(3)
    for col, label, val in [
        (c1,"Empresas Catalogadas",          f"{len(df):,}".replace(",",".")),
        (c2,"Distritos Considerados",        str(df["Distrito"].nunique())),
        (c3,"Especializações Identificadas", str(df["Ramo_Nome"].nunique())),
    ]:
        with col: st.markdown(kpi_card(label,val), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_e, col_d = st.columns([1, 2.5])

    with col_e:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_presenca(df,230), use_container_width=True)
        st.markdown('<div class="section-title">Porte das Empresas</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_porte(df,190,"h"), use_container_width=True)

    with col_d:
        st.markdown('<div class="section-title">Distribuição Geográfica por Distrito</div>',
                    unsafe_allow_html=True)
        if ("Lat" in df.columns and "Lon" in df.columns and df["Lat"].notna().sum() > 10):
            df_m = df.dropna(subset=["Lat","Lon"])
            fig_m = px.scatter_mapbox(df_m, lat="Lat", lon="Lon", color="Distrito",
                                      size_max=12, zoom=12, height=440,
                                      mapbox_style="open-street-map",
                                      color_discrete_sequence=px.colors.qualitative.Bold,
                                      hover_data={"Lat":False,"Lon":False,"Porte":True,"Distrito":True})
            fig_m.update_layout(margin=dict(t=0,b=0,l=0,r=0), showlegend=False,
                                  paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_m, use_container_width=True)
        else:
            dist_c = df.groupby("Distrito").size().reset_index(name="Qtde").sort_values("Qtde",ascending=False)
            fig_b = px.bar(dist_c, x="Distrito", y="Qtde", text="Qtde", color="Qtde",
                           color_continuous_scale=[[0,COR_GOLD],[1,COR_NAVY]])
            fig_b.update_traces(textposition="outside")
            fig_b.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="Empresas")
            st.plotly_chart(_bl(fig_b,440), use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PÁGINA 2 — DISTRITOS & RAMOS
# ══════════════════════════════════════════════════════════════
def pagina_distritos(df):
    render_header()
    col_e, col_d = st.columns([1, 2.5])

    with col_e:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_presenca(df,200), use_container_width=True)

    with col_d:
        st.markdown('<div class="section-title">Empresas por Distrito</div>', unsafe_allow_html=True)
        dist_c = df.groupby("Distrito").size().reset_index(name="Qtde")
        fig_t = px.treemap(dist_c, path=["Distrito"], values="Qtde", color="Qtde",
                           color_continuous_scale=[[0,"#1a237e"],[.15,"#283593"],[.3,"#5c6bc0"],
                               [.45,"#00897b"],[.6,"#43a047"],[.75,"#e53935"],[.9,"#8e24aa"],[1,"#00acc1"]])
        fig_t.update_traces(textinfo="label+value", textfont_size=13, textposition="bottom left")
        fig_t.update_layout(coloraxis_showscale=False, margin=dict(t=5,b=5,l=5,r=5),
                             height=280, paper_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_t, use_container_width=True)

    c1,c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">Tempo de Existência</div>', unsafe_allow_html=True)
        ordem = ["Mais de 20 anos","0 a 2 anos","11 a 20 anos","6 a 10 anos","3 a 5 anos","Sem dados"]
        fi = df["Faixa_Idade"].value_counts().reindex(ordem).dropna().reset_index()
        fi.columns = ["Faixa","Qtde"]
        fig_fi = px.bar(fi, y="Faixa", x="Qtde", orientation="h", text="Qtde", color="Qtde",
                        color_continuous_scale=[[0,"#90CAF9"],[.5,"#1565C0"],[1,COR_NAVY]])
        fig_fi.update_traces(textposition="outside", textfont_size=12)
        fig_fi.update_layout(coloraxis_showscale=False,
                              xaxis=dict(visible=False), yaxis_title="", showlegend=False)
        st.plotly_chart(_bl(fig_fi,270), use_container_width=True)

    with c2:
        st.markdown('<div class="section-title">Quantidade por Ramo</div>', unsafe_allow_html=True)
        ramo_c = df["Ramo_Principal"].value_counts().reset_index()
        ramo_c.columns = ["Ramo","Qtde"]
        ramo_c = ramo_c[ramo_c["Ramo"] != "Não informado"].head(10)
        fig_r = px.bar(ramo_c, x="Ramo", y="Qtde", text="Qtde",
                       color_discrete_sequence=[COR_GOLD])
        fig_r.update_traces(textposition="outside", textfont_size=11)
        fig_r.update_layout(xaxis=dict(title="",tickangle=-30,tickfont_size=10),
                             yaxis=dict(visible=False), showlegend=False)
        st.plotly_chart(_bl(fig_r,270), use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PÁGINA 3 — FATEC & SUB-RAMOS
# ══════════════════════════════════════════════════════════════
def pagina_fatec(df):
    render_header()
    c1,c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-title">Tempo da FATEC até a Empresa</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_tempo_fatec(df,240), use_container_width=True)
    with c2:
        st.markdown('<div class="section-title">Distância da FATEC até a Empresa</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_dist_fatec(df,240), use_container_width=True)

    col_f, col_t = st.columns([1, 3])

    with col_f:
        # ── Árvore Sub-Ramo → Especialização (igual PBI) ───────────
        st.markdown('<div class="section-title">Ramo de Atividade</div>', unsafe_allow_html=True)
        espec_sel = filtro_arvore(df, key_prefix="p3")

    with col_t:
        st.markdown('<div class="section-title">Sub-Divisão dos Ramos de Atividade</div>',
                    unsafe_allow_html=True)
        # Filtra por especializações marcadas; sem seleção → mostra tudo
        df_f = df[df["Ramo_Nome"].isin(espec_sel)] if espec_sel else df
        sub_c = (df_f.groupby(["Sub_Ramo","Ramo_Nome"])
                      .size().reset_index(name="Qtde")
                      .query("Sub_Ramo != 'Não informado' and Qtde > 0"))
        if not sub_c.empty:
            # Treemap 2 níveis: Sub-Ramo → Especialização (sem Ramo Principal)
            fig_s = px.treemap(
                sub_c,
                path=["Sub_Ramo","Ramo_Nome"],
                values="Qtde",
                color="Sub_Ramo",
                color_discrete_sequence=TREEMAP_COLORS,
            )
            fig_s.update_traces(
                textinfo="label+value",
                textfont_size=11,
                textposition="bottom left",
            )
            fig_s.update_layout(
                height=390,
                margin=dict(t=5,b=5,l=5,r=5),
                paper_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_s, use_container_width=True)
        else:
            st.info("Selecione especializações no painel à esquerda para filtrar a sub-divisão.")


# ══════════════════════════════════════════════════════════════
# PÁGINA 4 — MAPA POR SEGMENTO
# ══════════════════════════════════════════════════════════════
def pagina_mapa_ramos(df):
    render_header()
    col_f, col_m = st.columns([1, 3])

    with col_f:
        # ── Árvore Sub-Ramo → Especialização (igual PBI) ───────────
        st.markdown('<div class="section-title">Segmentação dos Ramos</div>',
                    unsafe_allow_html=True)
        espec_sel = filtro_arvore(df, key_prefix="p4")

    with col_m:
        df_f = df[df["Ramo_Nome"].isin(espec_sel)] if espec_sel else df
        tem_coords = ("Lat" in df_f.columns and "Lon" in df_f.columns
                      and df_f["Lat"].notna().sum() > 5)
        if tem_coords:
            df_m = df_f.dropna(subset=["Lat","Lon"])
            fig_m = px.scatter_mapbox(
                df_m, lat="Lat", lon="Lon", color="Sub_Ramo",
                hover_data={"Lat":False,"Lon":False,"Porte":True,"Distrito":True,
                            "Sub_Ramo":True,"Ramo_Nome":True},
                zoom=12, height=580, mapbox_style="open-street-map",
                color_discrete_sequence=px.colors.qualitative.Alphabet,
            )
            fig_m.update_layout(margin=dict(t=0,b=0,l=0,r=0),
                                  legend=dict(title="Sub-Ramo",font_size=10),
                                  paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_m, use_container_width=True)
        else:
            # fallback sem coords: barras de Sub-Ramo
            sub_c = df_f["Sub_Ramo"].value_counts().head(20).reset_index()
            sub_c.columns = ["Sub_Ramo","Qtde"]
            sub_c = sub_c[sub_c["Sub_Ramo"] != "Não informado"]
            fig_b = px.bar(sub_c, x="Qtde", y="Sub_Ramo", orientation="h",
                           text="Qtde", color="Sub_Ramo",
                           color_discrete_sequence=px.colors.qualitative.Bold)
            fig_b.update_traces(textposition="outside")
            fig_b.update_layout(showlegend=False,
                                  xaxis=dict(visible=False), yaxis_title="")
            st.plotly_chart(_bl(fig_b,580), use_container_width=True)


# ══════════════════════════════════════════════════════════════
# PÁGINA 5 — DISTRITOS POR SEGMENTO
# ══════════════════════════════════════════════════════════════
def pagina_distritos_segmento(df):
    render_header()
    col_f, col_t = st.columns([1, 3])

    with col_f:
        # ── Árvore Sub-Ramo → Especialização (igual PBI) ───────────
        st.markdown('<div class="section-title">Segmentação dos Ramos</div>',
                    unsafe_allow_html=True)
        espec_sel = filtro_arvore(df, key_prefix="p5")

    with col_t:
        df_f = df[df["Ramo_Nome"].isin(espec_sel)] if espec_sel else df
        st.markdown('<div class="section-title">Distritos que têm os Segmentos selecionados</div>',
                    unsafe_allow_html=True)
        dist_c = df_f.groupby("Distrito").size().reset_index(name="Qtde")
        if not dist_c.empty:
            fig_t = px.treemap(dist_c, path=["Distrito"], values="Qtde", color="Qtde",
                               color_continuous_scale=[[0,"#1a237e"],[.15,"#283593"],[.3,"#5c6bc0"],
                                   [.45,"#00897b"],[.6,"#43a047"],[.75,"#e53935"],
                                   [.9,"#8e24aa"],[1,"#00acc1"]])
            fig_t.update_traces(textinfo="label+value", textfont_size=14, textposition="bottom left")
            fig_t.update_layout(height=580, margin=dict(t=5,b=5,l=5,r=5),
                                  coloraxis_showscale=False, paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_t, use_container_width=True)
        else:
            st.info("Selecione segmentos no painel à esquerda para ver os distritos.")


# ══════════════════════════════════════════════════════════════
# PÁGINA 6 — IBGE / PERFIL
# ══════════════════════════════════════════════════════════════
def pagina_ibge(df, df_distrito):
    render_header()
    c1,c2,c3 = st.columns(3)
    with c1:
        st.markdown('<div class="section-title">Porte das Empresas</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_porte(df,230,"v"), use_container_width=True)
    with c2:
        st.markdown('<div class="section-title">Presença Digital</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_presenca(df,230), use_container_width=True)
    with c3:
        st.markdown('<div class="section-title">Expediente</div>', unsafe_allow_html=True)
        st.plotly_chart(grafico_expediente(df,230), use_container_width=True)

    st.markdown("---")
    st.markdown(f'<div style="font-size:1.05rem;font-weight:800;color:{COR_TEXTO};'
                f'margin-bottom:6px;">📊 Dados IBGE / Censo 2022</div>', unsafe_allow_html=True)
    if not df_distrito.empty:
        df_distrito.columns = df_distrito.columns.str.strip()
        st.dataframe(df_distrito, use_container_width=True, height=340)
    else:
        tab = df.groupby("Distrito").size().reset_index(name="Na Amostra").sort_values("Distrito")
        total = pd.DataFrame([{"Distrito":"Total","Na Amostra":len(df)}])
        st.dataframe(pd.concat([tab,total], ignore_index=True),
                     use_container_width=True, height=340)
    st.caption("Dados, exceto NA AMOSTRA, cruzando Prefeitura SP, SEADE e OBSERVASAMPA. "
               "Populações projetadas baseadas no Censo 2010.")


# ──────────────────────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────────────────────
def render_sidebar(df):
    l_meta = img_to_base64(LOGO_META) if LOGO_META else ""
    with st.sidebar:
        if l_meta:
            st.markdown(
                f'<div style="text-align:center;padding:10px 0 4px 0;">'
                f'<img src="{l_meta}" style="width:90px;height:90px;'
                f'object-fit:contain;border-radius:50%;"></div>',
                unsafe_allow_html=True)
        st.markdown("""
        <div style="text-align:center;padding:0 0 14px 0;">
            <div style="font-size:1.05rem;font-weight:800;letter-spacing:1px;">METADAY 2025</div>
            <div style="font-size:0.75rem;opacity:0.7;">Fatec Sebrae · Ciência de Dados</div>
        </div>""", unsafe_allow_html=True)
        st.markdown("---")

        pagina = st.radio("Navegação", options=[
            "1 · Visão Geral",
            "2 · Distritos & Ramos",
            "3 · FATEC & Sub-Ramos",
            "4 · Mapa por Segmento",
            "5 · Distritos por Segmento",
            "6 · Dados IBGE / Perfil",
        ], label_visibility="collapsed")

        st.markdown("---")
        st.markdown("**🔍 Filtros Globais**")

        distritos = sorted([d for d in df["Distrito"].dropna().unique()
                            if d not in ("Não informado","")])
        sel_dist  = st.multiselect("Distrito",  distritos, default=distritos)

        portes = sorted([p for p in df["Porte"].dropna().unique()
                         if p not in ("Não informado","")])
        sel_porte = st.multiselect("Porte", portes, default=portes)

        sel_pres = st.multiselect("Presença Digital", ["SIM","NÃO"], default=["SIM","NÃO"])

        st.markdown("---")
        mask = (df["Distrito"].isin(sel_dist) &
                df["Porte"].isin(sel_porte) &
                df["Presenca_Digital"].isin(sel_pres))
        st.caption(f"Total filtrado: **{mask.sum():,}** empresas")

    return pagina, sel_dist, sel_porte, sel_pres


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    caminho = encontrar_excel()
    if caminho:
        df_raw, df_dist = carregar_dados(caminho)
    else:
        render_header()
        st.info("📂 Coloque o `.xlsx` na mesma pasta que o `app.py` ou faça upload abaixo.")
        arq = st.file_uploader("Planilha Excel", type=["xlsx","xls"])
        if arq is None:
            st.stop()
        df_raw, df_dist = carregar_dados(arq)

    pagina, sel_dist, sel_porte, sel_pres = render_sidebar(df_raw)

    df = df_raw.copy()
    if sel_dist:  df = df[df["Distrito"].isin(sel_dist)]
    if sel_porte: df = df[df["Porte"].isin(sel_porte)]
    if sel_pres:  df = df[df["Presenca_Digital"].isin(sel_pres)]

    rotas = {
        "1": lambda: pagina_visao_geral(df, df_dist),
        "2": lambda: pagina_distritos(df),
        "3": lambda: pagina_fatec(df),
        "4": lambda: pagina_mapa_ramos(df),
        "5": lambda: pagina_distritos_segmento(df),
        "6": lambda: pagina_ibge(df, df_dist),
    }
    rotas[pagina[0]]()

if __name__ == "__main__":
    main()
