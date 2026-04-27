import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import requests
from github import Github

# ====================== FUNÇÃO KPI CARD PROFISSIONAL ======================
def render_kpi_card(label, value, delta=None, icon="📊", color="#1F4788"):
    """Renderiza card KPI profissional com HTML/CSS - substitui st.metric()"""
    delta_html = ""
    if delta:
        delta_val = str(delta).replace("%","").replace(",","").replace("+","").strip()
        try:
            delta_color = "#10B981" if float(delta_val) >= 0 else "#EF4444"
        except:
            delta_color = "#10B981" if "+" in str(delta) else "#EF4444"
        delta_html = f'<div style="color:{delta_color};font-size:0.875rem;font-weight:600;margin-top:0.5rem;">{delta}</div>'
    
    st.markdown(f"""
    <div style="background:white;padding:1.5rem;border-radius:15px;
                box-shadow:0 4px 12px rgba(0,0,0,0.05);border-left:4px solid {color};
                height:140px;display:flex;flex-direction:column;justify-content:space-between;">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
            <div style="font-size:0.8rem;color:#6B7280;font-weight:600;
                        text-transform:uppercase;letter-spacing:0.05em;">{label}</div>
            <div style="font-size:1.75rem;">{icon}</div>
        </div>
        <div>
            <div style="font-size:1.75rem;font-weight:700;color:#1F2937;">{value}</div>
            {delta_html}
        </div>
    </div>
    """, unsafe_allow_html=True)

# Configuração da página
st.set_page_config(
    page_title="Dashboard BI Medtextil", 
    layout="wide", 
    initial_sidebar_state="expanded",
    page_icon="https://i.imgur.com/gt3rgyL.png"  # Logo Medtextil
)

# ====================== CONFIGURAÇÃO DO ÍCONE ======================
# O Streamlit gerencia automaticamente o ícone via page_icon
# Nenhuma configuração adicional é necessária

# ====================== CSS CUSTOMIZADO - UX/UI PROFISSIONAL ======================
st.markdown("""
<style>
/* ── Importar fonte Inter ── */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

/* ── Reset e fonte base ── */
html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', Roboto, sans-serif !important;
}

/* ── Fundo cinza muito claro ── */
.stApp {
    background-color: var(--background-color) !important;
}

/* ── Sidebar limpa e profissional ── */
[data-testid="stSidebar"] {
    background: var(--secondary-background-color) !important;
    border-right: 1px solid rgba(128,128,128,0.2) !important;
    box-shadow: 2px 0 8px rgba(0,0,0,0.04) !important;
}
[data-testid="stSidebar"] .stMarkdown h1,
[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 {
    color: #4A7BC8 !important;
}

/* ── Botões de navegação da sidebar estilizados ── */
[data-testid="stSidebar"] [data-testid="stRadio"] {
    gap: 6px !important;
}

[data-testid="stSidebar"] [data-testid="stRadio"] label {
    background: #F8F9FA !important;
    border: 1px solid #E9ECEF !important;
    border-radius: 10px !important;
    padding: 12px 16px !important;
    margin: 0 !important;
    cursor: pointer !important;
    transition: all 0.2s ease !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    display: block !important;
    width: 100% !important;
}

[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
    background: #EEF2F7 !important;
    border-color: #4A7BC8 !important;
    transform: translateX(3px) !important;
}

/* Botão selecionado na sidebar */
[data-testid="stSidebar"] [data-testid="stRadio"] label[data-checked="true"] {
    background: linear-gradient(135deg, #1F4788 0%, #2D5AA0 100%) !important;
    border-color: #4A7BC8 !important;
    color: white !important;
    font-weight: 600 !important;
    box-shadow: 0 2px 8px rgba(31, 71, 136, 0.2) !important;
}

/* Esconder radio button padrão */
[data-testid="stSidebar"] [data-testid="stRadio"] input[type="radio"] {
    display: none !important;
}

/* ── Cards de métricas (st.metric) ── */
[data-testid="stMetric"] {
    background: var(--secondary-background-color) !important;
    border-radius: 12px !important;
    padding: 18px 20px !important;
    border-left: 4px solid #1F4788 !important;
    box-shadow: 0 2px 8px rgba(31, 71, 136, 0.08), 0 1px 3px rgba(0,0,0,0.04) !important;
    transition: box-shadow 0.2s ease, transform 0.2s ease !important;
}
[data-testid="stMetric"]:hover {
    box-shadow: 0 6px 20px rgba(31, 71, 136, 0.14) !important;
    transform: translateY(-2px) !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.78rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.04em !important;
    color: #6C757D !important;
}
[data-testid="stMetricValue"] {
    font-size: 1.5rem !important;
    font-weight: 700 !important;
    color: #4A7BC8 !important;
}

/* ── Cor alternada para cards de destaque ── */
div[data-testid="column"]:nth-child(2) [data-testid="stMetric"] {
    border-left-color: #2E86AB !important;
}
div[data-testid="column"]:nth-child(3) [data-testid="stMetric"] {
    border-left-color: #28A745 !important;
}
div[data-testid="column"]:nth-child(4) [data-testid="stMetric"] {
    border-left-color: #F4A261 !important;
}

/* ── Botões modernos ── */
.stButton > button {
    border-radius: 8px !important;
    font-weight: 600 !important;
    font-size: 0.875rem !important;
    padding: 0.5rem 1.25rem !important;
    border: 1.5px solid #1F4788 !important;
    color: #4A7BC8 !important;
    background: var(--secondary-background-color) !important;
    transition: all 0.2s ease !important;
    box-shadow: 0 1px 4px rgba(31, 71, 136, 0.10) !important;
    letter-spacing: 0.01em !important;
}
.stButton > button:hover {
    background: #1F4788 !important;
    color: #FFFFFF !important;
    border-color: #4A7BC8 !important;
    box-shadow: 0 4px 12px rgba(31, 71, 136, 0.25) !important;
    transform: translateY(-1px) !important;
}
.stButton > button[kind="primary"] {
    background: #1F4788 !important;
    color: #FFFFFF !important;
    border-color: #4A7BC8 !important;
}
.stButton > button[kind="primary"]:hover {
    background: #163561 !important;
    border-color: #163561 !important;
}

/* ── Botão de Download ── */
[data-testid="stDownloadButton"] > button {
    border-radius: 8px !important;
    font-weight: 600 !important;
    background: var(--secondary-background-color) !important;
    color: #4A7BC8 !important;
    border: 1.5px solid #1F4788 !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #1F4788 !important;
    color: #FFFFFF !important;
}

/* ── Cards de navegação (Home) ── */
.stButton > button p {
    font-size: 0.9rem !important;
}

/* ── Títulos de página ── */
.stApp h1 {
    color: #2C5AA0 !important;
    font-weight: 700 !important;
    font-size: 1.75rem !important;
    letter-spacing: -0.02em !important;
}
.stApp h2 {
    color: #4A7BC8 !important;
    font-weight: 600 !important;
    font-size: 1.3rem !important;
}
.stApp h3 {
    color: #2E4A7A !important;
    font-weight: 600 !important;
}

/* ── Selectbox e inputs ── */
.stSelectbox > div > div,
.stTextInput > div > div > input,
.stDateInput > div > div > input,
.stNumberInput > div > div > input {
    border-radius: 8px !important;
    border-color: #DEE2E6 !important;
    font-size: 0.9rem !important;
}
.stSelectbox > div > div:focus-within,
.stTextInput > div > div:focus-within {
    border-color: #4A7BC8 !important;
    box-shadow: 0 0 0 2px rgba(31, 71, 136, 0.15) !important;
}

/* ── Tabs ── */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: var(--secondary-background-color) !important;
    border-radius: 10px !important;
    padding: 4px !important;
    border-bottom: none !important;
    gap: 4px !important;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius: 8px !important;
    font-weight: 500 !important;
    font-size: 0.875rem !important;
    color: #6C757D !important;
    padding: 6px 16px !important;
}
[data-testid="stTabs"] [aria-selected="true"] {
    background: #1F4788 !important;
    color: #FFFFFF !important;
    font-weight: 600 !important;
}

/* ── Dataframes / tabelas ── */
[data-testid="stDataFrame"] {
    border-radius: 10px !important;
    overflow: hidden !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06) !important;
}

/* ── Divisores ── */
hr {
    border-color: #E9ECEF !important;
    margin: 1rem 0 !important;
}

/* ── Alertas e info boxes ── */
[data-testid="stAlert"] {
    border-radius: 10px !important;
}

/* ── Sidebar logo area ── */
.sidebar-logo-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 16px 8px 8px 8px;
    border-bottom: 1px solid #E9ECEF;
    margin-bottom: 12px;
}
.sidebar-user-badge {
    background: #F0F4FF;
    border: 1px solid #C5D5F0;
    border-radius: 8px;
    padding: 8px 12px;
    font-size: 0.85rem;
    color: #4A7BC8;
    font-weight: 600;
    width: 100%;
    text-align: center;
    margin-top: 8px;
}

/* ── Barra de filtros no topo ── */
.filter-bar {
    background: #FFFFFF;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    border: 1px solid #E9ECEF;
}

/* ── Cabeçalho da página ── */
.page-header {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 4px;
}
.page-subtitle {
    color: #6C757D;
    font-size: 0.9rem;
    margin-bottom: 20px;
    margin-top: -4px;
}

/* ── Radio buttons estilo pill ── */
.stRadio > div {
    flex-direction: row !important;
    gap: 8px !important;
}
.stRadio > div label {
    border-radius: 20px !important;
    padding: 4px 14px !important;
    border: 1.5px solid #DEE2E6 !important;
    font-size: 0.875rem !important;
    cursor: pointer !important;
    transition: all 0.15s !important;
}
.stRadio > div label:has(input:checked) {
    background: #1F4788 !important;
    color: #FFFFFF !important;
    border-color: #4A7BC8 !important;
}

/* ── Spinner ── */
.stSpinner > div {
    border-top-color: #4A7BC8 !important;
}

/* ── Sidebar navigation radio ── */
[data-testid="stSidebar"] .stRadio > div {
    flex-direction: column !important;
    gap: 4px !important;
}
[data-testid="stSidebar"] .stRadio > div label {
    border-radius: 10px !important;
    padding: 11px 14px !important;
    border: none !important;
    font-size: 0.92rem !important;
    font-weight: 500 !important;
    color: #495057 !important;
    background: transparent !important;
    width: 100% !important;
    transition: all 0.15s ease !important;
    margin-bottom: 2px !important;
}
[data-testid="stSidebar"] .stRadio > div label:has(input:checked) {
    background: linear-gradient(135deg, #F0F4FF, #E4EDFC) !important;
    color: #4A7BC8 !important;
    border-left: 4px solid #1F4788 !important;
    border-radius: 0 10px 10px 0 !important;
    font-weight: 700 !important;
    box-shadow: 0 1px 4px rgba(31,71,136,0.10) !important;
}
[data-testid="stSidebar"] .stRadio > div label:hover {
    background: #F4F7FD !important;
    color: #4A7BC8 !important;
    transform: translateX(2px) !important;
}

/* ── Success/warning/error messages ── */
.stSuccess, .stInfo, .stWarning, .stError {
    border-radius: 8px !important;
    font-size: 0.875rem !important;
}

/* ── Métricas: rótulo uppercase corporativo ── */
[data-testid="stMetricLabel"] p {
    text-transform: uppercase !important;
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.06em !important;
    color: #8A96A8 !important;
}
[data-testid="stMetricValue"] {
    font-size: 1.45rem !important;
    font-weight: 700 !important;
    color: #2C5AA0 !important;
}

/* ── Subheaders sem emoji weight ── */
h2, h3 {
    font-weight: 600 !important;
    letter-spacing: -0.01em !important;
}

/* ── Expander painel de controle ── */
[data-testid="stExpander"] summary {
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    color: #6C757D !important;
}

/* ── Captions dos filtros ── */
.stCaption p {
    font-size: 0.7rem !important;
    color: #8A96A8 !important;
    margin-top: -4px !important;
    font-weight: 500 !important;
    letter-spacing: 0.03em !important;
}


/* ══════════════════════════════════════════════════════
   DARK MODE — sobrescreve cores fixas quando o tema
   escuro está ativo. Usa prefers-color-scheme E o
   atributo data-theme que o Streamlit injeta no <body>.
   ══════════════════════════════════════════════════════ */

/* Detectar dark mode via atributo do Streamlit */
@media (prefers-color-scheme: dark) {
    .stApp { background-color: #0E1117 !important; }
}

/* Streamlit injeta data-theme="dark" no root quando tema escuro ativo */
[data-theme="dark"] .stApp,
.stApp[data-theme="dark"] {
    background-color: #0E1117 !important;
}

/* Sidebar dark */
[data-theme="dark"] section[data-testid="stSidebar"],
[data-theme="dark"] [data-testid="stSidebar"] {
    background: #1A1D24 !important;
    border-right-color: #2D3139 !important;
}

/* Métricas dark */
[data-theme="dark"] [data-testid="stMetric"] {
    background: #1A1D24 !important;
    box-shadow: 0 1px 6px rgba(0,0,0,0.3) !important;
}
[data-theme="dark"] [data-testid="stMetricLabel"] p {
    color: #8A96A8 !important;
}
[data-theme="dark"] [data-testid="stMetricValue"] {
    color: #E8EDF5 !important;
}

/* Filtros dark */
[data-theme="dark"] div[style*="background:#F2F5FA"],
[data-theme="dark"] div[style*="background: #F2F5FA"] {
    background: #1A1D24 !important;
    border-color: #2D3139 !important;
}

/* Botões gerais dark */
[data-theme="dark"] .stButton > button {
    background: #1A1D24 !important;
    color: #A8C4E8 !important;
    border-color: #2D5AA0 !important;
}
[data-theme="dark"] .stButton > button:hover {
    background: #1F4788 !important;
    color: #FFFFFF !important;
}

/* Expander dark */
[data-theme="dark"] [data-testid="stExpander"] {
    background: #1A1D24 !important;
    border-color: #2D3139 !important;
}

/* Dataframes dark */
[data-theme="dark"] [data-testid="stDataFrame"] {
    background: #1A1D24 !important;
}

/* Tabs dark */
[data-theme="dark"] [data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: #1A1D24 !important;
}
[data-theme="dark"] [data-testid="stTabs"] [data-baseweb="tab"] {
    color: #8A96A8 !important;
}

/* Radio sidebar dark */
[data-theme="dark"] section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"] {
    color: #A8B4C4 !important;
}
[data-theme="dark"] section[data-testid="stSidebar"] .stRadio label[aria-checked="true"] {
    background: linear-gradient(90deg,#1A2A45,#1F3560) !important;
    border-left-color: #4A7BC8 !important;
}
[data-theme="dark"] section[data-testid="stSidebar"] .stRadio label[aria-checked="true"] p {
    color: #7EB3F7 !important;
}

/* Breadcrumb e textos de apoio dark */
[data-theme="dark"] .stMarkdown p {
    color: #C4CDD9 !important;
}

/* Inputs dark */
[data-theme="dark"] .stSelectbox > div > div,
[data-theme="dark"] .stTextInput > div > div > input,
[data-theme="dark"] .stDateInput > div > div > input {
    background: #1A1D24 !important;
    border-color: #2D3139 !important;
    color: #E0E6EF !important;
}

/* Download button dark */
[data-theme="dark"] [data-testid="stDownloadButton"] > button {
    background: #1A2A45 !important;
    color: #7EB3F7 !important;
    border-color: #2D5AA0 !important;
}
[data-theme="dark"] [data-testid="stDownloadButton"] > button:hover {
    background: #1F4788 !important;
    color: #FFFFFF !important;
}

/* Cabeçalho títulos dark */
[data-theme="dark"] .stApp h1 { color: #E8EDF5 !important; }
[data-theme="dark"] .stApp h2 { color: #A8C4E8 !important; }
[data-theme="dark"] .stApp h3 { color: #8AADD4 !important; }


/* ── Mobile (≤768px): layout compacto ── */
@media (max-width: 768px) {

    /* Cards da home: 2 por linha no mobile — seletor correto */
    div.home-grid > div[data-testid="stHorizontalBlock"] {
        flex-wrap: wrap !important;
        display: flex !important;
        gap: 6px !important;
    }
    div.home-grid > div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        width: calc(50% - 3px) !important;
        min-width: calc(50% - 3px) !important;
        max-width: calc(50% - 3px) !important;
        flex: 0 0 calc(50% - 3px) !important;
        box-sizing: border-box !important;
    }

    /* Filtros: esconder barra de labels (captions) no mobile */
    .filter-header-bar { display: none !important; }

    /* Colunas de filtros: empilhar verticalmente no mobile */
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="stSelectbox"]) {
        flex-wrap: wrap !important;
    }
    div[data-testid="stHorizontalBlock"]:has(div[data-testid="stSelectbox"])
        > div[data-testid="stVerticalBlock"] {
        width: calc(50% - 4px) !important;
        min-width: calc(50% - 4px) !important;
        flex: 0 0 calc(50% - 4px) !important;
    }

    /* Sidebar: fundo sólido no mobile — cores explícitas */
    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] > div,
    section[data-testid="stSidebar"] > div:first-child {
        background-color: #FFFFFF !important;
        background: #FFFFFF !important;
    }
    /* Dark mode sidebar mobile */
    @media (prefers-color-scheme: dark) {
        section[data-testid="stSidebar"],
        section[data-testid="stSidebar"] > div,
        section[data-testid="stSidebar"] > div:first-child {
            background-color: #1A1D24 !important;
            background: #1A1D24 !important;
        }
    }

    /* Gráficos: garantir que caibam na tela */
    div[data-testid="stPlotlyChart"] {
        overflow-x: auto !important;
    }
}

/* ── Expander de filtros discreto ── */
div[data-testid="stExpander"] details {
    border: 1px solid rgba(128,128,128,0.15) !important;
    border-radius: 8px !important;
    background: transparent !important;
    margin-bottom: 8px !important;
}
div[data-testid="stExpander"] details summary {
    font-size: 0.78rem !important;
    font-weight: 600 !important;
    color: #8A96A8 !important;
    padding: 6px 12px !important;
}
div[data-testid="stExpander"] details summary:hover {
    color: #4A7BC8 !important;
}


/* ── Dark mode via prefers-color-scheme (mais confiável no Streamlit) ── */
@media (prefers-color-scheme: dark) {
    .stApp h1, .stApp h2, .stApp h3 { color: #7EB3F7 !important; }
    [data-testid="stMetricValue"]    { color: #E8EDF5 !important; }
    [data-testid="stMetricLabel"] p  { color: #8A96A8 !important; }
    [data-testid="stSidebar"] .stRadio label[aria-checked="true"] p {
        color: #7EB3F7 !important;
    }
    /* Títulos dos módulos (h2 inline) */
    .stMarkdown h2 { color: #7EB3F7 !important; }
    /* Texto padrão */
    .stMarkdown p  { color: #C4CDD9 !important; }
    /* Cards */
    [data-testid="stMetric"] { background: var(--secondary-background-color) !important; }
    /* Breadcrumb */
    [data-testid="stSidebar"] { background: var(--secondary-background-color) !important; }
}

</style>
""", unsafe_allow_html=True)

# ====================== CONFIGURAÇÕES GITHUB ======================
GITHUB_REPO = "fabiosilvavendas-byte/CRM_Medtextil2.0"
GITHUB_FOLDER = "dados"  # ⭐ PASTA ONDE ESTÃO AS PLANILHAS
GITHUB_TOKEN = None  # Opcional: adicione token se repositório for privado

@st.cache_data(ttl=3600)
def listar_planilhas_github():
    """Lista todos os arquivos Excel da pasta 'dados' no repositório GitHub"""
    try:
        if GITHUB_TOKEN:
            g = Github(GITHUB_TOKEN, timeout=15)
        else:
            g = Github(timeout=15)
        
        repo = g.get_repo(GITHUB_REPO)
        # ⭐ BUSCAR NA PASTA 'dados'
        contents = repo.get_contents(GITHUB_FOLDER)
        
        planilhas = {
            'vendas': None,
            'inadimplencia': None,
            'vendas_produto': None,
            'produtos_agrupados': None,
            'pedidos_pendentes': None,
            'tabela_ne': None,
            'todas': []
        }
        
        for content in contents:
            if content.name.endswith(('.xlsx', '.xls')):
                info = {
                    'nome': content.name,
                    'url': content.download_url,
                    'path': content.path
                }
                planilhas['todas'].append(info)
                
                # Identificar planilha de vendas
                if 'CONSULTA_VENDEDORES' in content.name.upper():
                    planilhas['vendas'] = info
                
                # Identificar planilha de inadimplência
                if 'LANCAMENTO A RECEBER' in content.name.upper() or 'LANCAMENTO_A_RECEBER' in content.name.upper():
                    planilhas['inadimplencia'] = info
                
                # Identificar planilha de vendas por produto
                if 'VENDAS POR PRODUTO' in content.name.upper() and 'GERAL' in content.name.upper():
                    planilhas['vendas_produto'] = info
                
                # Identificar planilha de produtos agrupados
                if 'PRODUTOS_AGRUPADOS_COMPLETOS_CONCILIADOS' in content.name.upper():
                    planilhas['produtos_agrupados'] = info
                
                # Identificar planilha de pedidos pendentes
                if 'PEDIDOSPENDENTES' in content.name.upper().replace(' ', '').replace('_', ''):
                    planilhas['pedidos_pendentes'] = info

                # Identificar tabela NE
                if 'TABELA_NE' in content.name.upper().replace(' ', '_'):
                    planilhas['tabela_ne'] = info
        
        if not planilhas['todas']:
            st.warning(f"⚠️ Nenhuma planilha Excel encontrada na pasta '{GITHUB_FOLDER}'")
        
        return planilhas
    except Exception as e:
        st.error(f"❌ Erro ao conectar ao GitHub: {str(e)}")
        st.info(f"💡 Verificando: {GITHUB_REPO}/{GITHUB_FOLDER}")
        return {'vendas': None, 'inadimplencia': None, 'vendas_produto': None, 'produtos_agrupados': None, 'pedidos_pendentes': None, 'todas': []}

@st.cache_data(ttl=3600)
def carregar_planilha_github(url):
    """Carrega planilha diretamente do GitHub"""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        df = pd.read_excel(io.BytesIO(response.content))
        return df
    except requests.exceptions.Timeout:
        st.error("⏱️ Timeout ao carregar planilha. Tente novamente.")
        return None
    except requests.exceptions.RequestException as e:
        st.error(f"❌ Erro ao carregar planilha: {str(e)}")
        return None
    except Exception as e:
        st.error(f"❌ Erro ao processar planilha: {str(e)}")
        return None

# ====================== AUTENTICAÇÃO ======================
def check_password():
    """Sistema de autenticação com níveis de acesso"""
    
    # 🔐 CONFIGURAÇÃO DE USUÁRIOS E SENHAS
    USUARIOS = {
        "admin123": {
            "tipo": "administrador",
            "nome": "Administrador",
            "modulos": ["Dashboard", "Positivação", "Inadimplência", "Clientes sem Compra", "Histórico", "Preço Médio", "Pedidos Pendentes", "Rankings", "Performance de Vendedores", "Consulta Clientes"]
        },
        "colaborador123": {  # ⬅️ MUDE ESTA SENHA
            "tipo": "colaborador",
            "nome": "Colaborador",
            "modulos": ["Inadimplência", "Histórico", "Pedidos Pendentes", "Consulta Clientes"]
        }
    }
    
    def password_entered():
        senha = st.session_state.get("password_input", "")
        
        if senha in USUARIOS:
            st.session_state["password_correct"] = True
            st.session_state["usuario"] = USUARIOS[senha]
            st.session_state["senha_usada"] = senha
        else:
            st.session_state["password_correct"] = False
            st.session_state["show_error"] = True
            if "usuario" in st.session_state:
                del st.session_state["usuario"]

    # Tela de login estilizada
    def render_login(show_error=False):
        st.markdown("""
        <div style="max-width:420px;margin:60px auto 0 auto;background:#FFFFFF;border-radius:16px;
                    padding:40px 36px;box-shadow:0 8px 32px rgba(31,71,136,0.12);border-top:4px solid #1F4788;">
            <div style="text-align:center;margin-bottom:10px;">
                <img src="https://i.imgur.com/gt3rgyL.png" height="52"
                     style="border-radius:8px;" onerror="this.style.display='none'"/>
            </div>
            <div style="text-align:center;font-size:1.4rem;font-weight:700;color:#4A7BC8;margin-bottom:4px;">
                Medtextil BI
            </div>
            <div style="text-align:center;font-size:0.85rem;color:#6C757D;margin-bottom:8px;">
                Dashboard Comercial 2.0
            </div>
        </div>
        """, unsafe_allow_html=True)
        col_l, col_c, col_r = st.columns([1, 2, 1])
        with col_c:
            st.markdown("<br>", unsafe_allow_html=True)
            st.text_input("🔑 Senha de acesso", type="password", key="password_input", placeholder="Digite sua senha...")
            if st.button("Entrar →", use_container_width=True, type="primary"):
                password_entered()
            if show_error:
                st.error("Senha incorreta. Verifique e tente novamente.")

    if "password_correct" not in st.session_state:
        render_login(show_error=False)
        return False
    elif not st.session_state["password_correct"]:
        render_login(show_error=st.session_state.get("show_error", False))
        return False
    else:
        return True

# ====================== PROCESSAMENTO DE DADOS ======================
def calcular_prazo_historico(data_emissao, data_vencimento_str):
    """
    Calcula o prazo histórico CUMULATIVO em dias.
    NÃO conta o dia da emissão (começa a contar do dia seguinte).
    
    Exemplo:
    - Data Emissão: 18/12/2025 (não conta este dia)
    - Vencimentos: "15/01/2026; 22/01/2026; 29/01/2026; 05/02/2026"
    - Cálculo:
      * De 19/12 até 15/01 = 28 dias
      * De 19/12 até 22/01 = 35 dias
      * De 19/12 até 29/01 = 42 dias
      * De 19/12 até 05/02 = 49 dias
    - Resultado: "28/35/42/49"
    """
    try:
        if pd.isna(data_vencimento_str) or data_vencimento_str == '':
            return ''
        
        if pd.isna(data_emissao):
            return ''
        
        # Converter string para garantir que é texto e limpar
        data_vencimento_str = str(data_vencimento_str).strip()
        
        # Separar as datas por ponto e vírgula
        datas_vencimento = data_vencimento_str.split(';')
        
        prazos = []
        for data_venc_str in datas_vencimento:
            data_venc_str = data_venc_str.strip()
            if not data_venc_str:
                continue
            
            try:
                # Tentar converter para datetime com dayfirst=True (padrão brasileiro)
                data_venc = pd.to_datetime(data_venc_str, errors='coerce', dayfirst=True)
                
                # Normalizar para meia-noite (remove qualquer componente de hora)
                if pd.notna(data_venc):
                    data_venc = data_venc.normalize() if hasattr(data_venc, 'normalize') else data_venc.replace(hour=0, minute=0, second=0, microsecond=0)
                
                # Validar se a data é válida e razoável
                if pd.notna(data_venc):
                    # Verificar se a data está dentro de um intervalo razoável (ano entre 2020 e 2030)
                    if 2020 <= data_venc.year <= 2030:
                        # Calcular dias a partir do DIA SEGUINTE à emissão
                        # O Python já calcula correto: não conta o dia da emissão
                        diferenca = (data_venc - data_emissao).days
                        
                        # Validar prazo razoável (entre 1 e 365 dias)
                        if 1 <= diferenca <= 365:
                            prazos.append(str(diferenca))
            except:
                # Ignorar datas inválidas silenciosamente
                continue
        
        # Retornar prazos separados por "/"
        if prazos:
            return '/'.join(prazos)
        else:
            return ''
    except:
        return ''

def gerar_pdf_pedido(dados_cliente, dados_pedido, itens_pedido, observacao=''):
    """Gera PDF do pedido no formato da Medtextil"""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=15*mm, bottomMargin=15*mm)
    
    elements = []
    styles = getSampleStyleSheet()
    
    # Estilo customizado
    style_title = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        textColor=colors.HexColor('#1f4788'),
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    style_normal = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=9,
        fontName='Helvetica'
    )
    
    style_small = ParagraphStyle(
        'CustomSmall',
        parent=styles['Normal'],
        fontSize=8,
        fontName='Helvetica'
    )
    
    # CABEÇALHO - SOLUÇÃO COM CONTROLE TOTAL DE TAMANHO
    logo_adicionado = False
    try:
        # Buscar logo do GitHub
        logo_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{GITHUB_FOLDER}/logo.png"
        response = requests.get(logo_url)
        if response.status_code == 200:
            from PIL import Image as PILImage
            
            # Carregar imagem com PIL para ter controle total
            logo_buffer = io.BytesIO(response.content)
            pil_img = PILImage.open(logo_buffer)
            
            # Obter dimensões originais
            largura_original, altura_original = pil_img.size
            proporcao = largura_original / altura_original
            
            # DIMENSÕES FIXAS DO LOGO NO PDF
            altura_desejada_mm = 15  # Altura em milímetros
            largura_desejada_mm = altura_desejada_mm * proporcao
            
            # Criar Image do ReportLab com dimensões exatas
            logo_buffer.seek(0)  # Voltar ao início do buffer
            logo_img = Image(logo_buffer, width=largura_desejada_mm*mm, height=altura_desejada_mm*mm)
            
            # Texto ao lado - com width definido para não vazar
            texto_empresa = Paragraph(
                "<b>MEDTEXTIL PRODUTOS TEXTIL HOSPITALARES</b><br/>"
                "<font size=8>CNPJ: 40.357.820/0001-50  Inscrição Estadual: 16.390.286-0</font>",
                style_small
            )
            
            # Calcular colWidths baseado no logo real
            espaco_logo = largura_desejada_mm + 5  # Logo + margem de 5mm
            espaco_texto = 190 - espaco_logo  # Resto da página
            
            # Criar tabela com dimensões calculadas
            cabecalho_data = [[logo_img, texto_empresa]]
            cabecalho_table = Table(cabecalho_data, colWidths=[espaco_logo*mm, espaco_texto*mm])
            cabecalho_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (0, 0), 'TOP'),
                ('VALIGN', (1, 0), (1, 0), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            elements.append(cabecalho_table)
            elements.append(Spacer(1, 5*mm))
            logo_adicionado = True
            
    except Exception as e:
        # Se PIL não disponível ou erro, fallback simples
        try:
            logo_buffer = io.BytesIO(response.content)
            # Fallback: tamanho fixo conservador
            logo_img = Image(logo_buffer, width=30*mm, height=15*mm)
            
            texto_empresa = Paragraph(
                "<b>MEDTEXTIL PRODUTOS TEXTIL HOSPITALARES</b><br/>"
                "<font size=8>CNPJ: 40.357.820/0001-50  Inscrição Estadual: 16.390.286-0</font>",
                style_small
            )
            
            cabecalho_data = [[logo_img, texto_empresa]]
            cabecalho_table = Table(cabecalho_data, colWidths=[35*mm, 155*mm])
            cabecalho_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ]))
            elements.append(cabecalho_table)
            elements.append(Spacer(1, 5*mm))
            logo_adicionado = True
        except:
            pass
    
    if not logo_adicionado:
        # Fallback final para texto
        elements.append(Paragraph("<b>MEDTEXTIL PRODUTOS TEXTIL HOSPITALARES</b><br/>"
                                 "CNPJ: 40.357.820/0001-50 | Inscrição Estadual: 16.390.286-0", style_title))
        elements.append(Spacer(1, 5*mm))
    
    # REPRESENTANTE (tabela simples como no modelo)
    data_repr = [
        ['Representante', dados_cliente.get('representante', '')],
        ['CNPJ', dados_cliente.get('cnpj', '')]
    ]
    table_repr = Table(data_repr, colWidths=[40*mm, 150*mm])
    table_repr.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
    ]))
    elements.append(table_repr)
    elements.append(Spacer(1, 2*mm))
    
    # INFORMAÇÕES DO CLIENTE (layout simplificado)
    data_cliente = [
        [Paragraph("<b>Informações sobre o Cliente</b>", style_normal), '', '', ''],
        ['Cliente', dados_cliente.get('razao_social', ''), '', Paragraph("<b>Nome</b><br/>Fantasia:<br/>Inscr.<br/>Estadual:", style_small)],
        ['CNPJ:', dados_cliente.get('cnpj', ''), '', dados_cliente.get('nome_fantasia', '')],
        ['Telefone:', dados_cliente.get('telefone', ''), 'Email NF-e:', dados_cliente.get('email', '')],
        ['Endereço:', dados_cliente.get('endereco', ''), '', dados_cliente.get('ie', '')],
        ['Observação:', dados_cliente.get('obs_cliente', ''), '', '']
    ]
    
    table_cliente = Table(data_cliente, colWidths=[25*mm, 85*mm, 25*mm, 55*mm])
    table_cliente.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 1), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('SPAN', (0, 0), (-1, 0)),
        ('SPAN', (1, 1), (2, 1)),
        ('SPAN', (1, 5), (-1, 5)),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    elements.append(table_cliente)
    elements.append(Spacer(1, 2*mm))
    
    # INFORMAÇÕES DO PEDIDO (layout compacto)
    data_pedido_info = [
        ['Informações sobre Pedido Nº', dados_pedido.get('numero', ''), '', Paragraph("<b>Tabela de<br/>Preço:</b>", style_small)],
        ['Condições de Pagto:', dados_pedido.get('condicoes_pagto', ''), '', dados_pedido.get('tabela_preco', '')],
        ['Data da Venda:', dados_pedido.get('data_venda', ''), 'Tipo de<br/>Frete:', dados_pedido.get('tipo_frete', 'CIF')]
    ]
    
    table_pedido = Table(data_pedido_info, colWidths=[45*mm, 90*mm, 20*mm, 35*mm])
    table_pedido.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
    ]))
    elements.append(table_pedido)
    elements.append(Spacer(1, 3*mm))
    
    # DETALHE DO PEDIDO
    elements.append(Paragraph("<b>Detalhe do Pedido</b>", style_normal))
    elements.append(Spacer(1, 1*mm))
    
    # Cabeçalho da tabela de itens (cor azul igual ao modelo)
    header_itens = ['COD.', 'PRODUTO', 'PESO', 'CAIXA DE\nEMBARQUE', 'QTDE', 'VALOR', 'TOTAL']
    data_itens = [header_itens]
    
    # Adicionar itens
    for item in itens_pedido:
        descricao = str(item.get('descricao', ''))
        # Quebrar descrição se muito longa
        if len(descricao) > 45:
            descricao = descricao[:45] + '...'
        
        data_itens.append([
            str(item.get('codigo', '')),
            descricao,
            str(item.get('peso', '')),
            str(item.get('cx_embarque', '')),
            f"{item.get('quantidade', 0):.0f}",
            f"R$ {item.get('valor_unit', 0):.2f}",
            f"R$ {item.get('total', 0):.2f}"
        ])
    
    # Calcular totais
    total_qtde = sum([item.get('quantidade', 0) for item in itens_pedido])
    total_valor = sum([item.get('total', 0) for item in itens_pedido])
    
    # Linha de total (sem bordas superiores, fundo cinza)
    data_itens.append(['', '', '', '', f"{total_qtde:.0f}", '', f"R$ {total_valor:,.2f}"])
    
    col_widths = [12*mm, 76*mm, 15*mm, 18*mm, 12*mm, 22*mm, 25*mm]
    table_itens = Table(data_itens, colWidths=col_widths)
    
    # Estilo da tabela de itens (azul igual ao modelo)
    num_rows = len(data_itens)
    style_itens = [
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -2), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#003366')),  # Azul escuro
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (4, 1), (4, -1), 'CENTER'),  # Centralizar QTDE
        ('ALIGN', (5, 1), (7, -1), 'RIGHT'),   # Alinhar valores à direita
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, -1), (-1, -1), 9),
    ]
    
    table_itens.setStyle(TableStyle(style_itens))
    elements.append(table_itens)
    elements.append(Spacer(1, 3*mm))
    
    # RESUMO FINAL
    data_resumo = [
        ['Qtde Itens', 'Frete', 'Total Final'],
        [f"{total_qtde:.0f}", dados_pedido.get('tipo_frete', 'CIF'), f"R$ {total_valor:,.2f}"]
    ]
    
    table_resumo = Table(data_resumo, colWidths=[60*mm, 60*mm, 70*mm])
    table_resumo.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(table_resumo)
    elements.append(Spacer(1, 3*mm))
    
    # OBSERVAÇÃO
    if observacao:
        elements.append(Paragraph("<b>Observação</b>", style_normal))
        elements.append(Paragraph(observacao, style_small))
    
    # Construir PDF
    doc.build(elements)
    
    buffer.seek(0)
    return buffer.getvalue()

def calcular_comissao(preco_unit, preco_ref):
    """
    Calcula o percentual de comissão com base no desvio do PrecoUnit em relação ao preço de referência.
    
    Regras:
    - PrecoUnit >= 6% ACIMA  do preço ref → 4%
    - PrecoUnit igual ao preço ref (0% desconto) → 3%
    - PrecoUnit até 3% ABAIXO do preço ref → 2,5%
    - PrecoUnit 3% ou mais ABAIXO do preço ref → 2%
    """
    try:
        if pd.isna(preco_unit) or pd.isna(preco_ref) or preco_ref == 0:
            return ''
        
        preco_unit = float(preco_unit)
        preco_ref  = float(preco_ref)
        
        # Arredondar para 2 casas decimais (centavos) antes de comparar
        # Elimina ruído de ponto flutuante sem distorcer valores reais
        preco_unit = round(preco_unit, 2)
        preco_ref  = round(preco_ref,  2)
        
        variacao = round(((preco_unit - preco_ref) / preco_ref) * 100, 4)
        
        if variacao >= 6:     # >= 6% acima → 4%
            return '4%'
        elif variacao >= 0:   # >= 0% (igual ou acima) → 3%
            return '3%'
        elif variacao > -3:   # Entre 0% e -3% → 2,5%
            return '2,5%'
        else:                 # -3% ou mais abaixo → 2%
            return '2%'
    except:
        return ''


# ── Paleta institucional e helper de layout de gráficos ──────────────────
CORES_INST = ['#1F4788', '#2E86AB', '#28A745', '#F4A261', '#6C757D',
              '#163561', '#1B5E8A', '#1E7B34', '#C97A3A', '#495057']

def aplicar_layout_grafico(fig, height=None):
    """Aplica o estilo institucional Medtextil a qualquer figura Plotly."""
    layout_kwargs = dict(
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family='Inter, Segoe UI, Roboto, sans-serif', color='#495057', size=12),
        margin=dict(l=10, r=10, t=36, b=10),
        xaxis=dict(showgrid=False, showline=True, linecolor='#E9ECEF', tickfont=dict(size=11)),
        yaxis=dict(showgrid=True, gridcolor='#F0F0F0', showline=False, tickfont=dict(size=11)),
        coloraxis_showscale=False,
        hoverlabel=dict(bgcolor='#1F4788', font_color='white', font_size=12,
                        bordercolor='#1F4788'),
    )
    if height:
        layout_kwargs['height'] = height
    fig.update_layout(**layout_kwargs)
    return fig

@st.cache_data
def processar_dados(df):
    """Aplica as regras de negócio nos dados"""
    df['Valor_Real'] = df.apply(
        lambda row: row['TotalProduto'] if row['TipoMov'] == 'NF Venda' else -row['TotalProduto'],
        axis=1
    )
    # Converter DataEmissao com formato brasileiro e normalizar para meia-noite
    df['DataEmissao'] = pd.to_datetime(df['DataEmissao'], errors='coerce', dayfirst=True)
    # Normalizar para meia-noite (remove hora) para cálculos corretos de dias
    df['DataEmissao'] = df['DataEmissao'].dt.normalize()
    
    df['Mes'] = df['DataEmissao'].dt.month
    df['Ano'] = df['DataEmissao'].dt.year
    df['MesAno'] = df['DataEmissao'].dt.to_period('M').astype(str)
    
    # Calcular prazo histórico se a coluna DataVencimento existir
    if 'DataVencimento' in df.columns:
        df['PrazoHistorico'] = df.apply(
            lambda row: calcular_prazo_historico(row['DataEmissao'], row['DataVencimento']),
            axis=1
        )
    else:
        df['PrazoHistorico'] = ''
    
    return df

def obter_notas_unicas(df):
    """Remove duplicatas de Numero_NF mantendo apenas primeira ocorrência"""
    return df.drop_duplicates(subset=['Numero_NF'], keep='first')

def to_excel(df):
    """Converte DataFrame para Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

def to_excel_pedidos_pendentes(df):
    """Converte DataFrame de pedidos pendentes para Excel com abas por tipo de produto"""
    output = io.BytesIO()
    
    # Função para identificar tipo de produto pela descrição
    def identificar_tipo(descricao):
        if pd.isna(descricao):
            return 'OUTROS'
        
        descricao_upper = str(descricao).upper()
        
        if 'ATADURA' in descricao_upper:
            return 'ATADURAS'
        elif 'CAMPO' in descricao_upper:
            return 'CAMPO'
        elif ('GAZE' in descricao_upper and 'ROLO' in descricao_upper) or ('GAZE' in descricao_upper and 'CIRCULAR' in descricao_upper):
            return 'GAZE EM ROLO'
        elif 'NAO ESTERIL' in descricao_upper or 'NÃO ESTERIL' in descricao_upper or 'NÃO ESTÉRIL' in descricao_upper or 'NAO ESTÉRIL' in descricao_upper:
            return 'NÃO ESTERIL'
        elif 'ESTERIL' in descricao_upper or 'ESTÉRIL' in descricao_upper:
            return 'ESTERIL'
        else:
            return 'OUTROS'
    
    # Função para identificar se é HOSPITALAR ou FARMA (apenas para ATADURAS)
    def identificar_categoria(descricao, tipo):
        if tipo == 'ATADURAS':
            if pd.notna(descricao) and 'HOSP' in str(descricao).upper():
                return 'HOSPITALAR'
            else:
                return 'FARMA'
        return ''
    
    # Adicionar coluna de tipo
    df_export = df.copy()
    df_export['TipoProduto'] = df_export['Descricao'].apply(identificar_tipo)
    df_export['Categoria'] = df_export.apply(lambda row: identificar_categoria(row['Descricao'], row['TipoProduto']), axis=1)
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Definir ordem das abas
        tipos_ordem = ['ATADURAS', 'CAMPO', 'ESTERIL', 'NÃO ESTERIL', 'GAZE EM ROLO', 'OUTROS']
        
        for tipo in tipos_ordem:
            df_tipo = df_export[df_export['TipoProduto'] == tipo].copy()
            
            if len(df_tipo) > 0:
                # Para ATADURAS, ordenar por Categoria (HOSPITALAR primeiro)
                if tipo == 'ATADURAS':
                    df_tipo = df_tipo.sort_values('Categoria')
                
                # Remover colunas auxiliares antes de exportar
                colunas_para_remover = ['TipoProduto']
                # Manter coluna Categoria apenas para ATADURAS
                if tipo != 'ATADURAS':
                    colunas_para_remover.append('Categoria')
                
                df_tipo = df_tipo.drop(columns=[col for col in colunas_para_remover if col in df_tipo.columns])
                
                df_tipo.to_excel(writer, index=False, sheet_name=tipo)
    
    return output.getvalue()

def formatar_moeda(valor):
    """Formata valor para moeda brasileira"""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_dataframe_moeda(df, colunas_moeda):
    """Formata colunas de moeda em um dataframe para exibição"""
    df_formatado = df.copy()
    for col in colunas_moeda:
        if col in df_formatado.columns:
            df_formatado[col] = df_formatado[col].apply(lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00")
    return df_formatado

@st.cache_data
def processar_inadimplencia(df):
    """Processa dados de inadimplência"""
    # Padronizar nomes das colunas
    # Tentar várias variações de nomes de colunas
    rename_map = {
        'Funcionário': 'Vendedor',
        'Razão Social': 'Cliente',
        'N_Doc': 'NumeroDoc',
        'N Doc': 'NumeroDoc',
        'NDoc': 'NumeroDoc',
        'Numero Doc': 'NumeroDoc',
        'Dt.Vencimento': 'DataVencimento',
        'Dt Vencimento': 'DataVencimento',
        'Data Vencimento': 'DataVencimento',
        'Vr.Líquido': 'ValorLiquido',
        'Vr Líquido': 'ValorLiquido',
        'Valor Líquido': 'ValorLiquido',
        'Valor Liquido': 'ValorLiquido',
        'Conta/Caixa': 'Banco',
        'Conta Caixa': 'Banco',
        'UF': 'Estado'
    }
    
    df = df.rename(columns=rename_map)
    
    # Converter data de vencimento
    df['DataVencimento'] = pd.to_datetime(df['DataVencimento'], errors='coerce')
    
    # Calcular dias de atraso
    hoje = pd.Timestamp.now()
    df['DiasAtraso'] = (hoje - df['DataVencimento']).dt.days
    df['DiasAtraso'] = df['DiasAtraso'].apply(lambda x: max(0, x))  # Não mostrar valores negativos
    
    # Classificar inadimplência
    def classificar_inadimplencia(dias):
        if dias == 0:
            return 'A Vencer'
        elif dias <= 30:
            return '1-30 dias'
        elif dias <= 60:
            return '31-60 dias'
        elif dias <= 90:
            return '61-90 dias'
        else:
            return 'Acima de 90 dias'
    
    df['FaixaAtraso'] = df['DiasAtraso'].apply(classificar_inadimplencia)
    
    return df


# ====================== PROPOSTA PDF (HISTÓRICO DE CLIENTE) ======================
def gerar_proposta_pdf_historico(cliente_info_dict, historico_df, vendas_resumo):
    """
    Gera PDF de Proposta Comercial baseada no histórico de compras do cliente.
    Usa apenas a biblioteca fpdf2 (pip install fpdf2).
    Fallback: se fpdf2 não estiver disponível, usa ReportLab.
    """
    import io, requests
    from datetime import date

    razao    = str(cliente_info_dict.get('RazaoSocial', ''))
    cpf_cnpj = str(cliente_info_dict.get('CPF_CNPJ', ''))
    cidade   = str(cliente_info_dict.get('Cidade', ''))
    estado   = str(cliente_info_dict.get('Estado', ''))
    vendedor = str(cliente_info_dict.get('Vendedor', ''))
    hoje     = date.today().strftime('%d/%m/%Y')

    # ── Tentar fpdf2 primeiro ─────────────────────────────────────────────
    try:
        from fpdf import FPDF

        class PropostaPDF(FPDF):
            def header(self):
                # Logo
                try:
                    resp = requests.get("https://i.imgur.com/gt3rgyL.png", timeout=8)
                    if resp.status_code == 200:
                        tmp = io.BytesIO(resp.content)
                        tmp.seek(0)
                        import tempfile, os
                        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as f_tmp:
                            f_tmp.write(resp.content)
                            tmp_path = f_tmp.name
                        self.image(tmp_path, x=10, y=8, w=32)
                        os.unlink(tmp_path)
                except Exception:
                    pass
                self.set_xy(46, 10)
                self.set_font('Helvetica', 'B', 13)
                self.set_text_color(31, 71, 136)
                self.cell(0, 6, 'MEDTEXTIL PRODUTOS TEXTIL HOSPITALARES', ln=True)
                self.set_xy(46, 16)
                self.set_font('Helvetica', '', 8)
                self.set_text_color(100, 100, 100)
                self.cell(0, 5, 'CNPJ: 40.357.820/0001-50  |  IE: 16.390.286-0', ln=True)
                self.set_draw_color(31, 71, 136)
                self.set_line_width(0.6)
                self.line(10, 26, 200, 26)
                self.ln(4)

            def footer(self):
                self.set_y(-14)
                self.set_font('Helvetica', 'I', 7)
                self.set_text_color(160, 160, 160)
                self.cell(0, 6,
                    f'Medtextil — Proposta gerada em {hoje}  |  Pág. {self.page_no()}',
                    align='C')

        pdf = PropostaPDF()
        pdf.set_auto_page_break(auto=True, margin=18)
        pdf.add_page()
        pdf.set_margins(10, 10, 10)

        # ── Título ────────────────────────────────────────────────────────
        pdf.set_font('Helvetica', 'B', 14)
        pdf.set_text_color(31, 71, 136)
        pdf.cell(0, 8, 'PROPOSTA COMERCIAL', align='C', ln=True)
        pdf.set_font('Helvetica', '', 8)
        pdf.set_text_color(130, 130, 130)
        pdf.cell(0, 5, f'Emitida em {hoje}', align='C', ln=True)
        pdf.ln(4)

        # ── Dados do Cliente ──────────────────────────────────────────────
        pdf.set_fill_color(240, 244, 255)
        pdf.set_draw_color(200, 210, 230)
        pdf.set_font('Helvetica', 'B', 9)
        pdf.set_text_color(31, 71, 136)
        pdf.cell(0, 7, ' DADOS DO CLIENTE', fill=True, border=1, ln=True)

        pdf.set_font('Helvetica', '', 9)
        pdf.set_text_color(50, 50, 50)
        w1, w2 = 38, 152
        for label, valor in [
            ('Razão Social', razao),
            ('CPF / CNPJ',   cpf_cnpj),
            ('Cidade / UF',  f'{cidade} / {estado}'),
            ('Vendedor',     vendedor),
        ]:
            pdf.set_font('Helvetica', 'B', 8)
            pdf.cell(w1, 6, f'  {label}:', border='LB', fill=False)
            pdf.set_font('Helvetica', '', 8)
            pdf.cell(w2, 6, f'  {valor}', border='RB', ln=True)
        pdf.ln(5)

        # ── Resumo Financeiro ─────────────────────────────────────────────
        pdf.set_fill_color(240, 244, 255)
        pdf.set_font('Helvetica', 'B', 9)
        pdf.set_text_color(31, 71, 136)
        pdf.cell(0, 7, ' RESUMO FINANCEIRO (PERÍODO SELECIONADO)', fill=True, border=1, ln=True)

        pdf.set_font('Helvetica', '', 9)
        pdf.set_text_color(50, 50, 50)
        w3 = 95
        itens_resumo = list(vendas_resumo.items())
        for i in range(0, len(itens_resumo), 2):
            k1, v1 = itens_resumo[i]
            pdf.set_font('Helvetica', 'B', 8)
            pdf.cell(w3, 6, f'  {k1}:', border='LB')
            pdf.set_font('Helvetica', '', 8)
            if i + 1 < len(itens_resumo):
                k2, v2 = itens_resumo[i+1]
                pdf.cell(w3, 6, f'  {v1}', border='B')
                pdf.set_font('Helvetica', 'B', 8)
                pdf.cell(0, 6, f'  {k2}:', border='B')
            else:
                pdf.cell(0, 6, f'  {v1}', border='RB')
            pdf.ln()
        pdf.ln(5)

        # ── Tabela de Produtos ────────────────────────────────────────────
        pdf.set_fill_color(31, 71, 136)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font('Helvetica', 'B', 8)
        cols_w = [18, 72, 18, 22, 26, 26]
        cols_h = ['Código', 'Produto', 'Qtd', 'Prazo', 'Preço Unit.', 'Total']
        for cw, ch in zip(cols_w, cols_h):
            pdf.cell(cw, 7, ch, border=1, fill=True, align='C')
        pdf.ln()

        pdf.set_text_color(50, 50, 50)
        fill_row = False
        vendas_only = historico_df[historico_df['TipoMov'] == 'NF Venda'].copy()
        # Agrupar por produto para resumo
        grp_cols = ['CodigoProduto', 'NomeProduto']
        if 'PrazoHistorico' in vendas_only.columns:
            grp_cols_agg = {
                'Quantidade': 'sum',
                'PrecoUnit': 'mean',
                'TotalProduto': 'sum',
                'PrazoHistorico': 'first'
            }
        else:
            grp_cols_agg = {
                'Quantidade': 'sum',
                'PrecoUnit': 'mean',
                'TotalProduto': 'sum'
            }
        try:
            resumo_prod = vendas_only.groupby(
                ['CodigoProduto', 'NomeProduto'], as_index=False
            ).agg({k: v for k, v in grp_cols_agg.items() if k in vendas_only.columns})
            resumo_prod = resumo_prod.sort_values('TotalProduto', ascending=False)
        except Exception:
            resumo_prod = vendas_only[['CodigoProduto','NomeProduto','Quantidade','PrecoUnit','TotalProduto']].head(30)

        for _, row in resumo_prod.iterrows():
            pdf.set_fill_color(247, 249, 255) if fill_row else pdf.set_fill_color(255, 255, 255)
            pdf.set_font('Helvetica', '', 7)
            cod  = str(row.get('CodigoProduto', ''))[:8]
            nome = str(row.get('NomeProduto', ''))[:38]
            qtd  = f"{row.get('Quantidade', 0):,.0f}"
            prazo = str(row.get('PrazoHistorico', '-'))[:10] if 'PrazoHistorico' in row.index else '-'
            preco = f"R$ {row.get('PrecoUnit', 0):,.2f}"
            total = f"R$ {row.get('TotalProduto', 0):,.2f}"
            row_vals = [cod, nome, qtd, prazo, preco, total]
            aligns   = ['C', 'L', 'C', 'C', 'R', 'R']
            for cw, rv, al in zip(cols_w, row_vals, aligns):
                pdf.cell(cw, 6, rv, border=1, fill=True, align=al)
            pdf.ln()
            fill_row = not fill_row

        # Total geral
        total_geral = vendas_only['TotalProduto'].sum() if len(vendas_only) > 0 else 0
        pdf.set_fill_color(31, 71, 136)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font('Helvetica', 'B', 9)
        pdf.cell(sum(cols_w[:5]), 7, 'TOTAL GERAL', border=1, fill=True, align='R')
        pdf.cell(cols_w[5], 7, f'R$ {total_geral:,.2f}', border=1, fill=True, align='R')
        pdf.ln(8)

        # ── Rodapé da proposta ────────────────────────────────────────────
        pdf.set_font('Helvetica', 'I', 8)
        pdf.set_text_color(130, 130, 130)
        pdf.multi_cell(0, 5,
            'Esta proposta é baseada no histórico de compras do cliente e não representa um pedido confirmado. '
            'Valores sujeitos a alteração. Validade: 15 dias.')

        return pdf.output()

    # ── Fallback: ReportLab ───────────────────────────────────────────────
    except ImportError:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib import colors
        from reportlab.lib.units import mm
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4,
                                rightMargin=15*mm, leftMargin=15*mm,
                                topMargin=15*mm, bottomMargin=15*mm)
        styles = getSampleStyleSheet()
        azul = colors.HexColor('#1F4788')
        elements = []

        # Título
        p_title = ParagraphStyle('T', parent=styles['Heading1'],
                                 fontSize=16, textColor=azul, alignment=1)
        elements.append(Paragraph('PROPOSTA COMERCIAL', p_title))
        elements.append(Paragraph(f'<font size=9 color="grey">Medtextil — {hoje}</font>', styles['Normal']))
        elements.append(Spacer(1, 6*mm))

        # Dados do cliente
        dados = [['Razão Social', razao], ['CPF/CNPJ', cpf_cnpj],
                 ['Cidade/UF', f'{cidade}/{estado}'], ['Vendedor', vendedor]]
        t_dados = Table(dados, colWidths=[40*mm, 150*mm])
        t_dados.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(0,-1), colors.HexColor('#F0F4FF')),
            ('TEXTCOLOR', (0,0),(0,-1), azul),
            ('FONTNAME', (0,0),(0,-1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0),(-1,-1), 8),
            ('BOX', (0,0),(-1,-1), 0.5, colors.grey),
            ('INNERGRID', (0,0),(-1,-1), 0.3, colors.lightgrey),
            ('LEFTPADDING', (0,0),(-1,-1), 5),
        ]))
        elements.append(t_dados)
        elements.append(Spacer(1, 6*mm))

        # Resumo
        for k, v in vendas_resumo.items():
            elements.append(Paragraph(f'<b>{k}:</b> {v}', styles['Normal']))
        elements.append(Spacer(1, 4*mm))

        # Tabela de produtos
        vendas_only = historico_df[historico_df['TipoMov'] == 'NF Venda']
        header = ['Código', 'Produto', 'Qtd', 'Preço Unit.', 'Total']
        rows = [header]
        try:
            grp = vendas_only.groupby(['CodigoProduto','NomeProduto'], as_index=False).agg(
                {'Quantidade':'sum','PrecoUnit':'mean','TotalProduto':'sum'})
            grp = grp.sort_values('TotalProduto', ascending=False)
            for _, r in grp.iterrows():
                rows.append([str(r['CodigoProduto'])[:8], str(r['NomeProduto'])[:40],
                             f"{r['Quantidade']:,.0f}", f"R$ {r['PrecoUnit']:,.2f}",
                             f"R$ {r['TotalProduto']:,.2f}"])
        except Exception:
            pass
        total_g = vendas_only['TotalProduto'].sum() if len(vendas_only) > 0 else 0
        rows.append(['','','','Total Geral', f'R$ {total_g:,.2f}'])

        t_prod = Table(rows, colWidths=[20*mm, 80*mm, 18*mm, 28*mm, 28*mm])
        t_prod.setStyle(TableStyle([
            ('BACKGROUND', (0,0),(-1,0), azul),
            ('TEXTCOLOR', (0,0),(-1,0), colors.white),
            ('FONTNAME', (0,0),(-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0),(-1,-1), 7),
            ('ROWBACKGROUNDS', (0,1),(-1,-2), [colors.white, colors.HexColor('#F7F9FF')]),
            ('BACKGROUND', (0,-1),(-1,-1), colors.HexColor('#F0F4FF')),
            ('FONTNAME', (0,-1),(-1,-1), 'Helvetica-Bold'),
            ('BOX', (0,0),(-1,-1), 0.5, colors.grey),
            ('INNERGRID', (0,0),(-1,-1), 0.3, colors.lightgrey),
            ('LEFTPADDING', (0,0),(-1,-1), 3),
            ('ALIGN', (2,0),(-1,-1), 'RIGHT'),
        ]))
        elements.append(t_prod)
        doc.build(elements)
        buffer.seek(0)
        return buffer.getvalue()


# ====================== FILTROS DE DATA LOCAIS ======================
def renderizar_filtros_locais(key_prefix, label="📅 Ajustar Período"):
    """
    Expander compacto com date_inputs lado a lado.
    Retorna (data_inicial, data_final) — None se não preenchido.
    Usa key_prefix para evitar conflito de keys entre módulos.
    """
    data_ini = None
    data_fim = None
    with st.expander(label, expanded=False):
        c1, c2 = st.columns(2)
        with c1:
            data_ini = st.date_input(
                "De", value=None,
                key=f"local_ini_{key_prefix}",
                format="DD/MM/YYYY",
                label_visibility="visible"
            )
        with c2:
            data_fim = st.date_input(
                "Até", value=None,
                key=f"local_fim_{key_prefix}",
                format="DD/MM/YYYY",
                label_visibility="visible"
            )
    return data_ini, data_fim

# ====================== INÍCIO DO APP ======================
if not check_password():
    st.stop()

# Obter informações do usuário logado
usuario = st.session_state.get("usuario", {})
tipo_usuario = usuario.get("tipo", "")
nome_usuario = usuario.get("nome", "Usuário")
modulos_permitidos = usuario.get("modulos", [])

# ── Sidebar: Logo + Usuário + Navegação (clean) ──────────────────────────
with st.sidebar:
    # ── Logo centralizado ──
    st.markdown("""
    <div style="text-align:center;padding:20px 0 14px 0;border-bottom:1px solid #E9ECEF;margin-bottom:14px;">
        <img src="https://i.imgur.com/gt3rgyL.png" height="56"
             style="border-radius:10px;box-shadow:0 2px 8px rgba(31,71,136,0.18);"
             onerror="this.style.display='none'"/>
        <div style="font-size:0.8rem;font-weight:800;color:#4A7BC8;letter-spacing:0.1em;
                    text-transform:uppercase;margin-top:10px;">Medtextil</div>
        <div style="font-size:0.65rem;color:#ADB5BD;letter-spacing:0.06em;margin-top:1px;">
            BI Dashboard 2.0
        </div>
    </div>
    """, unsafe_allow_html=True)
    # ── Badge do usuário ──
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#F0F4FF,#E8EFFD);border:1px solid #C5D5F0;
                border-radius:10px;padding:10px 14px;font-size:0.83rem;color:#4A7BC8;
                font-weight:600;text-align:center;margin-bottom:6px;">
        👤 &nbsp;{nome_usuario}
    </div>
    """, unsafe_allow_html=True)
    # ── Botão sair compacto ──
    if st.button("🚪 Sair", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    st.markdown("<div style='margin-bottom:10px;'></div>", unsafe_allow_html=True)

# ── Cabeçalho principal ───────────────────────────────────────────────────
col_titulo, col_actions = st.columns([4, 1])
with col_titulo:
    st.markdown("# 📊 Dashboard Comercial")
    st.markdown('<p class="page-subtitle">Medtextil Produtos Textil Hospitalares — Análise de Vendas & BI</p>',
                unsafe_allow_html=True)

# ── Carregamento silencioso + Status no sidebar expander ─────────────────
with st.sidebar:
    with st.expander("🛠️ Status das Planilhas", expanded=False):
        with st.spinner("Conectando ao GitHub..."):
            planilhas_disponiveis = listar_planilhas_github()

        if planilhas_disponiveis['vendas']:
            st.success(f"✅ Vendas: {planilhas_disponiveis['vendas']['nome']}")
            url_planilha_vendas = planilhas_disponiveis['vendas']['url']
        else:
            st.error("❌ Planilha de vendas não encontrada")
            st.info("Procurando por arquivo com 'CONSULTA_VENDEDORES' no nome")

        if planilhas_disponiveis['inadimplencia']:
            st.success(f"✅ Inadimplência: {planilhas_disponiveis['inadimplencia']['nome']}")
        else:
            st.warning("⚠️ Inadimplência não encontrada (módulo desabilitado)")

        if planilhas_disponiveis.get('produtos_agrupados'):
            st.success(f"✅ Produtos: {planilhas_disponiveis['produtos_agrupados']['nome']}")

        if st.button("🔄 Recarregar Dados", use_container_width=True, key="btn_reload"):
            st.cache_data.clear()
            st.rerun()

# Validação crítica fora do expander (sem mensagem visual)
if not planilhas_disponiveis.get('vendas'):
    st.error("❌ Planilha de vendas não encontrada no GitHub. Verifique o repositório.")
    st.stop()

with st.spinner(""):
    df = carregar_planilha_github(url_planilha_vendas)

if df is None:
    st.error("❌ Não foi possível carregar os dados de vendas.")
    st.stop()

df = processar_dados(df)

# Carregar planilha de produtos para cálculo de comissão
if planilhas_disponiveis.get('produtos_agrupados'):
    df_ref_preco = carregar_planilha_github(planilhas_disponiveis['produtos_agrupados']['url'])
    if df_ref_preco is not None:
        df_ref_preco.columns = df_ref_preco.columns.str.upper()
        if 'ID_COD' in df_ref_preco.columns and 'PRECO' in df_ref_preco.columns:
            df_ref_preco = df_ref_preco[['ID_COD', 'PRECO']].rename(
                columns={'ID_COD': 'CodigoProduto', 'PRECO': 'PrecoRef'}
            )
            def normalizar_codigo(val):
                try:
                    return str(int(float(str(val).strip())))
                except:
                    return str(val).strip()
            df['CodigoProduto'] = df['CodigoProduto'].apply(normalizar_codigo)
            df_ref_preco['CodigoProduto'] = df_ref_preco['CodigoProduto'].apply(normalizar_codigo)
            df_ref_preco = df_ref_preco.drop_duplicates(subset=['CodigoProduto'], keep='first')
            df = df.merge(df_ref_preco, on='CodigoProduto', how='left')
            df['Comissao'] = df.apply(
                lambda row: calcular_comissao(row['PrecoUnit'], row['PrecoRef']),
                axis=1
            )
        else:
            df['PrecoRef'] = None
            df['Comissao'] = ''
else:
    df['PrecoRef'] = None
    df['Comissao'] = ''
''

# ── Filtros Globais — dentro de expander único ───────────────────────────
with st.expander("⚙️ Filtros", expanded=False):
    # Linha 1: Datas lado a lado
    fc1, fc2 = st.columns(2)
    with fc1:
        data_inicial = st.date_input("📅 Data Inicial", value=None,
                                     key="data_ini", format="DD/MM/YYYY")
    with fc2:
        data_final = st.date_input("📅 Data Final", value=None,
                                   key="data_fim", format="DD/MM/YYYY")
    # Linha 2: Vendedor e Estado lado a lado
    fc3, fc4 = st.columns(2)
    with fc3:
        vendedores = ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist())
        vendedor_filtro = st.selectbox("👤 Vendedor", vendedores, key="vend_global")
    with fc4:
        estados = ['Todos'] + sorted(df['Estado'].dropna().unique().tolist())
        estado_filtro = st.selectbox("🗺️ Estado", estados, key="est_global")
    # Linha 3: Mês e Ano lado a lado
    fc5, fc6 = st.columns(2)
    with fc5:
        meses_opcoes = ['Todos'] + list(range(1, 13))
        mes_filtro = st.selectbox("📆 Mês", meses_opcoes, key="mes_global")
    with fc6:
        anos_opcoes = ['Todos'] + sorted(df['Ano'].dropna().unique().tolist(), reverse=True)
        ano_filtro = st.selectbox("🗓️ Ano", anos_opcoes, key="ano_global")
df_filtrado = df.copy()

if data_inicial:
    df_filtrado = df_filtrado[df_filtrado['DataEmissao'] >= pd.to_datetime(data_inicial)]
if data_final:
    df_filtrado = df_filtrado[df_filtrado['DataEmissao'] <= pd.to_datetime(data_final)]
if vendedor_filtro != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['Vendedor'] == vendedor_filtro]
if estado_filtro != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['Estado'] == estado_filtro]
if mes_filtro != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['Mes'] == mes_filtro]
if ano_filtro != 'Todos':
    df_filtrado = df_filtrado[df_filtrado['Ano'] == ano_filtro]

notas_unicas = obter_notas_unicas(df_filtrado)

st.sidebar.markdown("---")

# ====================== NAVEGAÇÃO ======================

_DESC = {
    "Dashboard":          "Visão geral de faturamento",
    "Positivação":        "Clientes atendidos no período",
    "Inadimplência":      "Títulos em aberto e atrasos",
    "Clientes sem Compra":"Clientes inativos para reativar",
    "Histórico":          "Consulta por cliente ou vendedor",
    "Preço Médio":        "Análise de preços por produto",
    "Pedidos Pendentes":  "Itens aguardando faturamento",
    "Rankings":           "Top vendedores e clientes",
    "Performance de Vendedores": "Painel completo de KPIs por vendedor",
}
_INFO_CARD = {}  # preenchido depois dos dados

if 'menu_option' not in st.session_state:
    # Admin vai para home; colaborador vai direto para primeiro módulo
    _tipo_usuario = usuario.get('tipo', 'administrador')
    if _tipo_usuario == 'administrador':
        st.session_state.menu_option = '__home__'
    else:
        _primeiro_modulo = modulos_permitidos[0] if modulos_permitidos else 'Inadimplência'
        st.session_state.menu_option = _primeiro_modulo

modulos_visiveis = modulos_permitidos if modulos_permitidos else [
    "Dashboard","Positivação","Inadimplência","Clientes sem Compra",
    "Histórico","Preço Médio","Pedidos Pendentes","Rankings","Performance de Vendedores"
]

# ══════════════════════════════════════════════════════════════════
# CSS GLOBAL — injetado uma única vez no topo do fluxo de página
# Cobre: sidebar radio, cards home (overlay + visual), métricas,
# gráficos, botões, filtros — tudo em um único bloco.
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

/* ── Base ── */
html, body, [class*="css"] { font-family: 'Inter','Segoe UI',sans-serif !important; }
.stApp { background-color: var(--background-color) !important; }

/* ── Remove padding padrão das colunas (alinha cards) ── */
div[data-testid="stHorizontalBlock"] { gap: 10px !important; }
div[data-testid="stHorizontalBlock"] > div[data-testid="stVerticalBlock"] {
    padding: 0 !important;
    margin: 0 !important;
    min-width: 0 !important;
}
div[data-testid="column"] { padding: 0 !important; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background: var(--secondary-background-color) !important;
    border-right: 1px solid rgba(128,128,128,0.2) !important;
}
section[data-testid="stSidebar"] .stRadio > label { display:none !important; }
section[data-testid="stSidebar"] .stRadio > div {
    display: flex !important; flex-direction: column !important; gap: 1px !important;
}
section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"] {
    padding: 9px 12px 9px 10px !important;
    border-radius: 0 10px 10px 0 !important;
    border-left: 3px solid transparent !important;
    margin: 0 !important; cursor: pointer !important;
    transition: all 0.15s !important; align-items: center !important;
}
section[data-testid="stSidebar"] .stRadio label[data-baseweb="radio"]:hover {
    background: rgba(31,71,136,0.08) !important; border-left-color: #8EB3E8 !important;
}
section[data-testid="stSidebar"] .stRadio label[aria-checked="true"] {
    background: linear-gradient(90deg,#EEF3FC,#F4F7FD) !important;
    border-left-color: #4A7BC8 !important;
}
section[data-testid="stSidebar"] .stRadio label[aria-checked="true"] p {
    color: #4A7BC8 !important; font-weight: 600 !important;
}
section[data-testid="stSidebar"] .stRadio div[class*="RadioMark"],
section[data-testid="stSidebar"] .stRadio div[class*="RadioMarkFill"],
section[data-testid="stSidebar"] .stRadio svg { display:none !important; }
section[data-testid="stSidebar"] .stRadio div[data-testid="stMarkdownContainer"] p {
    font-size: 0.88rem !important; color: #495057 !important;
    padding: 0 !important; margin: 0 !important;
}

/* ── HOME: card visual (div.med-card) ── */
div.med-card {
    background: #FFFFFF;
    border: 1px solid #E4E9F0;
    border-radius: 14px;
    padding: 20px 18px 16px;
    min-height: 138px;
    box-shadow: 0 1px 5px rgba(31,71,136,0.06);
    transition: box-shadow 0.18s, transform 0.18s, border-color 0.18s;
    cursor: pointer;
    position: relative;
    box-sizing: border-box;
}
div.med-card:hover {
    border-color: #B8CDF0 !important;
    box-shadow: 0 7px 22px rgba(31,71,136,0.14) !important;
    transform: translateY(-3px);
}

/* ── HOME: botão overlay invisível sobre o card ── */
div.med-card-col div[data-testid="stButton"] > button {
    position: relative !important;
    display: block !important;
    width: 100% !important;
    height: 138px !important;
    margin-top: -148px !important;
    opacity: 0 !important;
    cursor: pointer !important;
    border: none !important;
    background: transparent !important;
    z-index: 99 !important;
    padding: 0 !important;
}

div.med-card .mc-icon {
    width: 38px; height: 38px;
    background: #EEF3FC; border-radius: 9px;
    display: flex; align-items: center; justify-content: center;
    margin-bottom: 11px; color: #4A7BC8; font-size: 16px;
}
div.med-card .mc-title {
    font-size: 0.94rem; font-weight: 700;
    color: #2C5AA0; margin-bottom: 4px; letter-spacing: -0.01em;
}
div.med-card .mc-desc {
    font-size: 0.75rem; color: #6C757D; line-height: 1.4; margin-bottom: 9px;
}
div.med-card .mc-info {
    font-size: 0.70rem; color: #ADB5BD;
    border-top: 1px solid #F0F2F5; padding-top: 7px;
}

/* ── HOME: botão overlay invisível sobre o card ──
   O Streamlit renderiza o st.button logo após o st.markdown na mesma coluna.
   Usamos margin-top negativo para "subir" o botão sobre o card,
   e deixamos opacity:0 para ser invisível mas clicável. ── */
div.med-card-col div[data-testid="stButton"] > button {
    position: relative !important;
    display: block !important;
    width: 100% !important;
    height: 138px !important;
    margin-top: -148px !important;
    opacity: 0 !important;
    cursor: pointer !important;
    border: none !important;
    background: transparent !important;
    z-index: 99 !important;
    padding: 0 !important;
}

/* ── Métricas ── */
[data-testid="stMetric"] {
    background: var(--secondary-background-color) !important; border-radius: 12px !important;
    padding: 16px 18px !important; border-left: 4px solid #1F4788 !important;
    box-shadow: 0 1px 6px rgba(31,71,136,0.07) !important;
}
[data-testid="stMetricLabel"] p {
    font-size: 0.71rem !important; font-weight: 700 !important;
    text-transform: uppercase !important; letter-spacing: 0.06em !important;
    color: #8A96A8 !important;
}
[data-testid="stMetricValue"] {
    font-size: 1.4rem !important; font-weight: 700 !important; color: #2C5AA0 !important;
}

/* ── Botões gerais ── */
.stButton > button {
    border-radius: 8px !important; font-weight: 600 !important;
    font-size: 0.875rem !important;
}

/* ── Filtros ── */
div[data-testid="stHorizontalBlock"].filter-bar { background: #F2F5FA !important; }

/* ── Tabs ── */
[data-testid="stTabs"] [aria-selected="true"] {
    background: #1F4788 !important; color: #FFFFFF !important; font-weight: 600 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Sidebar: st.radio com labels com ícone Unicode ────────────────────────
_ICONES_NAV = {
    "Dashboard":"▦","Positivação":"✓","Inadimplência":"⚠",
    "Clientes sem Compra":"＋","Histórico":"◷","Preço Médio":"＄",
    "Pedidos Pendentes":"▣","Rankings":"▲","Performance de Vendedores":"★",
}
_ICONES_CARD = {
    "Dashboard":"▦","Positivação":"✓","Inadimplência":"⚠",
    "Clientes sem Compra":"＋","Histórico":"◷","Preço Médio":"＄",
    "Pedidos Pendentes":"▣","Rankings":"▲","Performance de Vendedores":"★",
}

with st.sidebar:
    st.markdown("""<div style="font-size:0.62rem;font-weight:700;color:#ADB5BD;
        letter-spacing:0.1em;text-transform:uppercase;
        margin-bottom:5px;padding-left:4px;">Navegação</div>""",
        unsafe_allow_html=True)

    # Botão Início
    if st.button("🏠  Início", key="nav_home", use_container_width=True, 
                 type="primary" if st.session_state.menu_option == '__home__' else "secondary"):
        st.session_state.menu_option = '__home__'
        st.rerun()
    
    # Botões dos módulos
    for modulo in modulos_visiveis:
        icone = _ICONES_NAV.get(modulo, '•')
        is_selected = (st.session_state.menu_option == modulo)
        
        if st.button(f"{icone}  {modulo}", 
                    key=f"nav_{modulo}", 
                    use_container_width=True,
                    type="primary" if is_selected else "secondary"):
            st.session_state.menu_option = modulo
            st.rerun()

    st.sidebar.markdown("---")
    _cons_sel = st.session_state.menu_option == 'Consulta Clientes'
    if st.button("🔎  Consulta Tabela",
                 key="nav_consulta_clientes",
                 use_container_width=True,
                 type="primary" if _cons_sel else "secondary"):
        st.session_state.menu_option = 'Consulta Clientes'
        st.rerun()

    st.sidebar.markdown("---")
    st.sidebar.markdown("""<div style="font-size:0.62rem;font-weight:700;color:#ADB5BD;
        letter-spacing:0.1em;text-transform:uppercase;
        margin-bottom:5px;padding-left:4px;">Relatório Semanal</div>""",
        unsafe_allow_html=True)

    if st.button("📦 Gerar Relatório Semanal", key="btn_semanal", use_container_width=True):
        import zipfile, io as _io
        from datetime import date

        _hoje = pd.Timestamp.now()
        _inicio_mes = _hoje.replace(day=1)

        _zip_buf = _io.BytesIO()

        with st.sidebar.spinner("Gerando relatórios..."):
            try:
                # ── Carregar inadimplência ──
                _df_inad_sem = None
                if planilhas_disponiveis.get('inadimplencia'):
                    _raw_inad = carregar_planilha_github(planilhas_disponiveis['inadimplencia']['url'])
                    if _raw_inad is not None:
                        _df_inad_sem = processar_inadimplencia(_raw_inad)

                # ── Carregar pedidos pendentes ──
                _df_pend_sem = None
                if planilhas_disponiveis.get('pedidos_pendentes'):
                    try:
                        import zipfile as _zf, xml.etree.ElementTree as _ET
                        _resp_p = requests.get(planilhas_disponiveis['pedidos_pendentes']['url'])
                        _ef = _io.BytesIO(_resp_p.content)
                        _data_p = []
                        _cur_cli = None
                        with _zf.ZipFile(_ef) as _z:
                            with _z.open('xl/sharedStrings.xml') as _f:
                                _st = _ET.parse(_f)
                                _ns_s = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                                _ss = [si.text if si.text else '' for si in _st.findall('.//ss:t', _ns_s)]
                            with _z.open('xl/worksheets/sheet1.xml') as _f:
                                _sh = _ET.parse(_f)
                                _ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                                for _row in _sh.findall('.//ss:row', _ns):
                                    _rd = {}
                                    for _cell in _row.findall('.//ss:c', _ns):
                                        _ref = _cell.get('r', '')
                                        _col = ''.join([c for c in _ref if c.isalpha()])
                                        _ve = _cell.find('.//ss:v', _ns)
                                        if _ve is not None and _ve.text:
                                            if _cell.get('t', 'n') == 's':
                                                _idx = int(_ve.text)
                                                _rd[_col] = _ss[_idx] if _idx < len(_ss) else _ve.text
                                            else:
                                                _rd[_col] = _ve.text
                                    if not _rd:
                                        continue
                                    _ca, _cb = _rd.get('A', ''), _rd.get('B', '')
                                    if _ca and not _cb and 'N° do pedido' not in _ca and 'Valor Total' not in _ca and _ca != 'Subgrupo:':
                                        _cur_cli = _ca
                                    elif 'N° do pedido' in _ca:
                                        _desc = _rd.get('C', '')
                                        if _desc and ' - ' in _desc:
                                            try:
                                                _qtdc = float(_rd.get('D', 0))
                                                _vunit = float(_rd.get('E', 0))
                                                _qtde = float(_rd.get('H', 0))
                                                _qtdp = _qtdc - _qtde
                                                _dt_v = _rd.get('G', '')
                                                _dt_em = (pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(_dt_v))) if _dt_v else None
                                                _data_p.append({
                                                    'Cliente': _cur_cli,
                                                    'NumeroPedido': _cb,
                                                    'CodigoProduto': _desc.split(' - ')[0].strip(),
                                                    'Descricao': _desc,
                                                    'QtdContratada': _qtdc,
                                                    'QtdEntregue': _qtde,
                                                    'QtdPendente': _qtdp,
                                                    'ValorUnit': _vunit,
                                                    'ValorPendente': _qtdp * _vunit,
                                                    'DataEmissao': _dt_em,
                                                    'Vendedor': _rd.get('J', ''),
                                                    'PercEntregue': float(_rd.get('I', 0))
                                                })
                                            except:
                                                continue
                        _df_pend_sem = pd.DataFrame(_data_p)
                        # Filtrar: primeiro dia do mês até hoje
                        if len(_df_pend_sem) > 0 and 'DataEmissao' in _df_pend_sem.columns:
                            _df_pend_sem = _df_pend_sem[_df_pend_sem['QtdPendente'] > 0]
                            _df_pend_sem = _df_pend_sem[
                                (_df_pend_sem['DataEmissao'] >= _inicio_mes) &
                                (_df_pend_sem['DataEmissao'] <= _hoje)
                            ]
                    except:
                        _df_pend_sem = None

                # ── Faturados: primeiro dia do mês até hoje ──
                _df_fat_sem = df[
                    (df['TipoMov'] == 'NF Venda') &
                    (df['DataEmissao'] >= _inicio_mes) &
                    (df['DataEmissao'] <= _hoje)
                ].copy()

                # ── Lista de vendedores ativos ──
                _vendedores_ativos = sorted(df[
                    (df['TipoMov'] == 'NF Venda') &
                    (df['DataEmissao'] >= _inicio_mes)
                ]['Vendedor'].dropna().unique().tolist())

                with zipfile.ZipFile(_zip_buf, 'w', zipfile.ZIP_DEFLATED) as _zout:
                    for _vend in _vendedores_ativos:
                        _vend_pasta = _vend.upper().replace(' ', '_')
                        _prefixo = f"RELATORIO SEMANAL/{_vend_pasta}/"

                        # ── 1. FATURADOS ──
                        _df_v_fat = _df_fat_sem[_df_fat_sem['Vendedor'] == _vend].copy()
                        if len(_df_v_fat) > 0:
                            _cols = [c for c in ['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado', 'Vendedor',
                                                  'DataEmissao', 'Numero_NF', 'TipoMov',
                                                  'CodigoProduto', 'NomeProduto', 'Quantidade', 'PrecoUnit',
                                                  'TotalProduto', 'Valor_Real'] if c in _df_v_fat.columns]
                            _df_exp_f = _df_v_fat[_cols].copy()
                            _df_exp_f['DataEmissao'] = _df_exp_f['DataEmissao'].dt.strftime('%d/%m/%Y')

                            # Aba FATURAMENTO TOTAL: dedup Numero_NF + soma
                            _cols_oc = ['CodigoProduto', 'NomeProduto', 'Quantidade', 'PrecoUnit', 'TotalProduto', 'Valor_Real']
                            _cols_ft = [c for c in _cols if c not in _cols_oc]
                            _df_ft = _df_v_fat.drop_duplicates(subset=['Numero_NF'], keep='first')[_cols_ft + ['TotalProduto']].copy()
                            _df_ft['DataEmissao'] = _df_ft['DataEmissao'].dt.strftime('%d/%m/%Y')
                            _soma_ft = _df_ft['TotalProduto'].sum()
                            _ln_tot = {c: '' for c in _df_ft.columns}
                            _ln_tot['TotalProduto'] = _soma_ft
                            _ln_tot['RazaoSocial'] = 'TOTAL'
                            _df_ft = pd.concat([_df_ft, pd.DataFrame([_ln_tot])], ignore_index=True)

                            _buf_f = _io.BytesIO()
                            with pd.ExcelWriter(_buf_f, engine='xlsxwriter') as _wr:
                                _wb = _wr.book
                                _df_exp_f.to_excel(_wr, index=False, sheet_name='PRODUTOS POR CLIENTE')
                                _ws1 = _wr.sheets['PRODUTOS POR CLIENTE']
                                if len(_df_exp_f) > 0:
                                    _ws1.add_table(0, 0, len(_df_exp_f), len(_df_exp_f.columns)-1, {
                                        'name': f'TblPC_{_vend_pasta[:10]}',
                                        'style': 'Table Style Medium 2',
                                        'columns': [{'header': c} for c in _df_exp_f.columns]
                                    })
                                _df_ft.to_excel(_wr, index=False, sheet_name='FATURAMENTO TOTAL')
                                _ws2 = _wr.sheets['FATURAMENTO TOTAL']
                                _nft = len(_df_ft) - 1
                                if _nft > 0:
                                    _ws2.add_table(0, 0, _nft, len(_df_ft.columns)-1, {
                                        'name': f'TblFT_{_vend_pasta[:10]}',
                                        'style': 'Table Style Medium 2',
                                        'columns': [{'header': c} for c in _df_ft.columns]
                                    })
                                _fmt_b = _wb.add_format({'bold': True, 'num_format': '#,##0.00'})
                                _sc = list(_df_ft.columns).index('TotalProduto')
                                _ws2.write(_nft + 1, _sc, _soma_ft, _fmt_b)
                            _zout.writestr(_prefixo + f"{_vend_pasta}_FATURADOS.xlsx", _buf_f.getvalue())

                        # ── 2. PENDENTES ──
                        if _df_pend_sem is not None and len(_df_pend_sem) > 0:
                            _df_v_pend = _df_pend_sem[_df_pend_sem['Vendedor'] == _vend].copy()
                            if len(_df_v_pend) > 0:
                                _buf_p = _io.BytesIO()
                                with pd.ExcelWriter(_buf_p, engine='xlsxwriter') as _wr:
                                    _df_v_pend.to_excel(_wr, index=False, sheet_name='PENDENTES')
                                    _wsp = _wr.sheets['PENDENTES']
                                    _wsp.add_table(0, 0, len(_df_v_pend), len(_df_v_pend.columns)-1, {
                                        'name': f'TblPend_{_vend_pasta[:10]}',
                                        'style': 'Table Style Medium 2',
                                        'columns': [{'header': c} for c in _df_v_pend.columns]
                                    })
                                _zout.writestr(_prefixo + f"PENDENTES_{_vend_pasta}.xlsx", _buf_p.getvalue())

                        # ── 3. INADIMPLÊNCIA ──
                        if _df_inad_sem is not None and len(_df_inad_sem) > 0:
                            _df_v_inad = _df_inad_sem[_df_inad_sem['Vendedor'] == _vend].copy()
                            if len(_df_v_inad) > 0:
                                _cols_inad = [c for c in ['Vendedor', 'Cliente', 'NumeroDoc', 'DataVencimento',
                                                           'ValorLiquido', 'DiasAtraso', 'FaixaAtraso', 'Banco', 'Estado']
                                              if c in _df_v_inad.columns]
                                _df_v_inad = _df_v_inad[_cols_inad].copy()
                                if 'DataVencimento' in _df_v_inad.columns:
                                    _df_v_inad['DataVencimento'] = _df_v_inad['DataVencimento'].dt.strftime('%d/%m/%Y')
                                _buf_i = _io.BytesIO()
                                with pd.ExcelWriter(_buf_i, engine='xlsxwriter') as _wr:
                                    _df_v_inad.to_excel(_wr, index=False, sheet_name='INADIMPLENCIA')
                                    _wsi = _wr.sheets['INADIMPLENCIA']
                                    _wsi.add_table(0, 0, len(_df_v_inad), len(_df_v_inad.columns)-1, {
                                        'name': f'TblInad_{_vend_pasta[:10]}',
                                        'style': 'Table Style Medium 2',
                                        'columns': [{'header': c} for c in _df_v_inad.columns]
                                    })
                                _zout.writestr(_prefixo + f"INADIMPLENCIA_{_vend_pasta}.xlsx", _buf_i.getvalue())

                st.session_state['_zip_semanal'] = _zip_buf.getvalue()
                st.session_state['_zip_semanal_nome'] = f"RELATORIO_SEMANAL_{_hoje.strftime('%d-%m-%Y')}.zip"
                st.rerun()

            except Exception as _e:
                st.sidebar.error(f"Erro: {_e}")

    if st.session_state.get('_zip_semanal'):
        st.sidebar.download_button(
            "💾 Baixar ZIP Semanal",
            st.session_state['_zip_semanal'],
            st.session_state.get('_zip_semanal_nome', 'RELATORIO_SEMANAL.zip'),
            "application/zip",
            key="download_zip_semanal"
        )


if st.session_state.menu_option == '__home__':
    usuario_info = st.session_state.get("usuario", {})

    # ── Calcular previews dos cards ──────────────────────────────────────
    _now   = pd.Timestamp.now()
    _mes   = _now.month
    _ano   = _now.year
    _6m    = _now - pd.DateOffset(months=6)

    # Dashboard
    try:
        vendas_mes = notas_unicas[
            (notas_unicas['DataEmissao'].dt.month == _mes) &
            (notas_unicas['DataEmissao'].dt.year  == _ano)
        ]['Valor_Real'].sum()
    except:
        vendas_mes = 0

    # Positivação
    try:
        _base_total   = df['CPF_CNPJ'].nunique()
        _posit_mes    = df[
            (df['TipoMov']=='NF Venda') &
            (df['DataEmissao'].dt.month==_mes) &
            (df['DataEmissao'].dt.year==_ano)
        ]['CPF_CNPJ'].nunique()
        _info_posit   = f"{_base_total:,} na base · {_posit_mes:,} positivados no mês"
    except:
        _info_posit   = "Base de clientes"

    # Inadimplência — carregada separadamente, usar placeholder se não disponível
    try:
        if planilhas_disponiveis.get('inadimplencia'):
            _df_inad = carregar_planilha_github(planilhas_disponiveis['inadimplencia']['url'])
            if _df_inad is not None:
                _df_inad = processar_inadimplencia(_df_inad)
                _val_inad = _df_inad['ValorLiquido'].sum()
                _cli_inad = _df_inad['Cliente'].nunique()
                _info_inad = f"R$ {_val_inad:,.0f} · {_cli_inad:,} clientes"
            else:
                _info_inad = "Dados não disponíveis"
        else:
            _info_inad = "Planilha não configurada"
    except:
        _info_inad = "Títulos em aberto e atrasos"

    # Clientes sem compra há mais de 6 meses
    try:
        _ultima_compra = df[df['TipoMov']=='NF Venda'].groupby('CPF_CNPJ')['DataEmissao'].max()
        _sem_6m = (_ultima_compra < _6m).sum()
        _info_churn = f"{_sem_6m:,} clientes sem compra há +6 meses"
    except:
        _info_churn = "Clientes inativos"

    # Pedidos pendentes — carregado separadamente
    try:
        if planilhas_disponiveis.get('pedidos_pendentes'):
            _resp = requests.get(planilhas_disponiveis['pedidos_pendentes']['url'], timeout=15)
            import zipfile, xml.etree.ElementTree as ET
            from io import BytesIO as _BytesIO
            _ef = _BytesIO(_resp.content)
            with zipfile.ZipFile(_ef) as _z:
                with _z.open('xl/sharedStrings.xml') as _f:
                    _st = ET.parse(_f)
                    _ns = {'ss':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    _ss = [s.text or '' for s in _st.findall('.//ss:t',_ns)]
                with _z.open('xl/worksheets/sheet1.xml') as _f:
                    _sh = ET.parse(_f)
                    _val_pend = 0.0; _cli_pend = set()
                    for _row in _sh.findall('.//ss:row',_ns):
                        _rd = {}
                        for _c in _row.findall('.//ss:c',_ns):
                            _ref=''.join([x for x in _c.get('r','') if x.isalpha()])
                            _v=_c.find('.//ss:v',_ns)
                            if _v is not None and _v.text:
                                _rd[_ref]=_ss[int(_v.text)] if _c.get('t')=='s' else _v.text
                        _ca=_rd.get('A',''); _cb=_rd.get('B','')
                        if _ca and not _cb and 'N° do pedido' not in _ca:
                            _cur_cli=_ca
                        elif 'N° do pedido' in _ca:
                            try:
                                _qc=float(_rd.get('D',0)); _vu=float(_rd.get('E',0))
                                _qe=float(_rd.get('H',0)); _qp=_qc-_qe
                                _val_pend+=_qp*_vu
                                if hasattr(_cur_cli,'__len__'): _cli_pend.add(_cur_cli)
                            except: pass
            _info_pend = f"R$ {_val_pend:,.0f} · {len(_cli_pend):,} clientes"
        else:
            _info_pend = "Aguardando faturamento"
    except:
        _info_pend = "Aguardando faturamento"

    # Rankings — top 3 vendedores DO MÊS ATUAL
    try:
        # Filtrar apenas vendas do mês atual (mesma lógica do Dashboard)
        vendas_mes_atual = notas_unicas[
            (notas_unicas['TipoMov'] == 'NF Venda') &
            (notas_unicas['DataEmissao'].dt.month == _mes) &
            (notas_unicas['DataEmissao'].dt.year == _ano)
        ]
        _rank = vendas_mes_atual.groupby('Vendedor')['TotalProduto'].sum().nlargest(3)
        _info_rank = " · ".join([f"{v.split()[0]} R${r:,.0f}" for v,r in _rank.items()])
    except:
        _info_rank = "Top vendedores e clientes"

    cards_data = [
        {'nome':'Dashboard',          'info':f'R$ {vendas_mes:,.0f} no mês atual'},
        {'nome':'Positivação',         'info':_info_posit},
        {'nome':'Inadimplência',       'info':_info_inad},
        {'nome':'Clientes sem Compra', 'info':_info_churn},
        {'nome':'Histórico',           'info':'Por cliente ou vendedor'},
        {'nome':'Preço Médio',         'info':'Análise por produto'},
        {'nome':'Pedidos Pendentes',   'info':_info_pend},
        {'nome':'Rankings',            'info':_info_rank},
        {'nome':'Performance de Vendedores', 'info':'Análise completa por vendedor'},
    ]
    cards_visiveis = [c for c in cards_data if c['nome'] in modulos_visiveis]

    st.markdown(f"""
    <div style="margin-bottom:20px;padding-bottom:14px;border-bottom:1px solid #E9ECEF;">
        <div style="font-size:1.45rem;font-weight:600;color:#2C5AA0;margin-bottom:3px;">
            Olá, {usuario_info.get('nome','Usuário')}
        </div>
        <div style="color:#8A96A8;font-size:0.87rem;">
            Selecione um módulo abaixo para iniciar a análise.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Grid 4 colunas — técnica de botão overlay que FUNCIONA
    try:
        from streamlit_card import card as st_card
        _USE_CARD_LIB = True
    except ImportError:
        _USE_CARD_LIB = False

    # ── Detectar mobile via streamlit-js-eval ───────────────────────────
    try:
        from streamlit_js_eval import streamlit_js_eval
        _screen_width = streamlit_js_eval(
            js_expressions="window.innerWidth",
            key="screen_width"
        )
        _is_mobile = bool(_screen_width and int(_screen_width) <= 768)
    except Exception:
        _is_mobile = False

    _n_cols = 2 if _is_mobile else 4

    for row_start in range(0, len(cards_visiveis), _n_cols):
        row = cards_visiveis[row_start:row_start+_n_cols]
        cols = st.columns(_n_cols)
        for j, c in enumerate(row):
            with cols[j]:
                nome = c['nome']
                desc = _DESC.get(nome, '')
                info = c['info']
                ic   = _ICONES_CARD.get(nome, '•')

                if _USE_CARD_LIB:
                    clicked = st_card(
                        title=f"{ic}  {nome}",
                        text=[desc, info],
                        key=f"hc_{nome}",
                        styles={
                            "card": {
                                "width": "100%",
                                "height": "190px",
                                "background-color": "var(--secondary-background-color)",
                                "border": "1px solid #E4E9F0",
                                "border-radius": "14px",
                                "box-shadow": "0 1px 6px rgba(31,71,136,0.07)",
                                "cursor": "pointer",
                                "padding": "22px 20px 18px 20px",
                                "transition": "box-shadow 0.18s ease, transform 0.18s ease, border-color 0.18s ease",
                                "font-family": "'Inter', 'Segoe UI', Roboto, sans-serif",
                                "margin": "0",
                            },
                            "title": {
                                "font-size": "1rem",
                                "font-weight": "700",
                                "color": "#2C5AA0",
                                "font-family": "'Inter', 'Segoe UI', Roboto, sans-serif",
                                "letter-spacing": "-0.01em",
                                "margin-bottom": "6px",
                                "white-space": "nowrap",
                                "overflow": "hidden",
                                "text-overflow": "ellipsis",
                            },
                            "text": {
                                "font-size": "0.78rem",
                                "color": "#6C757D",
                                "font-family": "'Inter', 'Segoe UI', Roboto, sans-serif",
                                "line-height": "1.45",
                                "font-weight": "400",
                            },
                            "filter": {"background": "rgba(0,0,0,0)"},
                        }
                    )
                    if clicked:
                        st.session_state.menu_option = nome
                        st.rerun()
                else:
                    st.markdown(f"""
                    <div style="background:var(--secondary-background-color);border:1px solid rgba(128,128,128,0.2);
                                border-radius:14px;padding:20px 18px;min-height:148px;
                                box-shadow:0 1px 6px rgba(31,71,136,0.07);
                                font-family:'Inter','Segoe UI',sans-serif;">
                        <div style="font-size:1rem;margin-bottom:10px;">{ic}</div>
                        <div style="font-size:0.95rem;font-weight:700;color:#2C5AA0;
                                    margin-bottom:5px;">{nome}</div>
                        <div style="font-size:0.76rem;color:#6C757D;
                                    margin-bottom:8px;">{desc}</div>
                        <div style="font-size:0.70rem;color:#ADB5BD;
                                    border-top:1px solid #F0F2F5;padding-top:7px;">{info}</div>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button(f"Abrir {nome}", key=f"hc_{nome}",
                                 use_container_width=True):
                        st.session_state.menu_option = nome
                        st.rerun()

        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
    st.stop()

# ── Módulo ativo ──────────────────────────────────────────────────────────
menu = st.session_state.menu_option

st.markdown(f"""
<div style="font-size:0.74rem;color:#ADB5BD;margin-bottom:14px;
            padding-bottom:10px;border-bottom:1px solid #F0F2F5;">
    <span style="color:#6C757D;">Início</span>
    <span style="margin:0 6px;color:#D0D5DE;">›</span>
    <span style="color:#4A7BC8;font-weight:600;">{menu}</span>
</div>
""", unsafe_allow_html=True)

if menu not in modulos_permitidos:
    st.markdown("""
    <div style="background:#FFF3F3;border:1px solid #F5C6CB;border-radius:10px;
                padding:16px 20px;color:#721C24;font-size:0.9rem;">
        Acesso negado. Você não tem permissão para acessar este módulo.
    </div>""", unsafe_allow_html=True)
    st.stop()
# ====================== DASHBOARD ======================
if menu == "Dashboard":
    # KPIs principais com cards customizados
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        vendas_brutas = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda']['TotalProduto'].sum()
        render_kpi_card("Faturamento Bruto", f"R$ {vendas_brutas:,.0f}", icon="💰", color="#1F4788")
    
    with col2:
        faturamento_liquido = notas_unicas['Valor_Real'].sum()
        render_kpi_card("Faturamento Líquido", f"R$ {faturamento_liquido:,.0f}", icon="💵", color="#10B981")
    
    with col3:
        clientes_unicos = df_filtrado['CPF_CNPJ'].nunique()
        render_kpi_card("Clientes Únicos", f"{clientes_unicos:,}", icon="👥", color="#F59E0B")
    
    with col4:
        total_notas = len(notas_unicas[notas_unicas['TipoMov'] == 'NF Venda'])
        render_kpi_card("Notas de Venda", f"{total_notas:,}", icon="📄", color="#EF4444")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Segunda linha de KPIs
    col1b, col2b, col3b, col4b = st.columns(4)
    
    with col1b:
        total_devolucoes = notas_unicas[notas_unicas['TipoMov'] == 'NF Dev.Venda']['TotalProduto'].sum()
        render_kpi_card("Devoluções", f"R$ {total_devolucoes:,.0f}", icon="↩️", color="#E5E7EB")
    
    with col2b:
        ticket_medio = vendas_brutas / clientes_unicos if clientes_unicos > 0 else 0
        render_kpi_card("Ticket Médio", f"R$ {ticket_medio:,.0f}", icon="🎯", color="#E5E7EB")
    
    with col3b:
        qtd_notas_dev = len(notas_unicas[notas_unicas['TipoMov'] == 'NF Dev.Venda'])
        render_kpi_card("Notas Devolução", f"{qtd_notas_dev:,}", icon="📋", color="#E5E7EB")
    
    with col4b:
        taxa_devolucao = (total_devolucoes / vendas_brutas * 100) if vendas_brutas > 0 else 0
        render_kpi_card("Taxa Devolução", f"{taxa_devolucao:.1f}%", icon="📊", color="#E5E7EB")
    
    st.markdown("---")

    # Linha 1: 3 gráficos
    col5, col6, col7 = st.columns(3)

    with col5:
        st.subheader("📈 Evolução de Vendas")
        vendas_tempo = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda'].groupby('MesAno')['TotalProduto'].sum().reset_index().sort_values('MesAno')
        if len(vendas_tempo) > 0:
            fig_linha = px.line(vendas_tempo, x='MesAno', y='TotalProduto',
                labels={'MesAno': 'Período', 'TotalProduto': 'Valor (R$)'})
            fig_linha.update_traces(line_color='#1F4788', line_width=3, mode='lines+markers', marker=dict(size=5, color='#1F4788'))
            fig_linha.update_layout(xaxis_title="Período", yaxis_title="Valor (R$)", hovermode='x unified')
            fig_linha = aplicar_layout_grafico(fig_linha)
            st.plotly_chart(fig_linha, use_container_width=True)
        else:
            st.info("Sem dados para exibir")

    with col6:
        st.subheader("🗺️ Top 10 Estados")
        vendas_estado = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda'].groupby('Estado')['TotalProduto'].sum().reset_index().sort_values('TotalProduto', ascending=False).head(10)
        fig_bar = px.bar(vendas_estado, x='Estado', y='TotalProduto',
            labels={'Estado': 'Estado', 'TotalProduto': 'Valor (R$)'},
            color_discrete_sequence=['#2E86AB'])
        fig_bar = aplicar_layout_grafico(fig_bar)
        st.plotly_chart(fig_bar, use_container_width=True)

    with col7:
        st.subheader("👥 Positivação por Vendedor")
        atendidos = df_filtrado[df_filtrado['TipoMov'] == 'NF Venda'].groupby('Vendedor')['CPF_CNPJ'].nunique().reset_index()
        atendidos.columns = ['Vendedor', 'Clientes']
        atendidos = atendidos.sort_values('Clientes', ascending=False).head(10)
        fig_posit = px.bar(atendidos, x='Vendedor', y='Clientes',
            labels={'Vendedor': 'Vendedor', 'Clientes': 'Clientes Atendidos'},
            color_discrete_sequence=['#1F4788'])
        fig_posit = aplicar_layout_grafico(fig_posit)
        st.plotly_chart(fig_posit, use_container_width=True)

    st.markdown("---")

    # Linha 2: 3 gráficos
    col8, col9, col10 = st.columns(3)

    with col8:
        st.subheader("🏆 Top 10 Clientes")
        ranking_clientes = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda'].groupby('RazaoSocial')['TotalProduto'].sum().reset_index().sort_values('TotalProduto', ascending=False).head(10)
        fig_clientes = px.bar(ranking_clientes, x='TotalProduto', y='RazaoSocial', orientation='h',
            labels={'RazaoSocial': 'Cliente', 'TotalProduto': 'Valor (R$)'},
            color_discrete_sequence=['#4A7BC8'])
        fig_clientes = aplicar_layout_grafico(fig_clientes)
        st.plotly_chart(fig_clientes, use_container_width=True)

    with col9:
        st.subheader("⚠️ Clientes sem Compra")
        _com_venda = set(df_filtrado[df_filtrado['TipoMov'] == 'NF Venda']['CPF_CNPJ'].unique())
        _todos = df.sort_values('DataEmissao').groupby('CPF_CNPJ').last().reset_index()
        _vhist = df[df['TipoMov'] == 'NF Venda'].groupby('CPF_CNPJ')['TotalProduto'].sum().reset_index()
        _vhist.columns = ['CPF_CNPJ', 'ValorHistorico']
        _todos = pd.merge(_todos, _vhist, on='CPF_CNPJ', how='left').fillna(0)
        _sem = _todos[~_todos['CPF_CNPJ'].isin(_com_venda)].sort_values('ValorHistorico', ascending=False).head(10)
        fig_churn = px.bar(_sem, x='ValorHistorico', y='RazaoSocial', orientation='h',
            labels={'RazaoSocial': 'Cliente', 'ValorHistorico': 'Valor Histórico (R$)'},
            color_discrete_sequence=['#1F4788'])
        fig_churn = aplicar_layout_grafico(fig_churn)
        st.plotly_chart(fig_churn, use_container_width=True)

    with col10:
        st.subheader("📊 Ranking de Vendedores")
        ranking_vendedores = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda'].groupby('Vendedor')['TotalProduto'].sum().reset_index().sort_values('TotalProduto', ascending=False).head(10)
        fig_rank_vend = px.bar(ranking_vendedores, x='TotalProduto', y='Vendedor', orientation='h',
            labels={'Vendedor': 'Vendedor', 'TotalProduto': 'Valor Total (R$)'},
            color_discrete_sequence=['#163561'])
        fig_rank_vend = aplicar_layout_grafico(fig_rank_vend)
        st.plotly_chart(fig_rank_vend, use_container_width=True)

# ====================== POSITIVAÇÃO ======================
elif menu == "Positivação":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Relatório de Positivação</h2>', unsafe_allow_html=True)

    # ── KPIs do mês vigente no topo ───────────────────────────────────────
    _mes_atual = pd.Timestamp.now().month
    _ano_atual = pd.Timestamp.now().year
    _vendas_mes = df_filtrado[
        (df_filtrado['TipoMov'] == 'NF Venda') &
        (df_filtrado['DataEmissao'].dt.month == _mes_atual) &
        (df_filtrado['DataEmissao'].dt.year == _ano_atual)
    ]
    _posit_mes    = _vendas_mes['CPF_CNPJ'].nunique()
    _total_base   = df['CPF_CNPJ'].nunique()
    _perc_posit   = (_posit_mes / _total_base * 100) if _total_base > 0 else 0

    _kp1, _kp2, _kp3 = st.columns(3)
    with _kp1:
        st.metric("Positivados no Mês", f"{_posit_mes:,} clientes",
                  help="Clientes com ao menos uma compra no mês vigente")
    with _kp2:
        st.metric("Total da Base", f"{_total_base:,} clientes",
                  help="Total de clientes únicos na base")
    with _kp3:
        st.metric("% da Base Positivada", f"{_perc_posit:.1f}%",
                  help="Percentual da base que comprou no mês vigente")

    st.markdown("---")

    tab1, tab2, tab3_fat = st.tabs(["📊 Por Vendedor", "🗺️ Por Estado", "🧾 Pedidos Faturados"])
    
    with tab1:
        base_vendedor = df.groupby('Vendedor')['CPF_CNPJ'].nunique().reset_index()
        base_vendedor.columns = ['Vendedor', 'TotalBase']
        
        vendas_periodo = df_filtrado[df_filtrado['TipoMov'] == 'NF Venda']
        atendidos = vendas_periodo.groupby('Vendedor')['CPF_CNPJ'].nunique().reset_index()
        atendidos.columns = ['Vendedor', 'QtdAtendidos']
        
        valor_vendedor = obter_notas_unicas(vendas_periodo).groupby('Vendedor')['Valor_Real'].sum().reset_index()
        valor_vendedor.columns = ['Vendedor', 'ValorTotal']
        
        relatorio_positivacao = pd.merge(base_vendedor, atendidos, on='Vendedor', how='left')
        relatorio_positivacao = pd.merge(relatorio_positivacao, valor_vendedor, on='Vendedor', how='left')
        relatorio_positivacao['QtdAtendidos'] = relatorio_positivacao['QtdAtendidos'].fillna(0).astype(int)
        relatorio_positivacao['ValorTotal'] = relatorio_positivacao['ValorTotal'].fillna(0)
        relatorio_positivacao['Percentual'] = (relatorio_positivacao['QtdAtendidos'] / relatorio_positivacao['TotalBase'] * 100).round(1)
        relatorio_positivacao = relatorio_positivacao.sort_values('QtdAtendidos', ascending=False)
        
        fig_posit_vend = px.bar(
            relatorio_positivacao.head(15),
            x='Vendedor',
            y='Percentual',
            labels={'Vendedor': 'Vendedor', 'Percentual': 'Positivação (%)'},
            color='Percentual',
            color_discrete_sequence=['#1F4788'],
            title='Top 15 Vendedores - Taxa de Positivação'
        )
        fig_posit_vend = aplicar_layout_grafico(fig_posit_vend)
        st.plotly_chart(fig_posit_vend, use_container_width=True)
        
        # Formatar para exibição
        relatorio_positivacao_display = formatar_dataframe_moeda(relatorio_positivacao, ['ValorTotal'])
        st.dataframe(relatorio_positivacao_display, use_container_width=True)
        
        st.download_button(
            "📥 Exportar Positivação por Vendedor",
            to_excel(relatorio_positivacao),
            "positivacao_vendedor.xlsx",
            "application/vnd.ms-excel"
        )
        
        st.markdown("---")
        
        st.subheader("📋 Detalhamento de Clientes")
        vendedor_selecionado = st.selectbox(
            "Selecione o vendedor",
            relatorio_positivacao['Vendedor'].tolist()
        )
        
        if vendedor_selecionado:
            notas_vendedor = obter_notas_unicas(vendas_periodo[vendas_periodo['Vendedor'] == vendedor_selecionado])
            
            clientes_vendedor = notas_vendedor.groupby(['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado']).agg({
                'Valor_Real': 'sum'
            }).reset_index()
            clientes_vendedor.columns = ['CPF/CNPJ', 'Razão Social', 'Cidade', 'Estado', 'Valor Total']
            clientes_vendedor = clientes_vendedor.sort_values('Valor Total', ascending=False)
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Clientes Atendidos", len(clientes_vendedor))
            with col2:
                st.metric("Valor Total", f"R$ {clientes_vendedor['Valor Total'].sum():,.2f}")
            
            # Formatar para exibição
            clientes_vendedor_display = formatar_dataframe_moeda(clientes_vendedor, ['Valor Total'])
            st.dataframe(clientes_vendedor_display, use_container_width=True)
            
            st.download_button(
                f"📥 Exportar Clientes - {vendedor_selecionado}",
                to_excel(clientes_vendedor),
                f"clientes_{vendedor_selecionado}.xlsx",
                "application/vnd.ms-excel"
            )
    
    with tab2:
        st.subheader("🗺️ Positivação por Estado")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            vendedor_estado_filtro = st.selectbox(
                "Filtrar por Vendedor",
                ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist()),
                key="vend_estado"
            )
        with col_f2:
            ano_estado_filtro = st.selectbox(
                "Filtrar por Ano",
                ['Todos'] + sorted(df['Ano'].dropna().unique().tolist(), reverse=True),
                key="ano_estado"
            )
        
        df_estado_filtrado = df_filtrado.copy()
        if vendedor_estado_filtro != 'Todos':
            df_estado_filtrado = df_estado_filtrado[df_estado_filtrado['Vendedor'] == vendedor_estado_filtro]
        if ano_estado_filtro != 'Todos':
            df_estado_filtrado = df_estado_filtrado[df_estado_filtrado['Ano'] == ano_estado_filtro]
        
        base_estado = df.groupby('Estado')['CPF_CNPJ'].nunique().reset_index()
        base_estado.columns = ['Estado', 'TotalBase']
        
        vendas_estado = df_estado_filtrado[df_estado_filtrado['TipoMov'] == 'NF Venda']
        atendidos_estado = vendas_estado.groupby('Estado')['CPF_CNPJ'].nunique().reset_index()
        atendidos_estado.columns = ['Estado', 'QtdAtendidos']
        
        valor_estado = obter_notas_unicas(vendas_estado).groupby('Estado')['Valor_Real'].sum().reset_index()
        valor_estado.columns = ['Estado', 'ValorTotal']
        
        relatorio_estado = pd.merge(base_estado, atendidos_estado, on='Estado', how='left')
        relatorio_estado = pd.merge(relatorio_estado, valor_estado, on='Estado', how='left')
        relatorio_estado['QtdAtendidos'] = relatorio_estado['QtdAtendidos'].fillna(0).astype(int)
        relatorio_estado['ValorTotal'] = relatorio_estado['ValorTotal'].fillna(0)
        relatorio_estado['Percentual'] = (relatorio_estado['QtdAtendidos'] / relatorio_estado['TotalBase'] * 100).round(1)
        relatorio_estado = relatorio_estado.sort_values('Percentual', ascending=False)
        
        fig_posit_estado = px.bar(
            relatorio_estado.head(15),
            x='Estado',
            y='Percentual',
            labels={'Estado': 'Estado', 'Percentual': 'Positivação (%)'},
            color='Percentual',
            color_discrete_sequence=['#2E86AB'],
            title='Top 15 Estados - Taxa de Positivação'
        )
        fig_posit_estado = aplicar_layout_grafico(fig_posit_estado)
        st.plotly_chart(fig_posit_estado, use_container_width=True)
        
        # Formatar para exibição
        relatorio_estado_display = formatar_dataframe_moeda(relatorio_estado, ['ValorTotal'])
        st.dataframe(relatorio_estado_display, use_container_width=True)
        
        st.download_button(
            "📥 Exportar Positivação por Estado",
            to_excel(relatorio_estado),
            "positivacao_estado.xlsx",
            "application/vnd.ms-excel"
        )

    with tab3_fat:
        st.subheader("🧾 Relatório de Pedidos Faturados")

        # ── Filtros locais ──
        _fc1, _fc2, _fc3 = st.columns(3)
        with _fc1:
            _fat_vendedores = ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist())
            _fat_vend = st.selectbox("👤 Vendedor", _fat_vendedores, key="fat_vend")
        with _fc2:
            _fat_regioes = ['Todos'] + sorted(df['Estado'].dropna().unique().tolist())
            _fat_reg = st.selectbox("🗺️ Estado/Região", _fat_regioes, key="fat_reg")
        with _fc3:
            _fat_col1, _fat_col2 = st.columns(2)
            with _fat_col1:
                _fat_di = st.date_input("📅 De", value=None, key="fat_di", format="DD/MM/YYYY")
            with _fat_col2:
                _fat_df = st.date_input("📅 Até", value=None, key="fat_df", format="DD/MM/YYYY")

        # ── Aplicar filtros ──
        _df_fat = df[df['TipoMov'] == 'NF Venda'].copy()
        if _fat_vend != 'Todos':
            _df_fat = _df_fat[_df_fat['Vendedor'] == _fat_vend]
        if _fat_reg != 'Todos':
            _df_fat = _df_fat[_df_fat['Estado'] == _fat_reg]
        if _fat_di:
            _df_fat = _df_fat[_df_fat['DataEmissao'] >= pd.to_datetime(_fat_di)]
        if _fat_df:
            _df_fat = _df_fat[_df_fat['DataEmissao'] <= pd.to_datetime(_fat_df)]

        if len(_df_fat) == 0:
            st.info("Nenhum registro encontrado com os filtros selecionados.")
        else:
            # ── KPIs ──
            _fk1, _fk2, _fk3 = st.columns(3)
            with _fk1:
                st.metric("Total Faturado", f"R$ {_df_fat['TotalProduto'].sum():,.2f}")
            with _fk2:
                st.metric("Notas Fiscais", obter_notas_unicas(_df_fat)['Numero_NF'].nunique())
            with _fk3:
                st.metric("Clientes", _df_fat['CPF_CNPJ'].nunique())

            # ── Colunas para exibição/exportação ──
            _cols_base = ['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado', 'Vendedor',
                          'DataEmissao', 'Numero_NF', 'TipoMov',
                          'CodigoProduto', 'NomeProduto', 'Quantidade', 'PrecoUnit',
                          'TotalProduto', 'Valor_Real']
            _cols_disp = [c for c in _cols_base if c in _df_fat.columns]
            _df_fat_disp = _df_fat[_cols_disp].copy()
            _df_fat_disp['DataEmissao'] = _df_fat_disp['DataEmissao'].dt.strftime('%d/%m/%Y')
            st.dataframe(_df_fat_disp, use_container_width=True, height=400)

            # ── Exportar Excel com duas abas como tabela ──
            def _gerar_excel_faturado(df_src, nome_vend):
                import io
                _cols = [c for c in ['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado', 'Vendedor',
                                     'DataEmissao', 'Numero_NF', 'TipoMov',
                                     'CodigoProduto', 'NomeProduto', 'Quantidade', 'PrecoUnit',
                                     'TotalProduto', 'Valor_Real'] if c in df_src.columns]
                _df_exp = df_src[_cols].copy()
                _df_exp['DataEmissao'] = _df_exp['DataEmissao'].dt.strftime('%d/%m/%Y')

                # Aba FATURAMENTO TOTAL: dedup por Numero_NF + soma TotalProduto
                _cols_ocultas = ['CodigoProduto', 'NomeProduto', 'Quantidade', 'PrecoUnit', 'TotalProduto', 'Valor_Real']
                _cols_fat_total = [c for c in _cols if c not in _cols_ocultas]
                _df_fat_total = (
                    df_src.drop_duplicates(subset=['Numero_NF'], keep='first')
                    [_cols_fat_total + ['TotalProduto']]
                    .copy()
                )
                _df_fat_total['DataEmissao'] = _df_fat_total['DataEmissao'].dt.strftime('%d/%m/%Y')
                _soma = _df_fat_total['TotalProduto'].sum()
                _linha_total = {c: '' for c in _df_fat_total.columns}
                _linha_total['TotalProduto'] = _soma
                _linha_total['RazaoSocial'] = 'TOTAL'
                _df_fat_total = pd.concat([_df_fat_total, pd.DataFrame([_linha_total])], ignore_index=True)

                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    wb = writer.book

                    # ── Aba 1: PRODUTOS POR CLIENTE ──
                    _df_exp.to_excel(writer, index=False, sheet_name='PRODUTOS POR CLIENTE')
                    ws1 = writer.sheets['PRODUTOS POR CLIENTE']
                    if len(_df_exp) > 0:
                        ws1.add_table(0, 0, len(_df_exp), len(_df_exp.columns) - 1, {
                            'name': 'TblProdutosCliente',
                            'style': 'Table Style Medium 2',
                            'columns': [{'header': c} for c in _df_exp.columns]
                        })

                    # ── Aba 2: FATURAMENTO TOTAL ──
                    _df_fat_total.to_excel(writer, index=False, sheet_name='FATURAMENTO TOTAL')
                    ws2 = writer.sheets['FATURAMENTO TOTAL']
                    _nrows_ft = len(_df_fat_total) - 1  # última linha é total, fora da tabela
                    if _nrows_ft > 0:
                        ws2.add_table(0, 0, _nrows_ft, len(_df_fat_total.columns) - 1, {
                            'name': 'TblFaturamentoTotal',
                            'style': 'Table Style Medium 2',
                            'columns': [{'header': c} for c in _df_fat_total.columns]
                        })
                    # Linha de soma logo após a tabela
                    _fmt_bold = wb.add_format({'bold': True, 'num_format': '#,##0.00'})
                    _soma_row = _nrows_ft + 1
                    _soma_col = list(_df_fat_total.columns).index('TotalProduto')
                    ws2.write(_soma_row, _soma_col, _soma, _fmt_bold)

                return output.getvalue()

            _nome_arquivo_fat = f"{_fat_vend.upper().replace(' ', '_')}.xlsx" if _fat_vend != 'Todos' else "FATURADO_GERAL.xlsx"

            st.download_button(
                "📥 Exportar Pedidos Faturados",
                _gerar_excel_faturado(_df_fat, _fat_vend),
                _nome_arquivo_fat,
                "application/vnd.ms-excel",
                key="download_fat"
            )

# ====================== INADIMPLÊNCIA ======================
elif menu == "Inadimplência":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Relatório de Inadimplência</h2>', unsafe_allow_html=True)
    
    # Verificar se a planilha de inadimplência existe
    if not planilhas_disponiveis['inadimplencia']:
        st.error("❌ Planilha de inadimplência não encontrada")
        st.info("💡 Para usar este módulo, adicione no GitHub um arquivo com 'LANCAMENTO A RECEBER' no nome")
        st.info(f"📂 Local: {GITHUB_REPO}/{GITHUB_FOLDER}/")
        st.info("📋 Colunas necessárias: Funcionário, Razão Social, N_Doc, Dt.Vencimento, Vr.Líquido, Conta/Caixa, UF")
        st.stop()
    
    # Carregar dados de inadimplência
    with st.spinner("📥 Carregando dados de inadimplência..."):
        df_inadimplencia = carregar_planilha_github(planilhas_disponiveis['inadimplencia']['url'])
    
    if df_inadimplencia is not None and len(df_inadimplencia) > 0:
        df_inadimplencia = processar_inadimplencia(df_inadimplencia)
        
        
        # ========== FILTROS ==========
        st.subheader("🔍 Filtros")
        col_f1, col_f2 = st.columns(2)
        
        with col_f1:
            vendedores_inad = ['Todos'] + sorted(df_inadimplencia['Vendedor'].dropna().unique().tolist())
            vendedor_inad_filtro = st.selectbox("Vendedor", vendedores_inad, key="vend_inad")
        
        with col_f2:
            estados_inad = ['Todos'] + sorted(df_inadimplencia['Estado'].dropna().unique().tolist())
            estado_inad_filtro = st.selectbox("Estado", estados_inad, key="est_inad")
        
        data_inicial_inad, data_final_inad = renderizar_filtros_locais("inad", "📅 Ajustar Período de Vencimento")
        
        # Aplicar filtros
        df_inad_filtrado = df_inadimplencia.copy()
        
        if vendedor_inad_filtro != 'Todos':
            df_inad_filtrado = df_inad_filtrado[df_inad_filtrado['Vendedor'] == vendedor_inad_filtro]
        if estado_inad_filtro != 'Todos':
            df_inad_filtrado = df_inad_filtrado[df_inad_filtrado['Estado'] == estado_inad_filtro]
        if data_inicial_inad:
            df_inad_filtrado = df_inad_filtrado[df_inad_filtrado['DataVencimento'] >= pd.to_datetime(data_inicial_inad)]
        if data_final_inad:
            df_inad_filtrado = df_inad_filtrado[df_inad_filtrado['DataVencimento'] <= pd.to_datetime(data_final_inad)]
        
        st.markdown("---")
        
        # ========== CARDS DE RESUMO ==========
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            total_inadimplencia = df_inad_filtrado['ValorLiquido'].sum()
            st.metric("Total em Aberto", f"R$ {total_inadimplencia:,.2f}")
        
        with col2:
            qtd_titulos = len(df_inad_filtrado)
            st.metric("Qtd. Títulos", f"{qtd_titulos:,}")
        
        with col3:
            clientes_inadimplentes = df_inad_filtrado['Cliente'].nunique()
            st.metric("Clientes Inadimplentes", f"{clientes_inadimplentes:,}")
        
        with col4:
            atraso_medio = df_inad_filtrado['DiasAtraso'].mean()
            st.metric("Atraso Médio", f"{atraso_medio:.0f} dias")
        
        st.markdown("---")
        
        # ========== GRÁFICOS — 4 por linha ==========
        col5, col6, col7, col8 = st.columns(4)

        with col5:
            st.markdown("**📊 Por Faixa de Atraso**")
            ordem_faixas = ['A Vencer', '1-30 dias', '31-60 dias', '61-90 dias', 'Acima de 90 dias']
            inad_por_faixa = df_inad_filtrado.groupby('FaixaAtraso')['ValorLiquido'].sum().reset_index()
            inad_por_faixa['FaixaAtraso'] = pd.Categorical(inad_por_faixa['FaixaAtraso'], categories=ordem_faixas, ordered=True)
            inad_por_faixa = inad_por_faixa.sort_values('FaixaAtraso')
            fig_faixa = px.bar(inad_por_faixa, x='FaixaAtraso', y='ValorLiquido',
                labels={'FaixaAtraso': '', 'ValorLiquido': 'R$'},
                color_discrete_sequence=['#1F4788'])
            fig_faixa = aplicar_layout_grafico(fig_faixa, height=280)
            st.plotly_chart(fig_faixa, use_container_width=True)

        with col6:
            st.markdown("**🏦 Por Banco**")
            inad_por_banco = df_inad_filtrado.groupby('Banco')['ValorLiquido'].sum().reset_index()
            inad_por_banco = inad_por_banco.sort_values('ValorLiquido', ascending=False).head(10)
            fig_banco = px.bar(inad_por_banco, x='ValorLiquido', y='Banco', orientation='h',
                labels={'Banco': '', 'ValorLiquido': 'R$'},
                color_discrete_sequence=['#1F4788'])
            fig_banco = aplicar_layout_grafico(fig_banco, height=280)
            st.plotly_chart(fig_banco, use_container_width=True)

        with col7:
            st.markdown("**👤 Top Vendedores**")
            if 'NumeroDoc' not in df_inad_filtrado.columns:
                possiveis_nomes = [col for col in df_inad_filtrado.columns if 'DOC' in col.upper() or 'NUMERO' in col.upper()]
                df_inad_filtrado['NumeroDoc'] = df_inad_filtrado[possiveis_nomes[0]] if possiveis_nomes else 1
            inad_por_vendedor = df_inad_filtrado.groupby('Vendedor').agg(
                {'ValorLiquido': 'sum', 'NumeroDoc': 'count'}).reset_index()
            inad_por_vendedor.columns = ['Vendedor', 'Valor', 'QtdTitulos']
            inad_por_vendedor = inad_por_vendedor.sort_values('Valor', ascending=False).head(10)
            fig_vend_inad = px.bar(inad_por_vendedor, x='Valor', y='Vendedor', orientation='h',
                labels={'Vendedor': '', 'Valor': 'R$'},
                color_discrete_sequence=['#4A7BC8'])
            fig_vend_inad = aplicar_layout_grafico(fig_vend_inad, height=280)
            st.plotly_chart(fig_vend_inad, use_container_width=True)

        with col8:
            st.markdown("**🗺️ Top Estados**")
            inad_por_estado = df_inad_filtrado.groupby('Estado')['ValorLiquido'].sum().reset_index()
            inad_por_estado = inad_por_estado.sort_values('ValorLiquido', ascending=False).head(10)
            fig_est_inad = px.bar(inad_por_estado, x='ValorLiquido', y='Estado', orientation='h',
                labels={'Estado': '', 'ValorLiquido': 'R$'},
                color_discrete_sequence=['#163561'])
            fig_est_inad = aplicar_layout_grafico(fig_est_inad, height=280)
            st.plotly_chart(fig_est_inad, use_container_width=True)
        
        st.markdown("---")
        
        # ========== TABELA DETALHADA ==========
        st.subheader("📋 Detalhamento dos Títulos")
        
        # Preparar dados para exibição
        df_detalhado = df_inad_filtrado[[
            'Vendedor', 'Cliente', 'NumeroDoc', 'DataVencimento', 
            'ValorLiquido', 'DiasAtraso', 'FaixaAtraso', 'Banco', 'Estado'
        ]].copy()
        
        # Formatar data
        df_detalhado['DataVencimento'] = df_detalhado['DataVencimento'].dt.strftime('%d/%m/%Y')
        
        # Formatar valores para exibição
        df_detalhado_display = df_detalhado.copy()
        df_detalhado_display['ValorLiquido'] = df_detalhado_display['ValorLiquido'].apply(
            lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00"
        )
        
        # Renomear colunas
        df_detalhado_display = df_detalhado_display.rename(columns={
            'Vendedor': 'Vendedor',
            'Cliente': 'Cliente',
            'NumeroDoc': 'Nº Documento',
            'DataVencimento': 'Vencimento',
            'ValorLiquido': 'Valor em Aberto',
            'DiasAtraso': 'Dias Atraso',
            'FaixaAtraso': 'Faixa',
            'Banco': 'Banco',
            'Estado': 'UF'
        })
        
        # Ordenar por dias de atraso (maior para menor)
        df_detalhado_display = df_detalhado_display.sort_values('Dias Atraso', ascending=False)
        
        st.dataframe(df_detalhado_display, use_container_width=True, height=400)
        
        # Botão de download
        _nome_inad = (
            f"{vendedor_inad_filtro.upper().replace(' ', '_')}_INADIMPLENCIA.xlsx"
            if vendedor_inad_filtro != 'Todos'
            else "RELATORIO_INADIMPLENCIA.xlsx"
        )
        st.download_button(
            "📥 Exportar Relatório Completo",
            to_excel(df_detalhado),
            _nome_inad,
            "application/vnd.ms-excel"
        )

# ====================== CLIENTES SEM COMPRA ======================
elif menu == "Clientes sem Compra":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Clientes sem Compra no Período</h2>', unsafe_allow_html=True)

    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        vendedor_churn_filtro = st.selectbox(
            "Filtrar por Vendedor",
            ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist()),
            key="vend_churn"
        )
    with col_f2:
        estado_churn_filtro = st.selectbox(
            "Filtrar por Estado",
            ['Todos'] + sorted(df['Estado'].dropna().unique().tolist()),
            key="est_churn"
        )
    with col_f3:
        ordem = st.selectbox(
            "Ordenar por",
            ["Valor Histórico (Maior)", "Valor Histórico (Menor)", "Nome (A-Z)", "Última Compra (Mais Recente)"],
            key="ordem_churn"
        )
    with col_f4:
        busca_cliente_churn = st.text_input(
            "🔍 Buscar Cliente",
            placeholder="Digite o nome...",
            key="busca_churn"
        )

    # ── Lógica de período ─────────────────────────────────────────────────
    if data_inicial and data_final:
        _label_periodo = f"{data_inicial.strftime('%d/%m/%Y')} a {data_final.strftime('%d/%m/%Y')}"
        _df_periodo = df[
            (df['DataEmissao'] >= pd.to_datetime(data_inicial)) &
            (df['DataEmissao'] <= pd.to_datetime(data_final))
        ]
    elif data_inicial:
        _label_periodo = f"A partir de {data_inicial.strftime('%d/%m/%Y')}"
        _df_periodo = df[df['DataEmissao'] >= pd.to_datetime(data_inicial)]
    elif data_final:
        _label_periodo = f"Até {data_final.strftime('%d/%m/%Y')}"
        _df_periodo = df[df['DataEmissao'] <= pd.to_datetime(data_final)]
    else:
        _mes_now = pd.Timestamp.now().month
        _ano_now = pd.Timestamp.now().year
        _label_periodo = f"Mês vigente ({_mes_now:02d}/{_ano_now})"
        _df_periodo = df[
            (df['DataEmissao'].dt.month == _mes_now) &
            (df['DataEmissao'].dt.year == _ano_now)
        ]

    st.info(f"📅 Período analisado: **{_label_periodo}** — clientes da base que não realizaram compras neste período")

    clientes_com_venda = set(_df_periodo[_df_periodo['TipoMov'] == 'NF Venda']['CPF_CNPJ'].unique())
    
    # Pegamos a última compra de cada cliente (DataEmissao)
    todos_clientes = df.sort_values('DataEmissao').groupby('CPF_CNPJ').last().reset_index()
    
    valor_historico = df[df['TipoMov'] == 'NF Venda'].groupby('CPF_CNPJ')['TotalProduto'].sum().reset_index()
    valor_historico.columns = ['CPF_CNPJ', 'ValorHistorico']
    
    todos_clientes = pd.merge(todos_clientes, valor_historico, on='CPF_CNPJ', how='left')
    todos_clientes['ValorHistorico'] = todos_clientes['ValorHistorico'].fillna(0)
    
    clientes_sem_compra = todos_clientes[~todos_clientes['CPF_CNPJ'].isin(clientes_com_venda)]
    
    # ADICIONADO: DataEmissao incluída na seleção de colunas
    clientes_sem_compra = clientes_sem_compra[['RazaoSocial', 'CPF_CNPJ', 'Vendedor', 'Cidade', 'Estado', 'ValorHistorico', 'DataEmissao']]
    
    if vendedor_churn_filtro != 'Todos':
        clientes_sem_compra = clientes_sem_compra[clientes_sem_compra['Vendedor'] == vendedor_churn_filtro]
    if estado_churn_filtro != 'Todos':
        clientes_sem_compra = clientes_sem_compra[clientes_sem_compra['Estado'] == estado_churn_filtro]
    
    if busca_cliente_churn and len(busca_cliente_churn) >= 2:
        clientes_sem_compra = clientes_sem_compra[
            clientes_sem_compra['RazaoSocial'].str.contains(busca_cliente_churn, case=False, na=False)
        ]
    
    # Lógica de Ordenação
    if ordem == "Valor Histórico (Maior)":
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=False)
    elif ordem == "Valor Histórico (Menor)":
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=True)
    elif ordem == "Nome (A-Z)":
        clientes_sem_compra = clientes_sem_compra.sort_values('RazaoSocial')
    elif ordem == "Última Compra (Mais Recente)":
        clientes_sem_compra = clientes_sem_compra.sort_values('DataEmissao', ascending=False)
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total de Clientes sem Compra", len(clientes_sem_compra))
    with col2:
        st.metric("Valor Potencial Perdido", f"R$ {clientes_sem_compra['ValorHistorico'].sum():,.2f}")
    with col3:
        ticket_medio_churn = clientes_sem_compra['ValorHistorico'].mean() if len(clientes_sem_compra) > 0 else 0
        st.metric("Ticket Médio Histórico", f"R$ {ticket_medio_churn:,.2f}")
    
    if len(clientes_sem_compra) > 0:
        top_churn = clientes_sem_compra.head(15)
        fig_churn = px.bar(
            top_churn,
            x='ValorHistorico',
            y='RazaoSocial',
            orientation='h',
            labels={'RazaoSocial': 'Cliente', 'ValorHistorico': 'Valor Histórico (R$)'},
            color='ValorHistorico',
            color_discrete_sequence=['#1F4788'],
            title='Top 15 Clientes sem Compra por Valor Histórico'
        )
        fig_churn = aplicar_layout_grafico(fig_churn)
        st.plotly_chart(fig_churn, use_container_width=True)
    
    # Preparar visualização
    clientes_sem_compra_display = formatar_dataframe_moeda(clientes_sem_compra, ['ValorHistorico'])
    
    # ADICIONADO: Formatação da data para o padrão brasileiro
    clientes_sem_compra_display['DataEmissao'] = pd.to_datetime(clientes_sem_compra_display['DataEmissao']).dt.strftime('%d/%m/%Y')
    
    # Renomear colunas para exibição
    clientes_sem_compra_display = clientes_sem_compra_display.rename(columns={
        'RazaoSocial': 'Razão Social',
        'CPF_CNPJ': 'CPF/CNPJ',
        'Vendedor': 'Vendedor',
        'Cidade': 'Cidade',
        'Estado': 'Estado',
        'ValorHistorico': 'Valor Histórico',
        'DataEmissao': 'Última Compra'  # Coluna renomeada para a tabela
    })
    
    st.dataframe(clientes_sem_compra_display, use_container_width=True, height=400)
    
    st.download_button(
        "📥 Exportar Clientes sem Compra",
        to_excel(clientes_sem_compra),
        "clientes_sem_compra.xlsx",
        "application/vnd.ms-excel"
    )

# ====================== HISTÓRICO ======================
elif menu == "Histórico":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Histórico de Vendas</h2>', unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4 = st.tabs(["👤 Por Cliente", "🧑‍💼 Por Vendedor", "📝 Pedidos", "📦 Por Produto"])
    
    # ========== ABA: POR CLIENTE ==========
    with tab1:
        st.subheader("Histórico de Vendas por Cliente")
        
        # Buscar cliente por CPF/CNPJ ou Nome
        col_busca1, col_busca2 = st.columns(2)
        
        with col_busca1:
            busca_tipo = st.radio("Buscar por:", ["Nome", "CPF/CNPJ"], horizontal=True, key="busca_tipo_cliente")
        
        with col_busca2:
            if busca_tipo == "Nome":
                busca_texto = st.text_input("Digite o nome do cliente", placeholder="Ex: Nome da Empresa", key="busca_nome_cliente")
            else:
                busca_texto = st.text_input("Digite o CPF/CNPJ", placeholder="Ex: 12345678901234", key="busca_cpf_cliente")
        
        cliente_selecionado = None
        cpf_cnpj = None
        
        if busca_texto and len(busca_texto) >= 3:
            if busca_tipo == "Nome":
                clientes_filtrados = df[df['RazaoSocial'].str.contains(busca_texto, case=False, na=False)][['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado']].drop_duplicates()
            else:
                clientes_filtrados = df[df['CPF_CNPJ'].str.contains(busca_texto, case=False, na=False)][['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado']].drop_duplicates()
            
            if len(clientes_filtrados) > 0:
                clientes_filtrados['Display'] = clientes_filtrados['RazaoSocial'] + " - " + clientes_filtrados['CPF_CNPJ'] + " (" + clientes_filtrados['Cidade'] + "/" + clientes_filtrados['Estado'] + ")"
                
                cliente_selecionado = st.selectbox(
                    f"📋 Clientes encontrados ({len(clientes_filtrados)}):",
                    options=clientes_filtrados['Display'].tolist(),
                    key="cliente_hist"
                )
                
                if cliente_selecionado:
                    cpf_cnpj = cliente_selecionado.split(' - ')[1].split(' (')[0]
            else:
                st.warning("❌ Nenhum cliente encontrado com esse critério")
        
        if cpf_cnpj:
            historico = df[df['CPF_CNPJ'] == cpf_cnpj].sort_values('DataEmissao', ascending=False)
            
            if len(historico) > 0:
                cliente_info = historico.iloc[0]
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Cliente", cliente_info['RazaoSocial'])
                with col2:
                    st.metric("CPF/CNPJ", cliente_info['CPF_CNPJ'])
                with col3:
                    st.metric("Cidade/Estado", f"{cliente_info['Cidade']}/{cliente_info['Estado']}")
                with col4:
                    st.metric("Total de Registros", len(historico))
                
                st.markdown("---")
                
                vendas_cliente = historico[historico['TipoMov'] == 'NF Venda']
                devolucoes_cliente = historico[historico['TipoMov'] == 'NF Dev.Venda']
                
                col5, col6, col7, col8 = st.columns(4)
                with col5:
                    st.metric("Total Vendas", f"R$ {vendas_cliente['TotalProduto'].sum():,.2f}")
                with col6:
                    st.metric("Total Devoluções", f"R$ {devolucoes_cliente['TotalProduto'].sum():,.2f}")
                with col7:
                    st.metric("Qtd Notas Vendas", len(vendas_cliente['Numero_NF'].unique()))
                with col8:
                    st.metric("Qtd Notas Devoluções", len(devolucoes_cliente['Numero_NF'].unique()))
                
                vendas_tempo_cliente = vendas_cliente.groupby('MesAno')['TotalProduto'].sum().reset_index()
                vendas_tempo_cliente = vendas_tempo_cliente.sort_values('MesAno')
                
                if len(vendas_tempo_cliente) > 0:
                    fig_hist = px.line(
                        vendas_tempo_cliente,
                        x='MesAno',
                        y='TotalProduto',
                        labels={'MesAno': 'Período', 'TotalProduto': 'Valor (R$)'},
                        title='Evolução de Compras'
                    )
                    fig_hist.update_traces(line_color='#28A745', line_width=3, mode='lines+markers', marker=dict(size=6, color='#28A745'))
                    fig_hist = aplicar_layout_grafico(fig_hist)
                    st.plotly_chart(fig_hist, use_container_width=True)
                
                st.markdown("---")
                
                st.subheader("📋 Detalhamento de Produtos")
                
                # Verificar se PrazoHistorico e Comissao existem no dataframe
                colunas_display = ['DataEmissao', 'TipoMov', 'Numero_NF', 'CodigoProduto', 'NomeProduto', 'Quantidade', 'PrecoUnit', 'TotalProduto']
                if 'PrazoHistorico' in historico.columns:
                    colunas_display.append('PrazoHistorico')
                if 'Comissao' in historico.columns:
                    colunas_display.append('Comissao')
                
                historico_display = historico[colunas_display].copy()
                historico_display['DataEmissao'] = historico_display['DataEmissao'].dt.strftime('%d/%m/%Y')
                
                # Formatar valores monetários
                historico_display['PrecoUnit'] = historico_display['PrecoUnit'].apply(lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00")
                historico_display['TotalProduto'] = historico_display['TotalProduto'].apply(lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00")
                
                # Renomear colunas
                colunas_rename = {
                    'DataEmissao': 'Data',
                    'TipoMov': 'Tipo',
                    'Numero_NF': 'Nota Fiscal',
                    'CodigoProduto': 'Código',
                    'NomeProduto': 'Produto',
                    'Quantidade': 'Qtd',
                    'PrecoUnit': 'Preço Unit.',
                    'TotalProduto': 'Total'
                }
                if 'PrazoHistorico' in historico_display.columns:
                    colunas_rename['PrazoHistorico'] = 'Prazo (dias)'
                if 'Comissao' in historico_display.columns:
                    colunas_rename['Comissao'] = 'Comissão%'
                
                historico_display = historico_display.rename(columns=colunas_rename)
                
                st.dataframe(historico_display, use_container_width=True, height=400)
                
                st.download_button(
                    "📥 Exportar Histórico Excel",
                    to_excel(historico),
                    f"historico_{cpf_cnpj}.xlsx",
                    "application/vnd.ms-excel",
                    key="dl_hist_excel"
                )

                st.markdown("---")
                st.markdown("""
                <div style="background:#F0F4FF;border:1px solid #C5D5F0;border-radius:10px;
                            padding:14px 18px;margin-bottom:8px;">
                    <div style="font-size:0.88rem;font-weight:700;color:#4A7BC8;margin-bottom:4px;">
                        Gerar Proposta Comercial PDF
                    </div>
                    <div style="font-size:0.78rem;color:#6C757D;">
                        Exporta os produtos do histórico do cliente em formato de proposta
                        comercial com cabeçalho Medtextil, dados do cliente e tabela de itens.
                    </div>
                </div>
                """, unsafe_allow_html=True)

                _col_pdf1, _col_pdf2 = st.columns([2, 1])
                with _col_pdf1:
                    _vendas_resumo = {
                        'Total de Vendas':   f"R$ {vendas_cliente['TotalProduto'].sum():,.2f}",
                        'Total Devoluções':  f"R$ {devolucoes_cliente['TotalProduto'].sum():,.2f}",
                        'Notas de Venda':    str(len(vendas_cliente['Numero_NF'].unique())),
                        'Clientes (CNPJ)':   cpf_cnpj,
                    }
                    _cliente_info_dict = {
                        'RazaoSocial': cliente_info.get('RazaoSocial',''),
                        'CPF_CNPJ':    cpf_cnpj,
                        'Cidade':      cliente_info.get('Cidade',''),
                        'Estado':      cliente_info.get('Estado',''),
                        'Vendedor':    cliente_info.get('Vendedor',''),
                    }
                    if st.button("Gerar Proposta PDF", key="btn_gerar_proposta",
                                 use_container_width=True, type="primary"):
                        with st.spinner("Gerando proposta..."):
                            try:
                                _pdf_bytes = gerar_proposta_pdf_historico(
                                    _cliente_info_dict, historico, _vendas_resumo
                                )
                                st.session_state['proposta_pdf_bytes'] = _pdf_bytes
                                st.session_state['proposta_pdf_nome'] = (
                                    f"Proposta_{razao_curta}_{hoje_str}.pdf"
                                    if 'razao_curta' in dir() else
                                    f"Proposta_{cpf_cnpj}.pdf"
                                )
                                st.success("Proposta gerada! Clique em Download abaixo.")
                            except Exception as _e:
                                st.error(f"Erro ao gerar proposta: {_e}")

                with _col_pdf2:
                    if st.session_state.get('proposta_pdf_bytes'):
                        import datetime as _dt
                        _nome_pdf = (
                            f"Proposta_Medtextil_{cliente_info.get('RazaoSocial','cliente')[:20].replace(' ','_')}"
                            f"_{_dt.date.today().strftime('%Y%m%d')}.pdf"
                        )
                        st.download_button(
                            "Download PDF",
                            data=st.session_state['proposta_pdf_bytes'],
                            file_name=_nome_pdf,
                            mime="application/pdf",
                            key="dl_proposta_pdf",
                            use_container_width=True
                        )
            else:
                st.warning("Nenhum registro encontrado para este cliente")
        else:
            st.info("👆 Digite pelo menos 3 caracteres para buscar um cliente")
    
    # ========== ABA: POR VENDEDOR ==========
    with tab2:
        st.subheader("Histórico de Vendas por Vendedor")
        
        # Filtros
        col_f1, = st.columns(1)
        
        with col_f1:
            vendedores_hist = ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist())
            vendedor_hist_filtro = st.selectbox("Vendedor", vendedores_hist, key="vend_hist")
        
        data_inicial_hist, data_final_hist = renderizar_filtros_locais("hist_vend", "📅 Ajustar Período")
        
        # Aplicar filtros
        df_hist_vendedor = df[df['TipoMov'] == 'NF Venda'].copy()
        
        if vendedor_hist_filtro != 'Todos':
            df_hist_vendedor = df_hist_vendedor[df_hist_vendedor['Vendedor'] == vendedor_hist_filtro]
        if data_inicial_hist:
            df_hist_vendedor = df_hist_vendedor[df_hist_vendedor['DataEmissao'] >= pd.to_datetime(data_inicial_hist)]
        if data_final_hist:
            df_hist_vendedor = df_hist_vendedor[df_hist_vendedor['DataEmissao'] <= pd.to_datetime(data_final_hist)]
        
        if len(df_hist_vendedor) > 0:
            # Obter notas únicas e agrupar por NF para somar o valor total
            notas_vendedor = obter_notas_unicas(df_hist_vendedor)
            
            # Preparar dados para exibição
            colunas_vendedor = ['DataEmissao', 'RazaoSocial', 'Numero_NF', 'TotalProduto', 'Vendedor']
            if 'PrazoHistorico' in notas_vendedor.columns:
                colunas_vendedor.append('PrazoHistorico')
            
            historico_vendedor = notas_vendedor[colunas_vendedor].copy()
            historico_vendedor = historico_vendedor.sort_values('DataEmissao', ascending=False)
            
            # Métricas resumidas
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total de Vendas", f"R$ {historico_vendedor['TotalProduto'].sum():,.2f}")
            with col2:
                st.metric("Quantidade de Notas", len(historico_vendedor))
            with col3:
                st.metric("Clientes Atendidos", historico_vendedor['RazaoSocial'].nunique())
            with col4:
                ticket_medio_vend = historico_vendedor['TotalProduto'].mean() if len(historico_vendedor) > 0 else 0
                st.metric("Ticket Médio", f"R$ {ticket_medio_vend:,.2f}")
            
            st.markdown("---")
            
            # Formatar para exibição
            historico_vendedor_display = historico_vendedor.copy()
            historico_vendedor_display['DataEmissao'] = historico_vendedor_display['DataEmissao'].dt.strftime('%d/%m/%Y')
            historico_vendedor_display['TotalProduto'] = historico_vendedor_display['TotalProduto'].apply(
                lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00"
            )
            
            # Renomear colunas
            colunas_rename_vendedor = {
                'DataEmissao': 'Data',
                'RazaoSocial': 'Cliente',
                'Numero_NF': 'Nota Fiscal',
                'TotalProduto': 'Valor Total',
                'Vendedor': 'Vendedor'
            }
            if 'PrazoHistorico' in historico_vendedor_display.columns:
                colunas_rename_vendedor['PrazoHistorico'] = 'Prazo (dias)'
            
            historico_vendedor_display = historico_vendedor_display.rename(columns=colunas_rename_vendedor)
            
            st.dataframe(historico_vendedor_display, use_container_width=True, height=400)
            
            # Botão de download
            st.download_button(
                "📥 Exportar Histórico de Vendas",
                to_excel(historico_vendedor),
                f"historico_vendedor_{vendedor_hist_filtro if vendedor_hist_filtro != 'Todos' else 'todos'}.xlsx",
                "application/vnd.ms-excel",
                key="download_hist_vend"
            )
        else:
            st.info("Nenhuma venda encontrada com os filtros selecionados")

    
    # ========== ABA: PEDIDOS ==========
    with tab3:
        st.subheader("📝 Gerar Pedido/Proposta")
        
        # Inicializar session_state para os itens do pedido
        if 'itens_pedido' not in st.session_state:
            st.session_state.itens_pedido = []
        
        # Carregar dados de produtos se disponível
        df_produtos_pedido = None
        if planilhas_disponiveis.get('produtos_agrupados'):
            with st.spinner("📥 Carregando catálogo de produtos..."):
                df_produtos_pedido = carregar_planilha_github(planilhas_disponiveis['produtos_agrupados']['url'])
                if df_produtos_pedido is not None:
                    df_produtos_pedido.columns = df_produtos_pedido.columns.str.upper()
        
        # SEÇÃO 1: DADOS DO CLIENTE
        st.markdown("### 👤 Informações do Cliente")
        
        col_cli1, col_cli2 = st.columns(2)
        
        with col_cli1:
            # Buscar cliente
            clientes_lista = sorted(df['RazaoSocial'].dropna().unique().tolist())
            cliente_selecionado = st.selectbox("Selecione o Cliente", [''] + clientes_lista, key="cliente_pedido")
        
        # Buscar dados do cliente
        dados_cliente = {}
        if cliente_selecionado:
            df_cliente = df[df['RazaoSocial'] == cliente_selecionado].iloc[0]
            dados_cliente = {
                'razao_social': df_cliente.get('RazaoSocial', ''),
                'cpf_cnpj': df_cliente.get('CPF_CNPJ', ''),
                'cidade': df_cliente.get('Cidade', ''),
                'estado': df_cliente.get('Estado', ''),
                'vendedor': df_cliente.get('Vendedor', '')
            }
        
        with col_cli2:
            representante = st.text_input("Representante", value=dados_cliente.get('vendedor', ''), key="representante_pedido")
        
        col_cli3, col_cli4, col_cli5 = st.columns(3)
        
        with col_cli3:
            nome_fantasia = st.text_input("Nome Fantasia", value=dados_cliente.get('razao_social', ''), key="fantasia_pedido")
        
        with col_cli4:
            cnpj_pedido = st.text_input("CNPJ", value=dados_cliente.get('cpf_cnpj', ''), key="cnpj_pedido")
        
        with col_cli5:
            insc_estadual = st.text_input("Inscrição Estadual", key="ie_pedido")
        
        col_cli6, col_cli7 = st.columns(2)
        
        with col_cli6:
            telefone_pedido = st.text_input("Telefone", key="tel_pedido")
        
        with col_cli7:
            email_pedido = st.text_input("Email NF-e", key="email_pedido")
        
        endereco_pedido = st.text_input("Endereço", value=f"{dados_cliente.get('cidade', '')}/{dados_cliente.get('estado', '')}" if dados_cliente else "", key="end_pedido")
        
        obs_cliente = st.text_area("Observação (Cliente)", key="obs_cli_pedido", height=80)
        
        st.markdown("---")
        
        # SEÇÃO 2: DADOS DO PEDIDO
        st.markdown("### 📋 Informações do Pedido")
        
        col_ped1, col_ped2, col_ped3, col_ped4 = st.columns(4)
        
        with col_ped1:
            num_pedido = st.text_input("Nº do Pedido", key="num_pedido")
        
        with col_ped2:
            tabela_preco = st.text_input("Tabela de Preço", key="tab_preco")
        
        with col_ped3:
            tipo_frete = st.selectbox("Tipo de Frete", ["CIF", "FOB"], key="tipo_frete")
        
        with col_ped4:
            data_venda = st.date_input("Data da Venda", value=pd.Timestamp.now(), key="data_venda")
        
        condicoes_pagto = st.text_input("Condições de Pagamento", key="cond_pagto")
        
        st.markdown("---")
        
        # SEÇÃO 3: ADICIONAR PRODUTOS
        st.markdown("### 🛒 Adicionar Produtos ao Pedido")
        
        col_prod1, col_prod2, col_prod3, col_prod4 = st.columns([2, 1, 1, 1])
        
        with col_prod1:
            # Buscar por código ou descrição
            tipo_busca_prod = st.radio("Buscar por:", ["Código", "Descrição"], horizontal=True, key="tipo_busca_prod")
            
            if tipo_busca_prod == "Código":
                if df_produtos_pedido is not None:
                    codigos = [''] + sorted(df_produtos_pedido['ID_COD'].dropna().astype(str).unique().tolist())
                    codigo_selecionado = st.selectbox("Código do Produto", codigos, key="cod_prod_pedido")
                else:
                    codigo_selecionado = st.text_input("Código do Produto", key="cod_prod_pedido")
            else:
                busca_desc = st.text_input("Descrição do Produto", key="desc_prod_pedido")
                codigo_selecionado = None
        
        # Buscar informações do produto
        produto_info = {}
        if df_produtos_pedido is not None and codigo_selecionado:
            prod = df_produtos_pedido[df_produtos_pedido['ID_COD'].astype(str) == str(codigo_selecionado)]
            if len(prod) > 0:
                prod = prod.iloc[0]
                # Montar descrição completa
                descricao_completa = f"{prod.get('GRUPO', '')} {prod.get('DESCRIÇÃO', '') or prod.get('DESCRICAO', '')} {prod.get('LINHA', '') or prod.get('LINHAS', '')}".strip()
                
                produto_info = {
                    'codigo': str(prod.get('ID_COD', '')),
                    'descricao': descricao_completa,
                    'peso': prod.get('GRAMATURA', ''),
                    'cx_embarque': prod.get('CX_EMB', ''),
                    'preco_ref': prod.get('PRECO', 0)
                }
                
                # Buscar último preço que o cliente comprou
                if cliente_selecionado:
                    hist_cliente = df[(df['RazaoSocial'] == cliente_selecionado) & 
                                     (df['CodigoProduto'].astype(str) == str(codigo_selecionado))]
                    if len(hist_cliente) > 0:
                        hist_cliente = hist_cliente.sort_values('DataEmissao', ascending=False)
                        produto_info['preco_sugerido'] = hist_cliente.iloc[0]['PrecoUnit']
                        produto_info['preco_historico'] = hist_cliente.iloc[0]['PrecoUnit']  # Para mostrar na tabela
                    else:
                        produto_info['preco_sugerido'] = prod.get('PRECO', 0)
                        produto_info['preco_historico'] = prod.get('PRECO', 0)
                else:
                    produto_info['preco_sugerido'] = prod.get('PRECO', 0)
                    produto_info['preco_historico'] = prod.get('PRECO', 0)
        
        with col_prod2:
            qtde_item = st.number_input("Quantidade", min_value=0, value=0, key="qtde_item_pedido")
        
        with col_prod3:
            valor_item = st.number_input("Valor Unit.", min_value=0.0, value=float(produto_info.get('preco_sugerido', 0)), format="%.2f", key="valor_item_pedido")
        
        with col_prod4:
            st.write("")
            st.write("")
            if st.button("➕ Adicionar Item", use_container_width=True, key="add_item_pedido"):
                if produto_info and qtde_item > 0:
                    # Calcular comissão
                    comissao = calcular_comissao(valor_item, produto_info.get('preco_ref', 0))
                    
                    item = {
                        'codigo': produto_info['codigo'],
                        'descricao': produto_info['descricao'],
                        'peso': produto_info.get('peso', ''),
                        'cx_embarque': produto_info.get('cx_embarque', ''),
                        'quantidade': qtde_item,
                        'valor_unit': valor_item,
                        'preco_historico': produto_info.get('preco_historico', 0),
                        'total': qtde_item * valor_item,
                        'comissao': comissao
                    }
                    st.session_state.itens_pedido.append(item)
                    st.success(f"✅ Item adicionado: {produto_info['descricao']}")
                    st.rerun()
        
        # PREVIEW EM TEMPO REAL - Mostrar antes de adicionar
        if produto_info and qtde_item > 0 and valor_item > 0:
            st.markdown("---")
            st.markdown("### 👁️ Preview do Item")
            
            # Calcular valores preview
            total_preview = qtde_item * valor_item
            comissao_preview = calcular_comissao(valor_item, produto_info.get('preco_ref', 0))
            preco_hist_preview = produto_info.get('preco_historico', 0)
            
            # Mostrar em tabela estilizada
            preview_data = {
                'Código': [produto_info['codigo']],
                'Produto': [produto_info['descricao'][:50]],
                'Peso': [produto_info.get('peso', '')],
                'Cx Embarque': [produto_info.get('cx_embarque', '')],
                'Qtde': [f"{qtde_item:,.0f}"],
                'Preço Histórico': [f"R$ {preco_hist_preview:,.2f}"],
                'Valor Unit.': [f"R$ {valor_item:,.2f}"],
                'Total': [f"R$ {total_preview:,.2f}"],
                'Comissão%': [comissao_preview]
            }
            
            df_preview = pd.DataFrame(preview_data)
            st.dataframe(df_preview, use_container_width=True, hide_index=True)
            
            # Comparação com preço histórico
            if preco_hist_preview > 0:
                variacao = ((valor_item - preco_hist_preview) / preco_hist_preview) * 100
                if variacao > 0:
                    st.info(f"📈 Valor {variacao:.1f}% **acima** do histórico (R$ {preco_hist_preview:,.2f})")
                elif variacao < 0:
                    st.warning(f"📉 Valor {abs(variacao):.1f}% **abaixo** do histórico (R$ {preco_hist_preview:,.2f})")
                else:
                    st.success(f"✅ Valor **igual** ao histórico (R$ {preco_hist_preview:,.2f})")
            
            st.markdown("---")
        
        # Mostrar produtos adicionados
        if st.session_state.itens_pedido:
            st.markdown("---")
            st.markdown("### 📦 Itens do Pedido")
            
            # Criar DataFrame dos itens
            df_itens = pd.DataFrame(st.session_state.itens_pedido)
            
            # Formatar para exibição
            df_itens_display = df_itens.copy()
            df_itens_display['preco_historico'] = df_itens_display['preco_historico'].apply(lambda x: f"R$ {x:,.2f}")
            df_itens_display['valor_unit'] = df_itens_display['valor_unit'].apply(lambda x: f"R$ {x:,.2f}")
            df_itens_display['total'] = df_itens_display['total'].apply(lambda x: f"R$ {x:,.2f}")
            
            df_itens_display = df_itens_display.rename(columns={
                'codigo': 'COD.',
                'descricao': 'PRODUTO',
                'peso': 'PESO',
                'cx_embarque': 'CAIXA EMBARQUE',
                'quantidade': 'QTDE',
                'preco_historico': 'PREÇO HISTÓRICO',
                'valor_unit': 'VALOR',
                'total': 'TOTAL',
                'comissao': 'COMISSÃO%'
            })
            
            st.dataframe(df_itens_display, use_container_width=True, height=300)
            
            # Métricas do pedido
            col_met1, col_met2, col_met3 = st.columns(3)
            
            with col_met1:
                total_itens = df_itens['quantidade'].sum()
                st.metric("Qtde Total de Itens", f"{total_itens:,.0f}")
            
            with col_met2:
                st.metric("Frete", tipo_frete)
            
            with col_met3:
                total_pedido = df_itens['total'].sum()
                st.metric("Total Final", f"R$ {total_pedido:,.2f}")
            
            # Observação final
            obs_pedido = st.text_area("Observação (Pedido)", key="obs_pedido", height=100)
            
            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                if st.button("🗑️ Limpar Pedido", use_container_width=True, key="limpar_pedido"):
                    st.session_state.itens_pedido = []
                    st.rerun()
            
            with col_btn2:
                if st.button("📄 Gerar PDF do Pedido", use_container_width=True, key="gerar_pdf_pedido", type="primary"):
                    try:
                        # Preparar dados para o PDF
                        dados_cliente_pdf = {
                            'representante': representante,
                            'razao_social': cliente_selecionado,
                            'nome_fantasia': nome_fantasia,
                            'cnpj': cnpj_pedido,
                            'ie': insc_estadual,
                            'telefone': telefone_pedido,
                            'email': email_pedido,
                            'endereco': endereco_pedido,
                            'obs_cliente': obs_cliente
                        }
                        
                        dados_pedido_pdf = {
                            'numero': num_pedido,
                            'tabela_preco': tabela_preco,
                            'tipo_frete': tipo_frete,
                            'data_venda': data_venda.strftime('%d/%m/%Y'),
                            'condicoes_pagto': condicoes_pagto
                        }
                        
                        # Gerar PDF
                        pdf_bytes = gerar_pdf_pedido(dados_cliente_pdf, dados_pedido_pdf, st.session_state.itens_pedido, obs_pedido)
                        
                        # Botão de download
                        st.download_button(
                            label="📥 Baixar PDF do Pedido",
                            data=pdf_bytes,
                            file_name=f"Pedido_{num_pedido or 'SN'}_{cliente_selecionado.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            key="download_pdf_pedido"
                        )
                        
                        st.success("✅ PDF gerado com sucesso!")
                        
                    except Exception as e:
                        st.error(f"❌ Erro ao gerar PDF: {str(e)}")
                        st.info("💡 Certifique-se de que a biblioteca ReportLab está instalada")
        else:
            st.info("ℹ️ Nenhum item adicionado ao pedido ainda. Use o formulário acima para adicionar produtos.")

    with tab4:
        st.markdown("#### 📦 Vendas por Produto")

        # ── Filtros ───────────────────────────────────────────────────────
        _tp1, _tp2 = st.columns(2)
        with _tp1:
            _tp_di = st.date_input("📅 Data Inicial", value=None,
                                   key="tp_data_ini", format="DD/MM/YYYY")
        with _tp2:
            _tp_df = st.date_input("📅 Data Final", value=None,
                                   key="tp_data_fim", format="DD/MM/YYYY")

        _tp3, _tp4 = st.columns(2)
        with _tp3:
            # Multiselect de produtos
            _prods_disponiveis = sorted(
                df[df['NomeProduto'].notna()]['NomeProduto'].unique().tolist()
            )
            _tp_prods = st.multiselect(
                "🔍 Filtrar por Produto(s)",
                options=_prods_disponiveis,
                placeholder="Todos os produtos",
                key="tp_produtos"
            )
        with _tp4:
            _codigos_disponiveis = sorted(
                df[df['CodigoProduto'].notna()]['CodigoProduto'].astype(str).unique().tolist()
            )
            _tp_cods = st.multiselect(
                "🔍 Filtrar por Código(s)",
                options=_codigos_disponiveis,
                placeholder="Todos os códigos",
                key="tp_codigos"
            )

        # ── Aplicar filtros ───────────────────────────────────────────────
        _df_tp = df[df['TipoMov'] == 'NF Venda'].copy()

        if _tp_di:
            _df_tp = _df_tp[_df_tp['DataEmissao'] >= pd.to_datetime(_tp_di)]
        if _tp_df:
            _df_tp = _df_tp[_df_tp['DataEmissao'] <= pd.to_datetime(_tp_df)]
        if _tp_prods:
            _df_tp = _df_tp[_df_tp['NomeProduto'].isin(_tp_prods)]
        if _tp_cods:
            _df_tp = _df_tp[_df_tp['CodigoProduto'].astype(str).isin(_tp_cods)]

        if len(_df_tp) == 0:
            st.info("Nenhuma venda encontrada com os filtros aplicados.")
        else:
            # ── KPIs ──────────────────────────────────────────────────────
            _m1, _m2, _m3, _m4 = st.columns(4)
            with _m1:
                st.metric("Total Faturado", f"R$ {_df_tp['TotalProduto'].sum():,.2f}")
            with _m2:
                st.metric("Qtd Vendida", f"{_df_tp['Quantidade'].sum():,.0f}")
            with _m3:
                _pm = (_df_tp['TotalProduto'].sum() / _df_tp['Quantidade'].sum()
                       if _df_tp['Quantidade'].sum() > 0 else 0)
                st.metric("Preço Médio", f"R$ {_pm:,.2f}")
            with _m4:
                st.metric("Produtos Únicos", f"{_df_tp['CodigoProduto'].nunique():,}")

            st.markdown("---")

            # ── Gráficos ──────────────────────────────────────────────────
            _g1, _g2 = st.columns(2)

            with _g1:
                st.markdown("**Top 10 por Faturamento**")
                _top_fat = (
                    _df_tp.groupby('NomeProduto')['TotalProduto']
                    .sum().reset_index()
                    .sort_values('TotalProduto', ascending=False).head(10)
                )
                _fig1 = px.bar(_top_fat, x='TotalProduto', y='NomeProduto',
                               orientation='h',
                               labels={'NomeProduto': '', 'TotalProduto': 'R$'},
                               color_discrete_sequence=['#1F4788'])
                _fig1 = aplicar_layout_grafico(_fig1, height=320)
                st.plotly_chart(_fig1, use_container_width=True)

            with _g2:
                st.markdown("**Top 10 por Quantidade**")
                _top_qtd = (
                    _df_tp.groupby('NomeProduto')['Quantidade']
                    .sum().reset_index()
                    .sort_values('Quantidade', ascending=False).head(10)
                )
                _fig2 = px.bar(_top_qtd, x='Quantidade', y='NomeProduto',
                               orientation='h',
                               labels={'NomeProduto': '', 'Quantidade': 'Unidades'},
                               color_discrete_sequence=['#2E86AB'])
                _fig2 = aplicar_layout_grafico(_fig2, height=320)
                st.plotly_chart(_fig2, use_container_width=True)

            # ── Evolução temporal ─────────────────────────────────────────
            st.markdown("**Evolução Mensal**")
            _prods_evo = ['Todos'] + sorted(_df_tp['NomeProduto'].dropna().unique().tolist())
            _prod_sel  = st.selectbox("Produto para evolução:", _prods_evo, key="tp_evo_prod")
            _df_evo    = _df_tp if _prod_sel == 'Todos' else _df_tp[_df_tp['NomeProduto'] == _prod_sel]

            if 'MesAno' in _df_evo.columns and len(_df_evo) > 0:
                _evo = _df_evo.groupby('MesAno').agg(
                    Faturamento=('TotalProduto', 'sum'),
                    Quantidade=('Quantidade', 'sum')
                ).reset_index().sort_values('MesAno')

                _e1, _e2 = st.columns(2)
                with _e1:
                    _fe = px.line(_evo, x='MesAno', y='Faturamento',
                                  labels={'MesAno': 'Período', 'Faturamento': 'R$'},
                                  color_discrete_sequence=['#1F4788'])
                    _fe.update_traces(line_width=2, mode='lines+markers', marker=dict(size=5))
                    _fe = aplicar_layout_grafico(_fe, height=240)
                    st.plotly_chart(_fe, use_container_width=True)
                with _e2:
                    _qe = px.line(_evo, x='MesAno', y='Quantidade',
                                  labels={'MesAno': 'Período', 'Quantidade': 'Unidades'},
                                  color_discrete_sequence=['#2E86AB'])
                    _qe.update_traces(line_width=2, mode='lines+markers', marker=dict(size=5))
                    _qe = aplicar_layout_grafico(_qe, height=240)
                    st.plotly_chart(_qe, use_container_width=True)

            # ── Tabela detalhada ──────────────────────────────────────────
            st.markdown("**Detalhamento por Produto**")
            _df_tab = (
                _df_tp.groupby(['CodigoProduto', 'NomeProduto'])
                .agg(QtdVendida=('Quantidade', 'sum'),
                     PrecoMedio=('PrecoUnit', 'mean'),
                     TotalFaturado=('TotalProduto', 'sum'))
                .reset_index()
                .sort_values('TotalFaturado', ascending=False)
            )
            _df_tab['PrecoMedio']    = _df_tab['PrecoMedio'].apply(lambda x: f"R$ {x:,.2f}")
            _df_tab['TotalFaturado'] = _df_tab['TotalFaturado'].apply(lambda x: f"R$ {x:,.2f}")
            _df_tab['QtdVendida']    = _df_tab['QtdVendida'].apply(lambda x: f"{x:,.0f}")
            _df_tab = _df_tab.rename(columns={
                'CodigoProduto': 'Código',
                'NomeProduto':   'Produto',
                'QtdVendida':    'Qtd Vendida',
                'PrecoMedio':    'Preço Médio',
                'TotalFaturado': 'Total (R$)',
            })
            st.dataframe(_df_tab, use_container_width=True, height=350)

            st.download_button(
                "📥 Exportar Vendas por Produto",
                to_excel(_df_tp),
                "historico_por_produto.xlsx",
                "application/vnd.ms-excel",
                key="dl_hist_produto"
            )


# ====================== PREÇO MÉDIO ======================
elif menu == "Preço Médio":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Análise de Preço Médio por Produto</h2>', unsafe_allow_html=True)
    
    # Verificar se as planilhas necessárias existem
    if not planilhas_disponiveis['vendas_produto']:
        st.error("❌ Planilha 'Vendas por produto - GERAL.xlsx' não encontrada")
        st.info("💡 Adicione no GitHub um arquivo com 'VENDAS POR PRODUTO' e 'GERAL' no nome")
        st.info(f"📂 Local: {GITHUB_REPO}/{GITHUB_FOLDER}/")
        st.info("📋 Colunas necessárias: CODPRODUTO, TOTQTD, PRECOUNITMEDIO, TOTLIQUIDO")
        st.stop()
    
    if not planilhas_disponiveis['produtos_agrupados']:
        st.error("❌ Planilha 'Produtos_Agrupados_Completos_conciliados.xlsx' não encontrada")
        st.info("💡 Adicione no GitHub um arquivo com 'PRODUTOS_AGRUPADOS_COMPLETOS_CONCILIADOS' no nome")
        st.info(f"📂 Local: {GITHUB_REPO}/{GITHUB_FOLDER}/")
        st.info("📋 Colunas necessárias: ID_COD, Grupo, Descrição, Linha, Gramatura")
        st.stop()
    
    # Carregar planilhas
    with st.spinner("📥 Carregando dados de vendas por produto..."):
        df_vendas_produto = carregar_planilha_github(planilhas_disponiveis['vendas_produto']['url'])
    
    with st.spinner("📥 Carregando dados de produtos..."):
        df_produtos = carregar_planilha_github(planilhas_disponiveis['produtos_agrupados']['url'])
    
    if df_vendas_produto is None or df_produtos is None:
        st.error("❌ Erro ao carregar uma ou mais planilhas")
        st.stop()
    
    # Padronizar nomes das colunas (case-insensitive)
    df_vendas_produto.columns = df_vendas_produto.columns.str.upper()
    df_produtos.columns = df_produtos.columns.str.upper()
    
    # IMPORTANTE: Se a planilha de vendas já tiver NOMEPRODUTO, remover para usar apenas o da planilha de produtos
    if 'NOMEPRODUTO' in df_vendas_produto.columns:
        df_vendas_produto = df_vendas_produto.drop(columns=['NOMEPRODUTO'])
    
    # Verificar se as colunas necessárias existem
    colunas_vendas_necessarias = ['CODPRODUTO', 'TOTQTD', 'PRECOUNITMEDIO', 'TOTLIQUIDO']
    colunas_produtos_necessarias = ['ID_COD', 'GRUPO', 'DESCRIÇÃO', 'LINHA', 'GRAMATURA']
    
    # Verificar colunas alternativas
    if 'DESCRIÇÃO' not in df_produtos.columns and 'DESCRICAO' in df_produtos.columns:
        df_produtos = df_produtos.rename(columns={'DESCRICAO': 'DESCRIÇÃO'})
    
    if 'LINHA' not in df_produtos.columns and 'LINHAS' in df_produtos.columns:
        df_produtos = df_produtos.rename(columns={'LINHAS': 'LINHA'})
    
    if 'GRUPO' not in df_produtos.columns and 'GRUPOS' in df_produtos.columns:
        df_produtos = df_produtos.rename(columns={'GRUPOS': 'GRUPO'})
    
    faltando_vendas = [col for col in colunas_vendas_necessarias if col not in df_vendas_produto.columns]
    faltando_produtos = [col for col in colunas_produtos_necessarias if col not in df_produtos.columns]
    
    if faltando_vendas:
        st.error(f"❌ Colunas faltando na planilha de vendas: {', '.join(faltando_vendas)}")
        st.info(f"📋 Colunas encontradas: {', '.join(df_vendas_produto.columns.tolist())}")
        st.stop()
    
    if faltando_produtos:
        st.error(f"❌ Colunas faltando na planilha de produtos: {', '.join(faltando_produtos)}")
        st.info(f"📋 Colunas encontradas: {', '.join(df_produtos.columns.tolist())}")
        st.stop()
    
    # Criar a descrição concatenada na planilha de produtos
    # Limpar espaços e garantir que não fique vazio
    df_produtos['GRUPO_LIMPO'] = df_produtos['GRUPO'].fillna('').astype(str).str.strip()
    df_produtos['DESCRIÇÃO_LIMPO'] = df_produtos['DESCRIÇÃO'].fillna('').astype(str).str.strip()
    df_produtos['LINHA_LIMPO'] = df_produtos['LINHA'].fillna('').astype(str).str.strip()
    
    df_produtos['NOMEPRODUTO'] = (
        df_produtos['GRUPO_LIMPO'] + ' ' +
        df_produtos['DESCRIÇÃO_LIMPO'] + ' ' +
        df_produtos['LINHA_LIMPO']
    ).str.strip()
    
    # Se NOMEPRODUTO ficar vazio, usar o ID_COD como descrição
    df_produtos.loc[df_produtos['NOMEPRODUTO'] == '', 'NOMEPRODUTO'] = 'Produto ' + df_produtos['ID_COD'].astype(str)
    
    # Remover colunas temporárias
    df_produtos = df_produtos.drop(columns=['GRUPO_LIMPO', 'DESCRIÇÃO_LIMPO', 'LINHA_LIMPO'])
    
    # Renomear ID_COD para CODPRODUTO para facilitar o merge
    df_produtos = df_produtos.rename(columns={'ID_COD': 'CODPRODUTO'})
    
    # Adicionar coluna DATA com o mês/ano atual (já que não existe na planilha)
    data_atual = pd.Timestamp.now()
    df_vendas_produto['DATA'] = data_atual
    df_vendas_produto['Mes'] = data_atual.month
    df_vendas_produto['Ano'] = data_atual.year
    df_vendas_produto['MesAno'] = data_atual.strftime('%Y-%m')
    
    # Fazer o merge (PROCV) entre as planilhas
    # Garantir que as colunas necessárias existam antes do merge
    colunas_merge = ['CODPRODUTO', 'NOMEPRODUTO', 'GRAMATURA']
    for col in colunas_merge:
        if col not in df_produtos.columns:
            if col == 'NOMEPRODUTO':
                df_produtos['NOMEPRODUTO'] = 'Produto não catalogado'
            elif col == 'GRAMATURA':
                df_produtos['GRAMATURA'] = 0
    
    df_preco_medio = pd.merge(
        df_vendas_produto,
        df_produtos[colunas_merge],
        on='CODPRODUTO',
        how='left'
    )
    
    # Preencher valores nulos após o merge
    produtos_nao_catalogados = 0
    if 'NOMEPRODUTO' in df_preco_medio.columns:
        produtos_nao_catalogados = df_preco_medio['NOMEPRODUTO'].isna().sum()
        df_preco_medio['NOMEPRODUTO'] = df_preco_medio['NOMEPRODUTO'].fillna('Produto não catalogado - Código: ' + df_preco_medio['CODPRODUTO'].astype(str))
    else:
        df_preco_medio['NOMEPRODUTO'] = 'Produto não catalogado - Código: ' + df_preco_medio['CODPRODUTO'].astype(str)
        produtos_nao_catalogados = len(df_preco_medio)
    
    if 'GRAMATURA' in df_preco_medio.columns:
        df_preco_medio['GRAMATURA'] = df_preco_medio['GRAMATURA'].fillna(0)
    else:
        df_preco_medio['GRAMATURA'] = 0
    
    
    
    st.markdown("---")
    
    # ========== FILTROS ==========
    st.subheader("🔍 Filtros")
    
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        anos_preco = ['Todos'] + sorted(df_preco_medio['Ano'].dropna().unique().tolist(), reverse=True)
        ano_preco_filtro = st.selectbox("Ano", anos_preco, key="ano_preco")
    
    with col_f2:
        meses_preco = ['Todos'] + list(range(1, 13))
        mes_preco_filtro = st.selectbox("Mês", meses_preco, key="mes_preco")
    
    with col_f3:
        busca_cod = st.text_input("🔍 Buscar Código", placeholder="Digite o código...", key="busca_cod_preco")
    
    with col_f4:
        busca_nome = st.text_input("🔍 Buscar Produto", placeholder="Digite o nome...", key="busca_nome_preco")
    
    # Aplicar filtros
    df_preco_filtrado = df_preco_medio.copy()
    
    if ano_preco_filtro != 'Todos':
        df_preco_filtrado = df_preco_filtrado[df_preco_filtrado['Ano'] == ano_preco_filtro]
    if mes_preco_filtro != 'Todos':
        df_preco_filtrado = df_preco_filtrado[df_preco_filtrado['Mes'] == mes_preco_filtro]
    if busca_cod and len(busca_cod) >= 2:
        df_preco_filtrado = df_preco_filtrado[
            df_preco_filtrado['CODPRODUTO'].astype(str).str.contains(busca_cod, case=False, na=False)
        ]
    if busca_nome and len(busca_nome) >= 2:
        df_preco_filtrado = df_preco_filtrado[
            df_preco_filtrado['NOMEPRODUTO'].str.contains(busca_nome, case=False, na=False)
        ]
    
    st.markdown("---")
    
    # ========== MÉTRICAS GERAIS ==========
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_vendido = df_preco_filtrado['TOTLIQUIDO'].sum()
        st.metric("Total Vendido", f"R$ {total_vendido:,.2f}")
    
    with col2:
        qtd_total = df_preco_filtrado['TOTQTD'].sum()
        st.metric("Qtd. Total Vendida", f"{qtd_total:,.0f}")
    
    with col3:
        # CORREÇÃO: Média ponderada = Total Vendido / Quantidade Total
        if df_preco_filtrado['TOTQTD'].sum() > 0:
            preco_medio_geral = df_preco_filtrado['TOTLIQUIDO'].sum() / df_preco_filtrado['TOTQTD'].sum()
        else:
            preco_medio_geral = 0
        st.metric("Preço Médio Geral", f"R$ {preco_medio_geral:,.2f}")
    
    with col4:
        produtos_unicos = df_preco_filtrado['CODPRODUTO'].nunique()
        st.metric("Produtos Únicos", f"{produtos_unicos:,}")
    
    st.markdown("---")
    
    # ========== GRÁFICOS ==========
    col5, col6 = st.columns(2)
    
    with col5:
        st.subheader("📊 Top 10 Produtos por Faturamento")
        
        top_faturamento = df_preco_filtrado.groupby('NOMEPRODUTO')['TOTLIQUIDO'].sum().reset_index()
        top_faturamento = top_faturamento.sort_values('TOTLIQUIDO', ascending=False).head(10)
        
        fig_fat = px.bar(
            top_faturamento,
            x='TOTLIQUIDO',
            y='NOMEPRODUTO',
            orientation='h',
            labels={'NOMEPRODUTO': 'Produto', 'TOTLIQUIDO': 'Faturamento (R$)'},
            color='TOTLIQUIDO',
            color_discrete_sequence=['#1F4788']
        )
        fig_fat = aplicar_layout_grafico(fig_fat)
        st.plotly_chart(fig_fat, use_container_width=True)
    
    with col6:
        st.subheader("📈 Top 10 Produtos por Quantidade")
        
        top_quantidade = df_preco_filtrado.groupby('NOMEPRODUTO')['TOTQTD'].sum().reset_index()
        top_quantidade = top_quantidade.sort_values('TOTQTD', ascending=False).head(10)
        
        fig_qtd = px.bar(
            top_quantidade,
            x='TOTQTD',
            y='NOMEPRODUTO',
            orientation='h',
            labels={'NOMEPRODUTO': 'Produto', 'TOTQTD': 'Quantidade Vendida'},
            color='TOTQTD',
            color_discrete_sequence=['#2E86AB']
        )
        fig_qtd = aplicar_layout_grafico(fig_qtd)
        st.plotly_chart(fig_qtd, use_container_width=True)
    
    st.markdown("---")
    
    # ========== ANÁLISE DE PREÇO MÉDIO POR PERÍODO ==========
    st.subheader("📅 Evolução de Preço Médio")
    
    # Permitir seleção de produto específico
    produtos_lista = ['Todos'] + sorted(df_preco_filtrado['NOMEPRODUTO'].unique().tolist())
    produto_selecionado = st.selectbox(
        "Selecione um produto para ver evolução de preço:",
        produtos_lista,
        key="produto_evolucao"
    )
    
    if produto_selecionado != 'Todos':
        df_evolucao = df_preco_filtrado[df_preco_filtrado['NOMEPRODUTO'] == produto_selecionado]
    else:
        df_evolucao = df_preco_filtrado.copy()
    
    if len(df_evolucao) > 0:
        evolucao_preco = df_evolucao.groupby('MesAno').agg({
            'PRECOUNITMEDIO': 'mean',
            'TOTQTD': 'sum',
            'TOTLIQUIDO': 'sum'
        }).reset_index()
        evolucao_preco = evolucao_preco.sort_values('MesAno')
        
        fig_evolucao = px.line(
            evolucao_preco,
            x='MesAno',
            y='PRECOUNITMEDIO',
            labels={'MesAno': 'Período', 'PRECOUNITMEDIO': 'Preço Médio (R$)'},
            title=f'Evolução do Preço Médio - {produto_selecionado}'
        )
        fig_evolucao.update_traces(line_color='#FF6B6B', line_width=3)
        fig_evolucao = aplicar_layout_grafico(fig_evolucao)
        st.plotly_chart(fig_evolucao, use_container_width=True)
    
    st.markdown("---")
    
    # ========== TABELA DETALHADA ==========
    st.subheader("📋 Detalhamento de Preços")
    
    # Preparar dados para exibição
    df_detalhado = df_preco_filtrado[[
        'CODPRODUTO', 'NOMEPRODUTO', 'GRAMATURA', 'TOTQTD', 
        'PRECOUNITMEDIO', 'TOTLIQUIDO', 'DATA'
    ]].copy()
    
    # Ordenar por data (mais recente primeiro)
    df_detalhado = df_detalhado.sort_values('DATA', ascending=False)
    
    # Formatar para exibição
    df_detalhado_display = df_detalhado.copy()
    
    # Formatar data no padrão brasileiro com mês atual
    mes_atual = data_atual.month
    ano_atual = data_atual.year
    df_detalhado_display['DATA'] = f"{mes_atual:02d}/{ano_atual}"
    
    df_detalhado_display['PRECOUNITMEDIO'] = df_detalhado_display['PRECOUNITMEDIO'].apply(
        lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00"
    )
    df_detalhado_display['TOTLIQUIDO'] = df_detalhado_display['TOTLIQUIDO'].apply(
        lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00"
    )
    
    # Renomear colunas
    df_detalhado_display = df_detalhado_display.rename(columns={
        'CODPRODUTO': 'Código',
        'NOMEPRODUTO': 'Nome do Produto',
        'GRAMATURA': 'Gramatura',
        'TOTQTD': 'Qtd Vendida',
        'PRECOUNITMEDIO': 'Preço Médio Unit.',
        'TOTLIQUIDO': 'Total Líquido',
        'DATA': 'Período (Mês/Ano)'
    })
    
    st.dataframe(df_detalhado_display, use_container_width=True, height=400)
    
    # Botão de download
    st.download_button(
        "📥 Exportar Relatório de Preços",
        to_excel(df_detalhado),
        "relatorio_preco_medio.xlsx",
        "application/vnd.ms-excel",
        key="download_preco_medio"
    )

# ====================== PEDIDOS PENDENTES ======================
elif menu == "Pedidos Pendentes":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Pedidos Pendentes de Faturamento</h2>', unsafe_allow_html=True)
    
    # Verificar se a planilha existe
    if not planilhas_disponiveis.get('pedidos_pendentes'):
        st.error("❌ Planilha 'PEDIDOSPENDENTES.xlsx' não encontrada")
        st.info("💡 Adicione no GitHub um arquivo com 'PEDIDOSPENDENTES' no nome")
        st.info(f"📂 Local: {GITHUB_REPO}/{GITHUB_FOLDER}/")
        st.stop()
    
    # Carregar planilha
    with st.spinner("📥 Carregando pedidos pendentes..."):
        try:
            import zipfile
            import xml.etree.ElementTree as ET
            from io import BytesIO
            
            # Baixar arquivo
            response = requests.get(planilhas_disponiveis['pedidos_pendentes']['url'])
            excel_file = BytesIO(response.content)
            
            # Extrair shared strings
            with zipfile.ZipFile(excel_file) as z:
                with z.open('xl/sharedStrings.xml') as f:
                    strings_tree = ET.parse(f)
                    ns_str = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    shared_strings = [si.text if si.text else '' for si in strings_tree.findall('.//ss:t', ns_str)]
                
                # Extrair sheet
                with z.open('xl/worksheets/sheet1.xml') as f:
                    sheet_tree = ET.parse(f)
                    ns = {'ss': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
                    
                    # Parsear dados
                    data = []
                    current_client = None
                    current_pedido = None
                    
                    for row in sheet_tree.findall('.//ss:row', ns):
                        row_data = {}
                        for cell in row.findall('.//ss:c', ns):
                            ref = cell.get('r', '')
                            col = ''.join([c for c in ref if c.isalpha()])
                            v_elem = cell.find('.//ss:v', ns)
                            if v_elem is not None and v_elem.text:
                                val_type = cell.get('t', 'n')
                                if val_type == 's':
                                    idx = int(v_elem.text)
                                    value = shared_strings[idx] if idx < len(shared_strings) else v_elem.text
                                else:
                                    value = v_elem.text
                                row_data[col] = value
                        
                        if not row_data:
                            continue
                        
                        # Detectar tipo de linha
                        col_a = row_data.get('A', '')
                        col_b = row_data.get('B', '')
                        
                        # Linha de cliente (apenas coluna A preenchida com nome)
                        if col_a and not col_b and 'N° do pedido' not in col_a and 'Valor Total' not in col_a and col_a != 'Subgrupo:':
                            current_client = col_a
                        
                        # Linha de pedido (tem "N° do pedido:")
                        elif 'N° do pedido' in col_a:
                            current_pedido = col_b
                            
                            # Extrair dados do produto
                            descricao = row_data.get('C', '')
                            if descricao and ' - ' in descricao:
                                # Extrair código do produto (ex: "476 - ATADURA...")
                                codigo_produto = descricao.split(' - ')[0].strip()
                                
                                try:
                                    qtd_contratada = float(row_data.get('D', 0))
                                    valor_unit = float(row_data.get('E', 0))  # Corrigido: E é o valor unitário
                                    qtd_entregue = float(row_data.get('H', 0))  # Corrigido: H é qtd entregue
                                    qtd_pendente = qtd_contratada - qtd_entregue
                                    valor_pendente = qtd_pendente * valor_unit
                                    
                                    # Converter data de emissão (coluna G)
                                    dt_emissao_val = row_data.get('G', '')
                                    if dt_emissao_val:
                                        try:
                                            # Data vem como número (days since 1900)
                                            dt_emissao = pd.Timestamp('1899-12-30') + pd.Timedelta(days=float(dt_emissao_val))
                                        except:
                                            dt_emissao = None
                                    else:
                                        dt_emissao = None
                                    
                                    data.append({
                                        'Cliente': current_client,
                                        'NumeroPedido': current_pedido,
                                        'CodigoProduto': codigo_produto,
                                        'Descricao': descricao,
                                        'QtdContratada': qtd_contratada,
                                        'QtdEntregue': qtd_entregue,
                                        'QtdPendente': qtd_pendente,
                                        'ValorUnit': valor_unit,
                                        'ValorPendente': valor_pendente,
                                        'DataEmissao': dt_emissao,
                                        'Vendedor': row_data.get('J', ''),  # Corrigido: J é o vendedor
                                        'PercEntregue': float(row_data.get('I', 0))  # Corrigido: I é % entregue
                                    })
                                except:
                                    continue
            
            df_pendentes = pd.DataFrame(data)
            
            if len(df_pendentes) == 0:
                st.warning("⚠️ Nenhum pedido pendente encontrado na planilha")
                st.stop()
            
            
        except Exception as e:
            st.error(f"❌ Erro ao processar planilha: {str(e)}")
            st.stop()
    
    st.markdown("---")
    
    # Filtros
    st.subheader("🔍 Filtros")
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        clientes_pend = ['Todos'] + sorted(df_pendentes['Cliente'].dropna().unique().tolist())
        cliente_pend_filtro = st.selectbox("Cliente", clientes_pend, key="cli_pend")
    
    with col_f2:
        vendedores_pend = ['Todos'] + sorted(df_pendentes['Vendedor'].dropna().unique().tolist())
        vendedor_pend_filtro = st.selectbox("Vendedor", vendedores_pend, key="vend_pend")
    
    with col_f3:
        busca_produto = st.text_input("🔍 Buscar Produto", placeholder="Digite código ou descrição", key="busca_prod_pend")
    
    with col_f4:
        apenas_pendentes = st.checkbox("Apenas com pendência", value=True, key="apenas_pend")
    
    # Aplicar filtros
    df_pend_filtrado = df_pendentes.copy()
    
    if cliente_pend_filtro != 'Todos':
        df_pend_filtrado = df_pend_filtrado[df_pend_filtrado['Cliente'] == cliente_pend_filtro]
    if vendedor_pend_filtro != 'Todos':
        df_pend_filtrado = df_pend_filtrado[df_pend_filtrado['Vendedor'] == vendedor_pend_filtro]
    if busca_produto:
        df_pend_filtrado = df_pend_filtrado[
            df_pend_filtrado['CodigoProduto'].str.contains(busca_produto, case=False, na=False) |
            df_pend_filtrado['Descricao'].str.contains(busca_produto, case=False, na=False)
        ]
    if apenas_pendentes:
        df_pend_filtrado = df_pend_filtrado[df_pend_filtrado['QtdPendente'] > 0]
    
    st.markdown("---")
    
    # Métricas
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_pendente = df_pend_filtrado['ValorPendente'].sum()
        st.metric("Valor Total Pendente", f"R$ {total_pendente:,.2f}")
    
    with col2:
        qtd_pendente = df_pend_filtrado['QtdPendente'].sum()
        st.metric("Qtd. Total Pendente", f"{qtd_pendente:,.0f}")
    
    with col3:
        pedidos_unicos = df_pend_filtrado['NumeroPedido'].nunique()
        st.metric("Pedidos Únicos", f"{pedidos_unicos:,}")
    
    with col4:
        perc_medio = df_pend_filtrado['PercEntregue'].mean() if len(df_pend_filtrado) > 0 else 0
        st.metric("% Médio Entregue", f"{perc_medio:.1f}%")
    
    st.markdown("---")
    
    # Gráficos
    col5, col6 = st.columns(2)
    
    with col5:
        st.subheader("🏢 Top 10 Clientes - Valor Pendente")
        top_clientes = df_pend_filtrado.groupby('Cliente')['ValorPendente'].sum().reset_index()
        top_clientes = top_clientes.sort_values('ValorPendente', ascending=False).head(10)
        
        fig_cli = px.bar(
            top_clientes,
            x='ValorPendente',
            y='Cliente',
            orientation='h',
            labels={'Cliente': 'Cliente', 'ValorPendente': 'Valor Pendente (R$)'},
            color='ValorPendente',
            color_discrete_sequence=['#1F4788']
        )
        fig_cli = aplicar_layout_grafico(fig_cli)
        st.plotly_chart(fig_cli, use_container_width=True)
    
    with col6:
        st.subheader("👤 Top 10 Vendedores - Valor Pendente")
        top_vend = df_pend_filtrado.groupby('Vendedor')['ValorPendente'].sum().reset_index()
        top_vend = top_vend.sort_values('ValorPendente', ascending=False).head(10)
        
        fig_vend = px.bar(
            top_vend,
            x='ValorPendente',
            y='Vendedor',
            orientation='h',
            labels={'Vendedor': 'Vendedor', 'ValorPendente': 'Valor Pendente (R$)'},
            color='ValorPendente',
            color_discrete_sequence=['#4A7BC8']
        )
        fig_vend = aplicar_layout_grafico(fig_vend)
        st.plotly_chart(fig_vend, use_container_width=True)
    
    st.markdown("---")
    
    # Tabela detalhada
    st.subheader("📋 Detalhamento de Pedidos Pendentes")
    
    # Preparar dados para exibição
    df_pend_display = df_pend_filtrado[[
        'Cliente', 'NumeroPedido', 'CodigoProduto', 'Descricao', 
        'QtdContratada', 'QtdEntregue', 'QtdPendente',
        'ValorUnit', 'ValorPendente', 'PercEntregue', 'DataEmissao', 'Vendedor'
    ]].copy()
    
    # Formatar valores
    df_pend_display['ValorUnit'] = df_pend_display['ValorUnit'].apply(
        lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00"
    )
    df_pend_display['ValorPendente'] = df_pend_display['ValorPendente'].apply(
        lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00"
    )
    df_pend_display['DataEmissao'] = df_pend_display['DataEmissao'].apply(
        lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ''
    )
    df_pend_display['PercEntregue'] = df_pend_display['PercEntregue'].apply(
        lambda x: f"{x:.1f}%" if pd.notnull(x) else "0%"
    )
    
    # Renomear colunas
    df_pend_display = df_pend_display.rename(columns={
        'Cliente': 'Cliente',
        'NumeroPedido': 'N° Pedido',
        'CodigoProduto': 'Código',
        'Descricao': 'Descrição',
        'QtdContratada': 'Qtd Contratada',
        'QtdEntregue': 'Qtd Entregue',
        'QtdPendente': 'Qtd Pendente',
        'ValorUnit': 'Valor Unit.',
        'ValorPendente': 'Valor Pendente',
        'PercEntregue': '% Entregue',
        'DataEmissao': 'Data Emissão',
        'Vendedor': 'Vendedor'
    })
    
    st.dataframe(df_pend_display, use_container_width=True, height=400)
    
    # Botão de download — nome do arquivo reflete o vendedor filtrado
    _nome_arquivo_pend = (
        f"{vendedor_pend_filtro.upper().replace(' ', '_')}_PENDENTES.xlsx"
        if vendedor_pend_filtro != 'Todos'
        else "PEDIDOS_PENDENTES.xlsx"
    )
    st.download_button(
        "📥 Exportar Pedidos Pendentes (Separado por Tipo)",
        to_excel_pedidos_pendentes(df_pend_filtrado),
        _nome_arquivo_pend,
        "application/vnd.ms-excel",
        key="download_pendentes"
    )

    # ===== PREVISÃO DE FATURAMENTO POR CAPACIDADE PRODUTIVA =====
    st.markdown("---")
    st.markdown("### 🏭 Previsão de Faturamento por Capacidade Produtiva")
    st.caption("Converte unidades pendentes em caixas e estima datas de conclusão com base na capacidade diária de cada produto.")

    import math
    from datetime import date, timedelta

    # Capacidade produtiva por produto (caixas/dia) — match por substring
    CAPACIDADE_PROD = {
        "CAMPO OPERATORIO 45X50": 80,
        "CAMPO OPERATÓRIO 25X28 C2": 23,
        "CAMPO OPERATÓRIO 25X28 C5": 20,
        "GAZE NÃO ESTERIL": 17,
        "GAZE ESTERIL": 80,
        "ATADURA FARMA": 18,
        "ATADURA CONJUGADA": 15,
        "GAZE CIRCULAR": 50,
    }

    def identificar_capacidade(descricao):
        """Match por substring, case-insensitive."""
        if not descricao or str(descricao).lower() == 'nan':
            return None, "SEM CAPACIDADE"
        desc_upper = str(descricao).upper().strip()
        for chave, cap in CAPACIDADE_PROD.items():
            if chave.upper() in desc_upper:
                return cap, chave
        return None, "SEM CAPACIDADE"

    def adicionar_dias_uteis(data_inicio, dias):
        """Avança N dias úteis (seg–sáb), ignorando domingo."""
        atual = data_inicio
        contados = 0
        while contados < dias:
            atual += timedelta(days=1)
            if atual.weekday() != 6:  # 6 = domingo
                contados += 1
        return atual

    # Carregar produtos_agrupados para obter CX_EMB e PRECO via ID_COD
    _df_prod_prev = None
    if planilhas_disponiveis.get('produtos_agrupados'):
        with st.spinner("Carregando dados de produtos para previsão..."):
            _df_prod_prev = carregar_planilha_github(
                planilhas_disponiveis['produtos_agrupados']['url']
            )
        if _df_prod_prev is not None:
            _df_prod_prev.columns = _df_prod_prev.columns.str.upper().str.strip()

    if _df_prod_prev is None:
        st.warning("⚠️ Planilha Produtos_Agrupados não disponível. Previsão desabilitada.")
    else:
        # Normalizar ID_COD
        def _norm_cod(v):
            try:
                return str(int(float(str(v).strip())))
            except:
                return str(v).strip()

        _df_prod_prev['ID_COD_N'] = _df_prod_prev['ID_COD'].apply(_norm_cod)

        # Colunas necessárias
        _cx_col    = next((c for c in _df_prod_prev.columns if 'CX_EMB' in c), None)
        _preco_col = next((c for c in _df_prod_prev.columns if 'PRECO' in c or 'PREÇO' in c), None)
        _desc_col  = next((c for c in _df_prod_prev.columns if 'DESCRI' in c or 'GRUPO' in c), None)

        if not _cx_col or not _preco_col:
            st.warning(f"⚠️ Colunas CX_EMB ou PRECO não encontradas. Colunas disponíveis: {_df_prod_prev.columns.tolist()}")
        else:
            # Preparar base de pendentes com ID_COD normalizado
            _df_base = df_pend_filtrado.copy()
            _df_base['COD_N'] = _df_base['CodigoProduto'].apply(_norm_cod)

            # Merge com produtos
            _df_merge = _df_base.merge(
                _df_prod_prev[['ID_COD_N', _cx_col, _preco_col] + ([_desc_col] if _desc_col else [])],
                left_on='COD_N',
                right_on='ID_COD_N',
                how='left'
            )

            # Converter unidades → caixas (ceil, evitar div/0)
            def _calc_caixas(row):
                try:
                    cx = float(row[_cx_col])
                    if cx <= 0 or pd.isna(cx):
                        return None
                    return math.ceil(float(row['QtdPendente']) / cx)
                except:
                    return None

            _df_merge['CAIXAS_NECESSARIAS'] = _df_merge.apply(_calc_caixas, axis=1)

            # Identificar capacidade produtiva
            _desc_vals = _df_merge[_desc_col].tolist() if _desc_col else [''] * len(_df_merge)
            _caps = [identificar_capacidade(d) for d in _desc_vals]
            _df_merge['CAPACIDADE_DIA']  = [c[0] for c in _caps]
            _df_merge['GRUPO_PROD']      = [c[1] for c in _caps]

            # Calcular dias de produção (ceil, paralelo por produto)
            def _calc_dias(row):
                try:
                    if row['CAPACIDADE_DIA'] is None or row['CAIXAS_NECESSARIAS'] is None:
                        return None
                    return math.ceil(row['CAIXAS_NECESSARIAS'] / row['CAPACIDADE_DIA'])
                except:
                    return None

            _df_merge['DIAS_PRODUCAO'] = _df_merge.apply(_calc_dias, axis=1)

            # Calcular data prevista
            _hoje = date.today()

            def _calc_data(dias):
                if dias is None or pd.isna(dias):
                    return None
                return adicionar_dias_uteis(_hoje, int(dias))

            _df_merge['DATA_PREVISTA'] = _df_merge['DIAS_PRODUCAO'].apply(_calc_data)
            _df_merge['PREVISAO_FORMATADA'] = _df_merge.apply(
                lambda r: f"{r['DATA_PREVISTA'].strftime('%d/%m/%Y')} ({int(r['DIAS_PRODUCAO'])} dias)"
                if r['DATA_PREVISTA'] is not None else "SEM CAPACIDADE",
                axis=1
            )

            # Valor total por linha
            def _calc_valor(row):
                try:
                    return float(row['QtdPendente']) * float(row[_preco_col])
                except:
                    return 0.0

            _df_merge['VALOR_TOTAL'] = _df_merge.apply(_calc_valor, axis=1)

            # ── KPIs ────────────────────────────────────────────────────
            _kp1, _kp2, _kp3 = st.columns(3)
            _com_prev  = _df_merge[_df_merge['DATA_PREVISTA'].notna()]
            _sem_prev  = _df_merge[_df_merge['DATA_PREVISTA'].isna()]
            _dias_pond = (
                (_com_prev['DIAS_PRODUCAO'] * _com_prev['VALOR_TOTAL']).sum()
                / _com_prev['VALOR_TOTAL'].sum()
            ) if _com_prev['VALOR_TOTAL'].sum() > 0 else 0

            with _kp1:
                st.metric("Valor Total Previsto", f"R$ {_com_prev['VALOR_TOTAL'].sum():,.2f}")
            with _kp2:
                st.metric("Média Ponderada de Dias", f"{_dias_pond:.1f} dias")
            with _kp3:
                st.metric("Itens sem Capacidade", f"{len(_sem_prev):,}")

            st.markdown("---")

            # ── Faturamento agrupado por data ────────────────────────────
            st.markdown("**Faturamento Previsto por Data de Conclusão**")
            _fat_data = (
                _com_prev.groupby('PREVISAO_FORMATADA')['VALOR_TOTAL']
                .sum()
                .reset_index()
                .sort_values('PREVISAO_FORMATADA')
            )
            _fat_data.columns = ['Data Prevista', 'Faturamento (R$)']

            _fig_fat = px.bar(
                _fat_data,
                x='Data Prevista',
                y='Faturamento (R$)',
                labels={'Data Prevista': 'Data', 'Faturamento (R$)': 'R$'},
                color_discrete_sequence=['#1F4788']
            )
            _fig_fat = aplicar_layout_grafico(_fig_fat, height=300)
            st.plotly_chart(_fig_fat, use_container_width=True)

            # ── Tabela de previsão ───────────────────────────────────────
            st.markdown("**Relatório Detalhado de Previsão**")
            _cols_show = ['Cliente', 'CodigoProduto', 'Descricao',
                          'QtdPendente', 'CAIXAS_NECESSARIAS',
                          'CAPACIDADE_DIA', 'DIAS_PRODUCAO',
                          'PREVISAO_FORMATADA', 'VALOR_TOTAL']
            _cols_show = [c for c in _cols_show if c in _df_merge.columns]
            _df_show   = _df_merge[_cols_show].copy()
            _df_show['VALOR_TOTAL'] = _df_show['VALOR_TOTAL'].apply(
                lambda x: f"R$ {x:,.2f}" if pd.notnull(x) else "R$ 0,00"
            )
            _df_show = _df_show.rename(columns={
                'CodigoProduto':      'Código',
                'Descricao':          'Produto',
                'QtdPendente':        'Qtd (un)',
                'CAIXAS_NECESSARIAS': 'Caixas',
                'CAPACIDADE_DIA':     'Cap/Dia',
                'DIAS_PRODUCAO':      'Dias',
                'PREVISAO_FORMATADA': 'Previsão',
                'VALOR_TOTAL':        'Faturamento',
            })
            st.dataframe(_df_show, use_container_width=True, height=380)

            # ── Downloads ───────────────────────────────────────────────
            def _gerar_relatorio_previsao(df_merge, df_prod_prev, cx_col, preco_col, desc_col):
                """
                Gera Excel com abas separadas por grupo de produto.
                Colunas: Cliente, Código, Volumes(cx), Descrição, Contratado,
                         Entregue, Pendente, Valor Unit, Valor Pendente,
                         Data Emissão, Dias Pendentes, Vendedor, %Entregue,
                         Previsão(branco), Categoria, Observações(branco)
                """
                import math
                from datetime import date
                from io import BytesIO
                import openpyxl
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                from openpyxl.utils import get_column_letter

                hoje = date.today()

                # Mapeamento de grupos → abas
                # Grupos: palavras-chave em ordem de prioridade
                # Regras por palavras-chave — ORDEM CRÍTICA: mais específico primeiro
                # NAO ESTERIL deve vir ANTES de ESTERIL para evitar falso match
                GRUPOS_REGRAS = [
                    ('Campo',       ['CAMPO OPERATORIO', 'CAMPO OPERATÓRIO', 'CAMPO OP']),
                    ('Tipo Queijo', ['GAZE CIRCULAR', 'QUEIJO', 'TIPO QUEIJO']),
                    ('Pacote',      ['NAO ESTERIL', 'NÃO ESTERIL', 'PACOTE']),  # ANTES de Esteril
                    ('Esteril',     ['GAZE ESTERIL', 'ESTERIL']),               # DEPOIS de Pacote
                ]

                def _norm(s):
                    """Remove acentos e normaliza para comparação segura."""
                    import unicodedata
                    return unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode('ascii').upper()

                def identificar_aba(descricao):
                    if not descricao:
                        return 'Outros'
                    d = _norm(descricao)

                    # Atadura — sempre primeiro
                    if 'ATADURA' in d:
                        return 'Atadura Hospitalar' if 'HOSPITALAR' in d else 'Atadura Farma'

                    # Demais grupos — chaves também normalizadas
                    for aba, chaves in GRUPOS_REGRAS:
                        for chave in chaves:
                            if _norm(chave) in d:
                                return aba

                    return 'Outros'

                def identificar_categoria(descricao, aba):
                    if 'Atadura' not in aba:
                        return ''
                    d = str(descricao).upper()
                    if 'HOSPITALAR' in d:
                        return 'Hospitalar'
                    if 'FARMA' in d:
                        return 'Farma'
                    return 'Farma'  # sem indicação → Farma

                def extrair_descricao_pura(descricao, codigo):
                    """Remove código e nome genérico — retorna só descrição adicional"""
                    if not descricao:
                        return ''
                    d = str(descricao).strip()
                    # Remover código se presente no início
                    cod_str = str(codigo).strip()
                    if d.startswith(cod_str):
                        d = d[len(cod_str):].strip(' -|')
                    return d

                # Calcular CX_EMB lookup
                cx_lookup = {}
                if df_prod_prev is not None and cx_col:
                    for _, row in df_prod_prev.iterrows():
                        try:
                            k = str(int(float(str(row['ID_COD_N'])))).strip()
                            v = float(row[cx_col])
                            if v > 0:
                                cx_lookup[k] = v
                        except:
                            pass

                # Gramatura lookup (mesma lógica do módulo consulta tabela)
                gram_col_g = next((c for c in df_prod_prev.columns if 'GRAMATUR' in c), None) if df_prod_prev is not None else None
                gram_lookup = {}
                if df_prod_prev is not None and gram_col_g:
                    for _, row in df_prod_prev.iterrows():
                        try:
                            k = str(int(float(str(row['ID_COD_N'])))).strip()
                            gv = str(row.get(gram_col_g, '')).strip()
                            if gv and gv.lower() not in ('nan', '0', '0.0', ''):
                                gram_lookup[k] = gv
                        except:
                            pass

                COLUNAS = [
                    'N° Pedido', 'Cliente', 'Código', 'Gramatura', 'Volumes (cx)', 'Descrição',
                    'Contratado', 'Entregue', 'Pendente',
                    'Valor Unitário', 'Valor Pendente',
                    'Data Emissão', 'Dias Pendentes', 'Vendedor',
                    '% Entregue', 'Previsão', 'Categoria', 'Observações'
                ]

                # Agrupar linhas por aba
                abas_data = {}
                for _, row in df_merge.iterrows():
                    desc_raw = str(row.get('Descricao', '') or '')
                    aba = identificar_aba(desc_raw)
                    if aba not in abas_data:
                        abas_data[aba] = []

                    cod = str(row.get('CodigoProduto', '')).strip()
                    try:
                        cod_n = str(int(float(cod)))
                    except:
                        cod_n = cod

                    cx = cx_lookup.get(cod_n, 0)
                    qtd_cont = float(row.get('QtdContratada', 0) or 0)
                    qtd_ent  = float(row.get('QtdEntregue', 0) or 0)
                    qtd_pend = float(row.get('QtdPendente', 0) or 0)
                    val_unit = float(row.get('ValorUnit', 0) or 0)
                    val_pend = val_unit * qtd_pend

                    # Volumes em caixas
                    volumes_cx = math.ceil(qtd_pend / cx) if cx > 0 else ''

                    # Dias pendentes
                    try:
                        dt_em = pd.to_datetime(row.get('DataEmissao')).date()
                        dias_pend = (hoje - dt_em).days
                    except:
                        dias_pend = ''

                    # % entregue
                    perc_ent = round((qtd_ent / qtd_cont * 100), 1) if qtd_cont > 0 else 0

                    desc_pura = extrair_descricao_pura(desc_raw, cod)
                    categoria = identificar_categoria(desc_raw, aba)

                    dt_em_fmt = ''
                    try:
                        dt_em_fmt = pd.to_datetime(row.get('DataEmissao')).strftime('%d/%m/%Y')
                    except:
                        pass

                    # Gramatura pelo código
                    gram_val = gram_lookup.get(cod_n, '')

                    abas_data[aba].append([
                        str(row.get('NumeroPedido', '')),  # N° Pedido
                        row.get('Cliente', ''),
                        cod,
                        gram_val,      # Gramatura
                        volumes_cx,
                        desc_pura,
                        qtd_cont,
                        qtd_ent,
                        qtd_pend,
                        val_unit,
                        val_pend,
                        dt_em_fmt,
                        dias_pend,
                        row.get('Vendedor', ''),
                        f"{perc_ent:.1f}%",
                        '',            # Previsão — branco
                        categoria,
                        '',            # Observações — branco
                    ])

                # Criar workbook
                wb = openpyxl.Workbook()
                wb.remove(wb.active)

                # Estilos
                HDR_FILL  = PatternFill("solid", fgColor="1F4788")
                HDR_FONT  = Font(bold=True, color="FFFFFF", size=10)
                HDR_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
                BORDER    = Border(
                    left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin')
                )
                ALT_FILL  = PatternFill("solid", fgColor="EEF3FC")

                ORDEM_ABAS = ['Atadura Farma', 'Atadura Hospitalar', 'Campo', 'Tipo Queijo', 'Esteril', 'Pacote', 'Outros']

                for nome_aba in ORDEM_ABAS:
                    linhas = abas_data.get(nome_aba, [])
                    ws = wb.create_sheet(title=nome_aba)

                    # Cabeçalho
                    ws.append(COLUNAS)
                    for col_idx, _ in enumerate(COLUNAS, 1):
                        cell = ws.cell(row=1, column=col_idx)
                        cell.fill   = HDR_FILL
                        cell.font   = HDR_FONT
                        cell.alignment = HDR_ALIGN
                        cell.border = BORDER

                    ws.row_dimensions[1].height = 30

                    # Dados
                    for r_idx, linha in enumerate(linhas, 2):
                        ws.append(linha)
                        fill = ALT_FILL if r_idx % 2 == 0 else PatternFill()
                        for col_idx in range(1, len(COLUNAS) + 1):
                            cell = ws.cell(row=r_idx, column=col_idx)
                            cell.border = BORDER
                            cell.alignment = Alignment(vertical="center")
                            if fill.fill_type:
                                cell.fill = fill
                            # Formatar moeda
                            if col_idx in (10, 11):
                                cell.number_format = 'R$ #,##0.00'
                            # Formatar números inteiros
                            if col_idx in (7, 8, 9):
                                cell.number_format = '#,##0'

                    # Larguras de coluna
                    larguras = [14, 30, 10, 12, 10, 35, 12, 12, 12, 14, 14, 14, 12, 20, 10, 14, 12, 20]
                    for i, larg in enumerate(larguras, 1):
                        ws.column_dimensions[get_column_letter(i)].width = larg

                    # Rodapé com totais (última linha)
                    if linhas:
                        ultima = len(linhas) + 2
                        ws.cell(ultima, 1, 'TOTAL').font = Font(bold=True)
                        # Somar Contratado, Entregue, Pendente, ValorPendente
                        for col_idx, nome_col in enumerate(COLUNAS, 1):
                            if nome_col in ('Contratado', 'Entregue', 'Pendente', 'Valor Pendente'):
                                total = sum(
                                    float(linha[col_idx - 1]) if isinstance(linha[col_idx - 1], (int, float)) else 0
                                    for linha in linhas
                                )
                                c = ws.cell(ultima, col_idx, total)
                                c.font = Font(bold=True)
                                c.number_format = '#,##0.00' if nome_col == 'Valor Pendente' else '#,##0'

                output = BytesIO()
                wb.save(output)
                return output.getvalue()

            _dc1, _dc2 = st.columns(2)
            with _dc1:
                st.download_button(
                    "📥 Relatório Final com Previsão",
                    _gerar_relatorio_previsao(
                        _df_merge, _df_prod_prev, _cx_col, _preco_col, _desc_col
                    ),
                    "RELATORIO_FINAL_COM_PREVISAO.xlsx",
                    "application/vnd.ms-excel",
                    key="dl_previsao_final"
                )
            with _dc2:
                st.download_button(
                    "📥 Faturamento por Data",
                    to_excel(_fat_data),
                    "FATURAMENTO_POR_DATA.xlsx",
                    "application/vnd.ms-excel",
                    key="dl_fat_data"
                )


    # =====================================================================
    # CONCILIAÇÃO: Relatório Atual + Relatório Anterior com Observações
    # =====================================================================
    st.markdown("---")
    st.markdown("### 🔀 Conciliar Relatórios")
    st.caption(
        "Carregue o **Relatório Atual** (recém gerado, sem observações) e o "
        "**Relatório Anterior** onde você preencheu Previsão e Observações. "
        "O sistema transfere as colunas preenchidas para o arquivo atualizado, "
        "vinculando pelo N° Pedido."
    )

    _cc1, _cc2 = st.columns(2)
    with _cc1:
        st.markdown("**1. Relatório Atual** (gerado agora, sem observações)")
        _f_atual = st.file_uploader(
            "Relatório Atual", type=["xlsx"], key="conc_atual",
            label_visibility="collapsed"
        )
    with _cc2:
        st.markdown("**2. Relatório Anterior** (com Previsão e Observações preenchidas)")
        _f_anterior = st.file_uploader(
            "Relatório Anterior", type=["xlsx"], key="conc_anterior",
            label_visibility="collapsed"
        )

    # Salvar bytes em session_state para sobreviver reruns
    if _f_atual:
        st.session_state['_conc_bytes_atual']    = _f_atual.read()
    if _f_anterior:
        st.session_state['_conc_bytes_anterior'] = _f_anterior.read()

    _bytes_atual    = st.session_state.get('_conc_bytes_atual')
    _bytes_anterior = st.session_state.get('_conc_bytes_anterior')

    if _bytes_atual and _bytes_anterior:
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            from openpyxl.utils import get_column_letter
            from io import BytesIO as _BIO

            # ── PASSO 1: extrair mapa {N°Pedido -> {previsao, obs}} do arquivo ANTERIOR ──
            _obs_map = {}
            _wb_ant  = openpyxl.load_workbook(_BIO(_bytes_anterior), data_only=True)

            for _ws_ant in _wb_ant.worksheets:
                _all_rows = list(_ws_ant.iter_rows(values_only=True))
                if len(_all_rows) < 2:
                    continue

                # cabeçalho: normalizar para maiúsculo sem espaços extras
                _hdr = [str(c).strip() if c is not None else '' for c in _all_rows[0]]

                # localizar colunas pelo nome exato (maiúsculo)
                _i_num = _i_prev = _i_obs = None
                for _i, _h in enumerate(_hdr):
                    _hu = _h.upper()
                    if _i_num  is None and any(x in _hu for x in ['N° PEDIDO','N PEDIDO','NUMERO','NUM']):
                        _i_num = _i
                    if _i_prev is None and 'PREV' in _hu:
                        _i_prev = _i
                    if _i_obs  is None and ('OBSERV' in _hu or _hu == 'OBS'):
                        _i_obs = _i

                # detectar também coluna Código e Cliente para chave composta (fallback)
                _i_cod = _i_cli = None
                for _i, _h in enumerate(_hdr):
                    _hu = _h.upper()
                    if _i_cod is None and any(x in _hu for x in ['CÓDIGO','CODIGO']) and 'N°' not in _hu:
                        _i_cod = _i
                    if _i_cli is None and 'CLIENTE' in _hu:
                        _i_cli = _i

                # se não tem N° Pedido E não tem nenhuma coluna útil, pular
                if _i_num is None and _i_cod is None:
                    continue

                for _row in _all_rows[1:]:
                    # chave primária: N° Pedido; fallback: Código|Cliente
                    if _i_num is not None:
                        _raw_num = _row[_i_num] if _i_num < len(_row) else None
                        _k = str(_raw_num).strip() if _raw_num not in (None, '', 'None') else ''
                    else:
                        _cod_v = str(_row[_i_cod]).strip() if _i_cod < len(_row) and _row[_i_cod] not in (None,'','None') else ''
                        _cli_v = str(_row[_i_cli]).strip() if _i_cli is not None and _i_cli < len(_row) and _row[_i_cli] not in (None,'','None') else ''
                        _k = f"{_cod_v}|{_cli_v}" if _cod_v else ''

                    if not _k or _k.upper().startswith('TOTAL'):
                        continue

                    _prev_v = ''
                    _obs_v  = ''
                    if _i_prev is not None and _i_prev < len(_row):
                        _rv = _row[_i_prev]
                        if _rv not in (None, '', 'None'):
                            _prev_v = str(_rv).strip()
                    if _i_obs is not None and _i_obs < len(_row):
                        _rv = _row[_i_obs]
                        if _rv not in (None, '', 'None'):
                            _obs_v = str(_rv).strip()

                    # não sobrescrever com vazio se já tem valor de outra aba
                    _ex = _obs_map.get(_k, {})
                    _obs_map[_k] = {
                        'previsao': _prev_v or _ex.get('previsao', ''),
                        'obs':      _obs_v  or _ex.get('obs', '')
                    }

            # ── PASSO 2: gramatura via tabela de produtos do GitHub ──────────
            _gram_map = {}
            if planilhas_disponiveis.get('produtos_agrupados'):
                _df_gram = carregar_planilha_github(planilhas_disponiveis['produtos_agrupados']['url'])
                if _df_gram is not None:
                    _df_gram.columns = _df_gram.columns.str.upper().str.strip()
                    _gc = next((c for c in _df_gram.columns if any(x in c for x in ['ID_COD','CODIGO','COD'])), None)
                    _gg = next((c for c in _df_gram.columns if 'GRAMATUR' in c), None)
                    if _gc and _gg:
                        for _, _gr in _df_gram.iterrows():
                            try:    _gk = str(int(float(str(_gr[_gc]).strip())))
                            except: _gk = str(_gr[_gc]).strip()
                            _gv = str(_gr[_gg]).strip()
                            if _gv and _gv.lower() not in ('nan','0','0.0',''):
                                _gram_map[_gk] = _gv

            # ── PASSO 3: copiar arquivo ATUAL aba a aba injetando os valores ──
            _wb_at  = openpyxl.load_workbook(_BIO(_bytes_atual))
            _wb_out = openpyxl.Workbook()
            _wb_out.remove(_wb_out.active)

            _S_HDR  = PatternFill("solid", fgColor="1F4788")
            _S_FONT = Font(bold=True, color="FFFFFF", size=10)
            _S_ALGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
            _S_BRD  = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'),  bottom=Side(style='thin'))
            _S_ALT  = PatternFill("solid", fgColor="EEF3FC")

            _total_aplicados = 0

            for _ws_src in _wb_at.worksheets:
                _src_rows = list(_ws_src.iter_rows(values_only=True))
                if len(_src_rows) < 2:
                    continue

                _hdr_src = [str(c).strip() if c is not None else '' for c in _src_rows[0]]

                # índices das colunas-chave no arquivo atual
                _si_num = _si_cod = _si_prev = _si_obs = _si_gram = None
                for _i, _h in enumerate(_hdr_src):
                    _hu = _h.upper()
                    if _si_num  is None and any(x in _hu for x in ['N° PEDIDO','N PEDIDO','NUMERO','NUM']):
                        _si_num = _i
                    if _si_cod  is None and any(x in _hu for x in ['CÓDIGO','CODIGO']) and 'N°' not in _hu:
                        _si_cod = _i
                    if _si_prev is None and 'PREV' in _hu:
                        _si_prev = _i
                    if _si_obs  is None and ('OBSERV' in _hu or _hu == 'OBS'):
                        _si_obs = _i
                    if _si_gram is None and 'GRAMATUR' in _hu:
                        _si_gram = _i

                # cabeçalho de saída = mesmo do atual (já tem N° Pedido e Gramatura
                # desde a correção anterior; se por acaso faltar, adiciona)
                _hdr_out = list(_hdr_src)
                if _si_num is None:
                    _hdr_out = ['N° Pedido'] + _hdr_out
                    _si_num = 0  # agora está na pos 0

                # garantir Gramatura depois de Código
                if _si_gram is None:
                    _pos_cod = next((i for i,h in enumerate(_hdr_out) if any(x in h.upper() for x in ['CÓDIGO','CODIGO']) and 'N°' not in h.upper()), None)
                    if _pos_cod is not None:
                        _hdr_out.insert(_pos_cod + 1, 'Gramatura')
                        _si_gram = _pos_cod + 1
                        # reajustar índices que deslocaram
                        if _si_prev is not None and _si_prev > _pos_cod: _si_prev += 1
                        if _si_obs  is not None and _si_obs  > _pos_cod: _si_obs  += 1
                    else:
                        _hdr_out.append('Gramatura')
                        _si_gram = len(_hdr_out) - 1

                # garantir Previsão
                if _si_prev is None:
                    # inserir antes de Observações se existir, senão no fim
                    _pos_obs = next((i for i,h in enumerate(_hdr_out) if 'OBSERV' in h.upper() or h.upper()=='OBS'), None)
                    if _pos_obs is not None:
                        _hdr_out.insert(_pos_obs, 'Previsão')
                        _si_prev = _pos_obs
                        _si_obs  = _pos_obs + 1
                    else:
                        _hdr_out.append('Previsão')
                        _si_prev = len(_hdr_out) - 1

                # garantir Observações
                if _si_obs is None:
                    _hdr_out.append('Observações')
                    _si_obs = len(_hdr_out) - 1

                # criar aba de saída
                _ws_out = _wb_out.create_sheet(title=_ws_src.title)

                # cabeçalho estilizado
                _ws_out.append(_hdr_out)
                for _ci in range(1, len(_hdr_out) + 1):
                    _c = _ws_out.cell(1, _ci)
                    _c.fill = _S_HDR; _c.font = _S_FONT
                    _c.alignment = _S_ALGN; _c.border = _S_BRD
                _ws_out.row_dimensions[1].height = 30

                # dados: linha a linha
                _ri = 2
                for _row_src in _src_rows[1:]:
                    _rv = list(_row_src)

                    # extrair N° Pedido desta linha
                    _num = ''
                    if _si_num is not None and _si_num < len(_rv) and _rv[_si_num] not in (None, '', 'None'):
                        _num = str(_rv[_si_num]).strip()

                    # extrair código para gramatura
                    _gram_val = ''
                    _ck = ''
                    if _si_cod is not None and _si_cod < len(_rv) and _rv[_si_cod] not in (None, ''):
                        try:    _ck = str(int(float(str(_rv[_si_cod]))))
                        except: _ck = str(_rv[_si_cod]).strip()
                        _gram_val = _gram_map.get(_ck, '')

                    # buscar previsão/obs: chave primária = N° Pedido
                    # fallback = Código|Cliente (para arquivos gerados antes da correção)
                    _prev_val = _obs_val = ''
                    _lookup_key = _num if _num and _num.upper() != 'TOTAL' else ''
                    if not _lookup_key and _ck:
                        _cli_v = ''
                        _si_cli = next((i for i,h in enumerate(_hdr_out) if 'CLIENTE' in h.upper()), None)
                        if _si_cli is not None and _si_cli < len(_rv):
                            _cli_v = str(_rv[_si_cli]).strip() if _rv[_si_cli] not in (None,'','None') else ''
                        _lookup_key = f"{_ck}|{_cli_v}"

                    if _lookup_key:
                        _entry = _obs_map.get(_lookup_key, {})
                        _prev_val = _entry.get('previsao', '')
                        _obs_val  = _entry.get('obs', '')
                        if _prev_val or _obs_val:
                            _total_aplicados += 1

                    # montar linha de saída na ordem do _hdr_out
                    _row_out = []
                    for _oi, _col_name in enumerate(_hdr_out):
                        if _oi == _si_prev:
                            _row_out.append(_prev_val)
                        elif _oi == _si_obs:
                            _row_out.append(_obs_val)
                        elif _oi == _si_gram:
                            _row_out.append(_gram_val)
                        else:
                            # encontrar índice correspondente no src pelo nome
                            _src_i = next(
                                (i for i, h in enumerate(_hdr_src)
                                 if str(h).strip().upper() == _col_name.upper()),
                                None
                            )
                            _row_out.append(_rv[_src_i] if _src_i is not None and _src_i < len(_rv) else '')

                    _ws_out.append(_row_out)

                    # estilos da linha
                    _fill = _S_ALT if _ri % 2 == 0 else PatternFill()
                    for _ci in range(1, len(_hdr_out) + 1):
                        _c = _ws_out.cell(_ri, _ci)
                        _c.border = _S_BRD
                        _c.alignment = Alignment(vertical="center")
                        if _fill.fill_type:
                            _c.fill = _fill
                        _hu = _hdr_out[_ci - 1].upper()
                        if any(x in _hu for x in ['VALOR UNIT', 'VALOR PEND']):
                            _c.number_format = 'R$ #,##0.00'
                        elif any(x in _hu for x in ['CONTRAT','ENTREGUE','PENDENTE','VOLUMES']):
                            try:
                                if _c.value not in (None, ''):
                                    _c.value = float(_c.value)
                                    _c.number_format = '#,##0'
                            except:
                                pass
                    _ri += 1

                # larguras
                _LARG = {
                    'N° PEDIDO': 14, 'CLIENTE': 30, 'CÓDIGO': 10, 'GRAMATURA': 12,
                    'VOLUMES': 10, 'DESCRI': 35, 'CONTRAT': 12, 'ENTREGUE': 12,
                    'PENDENTE': 12, 'VALOR UNIT': 14, 'VALOR PEND': 14,
                    'DATA': 14, 'DIAS': 12, 'VENDEDOR': 20, '%': 10,
                    'PREV': 18, 'CATEG': 12, 'OBSERV': 28, 'OBS': 28,
                }
                for _ci, _col_name in enumerate(_hdr_out, 1):
                    _hu = _col_name.upper()
                    _w = next((_v for _k, _v in _LARG.items() if _k in _hu), 12)
                    _ws_out.column_dimensions[get_column_letter(_ci)].width = _w

                # rodapé totais
                _ultima = _ri
                _ws_out.cell(_ultima, 1, 'TOTAL').font = Font(bold=True)
                for _ci, _col_name in enumerate(_hdr_out, 1):
                    _hu = _col_name.upper()
                    if any(x in _hu for x in ['CONTRAT','ENTREGUE','PENDENTE','VALOR PEND']):
                        _tot = 0
                        for _rr in range(2, _ultima):
                            try: _tot += float(_ws_out.cell(_rr, _ci).value or 0)
                            except: pass
                        _c = _ws_out.cell(_ultima, _ci, _tot)
                        _c.font = Font(bold=True)
                        _c.number_format = 'R$ #,##0.00' if 'VALOR' in _hu else '#,##0'

            # ── PASSO 4: gerar download ────────────────────────────────────
            _buf = _BIO()
            _wb_out.save(_buf)

            _n_prev = sum(1 for v in _obs_map.values() if v.get('previsao'))
            _n_obs  = sum(1 for v in _obs_map.values() if v.get('obs'))
            st.success(
                f"✅ Conciliação concluída — "
                f"{len(_obs_map)} pedidos lidos do arquivo anterior | "
                f"{_n_prev} com Previsão | {_n_obs} com Observações | "
                f"{_total_aplicados} linhas preenchidas no arquivo final"
            )

            # debug: mostrar amostra do mapa lido
            if _obs_map:
                with st.expander("🔍 Ver pedidos lidos do arquivo anterior"):
                    _sample = {k: v for k, v in list(_obs_map.items())[:20] if v.get('previsao') or v.get('obs')}
                    if _sample:
                        st.dataframe(
                            pd.DataFrame([{'N° Pedido': k, 'Previsão': v['previsao'], 'Observações': v['obs']} for k, v in _sample.items()])
                        )
                    else:
                        st.warning("Nenhum pedido com Previsão ou Observação encontrado no arquivo anterior.")

            st.download_button(
                "📥 Baixar Relatório Conciliado",
                _buf.getvalue(),
                "RELATORIO_CONCILIADO.xlsx",
                "application/vnd.ms-excel",
                key="dl_conciliado"
            )

        except Exception as _e:
            import traceback
            st.error(f"❌ Erro na conciliação: {_e}")
            st.code(traceback.format_exc())

    elif _bytes_atual or _bytes_anterior:
        st.info("⬆️ Carregue os dois arquivos para habilitar a conciliação.")

# ====================== PERFORMANCE DE VENDEDORES ======================
elif menu == "Performance de Vendedores":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">📈 Performance de Vendedores</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#6C757D;font-size:0.88rem;margin-bottom:20px;">Painel gerencial completo — análise individual e comparativa por vendedor</p>', unsafe_allow_html=True)

    # ── Filtros locais do módulo ──────────────────────────────────────────────
    with st.expander("⚙️ Filtros do Módulo", expanded=True):
        _pv_c1, _pv_c2, _pv_c3 = st.columns(3)
        with _pv_c1:
            _pv_vendedores_lista = ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist())
            _pv_vendedor = st.selectbox("👤 Vendedor", _pv_vendedores_lista, key="pv_vendedor")
        with _pv_c2:
            _pv_regioes_lista = ['Todas'] + sorted(df['Estado'].dropna().unique().tolist())
            _pv_regiao = st.selectbox("🗺️ Região (Estado)", _pv_regioes_lista, key="pv_regiao")
        with _pv_c3:
            _pv_periodo = st.selectbox(
                "📅 Período",
                ["Filtro Global", "Mês Atual", "Últimos 3 Meses", "Últimos 6 Meses", "Ano Atual", "Personalizado"],
                key="pv_periodo"
            )
        if _pv_periodo == "Personalizado":
            _pv_d1, _pv_d2 = st.columns(2)
            with _pv_d1:
                _pv_data_ini = st.date_input("De", value=None, key="pv_data_ini", format="DD/MM/YYYY")
            with _pv_d2:
                _pv_data_fim = st.date_input("Até", value=None, key="pv_data_fim", format="DD/MM/YYYY")
        else:
            _pv_data_ini = None
            _pv_data_fim = None

    # ── Aplicar filtros ───────────────────────────────────────────────────────
    _pv_now = pd.Timestamp.now()
    _pv_df = df.copy()

    # Filtro período
    if _pv_periodo == "Filtro Global":
        _pv_df = df_filtrado.copy()
    elif _pv_periodo == "Mês Atual":
        _pv_df = _pv_df[
            (_pv_df['DataEmissao'].dt.month == _pv_now.month) &
            (_pv_df['DataEmissao'].dt.year == _pv_now.year)
        ]
    elif _pv_periodo == "Últimos 3 Meses":
        _pv_df = _pv_df[_pv_df['DataEmissao'] >= (_pv_now - pd.DateOffset(months=3))]
    elif _pv_periodo == "Últimos 6 Meses":
        _pv_df = _pv_df[_pv_df['DataEmissao'] >= (_pv_now - pd.DateOffset(months=6))]
    elif _pv_periodo == "Ano Atual":
        _pv_df = _pv_df[_pv_df['DataEmissao'].dt.year == _pv_now.year]
    elif _pv_periodo == "Personalizado":
        if _pv_data_ini:
            _pv_df = _pv_df[_pv_df['DataEmissao'] >= pd.to_datetime(_pv_data_ini)]
        if _pv_data_fim:
            _pv_df = _pv_df[_pv_df['DataEmissao'] <= pd.to_datetime(_pv_data_fim)]

    # Filtro região
    if _pv_regiao != 'Todas':
        _pv_df = _pv_df[_pv_df['Estado'] == _pv_regiao]

    # Filtro vendedor
    if _pv_vendedor != 'Todos':
        _pv_df = _pv_df[_pv_df['Vendedor'] == _pv_vendedor]

    # Base apenas vendas e devoluções
    _pv_vendas = _pv_df[_pv_df['TipoMov'] == 'NF Venda'].copy()
    _pv_devol  = _pv_df[_pv_df['TipoMov'] == 'NF Dev.Venda'].copy()
    _pv_notas  = obter_notas_unicas(_pv_df)
    _pv_notas_v = _pv_notas[_pv_notas['TipoMov'] == 'NF Venda']
    _pv_notas_d = _pv_notas[_pv_notas['TipoMov'] == 'NF Dev.Venda']

    # ── KPIs Consolidados ─────────────────────────────────────────────────────
    _pv_fat_bruto   = _pv_notas_v['TotalProduto'].sum()
    _pv_fat_devol   = _pv_notas_d['TotalProduto'].sum()
    _pv_fat_liq     = _pv_fat_bruto - _pv_fat_devol
    _pv_clientes    = _pv_vendas['CPF_CNPJ'].nunique()
    _pv_qtd_notas   = len(_pv_notas_v)
    _pv_ticket      = _pv_fat_bruto / _pv_clientes if _pv_clientes > 0 else 0
    _pv_vol_total   = _pv_vendas['Quantidade'].sum() if 'Quantidade' in _pv_vendas.columns else 0

    # Prazo médio
    def _pv_prazo_medio(df_v):
        try:
            if 'PrazoHistorico' not in df_v.columns:
                return 0
            prazos = []
            for val in df_v['PrazoHistorico'].dropna():
                for p in str(val).split('/'):
                    try:
                        prazos.append(int(p))
                    except:
                        pass
            return sum(prazos) / len(prazos) if prazos else 0
        except:
            return 0

    _pv_prazo = _pv_prazo_medio(_pv_vendas)

    # Comissão média
    def _pv_comissao_media(df_v):
        try:
            if 'Comissao' not in df_v.columns:
                return "N/D"
            mapa = {'4%': 4.0, '3%': 3.0, '2,5%': 2.5, '2%': 2.0}
            vals = df_v['Comissao'].map(mapa).dropna()
            if len(vals) == 0:
                return "N/D"
            return f"{vals.mean():.2f}%"
        except:
            return "N/D"

    _pv_comissao = _pv_comissao_media(_pv_vendas)

    # ── Exibir KPI Cards ─────────────────────────────────────────────────────
    _pv_k1, _pv_k2, _pv_k3, _pv_k4 = st.columns(4)
    with _pv_k1:
        render_kpi_card("Faturamento Líquido", f"R$ {_pv_fat_liq:,.0f}", icon="💰", color="#1F4788")
    with _pv_k2:
        render_kpi_card("Faturamento Bruto", f"R$ {_pv_fat_bruto:,.0f}", icon="💵", color="#2E86AB")
    with _pv_k3:
        render_kpi_card("Devoluções", f"R$ {_pv_fat_devol:,.0f}", icon="↩️", color="#EF4444")
    with _pv_k4:
        render_kpi_card("Clientes Positivados", f"{_pv_clientes:,}", icon="👥", color="#28A745")

    st.markdown("<br>", unsafe_allow_html=True)

    _pv_k5, _pv_k6, _pv_k7, _pv_k8 = st.columns(4)
    with _pv_k5:
        render_kpi_card("Ticket Médio", f"R$ {_pv_ticket:,.0f}", icon="🎯", color="#F4A261")
    with _pv_k6:
        render_kpi_card("Volume Vendido", f"{_pv_vol_total:,.0f} un", icon="📦", color="#6C757D")
    with _pv_k7:
        render_kpi_card("Prazo Médio", f"{_pv_prazo:.0f} dias", icon="📅", color="#163561")
    with _pv_k8:
        render_kpi_card("Comissão Média", _pv_comissao, icon="💎", color="#1B5E8A")

    st.markdown("---")

    # ── Inadimplência por Vendedor ─────────────────────────────────────────
    _pv_df_inad = None
    _pv_inad_vendedor = 0
    _pv_inad_total = 0
    _pv_perc_inad = 0.0
    if planilhas_disponiveis.get('inadimplencia'):
        try:
            _pv_raw_inad = carregar_planilha_github(planilhas_disponiveis['inadimplencia']['url'])
            if _pv_raw_inad is not None:
                _pv_df_inad = processar_inadimplencia(_pv_raw_inad)
                if _pv_vendedor != 'Todos' and 'Vendedor' in _pv_df_inad.columns:
                    _pv_inad_vend_df = _pv_df_inad[_pv_df_inad['Vendedor'] == _pv_vendedor]
                else:
                    _pv_inad_vend_df = _pv_df_inad.copy()
                if _pv_regiao != 'Todas' and 'Estado' in _pv_df_inad.columns:
                    _pv_inad_vend_df = _pv_inad_vend_df[_pv_inad_vend_df['Estado'] == _pv_regiao]
                _pv_inad_vendedor = _pv_inad_vend_df['ValorLiquido'].sum() if 'ValorLiquido' in _pv_inad_vend_df.columns else 0
                _pv_inad_total    = _pv_df_inad['ValorLiquido'].sum() if 'ValorLiquido' in _pv_df_inad.columns else 0
                _pv_perc_inad     = (_pv_inad_vendedor / _pv_fat_bruto * 100) if _pv_fat_bruto > 0 else 0
        except:
            pass

    # Card de inadimplência
    _pv_ki1, _pv_ki2 = st.columns(2)
    with _pv_ki1:
        render_kpi_card(
            "Índice de Inadimplência (R$)",
            f"R$ {_pv_inad_vendedor:,.0f}",
            icon="⚠️",
            color="#EF4444" if _pv_inad_vendedor > 0 else "#28A745"
        )
    with _pv_ki2:
        render_kpi_card(
            "Inadimplência sobre Faturamento",
            f"{_pv_perc_inad:.1f}%",
            icon="📊",
            color="#EF4444" if _pv_perc_inad > 5 else "#F4A261"
        )

    st.markdown("---")

    # ── Tabs de análise ───────────────────────────────────────────────────────
    _pv_tab1, _pv_tab2, _pv_tab3, _pv_tab4 = st.tabs([
        "📊 Comparativo", "📈 Evolução Temporal", "🌐 Capilaridade", "🛒 Mix de Produtos"
    ])

    # ─── Tab 1: Comparativo de Vendedores ────────────────────────────────────
    with _pv_tab1:
        st.markdown("#### Comparativo de Desempenho por Vendedor")

        _pv_comp = _pv_notas_v.groupby('Vendedor').agg(
            FaturamentoBruto=('TotalProduto', 'sum'),
            QtdNotas=('Numero_NF', 'count'),
            ClientesAtendidos=('CPF_CNPJ', 'nunique'),
        ).reset_index()

        # Ticket médio por vendedor
        _pv_comp['TicketMedio'] = _pv_comp['FaturamentoBruto'] / _pv_comp['ClientesAtendidos'].replace(0, 1)

        # Volume por vendedor
        if 'Quantidade' in _pv_vendas.columns:
            _pv_vol_vend = _pv_vendas.groupby('Vendedor')['Quantidade'].sum().reset_index()
            _pv_vol_vend.columns = ['Vendedor', 'VolumeTotal']
            _pv_comp = _pv_comp.merge(_pv_vol_vend, on='Vendedor', how='left')
        else:
            _pv_comp['VolumeTotal'] = 0

        # Comissão média por vendedor
        if 'Comissao' in _pv_vendas.columns:
            _mapa_com = {'4%': 4.0, '3%': 3.0, '2,5%': 2.5, '2%': 2.0}
            _pv_vendas_c = _pv_vendas.copy()
            _pv_vendas_c['ComissaoNum'] = _pv_vendas_c['Comissao'].map(_mapa_com)
            _pv_com_vend = _pv_vendas_c.groupby('Vendedor')['ComissaoNum'].mean().reset_index()
            _pv_com_vend.columns = ['Vendedor', 'ComissaoMedia']
            _pv_comp = _pv_comp.merge(_pv_com_vend, on='Vendedor', how='left')
        else:
            _pv_comp['ComissaoMedia'] = None

        # Prazo médio por vendedor
        if 'PrazoHistorico' in _pv_vendas.columns:
            def _prazo_med_vend(series):
                prazos = []
                for val in series.dropna():
                    for p in str(val).split('/'):
                        try:
                            prazos.append(int(p))
                        except:
                            pass
                return sum(prazos) / len(prazos) if prazos else 0
            _pv_prazo_vend = _pv_vendas.groupby('Vendedor')['PrazoHistorico'].apply(_prazo_med_vend).reset_index()
            _pv_prazo_vend.columns = ['Vendedor', 'PrazoMedio']
            _pv_comp = _pv_comp.merge(_pv_prazo_vend, on='Vendedor', how='left')
        else:
            _pv_comp['PrazoMedio'] = 0

        _pv_comp = _pv_comp.sort_values('FaturamentoBruto', ascending=False)

        # Gráfico comparativo faturamento
        _pv_cv1, _pv_cv2 = st.columns(2)
        with _pv_cv1:
            _fig_fat = px.bar(
                _pv_comp.head(15),
                x='Vendedor', y='FaturamentoBruto',
                title='Faturamento Bruto por Vendedor',
                labels={'FaturamentoBruto': 'R$', 'Vendedor': ''},
                color='FaturamentoBruto',
                color_continuous_scale=['#A8C4E8', '#1F4788']
            )
            _fig_fat = aplicar_layout_grafico(_fig_fat, height=340)
            _fig_fat.update_traces(
                hovertemplate='<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>'
            )
            st.plotly_chart(_fig_fat, use_container_width=True)

        with _pv_cv2:
            _fig_cli = px.bar(
                _pv_comp.head(15),
                x='Vendedor', y='ClientesAtendidos',
                title='Clientes Atendidos por Vendedor',
                labels={'ClientesAtendidos': 'Clientes', 'Vendedor': ''},
                color='ClientesAtendidos',
                color_continuous_scale=['#B8E0C8', '#28A745']
            )
            _fig_cli = aplicar_layout_grafico(_fig_cli, height=340)
            _fig_cli.update_traces(
                hovertemplate='<b>%{x}</b><br>%{y} clientes<extra></extra>'
            )
            st.plotly_chart(_fig_cli, use_container_width=True)

        _pv_cv3, _pv_cv4 = st.columns(2)
        with _pv_cv3:
            _fig_ticket = px.bar(
                _pv_comp.head(15),
                x='Vendedor', y='TicketMedio',
                title='Ticket Médio por Vendedor',
                labels={'TicketMedio': 'R$', 'Vendedor': ''},
                color='TicketMedio',
                color_continuous_scale=['#F8D9B8', '#F4A261']
            )
            _fig_ticket = aplicar_layout_grafico(_fig_ticket, height=340)
            _fig_ticket.update_traces(
                hovertemplate='<b>%{x}</b><br>Ticket: R$ %{y:,.2f}<extra></extra>'
            )
            st.plotly_chart(_fig_ticket, use_container_width=True)

        with _pv_cv4:
            if _pv_comp['ComissaoMedia'].notna().any():
                _fig_com = px.bar(
                    _pv_comp[_pv_comp['ComissaoMedia'].notna()].head(15),
                    x='Vendedor', y='ComissaoMedia',
                    title='Comissão Média por Vendedor (%)',
                    labels={'ComissaoMedia': 'Comissão (%)', 'Vendedor': ''},
                    color='ComissaoMedia',
                    color_continuous_scale=['#C5D5F0', '#163561']
                )
                _fig_com = aplicar_layout_grafico(_fig_com, height=340)
                _fig_com.update_traces(
                    hovertemplate='<b>%{x}</b><br>Comissão: %{y:.2f}%<extra></extra>'
                )
                st.plotly_chart(_fig_com, use_container_width=True)
            else:
                st.info("Dados de comissão não disponíveis. Verifique se a planilha de produtos está carregada.")

        # Tabela consolidada
        st.markdown("#### Tabela Consolidada de Performance")
        _pv_comp_disp = _pv_comp.copy()
        _pv_comp_disp['FaturamentoBruto'] = _pv_comp_disp['FaturamentoBruto'].apply(formatar_moeda)
        _pv_comp_disp['TicketMedio']      = _pv_comp_disp['TicketMedio'].apply(formatar_moeda)
        _pv_comp_disp['VolumeTotal']      = _pv_comp_disp['VolumeTotal'].apply(lambda x: f"{x:,.0f} un")
        _pv_comp_disp['ComissaoMedia']    = _pv_comp_disp['ComissaoMedia'].apply(
            lambda x: f"{x:.2f}%" if pd.notnull(x) else "N/D"
        )
        _pv_comp_disp['PrazoMedio']       = _pv_comp_disp['PrazoMedio'].apply(lambda x: f"{x:.0f} dias")
        _pv_comp_disp.insert(0, 'Posição', range(1, len(_pv_comp_disp) + 1))
        _pv_comp_disp = _pv_comp_disp.rename(columns={
            'FaturamentoBruto': 'Faturamento',
            'QtdNotas':         'Nº Notas',
            'ClientesAtendidos':'Clientes',
            'TicketMedio':      'Ticket Médio',
            'VolumeTotal':      'Volume',
            'ComissaoMedia':    'Comissão Média',
            'PrazoMedio':       'Prazo Médio',
        })
        st.dataframe(_pv_comp_disp, use_container_width=True)

        st.download_button(
            "📥 Exportar Comparativo (Excel)",
            to_excel(_pv_comp),
            "performance_vendedores.xlsx",
            "application/vnd.ms-excel",
            key="pv_dl_comp"
        )

    # ─── Tab 2: Evolução Temporal ─────────────────────────────────────────────
    with _pv_tab2:
        st.markdown("#### Evolução de Vendas ao Longo do Tempo")

        _pv_evol = _pv_notas_v.groupby(['MesAno', 'Vendedor'])['TotalProduto'].sum().reset_index()
        _pv_evol = _pv_evol.sort_values('MesAno')

        if _pv_vendedor != 'Todos':
            # Exibir linha única do vendedor selecionado
            _pv_evol_filt = _pv_evol[_pv_evol['Vendedor'] == _pv_vendedor]
            if len(_pv_evol_filt) > 0:
                _fig_evol = px.line(
                    _pv_evol_filt, x='MesAno', y='TotalProduto',
                    title=f'Evolução Mensal — {_pv_vendedor}',
                    labels={'MesAno': 'Período', 'TotalProduto': 'R$'},
                    markers=True
                )
                _fig_evol.update_traces(line_color='#1F4788', line_width=3, marker=dict(size=7, color='#1F4788'))
                _fig_evol = aplicar_layout_grafico(_fig_evol, height=380)
                st.plotly_chart(_fig_evol, use_container_width=True)
            else:
                st.info("Nenhuma venda encontrada para este vendedor no período.")
        else:
            # Multi-linha: top 8 vendedores
            _top_vend = _pv_notas_v.groupby('Vendedor')['TotalProduto'].sum().nlargest(8).index.tolist()
            _pv_evol_top = _pv_evol[_pv_evol['Vendedor'].isin(_top_vend)]
            _fig_evol = px.line(
                _pv_evol_top, x='MesAno', y='TotalProduto',
                color='Vendedor',
                title='Evolução Mensal — Top 8 Vendedores',
                labels={'MesAno': 'Período', 'TotalProduto': 'R$', 'Vendedor': 'Vendedor'},
                markers=True,
                color_discrete_sequence=CORES_INST
            )
            _fig_evol.update_traces(line_width=2)
            _fig_evol = aplicar_layout_grafico(_fig_evol, height=420)
            st.plotly_chart(_fig_evol, use_container_width=True)

        # Evolução do ticket médio
        st.markdown("#### Evolução do Ticket Médio")
        _pv_tick_evol = _pv_notas_v.groupby(['MesAno', 'Vendedor']).agg(
            Fat=('TotalProduto', 'sum'),
            Cli=('CPF_CNPJ', 'nunique')
        ).reset_index()
        _pv_tick_evol['TicketMedio'] = _pv_tick_evol['Fat'] / _pv_tick_evol['Cli'].replace(0, 1)
        _pv_tick_evol = _pv_tick_evol.sort_values('MesAno')

        if _pv_vendedor != 'Todos':
            _pv_tick_filt = _pv_tick_evol[_pv_tick_evol['Vendedor'] == _pv_vendedor]
            _fig_tick = px.area(
                _pv_tick_filt, x='MesAno', y='TicketMedio',
                title=f'Ticket Médio Mensal — {_pv_vendedor}',
                labels={'MesAno': 'Período', 'TicketMedio': 'R$'},
                color_discrete_sequence=['#2E86AB']
            )
        else:
            _pv_tick_top = _pv_tick_evol[_pv_tick_evol['Vendedor'].isin(_top_vend[:5])]
            _fig_tick = px.line(
                _pv_tick_top, x='MesAno', y='TicketMedio',
                color='Vendedor',
                title='Ticket Médio Mensal — Top 5 Vendedores',
                labels={'MesAno': 'Período', 'TicketMedio': 'R$'},
                color_discrete_sequence=CORES_INST
            )
        _fig_tick = aplicar_layout_grafico(_fig_tick, height=340)
        st.plotly_chart(_fig_tick, use_container_width=True)

    # ─── Tab 3: Capilaridade ─────────────────────────────────────────────────
    with _pv_tab3:
        st.markdown("#### Análise de Capilaridade — Clientes Atendidos por Vendedor")

        _pv_cap = _pv_vendas.groupby(['Vendedor', 'Estado'])['CPF_CNPJ'].nunique().reset_index()
        _pv_cap.columns = ['Vendedor', 'Estado', 'Clientes']

        if _pv_vendedor != 'Todos':
            _pv_cap_filt = _pv_cap[_pv_cap['Vendedor'] == _pv_vendedor]
            _fig_cap = px.bar(
                _pv_cap_filt.sort_values('Clientes', ascending=False),
                x='Estado', y='Clientes',
                title=f'Clientes por Estado — {_pv_vendedor}',
                labels={'Estado': 'Estado', 'Clientes': 'Nº de Clientes'},
                color='Clientes',
                color_continuous_scale=['#B8E0C8', '#1E7B34']
            )
            _fig_cap = aplicar_layout_grafico(_fig_cap, height=380)
            st.plotly_chart(_fig_cap, use_container_width=True)

            # Mapa de calor vendedor × estado
            st.markdown("#### Heatmap de Faturamento por Estado")
            _pv_heat_v = _pv_notas_v[_pv_notas_v['Vendedor'] == _pv_vendedor].groupby('Estado')['TotalProduto'].sum().reset_index()
            _fig_heat = px.bar(
                _pv_heat_v.sort_values('TotalProduto', ascending=True).tail(15),
                x='TotalProduto', y='Estado', orientation='h',
                title=f'Faturamento por Estado — {_pv_vendedor}',
                labels={'TotalProduto': 'R$', 'Estado': ''},
                color='TotalProduto',
                color_continuous_scale=['#C5D5F0', '#1F4788']
            )
            _fig_heat = aplicar_layout_grafico(_fig_heat, height=380)
            st.plotly_chart(_fig_heat, use_container_width=True)

        else:
            # Heatmap vendedor × estado
            _pv_heat = _pv_vendas.groupby(['Vendedor', 'Estado'])['CPF_CNPJ'].nunique().reset_index()
            _pv_heat_pivot = _pv_heat.pivot(index='Vendedor', columns='Estado', values='CPF_CNPJ').fillna(0)

            _fig_heat = go.Figure(data=go.Heatmap(
                z=_pv_heat_pivot.values,
                x=_pv_heat_pivot.columns.tolist(),
                y=_pv_heat_pivot.index.tolist(),
                colorscale=[[0, '#FFFFFF'], [0.5, '#A8C4E8'], [1, '#1F4788']],
                hovertemplate='Vendedor: %{y}<br>Estado: %{x}<br>Clientes: %{z}<extra></extra>'
            ))
            _fig_heat.update_layout(
                title='Capilaridade — Clientes por Vendedor × Estado',
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(family='Inter, Segoe UI, sans-serif', size=11),
                margin=dict(l=10, r=10, t=40, b=10),
                height=480
            )
            st.plotly_chart(_fig_heat, use_container_width=True)

            # Bubble chart: faturamento vs clientes
            _pv_bubble = _pv_notas_v.groupby('Vendedor').agg(
                Fat=('TotalProduto', 'sum'),
                Cli=('CPF_CNPJ', 'nunique'),
                Notas=('Numero_NF', 'count')
            ).reset_index()
            _fig_bub = px.scatter(
                _pv_bubble,
                x='Cli', y='Fat',
                size='Notas', color='Vendedor',
                hover_name='Vendedor',
                title='Faturamento × Clientes Atendidos (tamanho = Qtd Notas)',
                labels={'Cli': 'Clientes Atendidos', 'Fat': 'Faturamento (R$)'},
                color_discrete_sequence=CORES_INST
            )
            _fig_bub = aplicar_layout_grafico(_fig_bub, height=420)
            st.plotly_chart(_fig_bub, use_container_width=True)

    # ─── Tab 4: Mix de Produtos ───────────────────────────────────────────────
    with _pv_tab4:
        st.markdown("#### Concentração de Vendas por Produto")

        _pv_mix = _pv_vendas.groupby('NomeProduto').agg(
            Total=('TotalProduto', 'sum'),
            Volume=('Quantidade', 'sum') if 'Quantidade' in _pv_vendas.columns else ('TotalProduto', 'count'),
            Clientes=('CPF_CNPJ', 'nunique')
        ).reset_index().sort_values('Total', ascending=False)

        _pv_tot_mix = _pv_mix['Total'].sum()
        _pv_mix['Participacao'] = (_pv_mix['Total'] / _pv_tot_mix * 100).round(2)
        _pv_mix['CumulativaPerc'] = _pv_mix['Participacao'].cumsum()

        # Top 15 pizza
        _pv_top15 = _pv_mix.head(15)
        _fig_pie = px.pie(
            _pv_top15, names='NomeProduto', values='Total',
            title='Top 15 Produtos por Faturamento',
            color_discrete_sequence=CORES_INST,
            hole=0.45
        )
        _fig_pie.update_traces(textposition='inside', textinfo='percent+label',
                               hovertemplate='<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>')
        _fig_pie.update_layout(
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Inter, Segoe UI, sans-serif', size=11),
            margin=dict(l=10, r=10, t=40, b=10),
            height=420,
            showlegend=False
        )
        st.plotly_chart(_fig_pie, use_container_width=True)

        # Curva ABC
        st.markdown("#### Curva ABC de Produtos")
        _fig_abc = go.Figure()
        _fig_abc.add_trace(go.Bar(
            x=_pv_top15['NomeProduto'], y=_pv_top15['Total'],
            name='Faturamento', marker_color='#1F4788',
            hovertemplate='<b>%{x}</b><br>R$ %{y:,.2f}<extra></extra>'
        ))
        _fig_abc.add_trace(go.Scatter(
            x=_pv_top15['NomeProduto'], y=_pv_top15['CumulativaPerc'].head(15),
            name='% Acumulado', yaxis='y2',
            line=dict(color='#EF4444', width=2, dash='dash'),
            marker=dict(size=5, color='#EF4444'),
            hovertemplate='%{x}<br>Acumulado: %{y:.1f}%<extra></extra>'
        ))
        _fig_abc.update_layout(
            title='Curva ABC — Top 15 Produtos',
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            font=dict(family='Inter, Segoe UI, sans-serif', size=11),
            margin=dict(l=10, r=10, t=40, b=10),
            height=380,
            yaxis=dict(title='Faturamento (R$)', showgrid=True, gridcolor='#F0F0F0'),
            yaxis2=dict(title='% Acumulado', overlaying='y', side='right', range=[0, 110]),
            legend=dict(orientation='h', y=1.08),
            hoverlabel=dict(bgcolor='#1F4788', font_color='white')
        )
        st.plotly_chart(_fig_abc, use_container_width=True)

        # Tabela de mix
        _pv_mix_disp = _pv_mix.copy()
        _pv_mix_disp['Total']        = _pv_mix_disp['Total'].apply(formatar_moeda)
        _pv_mix_disp['Participacao'] = _pv_mix_disp['Participacao'].apply(lambda x: f"{x:.2f}%")
        _pv_mix_disp['CumulativaPerc'] = _pv_mix_disp['CumulativaPerc'].apply(lambda x: f"{x:.2f}%")
        _pv_mix_disp = _pv_mix_disp.rename(columns={
            'NomeProduto':    'Produto',
            'Total':          'Faturamento',
            'Participacao':   '% Part.',
            'CumulativaPerc': '% Acum.',
            'Clientes':       'Clientes'
        })
        st.dataframe(_pv_mix_disp, use_container_width=True)

        st.download_button(
            "📥 Exportar Mix de Produtos (Excel)",
            to_excel(_pv_mix),
            "mix_produtos.xlsx",
            "application/vnd.ms-excel",
            key="pv_dl_mix"
        )

    # ── Geração de PDF ────────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("### 📄 Exportar Relatório PDF")
    st.markdown('<p style="color:#6C757D;font-size:0.84rem;">Gera um relatório executivo em PDF com base nos filtros aplicados.</p>', unsafe_allow_html=True)

    if st.button("🖨️ Gerar Relatório PDF", type="primary", key="pv_gerar_pdf"):
        try:
            from fpdf import FPDF
            import math

            class PerformancePDF(FPDF):
                def header(self):
                    # Logo
                    try:
                        import tempfile, os
                        _resp_logo = requests.get("https://i.imgur.com/gt3rgyL.png", timeout=8)
                        if _resp_logo.status_code == 200:
                            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as _ftmp:
                                _ftmp.write(_resp_logo.content)
                                _tmp_path = _ftmp.name
                            self.image(_tmp_path, x=10, y=8, w=28)
                            os.unlink(_tmp_path)
                    except Exception:
                        pass
                    self.set_xy(42, 10)
                    self.set_font('Helvetica', 'B', 13)
                    self.set_text_color(31, 71, 136)
                    self.cell(0, 6, 'MEDTEXTIL PRODUTOS TEXTIL HOSPITALARES', ln=True)
                    self.set_xy(42, 17)
                    self.set_font('Helvetica', '', 8)
                    self.set_text_color(108, 117, 125)
                    self.cell(0, 5, 'CNPJ: 40.357.820/0001-50  |  Dashboard Comercial BI 2.0', ln=True)
                    self.ln(4)
                    self.set_draw_color(31, 71, 136)
                    self.set_line_width(0.8)
                    self.line(10, self.get_y(), 200, self.get_y())
                    self.ln(3)

                def footer(self):
                    self.set_y(-14)
                    self.set_font('Helvetica', 'I', 7)
                    self.set_text_color(173, 181, 189)
                    self.cell(0, 5, f'Performance de Vendedores  ·  Gerado em {_pv_now.strftime("%d/%m/%Y %H:%M")}  ·  Pág. {self.page_no()}', align='C')

            pdf = PerformancePDF()
            pdf.set_auto_page_break(auto=True, margin=18)
            pdf.add_page()
            pdf.set_margins(12, 12, 12)

            # ── Título do relatório ───────────────────────────────────────
            pdf.set_font('Helvetica', 'B', 16)
            pdf.set_text_color(31, 71, 136)
            pdf.cell(0, 10, 'RELATÓRIO DE PERFORMANCE DE VENDEDORES', ln=True, align='C')
            pdf.set_font('Helvetica', '', 9)
            pdf.set_text_color(108, 117, 125)
            _pv_label_periodo = f"Período: {_pv_periodo}"
            if _pv_data_ini and _pv_periodo == "Personalizado":
                _pv_label_periodo = f"Período: {_pv_data_ini.strftime('%d/%m/%Y') if hasattr(_pv_data_ini,'strftime') else str(_pv_data_ini)} a {_pv_data_fim.strftime('%d/%m/%Y') if _pv_data_fim and hasattr(_pv_data_fim,'strftime') else '—'}"
            pdf.cell(0, 6, f"Vendedor: {_pv_vendedor}  |  Região: {_pv_regiao}  |  {_pv_label_periodo}", ln=True, align='C')
            pdf.ln(5)

            # ── Resumo Executivo ──────────────────────────────────────────
            pdf.set_fill_color(31, 71, 136)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font('Helvetica', 'B', 9)
            pdf.cell(0, 7, '  RESUMO EXECUTIVO', fill=True, border=0, ln=True)
            pdf.set_text_color(50, 50, 50)
            pdf.set_font('Helvetica', '', 8)

            _pv_kpis_pdf = [
                ('Faturamento Líquido',       f"R$ {_pv_fat_liq:,.2f}"),
                ('Faturamento Bruto',          f"R$ {_pv_fat_bruto:,.2f}"),
                ('Devoluções',                 f"R$ {_pv_fat_devol:,.2f}"),
                ('Clientes Positivados',       f"{_pv_clientes:,}"),
                ('Ticket Médio',               f"R$ {_pv_ticket:,.2f}"),
                ('Volume Vendido',             f"{_pv_vol_total:,.0f} un"),
                ('Prazo Médio de Venda',       f"{_pv_prazo:.0f} dias"),
                ('Comissão Média',             _pv_comissao),
                ('Índice de Inadimplência',    f"R$ {_pv_inad_vendedor:,.2f}"),
                ('Inadimplência / Fat. Bruto', f"{_pv_perc_inad:.1f}%"),
            ]
            w1, w2 = 85, 95
            fill_kpi = False
            for k, v in _pv_kpis_pdf:
                pdf.set_fill_color(240, 244, 255) if fill_kpi else pdf.set_fill_color(255, 255, 255)
                pdf.set_font('Helvetica', 'B', 8)
                pdf.cell(w1, 6, f'  {k}:', border='LB', fill=True)
                pdf.set_font('Helvetica', '', 8)
                pdf.cell(w2, 6, f'  {v}', border='RB', fill=True, ln=True)
                fill_kpi = not fill_kpi
            pdf.ln(6)

            # ── Tabela Comparativa de Vendedores ─────────────────────────
            pdf.set_fill_color(31, 71, 136)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font('Helvetica', 'B', 9)
            pdf.cell(0, 7, '  COMPARATIVO DE VENDEDORES', fill=True, border=0, ln=True)

            _pv_cols_pdf    = ['Vendedor', 'Faturamento (R$)', 'Clientes', 'Ticket Médio', 'Vol.', 'Prazo (d)', 'Comissão']
            _pv_widths_pdf  = [50, 32, 16, 28, 15, 20, 22]

            pdf.set_text_color(255, 255, 255)
            pdf.set_font('Helvetica', 'B', 7)
            for col, w in zip(_pv_cols_pdf, _pv_widths_pdf):
                pdf.cell(w, 7, col, border=1, fill=True, align='C')
            pdf.ln()

            _pv_comp_sorted = _pv_comp.sort_values('FaturamentoBruto', ascending=False).head(20)
            fill_row_pdf = False
            for _, row in _pv_comp_sorted.iterrows():
                pdf.set_fill_color(240, 244, 255) if fill_row_pdf else pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(50, 50, 50)
                pdf.set_font('Helvetica', '', 7)
                _vend_str = str(row['Vendedor'])[:22]
                _fat_str  = f"R$ {row['FaturamentoBruto']:,.2f}"
                _cli_str  = str(int(row['ClientesAtendidos']))
                _tick_str = f"R$ {row['TicketMedio']:,.2f}"
                _vol_str  = f"{row.get('VolumeTotal', 0):,.0f}"
                _prz_str  = f"{row.get('PrazoMedio', 0):.0f}"
                _com_str  = f"{row['ComissaoMedia']:.2f}%" if pd.notnull(row.get('ComissaoMedia')) else "N/D"
                _row_vals = [_vend_str, _fat_str, _cli_str, _tick_str, _vol_str, _prz_str, _com_str]
                _aligns   = ['L', 'R', 'C', 'R', 'C', 'C', 'C']
                for val, w, align in zip(_row_vals, _pv_widths_pdf, _aligns):
                    pdf.cell(w, 6, val, border=1, fill=True, align=align)
                pdf.ln()
                fill_row_pdf = not fill_row_pdf
            pdf.ln(6)

            # ── Top Produtos ──────────────────────────────────────────────
            pdf.set_fill_color(31, 71, 136)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font('Helvetica', 'B', 9)
            pdf.cell(0, 7, '  MIX DE PRODUTOS — TOP 15', fill=True, border=0, ln=True)

            _pv_mix_pdf_cols = ['Produto', 'Faturamento (R$)', '% Part.', '% Acum.', 'Clientes']
            _pv_mix_pdf_w    = [70, 32, 18, 18, 18]
            pdf.set_font('Helvetica', 'B', 7)
            for col, w in zip(_pv_mix_pdf_cols, _pv_mix_pdf_w):
                pdf.cell(w, 7, col, border=1, fill=True, align='C')
            pdf.ln()

            fill_mix = False
            for _, row in _pv_mix.head(15).iterrows():
                pdf.set_fill_color(240, 244, 255) if fill_mix else pdf.set_fill_color(255, 255, 255)
                pdf.set_text_color(50, 50, 50)
                pdf.set_font('Helvetica', '', 7)
                _prod_str = str(row['NomeProduto'])[:38]
                _fat_str  = f"R$ {row['Total']:,.2f}"
                _part_str = f"{row['Participacao']:.2f}%"
                _acum_str = f"{row['CumulativaPerc']:.2f}%"
                _cli_str  = str(int(row['Clientes']))
                _mix_vals = [_prod_str, _fat_str, _part_str, _acum_str, _cli_str]
                _mix_al   = ['L', 'R', 'C', 'C', 'C']
                for val, w, al in zip(_mix_vals, _pv_mix_pdf_w, _mix_al):
                    pdf.cell(w, 6, val, border=1, fill=True, align=al)
                pdf.ln()
                fill_mix = not fill_mix
            pdf.ln(6)

            # ── Inadimplência resumida ────────────────────────────────────
            if _pv_df_inad is not None and _pv_inad_vendedor > 0:
                pdf.set_fill_color(239, 68, 68)
                pdf.set_text_color(255, 255, 255)
                pdf.set_font('Helvetica', 'B', 9)
                pdf.cell(0, 7, '  INADIMPLÊNCIA', fill=True, border=0, ln=True)
                pdf.set_text_color(50, 50, 50)
                pdf.set_font('Helvetica', '', 8)
                pdf.set_fill_color(255, 240, 240)
                pdf.cell(95, 6, f'  Valor em Aberto: R$ {_pv_inad_vendedor:,.2f}', border='LB', fill=True)
                pdf.cell(90, 6, f'  % sobre Fat. Bruto: {_pv_perc_inad:.1f}%', border='RB', fill=True, ln=True)
                pdf.ln(4)

            # ── Rodapé do relatório ───────────────────────────────────────
            pdf.set_font('Helvetica', 'I', 7)
            pdf.set_text_color(173, 181, 189)
            pdf.multi_cell(0, 5,
                'Este relatório é gerado automaticamente com base nos dados filtrados do sistema BI Medtextil. '
                'As informações refletem o período e os filtros selecionados no momento da geração.')

            _pv_pdf_bytes = pdf.output()
            st.download_button(
                label="⬇️ Baixar PDF",
                data=bytes(_pv_pdf_bytes),
                file_name=f"performance_vendedores_{_pv_now.strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf",
                key="pv_dl_pdf_btn"
            )
            st.success("✅ PDF gerado com sucesso! Clique em 'Baixar PDF' para salvar.")

        except ImportError:
            # Fallback ReportLab
            try:
                from reportlab.lib.pagesizes import A4
                from reportlab.lib import colors
                from reportlab.lib.units import mm
                from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.enums import TA_CENTER, TA_LEFT

                _pv_buf = io.BytesIO()
                _pv_doc = SimpleDocTemplate(_pv_buf, pagesize=A4, rightMargin=15*mm, leftMargin=15*mm, topMargin=20*mm, bottomMargin=15*mm)
                _pv_styles = getSampleStyleSheet()
                _azul = colors.HexColor('#1F4788')
                _elements_rl = []

                _st_titulo = ParagraphStyle('T', parent=_pv_styles['Heading1'], fontSize=14, textColor=_azul, alignment=TA_CENTER)
                _st_normal = ParagraphStyle('N', parent=_pv_styles['Normal'], fontSize=8)

                _elements_rl.append(Paragraph('RELATÓRIO DE PERFORMANCE DE VENDEDORES', _st_titulo))
                _elements_rl.append(Paragraph(f"Vendedor: {_pv_vendedor}  |  Região: {_pv_regiao}  |  Período: {_pv_periodo}", _st_normal))
                _elements_rl.append(Spacer(1, 5*mm))

                # KPI Table
                _kpi_data = [['Indicador', 'Valor']]
                for k, v in _pv_kpis_pdf:
                    _kpi_data.append([k, v])
                _kpi_table = Table(_kpi_data, colWidths=[100*mm, 80*mm])
                _kpi_table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0),(-1,0), _azul),
                    ('TEXTCOLOR', (0,0),(-1,0), colors.white),
                    ('FONTNAME', (0,0),(-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0),(-1,-1), 8),
                    ('ROWBACKGROUNDS', (0,1),(-1,-1), [colors.white, colors.HexColor('#F0F4FF')]),
                    ('BOX', (0,0),(-1,-1), 0.5, colors.grey),
                    ('INNERGRID', (0,0),(-1,-1), 0.3, colors.lightgrey),
                    ('LEFTPADDING', (0,0),(-1,-1), 5),
                ]))
                _elements_rl.append(_kpi_table)
                _pv_doc.build(_elements_rl)
                _pv_buf.seek(0)

                st.download_button(
                    "⬇️ Baixar PDF",
                    data=_pv_buf.getvalue(),
                    file_name=f"performance_vendedores_{_pv_now.strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf",
                    key="pv_dl_pdf_rl"
                )
                st.success("✅ PDF gerado com sucesso!")
            except Exception as _e_rl:
                st.error(f"❌ Erro ao gerar PDF: {_e_rl}")
        except Exception as _e_pdf:
            st.error(f"❌ Erro ao gerar PDF: {_e_pdf}")
            st.info("💡 Verifique se a biblioteca fpdf2 está instalada: pip install fpdf2")

# ====================== RANKINGS ======================
elif menu == "Rankings":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Rankings</h2>', unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["📊 Vendedores", "👥 Clientes"])
    
    with tab1:
        st.subheader("Ranking de Vendedores por Valor")
        
        ranking_vendedores = notas_unicas.groupby('Vendedor').agg({
            'Valor_Real': 'sum',
            'Numero_NF': 'count',
            'CPF_CNPJ': 'nunique'
        }).reset_index()
        ranking_vendedores.columns = ['Vendedor', 'Valor Total', 'Qtd Notas', 'Qtd Clientes']
        ranking_vendedores = ranking_vendedores.sort_values('Valor Total', ascending=False)
        ranking_vendedores.insert(0, 'Posição', range(1, len(ranking_vendedores) + 1))
        
        fig_rank_vend = px.bar(
            ranking_vendedores.head(15),
            x='Vendedor',
            y='Valor Total',
            labels={'Vendedor': 'Vendedor', 'Valor Total': 'Valor Total (R$)'},
            color='Valor Total',
            color_discrete_sequence=['#163561'],
            title='Top 15 Vendedores por Valor'
        )
        fig_rank_vend = aplicar_layout_grafico(fig_rank_vend)
        st.plotly_chart(fig_rank_vend, use_container_width=True)
        
        # Formatar para exibição
        ranking_vendedores_display = formatar_dataframe_moeda(ranking_vendedores, ['Valor Total'])
        st.dataframe(ranking_vendedores_display, use_container_width=True)
        
        st.download_button(
            "📥 Exportar Ranking Vendedores",
            to_excel(ranking_vendedores),
            "ranking_vendedores.xlsx",
            "application/vnd.ms-excel"
        )
    
    with tab2:
        st.subheader("Ranking de Clientes por Valor")
        
        top_n = st.selectbox("Exibir Top:", [10, 20, 50, 100], key="top_clientes")
        
        ranking_clientes = notas_unicas.groupby(['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado']).agg({
            'Valor_Real': 'sum',
            'Numero_NF': 'count'
        }).reset_index()
        ranking_clientes.columns = ['CPF/CNPJ', 'Razão Social', 'Cidade', 'Estado', 'Valor Total', 'Qtd Notas']
        ranking_clientes = ranking_clientes.sort_values('Valor Total', ascending=False).head(top_n)
        ranking_clientes.insert(0, 'Posição', range(1, len(ranking_clientes) + 1))
        
        fig_rank_cli = px.bar(
            ranking_clientes.head(15),
            x='Valor Total',
            y='Razão Social',
            orientation='h',
            labels={'Razão Social': 'Cliente', 'Valor Total': 'Valor Total (R$)'},
            color='Valor Total',
            color_discrete_sequence=['#4A7BC8'],
            title=f'Top 15 Clientes por Valor'
        )
        fig_rank_cli = aplicar_layout_grafico(fig_rank_cli)
        st.plotly_chart(fig_rank_cli, use_container_width=True)
        
        # Formatar para exibição
        ranking_clientes_display = formatar_dataframe_moeda(ranking_clientes, ['Valor Total'])
        st.dataframe(ranking_clientes_display, use_container_width=True)
        
        st.download_button(
            "📥 Exportar Ranking Clientes",
            to_excel(ranking_clientes),
            f"ranking_top{top_n}_clientes.xlsx",
            "application/vnd.ms-excel"
        )


# ====================== CONSULTA CLIENTES ======================
elif menu == "Consulta Clientes":
    st.markdown('<h2 style="color:#4A7BC8;font-weight:700;margin-bottom:4px;font-size:1.35rem;">Consulta de Preços por Cliente</h2>', unsafe_allow_html=True)

    # ── Percentuais adicionais por estado ────────────────────────────────
    _PERC_ESTADO = {
        'AC': 6, 'RR': 6, 'RO': 6, 'AP': 6,
        'DF': 5, 'GO': 5,
        'MT': 5, 'MS': 5, 'TO': 5, 'AM': 5,
        'PA': 8,
        'RJ': 6, 'SP': 6, 'PR': 6,
        'RONDONIA': 20,
        'PR_DIRETA': 35,
    }
    _ESTADOS_OPCOES = [
        'Selecione o Estado',
        'AC (6%)', 'RR (6%)', 'RO (6%)', 'AP (6%)',
        'DF (5%)', 'GO (5%)',
        'MT (5%)', 'MS (5%)', 'TO (5%)', 'AM (5%)',
        'PA (8%)',
        'RJ (6%)', 'SP (6%)', 'PR (6%)',
        'RONDONIA (20%)',
        'PR - Venda Direta (35%)',
    ]
    _ESTADO_KEY_MAP = {
        'AC (6%)': ('AC', 6), 'RR (6%)': ('RR', 6), 'RO (6%)': ('RO', 6), 'AP (6%)': ('AP', 6),
        'DF (5%)': ('DF', 5), 'GO (5%)': ('GO', 5),
        'MT (5%)': ('MT', 5), 'MS (5%)': ('MS', 5), 'TO (5%)': ('TO', 5), 'AM (5%)': ('AM', 5),
        'PA (8%)': ('PA', 8),
        'RJ (6%)': ('RJ', 6), 'SP (6%)': ('SP', 6), 'PR (6%)': ('PR', 6),
        'RONDONIA (20%)': ('RONDONIA', 20),
        'PR - Venda Direta (35%)': ('PR_DIRETA', 35),
    }

    # ── Carregar tabela de preços ─────────────────────────────────────────
    _df_tabela = None
    
    # Tentar carregar produtos_agrupados primeiro (mais confiável)
    if planilhas_disponiveis.get('produtos_agrupados'):
        with st.spinner("Carregando catálogo de produtos..."):
            _df_tabela = carregar_planilha_github(planilhas_disponiveis['produtos_agrupados']['url'])
            if _df_tabela is not None:
                _df_tabela.columns = _df_tabela.columns.str.upper().str.strip()
                st.success("✅ Usando: Produtos Agrupados")
    
    # Se não conseguiu, tentar tabela_ne
    if _df_tabela is None and planilhas_disponiveis.get('tabela_ne'):
        with st.spinner("Carregando tabela NE..."):
            try:
                response = requests.get(planilhas_disponiveis['tabela_ne']['url'], timeout=15)
                content = io.BytesIO(response.content)
                
                # Tentar diferentes skiprows para encontrar o cabeçalho correto
                for skip in range(0, 10):
                    try:
                        df_test = pd.read_excel(content, skiprows=skip, nrows=5)
                        content.seek(0)  # Reset para próxima tentativa
                        
                        # Verificar se as colunas parecem ser cabeçalhos válidos
                        cols_str = [str(c).upper() for c in df_test.columns]
                        
                        # Se não tem UNNAMED, provavelmente achou o cabeçalho
                        if not any('UNNAMED' in c for c in cols_str):
                            # Verificar se tem colunas relevantes
                            has_code = any(x in ' '.join(cols_str) for x in ['COD', 'CODIGO', 'CÓDIGO'])
                            has_price = any(x in ' '.join(cols_str) for x in ['PRECO', 'PREÇO', 'VALOR', 'PRICE'])
                            
                            if has_code or has_price or len(cols_str) > 3:
                                # Parece ser o cabeçalho correto!
                                content.seek(0)
                                _df_tabela = pd.read_excel(content, skiprows=skip)
                                _df_tabela.columns = _df_tabela.columns.str.upper().str.strip()
                                _df_tabela = _df_tabela.dropna(how='all')
                                st.success(f"✅ Usando: Tabela NE (skiprows={skip})")
                                break
                    except:
                        continue
            except Exception as e:
                st.warning(f"Erro ao carregar tabela NE: {e}")

    if _df_tabela is None or len(_df_tabela) == 0:
        st.error("❌ Tabela de preços não encontrada")
        st.info("💡 Adicione 'Produtos_Agrupados_Completos_conciliados.xlsx' ou 'TABELA_NE_2026_CRM.xlsx' no GitHub")
        st.stop()

    # Verificar colunas disponíveis
    _cols = _df_tabela.columns.tolist()

    # Identificar coluna de código e preço (busca mais flexível)
    _cod_col   = next((c for c in _cols if any(x in c for x in ['ID_COD', 'CODIGO', 'CÓDIGO', 'COD', 'CÓD'])), None)
    _preco_col = next((c for c in _cols if any(x in c for x in ['PRECO', 'PREÇO', 'PRICE', 'VALOR', 'VLR'])), None)
    _desc_col  = next((c for c in _cols if any(x in c for x in ['DESCRI', 'DESCRIÇÃO', 'NOME', 'PRODUTO', 'GRUPO'])), None)

    if not _cod_col or not _preco_col:
        st.error(f"❌ Colunas necessárias não encontradas")
        st.info(f"📋 Colunas disponíveis: {_cols}")
        st.info(f"🔍 Procurando: Código (COD, CODIGO) e Preço (PRECO, PREÇO, VALOR)")
        with st.expander("🔍 Debug: Ver primeiras linhas da tabela"):
            st.dataframe(_df_tabela.head(10))
        st.stop()

    # ── Seleção de Estado ─────────────────────────────────────────────────
    st.markdown("#### Selecione o Estado")
    _estado_sel = st.selectbox(
        "Estado / Tabela",
        _ESTADOS_OPCOES,
        key="cc_estado",
        label_visibility="collapsed"
    )

    _perc_adicional = 0
    _estado_sigla   = ""
    if _estado_sel != 'Selecione o Estado':
        _estado_sigla, _perc_adicional = _ESTADO_KEY_MAP.get(_estado_sel, ('', 0))
        st.markdown(f"""
        <div style="background:#EEF3FC;border-radius:8px;padding:10px 14px;
                    margin-bottom:12px;font-size:0.85rem;color:#2C5AA0;">
            <b>Estado:</b> {_estado_sigla} &nbsp;·&nbsp;
            <b>Adicional:</b> {_perc_adicional}% &nbsp;·&nbsp;
            <b>Tabela base + {_perc_adicional}% = Tabela 3% comissão</b>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("#### Consulta de Produto")

    # ── Campo de código do produto ────────────────────────────────────────
    _codigos_lista = [''] + sorted(_df_tabela[_cod_col].dropna().astype(str).unique().tolist())
    _cc1, _cc2, _cc3 = st.columns([1, 2, 1])
    with _cc1:
        _cod_sel = st.selectbox("Código do Produto", _codigos_lista,
                                key="cc_codigo", label_visibility="visible")

    # ── Buscar produto e calcular preços ──────────────────────────────────
    _prod_row = None
    if _cod_sel:
        _match = _df_tabela[_df_tabela[_cod_col].astype(str) == str(_cod_sel)]
        if len(_match) > 0:
            _prod_row = _match.iloc[0]

    if _prod_row is not None:
        # Descrição
        _descricao = ""
        if _desc_col:
            _parts = []
            for _dc in [c for c in _cols if any(x in c for x in ['GRUPO','DESCRI','LINHA'])]:
                _v = str(_prod_row.get(_dc, '')).strip()
                if _v and _v.lower() not in ('nan', ''):
                    _parts.append(_v)
            _descricao = ' '.join(_parts) if _parts else str(_prod_row.get(_desc_col, ''))

        # Gramatura
        _gram_col = next((c for c in _cols if 'GRAMATUR' in c), None)
        _gramatura = ''
        if _gram_col:
            _gv = str(_prod_row.get(_gram_col, '')).strip()
            if _gv and _gv.lower() not in ('nan', '0', '0.0', ''):
                _gramatura = _gv

        with _cc2:
            st.text_input("Descrição", value=_descricao, disabled=True,
                          key=f"cc_desc_{_cod_sel}")

        with _cc3:
            st.text_input("Gramatura", value=_gramatura, disabled=True,
                          key=f"cc_gram_{_cod_sel}")

        # Preço base da tabela
        try:
            _preco_base = float(_prod_row.get(_preco_col, 0))
        except:
            _preco_base = 0.0

        # Tabela 3% comissão = preco_base * (1 + perc_adicional/100)
        _tab_3pct = _preco_base * (1 + _perc_adicional / 100) if _estado_sel != 'Selecione o Estado' else _preco_base
        # Tabela 4% comissão = tab_3pct * 1.06
        _tab_4pct = _tab_3pct * 1.06

        # Exibir preços calculados
        _pc1, _pc2, _pc3, _pc4 = st.columns(4)
        with _pc1:
            st.metric("Tabela Base", f"R$ {_preco_base:,.2f}",
                      help="Preço da tabela padrão sem adicional de estado")
        with _pc2:
            st.metric(f"Tabela 3% ({_perc_adicional}% estado)",
                      f"R$ {_tab_3pct:,.2f}",
                      help="Tabela base + percentual do estado = tabela comissão 3%")
        with _pc3:
            st.metric("Tabela 4%", f"R$ {_tab_4pct:,.2f}",
                      help="Tabela 3% + 6% = tabela comissão 4%")

        with _pc4:
            _val_neg = st.number_input("Valor Negociado (R$)",
                                       min_value=0.0,
                                       value=float(_tab_3pct),
                                       format="%.2f",
                                       key="cc_val_neg")

    # ── Calcular comissão sobre o valor negociado ─────────────────────
        # SOLUÇÃO FINAL: Comparação direta com margem de tolerância de 1 centavo
        if '_estado_sel' in locals() and _estado_sel:
            if _val_neg > 0 and _tab_3pct > 0:
                
                # Calculamos o valor exato que seria a tabela de 4% (3% + 6% de margem)
                # Aplicamos round para limpar o 7,77987 para 7,78
                _tabela_4_objetivo = round(_tab_3pct * 1.06, 2)
                _valor_digitado = round(_val_neg, 2)

                # FORÇAR A REGRA: Se o valor digitado for maior ou igual ao objetivo (com margem de 0.001)
                if _valor_digitado >= (_tabela_4_objetivo - 0.001):
                    _comissao_calc = '4%'
                    _variacao = round(((_valor_digitado - _tab_3pct) / _tab_3pct) * 100, 2)
                    _cor = "#10B981"; _msg = f"Comissão **4%** — objetivo de R$ {_tabela_4_objetivo:,.2f} atingido"
                else:
                    # Se não atingiu 4%, rodamos a função padrão para as outras faixas
                    _comissao_calc = calcular_comissao(_val_neg, _tab_3pct)
                    _variacao = round(((_val_neg - _tab_3pct) / _tab_3pct) * 100, 2)

                    if _comissao_calc == '3%':
                        _cor = "#2C5AA0"; _msg = "Comissão **3%** — valor igual ou acima da tabela do estado"
                    elif _comissao_calc == '2,5%':
                        _cor = "#F59E0B"; _msg = f"Comissão **2,5%** — valor {abs(_variacao):.1f}% abaixo (até 3%)"
                    elif _comissao_calc == '2%':
                        _cor = "#EF4444"; _msg = f"Comissão **2%** — valor {abs(_variacao):.1f}% abaixo (acima de 3%)"
                    else:
                        _cor = "#6B7280"; _msg = "Comissão não calculada"

                st.markdown(f"""
                <div style="background:{_cor}15;border-left:4px solid {_cor};
                            border-radius:8px;padding:12px 16px;margin-top:8px;">
                    <div style="font-size:1.1rem;font-weight:700;color:{_cor};">
                        Comissão: {_comissao_calc}
                    </div>
                    <div style="font-size:0.82rem;color:#6C757D;margin-top:3px;">{_msg}</div>
                    <div style="font-size:0.78rem;color:#ADB5BD;margin-top:4px;">
                        Valor negociado: R$ {_val_neg:,.2f} &nbsp;·&nbsp;
                        Tabela Estado (3%): R$ {round(_tab_3pct, 2):,.2f} &nbsp;·&nbsp;
                        Meta para 4%: R$ {_tabela_4_objetivo:,.2f}
                    </div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.info("Insira o valor negociado para calcular a comissão.")
        else:
            st.warning("Selecione um Estado para habilitar o cálculo de comissão.")
    else:
        if _cod_sel:
            st.warning(f"Produto {_cod_sel} não encontrado na tabela.")
        else:
            st.info("Selecione um código de produto para consultar os preços.")

st.markdown("""
<hr style="border-color:#E9ECEF;margin-top:32px;margin-bottom:12px;">
<div style="text-align:center;color:#ADB5BD;font-size:0.78rem;padding-bottom:16px;">
    Dashboard BI Medtextil 2.0 &nbsp;·&nbsp; Desenvolvido com Streamlit
    &nbsp;·&nbsp; <span style="color:#4A7BC8;font-weight:600;">Medtextil Produtos Textil Hospitalares</span>
</div>
""", unsafe_allow_html=True)
