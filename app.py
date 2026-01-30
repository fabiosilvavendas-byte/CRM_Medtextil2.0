import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import requests
from github import Github

# Configuração da página
st.set_page_config(
    page_title="Dashboard BI Medtextil", 
    layout="wide", 
    initial_sidebar_state="expanded",
    page_icon="📊"
)

# ====================== ÍCONE PERSONALIZADO PARA PWA/IPHONE ======================
# IMPORTANTE: Substitua a URL abaixo pelo link DIRETO da sua imagem
# Para pegar o link correto do Imgur:
# 1. Abra a imagem
# 2. Clique com botão direito na imagem
# 3. "Copiar endereço da imagem"
# 4. Deve ser algo como: https://i.imgur.com/XXXXX.png

LOGO_URL = "https://i.imgur.com/XXXXX.png"  # ⬅️ COLE AQUI O LINK DIRETO DA SUA LOGO

st.markdown(f"""
    <link rel="apple-touch-icon" sizes="180x180" href="{LOGO_URL}">
    <link rel="apple-touch-icon" sizes="152x152" href="{LOGO_URL}">
    <link rel="apple-touch-icon" sizes="120x120" href="{LOGO_URL}">
    <link rel="icon" type="image/png" sizes="192x192" href="{LOGO_URL}">
    <link rel="icon" type="image/png" sizes="32x32" href="{LOGO_URL}">
    <link rel="icon" type="image/png" sizes="16x16" href="{LOGO_URL}">
    <meta name="apple-mobile-web-app-title" content="Medtextil BI">
    <meta name="application-name" content="Medtextil BI">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
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
        
        if not planilhas['todas']:
            st.warning(f"⚠️ Nenhuma planilha Excel encontrada na pasta '{GITHUB_FOLDER}'")
        
        return planilhas
    except Exception as e:
        st.error(f"❌ Erro ao conectar ao GitHub: {str(e)}")
        st.info(f"💡 Verificando: {GITHUB_REPO}/{GITHUB_FOLDER}")
        return {'vendas': None, 'inadimplencia': None, 'vendas_produto': None, 'produtos_agrupados': None, 'todas': []}

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
    """Sistema de autenticação - ALTERE A SENHA AQUI"""
    
    # 🔐 ALTERE A SENHA AQUI (linha abaixo)
    SENHA_CORRETA = "admin123"  # ⬅️ MUDE AQUI PARA SUA SENHA
    
    def password_entered():
        if st.session_state.get("password_input", "") == SENHA_CORRETA:
            st.session_state["password_correct"] = True
        else:
            st.session_state["password_correct"] = False
            st.session_state["show_error"] = True

    if "password_correct" not in st.session_state:
        st.markdown("### 🔐 Login - Dashboard BI Medtextil")
        st.text_input("Senha", type="password", key="password_input")
        if st.button("🔓 Entrar", use_container_width=True):
            password_entered()
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔐 Login - Dashboard BI Medtextil")
        st.text_input("Senha", type="password", key="password_input")
        if st.button("🔓 Entrar", use_container_width=True):
            password_entered()
        if st.session_state.get("show_error", False):
            st.error("😕 Senha incorreta")
        return False
    else:
        return True

# ====================== PROCESSAMENTO DE DADOS ======================
@st.cache_data
def processar_dados(df):
    """Aplica as regras de negócio nos dados"""
    df['Valor_Real'] = df.apply(
        lambda row: row['TotalProduto'] if row['TipoMov'] == 'NF Venda' else -row['TotalProduto'],
        axis=1
    )
    df['DataEmissao'] = pd.to_datetime(df['DataEmissao'], errors='coerce')
    df['Mes'] = df['DataEmissao'].dt.month
    df['Ano'] = df['DataEmissao'].dt.year
    df['MesAno'] = df['DataEmissao'].dt.to_period('M').astype(str)
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
    df = df.rename(columns={
        'Funcionário': 'Vendedor',
        'Razão Social': 'Cliente',
        'N_Doc': 'NumeroDoc',
        'Dt.Vencimento': 'DataVencimento',
        'Vr.Líquido': 'ValorLiquido',
        'Conta/Caixa': 'Banco',
        'UF': 'Estado'
    })
    
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

# ====================== INÍCIO DO APP ======================
if not check_password():
    st.stop()

st.title("📊 Dashboard BI Medtextil 2.0")
st.markdown("---")

col_header1, col_header2 = st.columns([3, 1])

with col_header1:
    with st.spinner("🔄 Conectando ao GitHub..."):
        planilhas_disponiveis = listar_planilhas_github()
    
    if planilhas_disponiveis['vendas']:
        st.info(f"📊 Planilha de Vendas: **{planilhas_disponiveis['vendas']['nome']}**")
        url_planilha_vendas = planilhas_disponiveis['vendas']['url']
    else:
        st.error("❌ Planilha de vendas não encontrada")
        st.info("💡 Procurando por arquivo com 'CONSULTA_VENDEDORES' no nome")
        st.stop()
    
    if planilhas_disponiveis['inadimplencia']:
        st.info(f"💳 Planilha de Inadimplência: **{planilhas_disponiveis['inadimplencia']['nome']}**")
    else:
        st.warning("⚠️ Planilha de inadimplência não encontrada (módulo desabilitado)")

with col_header2:
    if st.button("🔄 Recarregar Dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

with st.spinner("📥 Carregando dados de vendas..."):
    df = carregar_planilha_github(url_planilha_vendas)

if df is None:
    st.error("❌ Não foi possível carregar os dados de vendas")
    st.stop()

df = processar_dados(df)
st.success(f"✅ Dados de vendas carregados: ({len(df):,} registros)")

# ====================== SIDEBAR - FILTROS GLOBAIS ======================
st.sidebar.header("🔍 Filtros Globais")

col1, col2 = st.sidebar.columns(2)
with col1:
    data_inicial = st.date_input(
        "Data Inicial", 
        value=None, 
        key="data_ini",
        format="DD/MM/YYYY"
    )
with col2:
    data_final = st.date_input(
        "Data Final", 
        value=None, 
        key="data_fim",
        format="DD/MM/YYYY"
    )

vendedores = ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist())
vendedor_filtro = st.sidebar.selectbox("Vendedor", vendedores, key="vend_global")

estados = ['Todos'] + sorted(df['Estado'].dropna().unique().tolist())
estado_filtro = st.sidebar.selectbox("Estado", estados, key="est_global")

col3, col4 = st.sidebar.columns(2)
with col3:
    meses_opcoes = ['Todos'] + list(range(1, 13))
    mes_filtro = st.selectbox("Mês", meses_opcoes, key="mes_global")
with col4:
    anos_opcoes = ['Todos'] + sorted(df['Ano'].dropna().unique().tolist(), reverse=True)
    ano_filtro = st.selectbox("Ano", anos_opcoes, key="ano_global")

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
menu = st.sidebar.radio(
    "📑 Navegação",
    ["Dashboard", "Positivação", "Inadimplência", "Clientes sem Compra", "Histórico", "Preço Médio", "Rankings"],
    index=0
)

# ====================== DASHBOARD ======================
if menu == "Dashboard":
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        # VENDAS BRUTAS = SOMASE(TipoMov="NF Venda", TotalProduto)
        vendas_brutas = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda']['TotalProduto'].sum()
        st.metric("💰 Faturamento Bruto", f"R$ {vendas_brutas:,.2f}")
    
    with col2:
        # FATURAMENTO LÍQUIDO = SOMA(Valor_Real) 
        # Valor_Real já negativiza as devoluções automaticamente
        faturamento_liquido = notas_unicas['Valor_Real'].sum()
        st.metric("💵 Faturamento Líquido", f"R$ {faturamento_liquido:,.2f}")
    
    with col3:
        clientes_unicos = df_filtrado['CPF_CNPJ'].nunique()
        st.metric("👥 Clientes Únicos", f"{clientes_unicos:,}")
    
    with col4:
        total_notas = len(notas_unicas[notas_unicas['TipoMov'] == 'NF Venda'])
        st.metric("📄 Notas de Venda", f"{total_notas:,}")
    
    # Segunda linha de métricas - Detalhamento
    col1b, col2b, col3b, col4b = st.columns(4)
    
    with col1b:
        # DEVOLUÇÕES = SOMASE(TipoMov="NF Dev.Venda", TotalProduto)
        total_devolucoes = notas_unicas[notas_unicas['TipoMov'] == 'NF Dev.Venda']['TotalProduto'].sum()
        st.metric("↩️ Devoluções", f"R$ {total_devolucoes:,.2f}")
    
    with col2b:
        ticket_medio = vendas_brutas / clientes_unicos if clientes_unicos > 0 else 0
        st.metric("🎯 Ticket Médio", f"R$ {ticket_medio:,.2f}")
    
    with col3b:
        qtd_notas_dev = len(notas_unicas[notas_unicas['TipoMov'] == 'NF Dev.Venda'])
        st.metric("📋 Notas Devolução", f"{qtd_notas_dev:,}")
    
    with col4b:
        taxa_devolucao = (total_devolucoes / vendas_brutas * 100) if vendas_brutas > 0 else 0
        st.metric("📊 Taxa Devolução", f"{taxa_devolucao:.1f}%")
    
    st.markdown("---")
    
    col5, col6 = st.columns(2)
    
    with col5:
        st.subheader("📈 Evolução de Vendas Brutas")
        # Filtra apenas vendas (sem devoluções) para o gráfico
        vendas_apenas = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda']
        vendas_tempo = vendas_apenas.groupby('MesAno')['TotalProduto'].sum().reset_index()
        vendas_tempo = vendas_tempo.sort_values('MesAno')
        
        if len(vendas_tempo) > 0:
            fig_linha = px.line(
                vendas_tempo, 
                x='MesAno', 
                y='TotalProduto',
                labels={'MesAno': 'Período', 'TotalProduto': 'Valor (R$)'},
                template='plotly_white'
            )
            fig_linha.update_traces(line_color='#1f77b4', line_width=3)
            fig_linha.update_layout(
                xaxis_title="Período",
                yaxis_title="Valor (R$)",
                hovermode='x unified'
            )
            st.plotly_chart(fig_linha, use_container_width=True)
        else:
            st.info("Sem dados para exibir no período selecionado")
    
    with col6:
        st.subheader("🗺️ Top 10 Estados")
        # Filtra apenas vendas (sem devoluções)
        vendas_estado_apenas = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda']
        vendas_estado = vendas_estado_apenas.groupby('Estado')['TotalProduto'].sum().reset_index()
        vendas_estado = vendas_estado.sort_values('TotalProduto', ascending=False).head(10)
        
        fig_bar = px.bar(
            vendas_estado, 
            x='Estado', 
            y='TotalProduto',
            labels={'Estado': 'Estado', 'TotalProduto': 'Valor (R$)'},
            template='plotly_white',
            color='TotalProduto',
            color_continuous_scale='Greens'
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    
    st.markdown("---")
    
    col7, col8 = st.columns(2)
    
    with col7:
        st.subheader("👥 Positivação por Vendedor")
        vendas_periodo = df_filtrado[df_filtrado['TipoMov'] == 'NF Venda']
        atendidos = vendas_periodo.groupby('Vendedor')['CPF_CNPJ'].nunique().reset_index()
        atendidos.columns = ['Vendedor', 'Clientes']
        atendidos = atendidos.sort_values('Clientes', ascending=False).head(10)
        
        fig_posit = px.bar(
            atendidos,
            x='Vendedor',
            y='Clientes',
            labels={'Vendedor': 'Vendedor', 'Clientes': 'Clientes Atendidos'},
            template='plotly_white',
            color='Clientes',
            color_continuous_scale='Blues'
        )
        st.plotly_chart(fig_posit, use_container_width=True)
    
    with col8:
        st.subheader("🏆 Top 10 Clientes")
        # Filtra apenas vendas
        vendas_clientes = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda']
        ranking_clientes = vendas_clientes.groupby('RazaoSocial')['TotalProduto'].sum().reset_index()
        ranking_clientes = ranking_clientes.sort_values('TotalProduto', ascending=False).head(10)
        
        fig_clientes = px.bar(
            ranking_clientes,
            x='TotalProduto',
            y='RazaoSocial',
            orientation='h',
            labels={'RazaoSocial': 'Cliente', 'TotalProduto': 'Valor (R$)'},
            template='plotly_white',
            color='TotalProduto',
            color_continuous_scale='Oranges'
        )
        st.plotly_chart(fig_clientes, use_container_width=True)
    
    st.markdown("---")
    
    col9, col10 = st.columns(2)
    
    with col9:
        st.subheader("⚠️ Clientes sem Compra (Top 10)")
        clientes_com_venda = set(df_filtrado[df_filtrado['TipoMov'] == 'NF Venda']['CPF_CNPJ'].unique())
        todos_clientes = df.sort_values('DataEmissao').groupby('CPF_CNPJ').last().reset_index()
        valor_historico = df[df['TipoMov'] == 'NF Venda'].groupby('CPF_CNPJ')['TotalProduto'].sum().reset_index()
        valor_historico.columns = ['CPF_CNPJ', 'ValorHistorico']
        
        todos_clientes = pd.merge(todos_clientes, valor_historico, on='CPF_CNPJ', how='left')
        todos_clientes['ValorHistorico'] = todos_clientes['ValorHistorico'].fillna(0)
        
        clientes_sem_compra = todos_clientes[~todos_clientes['CPF_CNPJ'].isin(clientes_com_venda)]
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=False).head(10)
        
        fig_churn = px.bar(
            clientes_sem_compra,
            x='ValorHistorico',
            y='RazaoSocial',
            orientation='h',
            labels={'RazaoSocial': 'Cliente', 'ValorHistorico': 'Valor Histórico (R$)'},
            template='plotly_white',
            color='ValorHistorico',
            color_continuous_scale='Reds'
        )
        st.plotly_chart(fig_churn, use_container_width=True)
    
    with col10:
        st.subheader("📊 Ranking de Vendedores")
        # Filtra apenas vendas
        vendas_vendedores = notas_unicas[notas_unicas['TipoMov'] == 'NF Venda']
        ranking_vendedores = vendas_vendedores.groupby('Vendedor')['TotalProduto'].sum().reset_index()
        ranking_vendedores = ranking_vendedores.sort_values('TotalProduto', ascending=False).head(10)
        
        fig_rank_vend = px.bar(
            ranking_vendedores,
            x='TotalProduto',
            y='Vendedor',
            orientation='h',
            labels={'Vendedor': 'Vendedor', 'TotalProduto': 'Valor Total (R$)'},
            template='plotly_white',
            color='TotalProduto',
            color_continuous_scale='Purples'
        )
        st.plotly_chart(fig_rank_vend, use_container_width=True)

# ====================== POSITIVAÇÃO ======================
elif menu == "Positivação":
    st.header("👥 Relatório de Positivação")
    
    tab1, tab2 = st.tabs(["📊 Por Vendedor", "🗺️ Por Estado"])
    
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
            template='plotly_white',
            color='Percentual',
            color_continuous_scale='Blues',
            title='Top 15 Vendedores - Taxa de Positivação'
        )
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
            template='plotly_white',
            color='Percentual',
            color_continuous_scale='Greens',
            title='Top 15 Estados - Taxa de Positivação'
        )
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

# ====================== INADIMPLÊNCIA ======================
elif menu == "Inadimplência":
    st.header("💳 Relatório de Inadimplência")
    
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
        
        st.success(f"✅ Dados carregados: {len(df_inadimplencia):,} títulos a receber")
        
        # ========== FILTROS ==========
        st.subheader("🔍 Filtros")
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        
        with col_f1:
            vendedores_inad = ['Todos'] + sorted(df_inadimplencia['Vendedor'].dropna().unique().tolist())
            vendedor_inad_filtro = st.selectbox("Vendedor", vendedores_inad, key="vend_inad")
        
        with col_f2:
            estados_inad = ['Todos'] + sorted(df_inadimplencia['Estado'].dropna().unique().tolist())
            estado_inad_filtro = st.selectbox("Estado", estados_inad, key="est_inad")
        
        with col_f3:
            data_inicial_inad = st.date_input(
                "Vencimento De", 
                value=None, 
                key="data_ini_inad",
                format="DD/MM/YYYY"
            )
        
        with col_f4:
            data_final_inad = st.date_input(
                "Vencimento Até", 
                value=None, 
                key="data_fim_inad",
                format="DD/MM/YYYY"
            )
        
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
            st.metric("💰 Total em Aberto", f"R$ {total_inadimplencia:,.2f}")
        
        with col2:
            qtd_titulos = len(df_inad_filtrado)
            st.metric("📄 Quantidade de Títulos", f"{qtd_titulos:,}")
        
        with col3:
            clientes_inadimplentes = df_inad_filtrado['Cliente'].nunique()
            st.metric("👥 Clientes Inadimplentes", f"{clientes_inadimplentes:,}")
        
        with col4:
            atraso_medio = df_inad_filtrado['DiasAtraso'].mean()
            st.metric("📅 Atraso Médio", f"{atraso_medio:.0f} dias")
        
        st.markdown("---")
        
        # ========== GRÁFICOS ==========
        col5, col6 = st.columns(2)
        
        with col5:
            st.subheader("📊 Inadimplência por Faixa de Atraso")
            
            # Ordenar faixas corretamente
            ordem_faixas = ['A Vencer', '1-30 dias', '31-60 dias', '61-90 dias', 'Acima de 90 dias']
            inad_por_faixa = df_inad_filtrado.groupby('FaixaAtraso')['ValorLiquido'].sum().reset_index()
            inad_por_faixa['FaixaAtraso'] = pd.Categorical(inad_por_faixa['FaixaAtraso'], categories=ordem_faixas, ordered=True)
            inad_por_faixa = inad_por_faixa.sort_values('FaixaAtraso')
            
            fig_faixa = px.bar(
                inad_por_faixa,
                x='FaixaAtraso',
                y='ValorLiquido',
                labels={'FaixaAtraso': 'Faixa de Atraso', 'ValorLiquido': 'Valor (R$)'},
                template='plotly_white',
                color='ValorLiquido',
                color_continuous_scale='Reds'
            )
            st.plotly_chart(fig_faixa, use_container_width=True)
        
        with col6:
            st.subheader("🏦 Inadimplência por Banco")
            inad_por_banco = df_inad_filtrado.groupby('Banco')['ValorLiquido'].sum().reset_index()
            inad_por_banco = inad_por_banco.sort_values('ValorLiquido', ascending=False).head(10)
            
            fig_banco = px.bar(
                inad_por_banco,
                x='ValorLiquido',
                y='Banco',
                orientation='h',
                labels={'Banco': 'Banco', 'ValorLiquido': 'Valor (R$)'},
                template='plotly_white',
                color='ValorLiquido',
                color_continuous_scale='Blues'
            )
            st.plotly_chart(fig_banco, use_container_width=True)
        
        st.markdown("---")
        
        col7, col8 = st.columns(2)
        
        with col7:
            st.subheader("👤 Top 10 Vendedores - Inadimplência")
            inad_por_vendedor = df_inad_filtrado.groupby('Vendedor').agg({
                'ValorLiquido': 'sum',
                'NumeroDoc': 'count'
            }).reset_index()
            inad_por_vendedor.columns = ['Vendedor', 'Valor', 'QtdTitulos']
            inad_por_vendedor = inad_por_vendedor.sort_values('Valor', ascending=False).head(10)
            
            fig_vend_inad = px.bar(
                inad_por_vendedor,
                x='Vendedor',
                y='Valor',
                labels={'Vendedor': 'Vendedor', 'Valor': 'Valor (R$)'},
                template='plotly_white',
                color='Valor',
                color_continuous_scale='Oranges'
            )
            st.plotly_chart(fig_vend_inad, use_container_width=True)
        
        with col8:
            st.subheader("🗺️ Top 10 Estados - Inadimplência")
            inad_por_estado = df_inad_filtrado.groupby('Estado')['ValorLiquido'].sum().reset_index()
            inad_por_estado = inad_por_estado.sort_values('ValorLiquido', ascending=False).head(10)
            
            fig_est_inad = px.bar(
                inad_por_estado,
                x='Estado',
                y='ValorLiquido',
                labels={'Estado': 'Estado', 'ValorLiquido': 'Valor (R$)'},
                template='plotly_white',
                color='ValorLiquido',
                color_continuous_scale='Purples'
            )
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
        st.download_button(
            "📥 Exportar Relatório Completo",
            to_excel(df_detalhado),
            "relatorio_inadimplencia.xlsx",
            "application/vnd.ms-excel"
        )

# ====================== CLIENTES SEM COMPRA ======================
elif menu == "Clientes sem Compra":
    st.header("⚠️ Clientes sem Compra no Período (Churn)")
    
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
    
    clientes_com_venda = set(df_filtrado[df_filtrado['TipoMov'] == 'NF Venda']['CPF_CNPJ'].unique())
    todos_clientes = df.sort_values('DataEmissao').groupby('CPF_CNPJ').last().reset_index()
    valor_historico = df[df['TipoMov'] == 'NF Venda'].groupby('CPF_CNPJ')['TotalProduto'].sum().reset_index()
    valor_historico.columns = ['CPF_CNPJ', 'ValorHistorico']
    
    todos_clientes = pd.merge(todos_clientes, valor_historico, on='CPF_CNPJ', how='left')
    todos_clientes['ValorHistorico'] = todos_clientes['ValorHistorico'].fillna(0)
    
    clientes_sem_compra = todos_clientes[~todos_clientes['CPF_CNPJ'].isin(clientes_com_venda)]
    clientes_sem_compra = clientes_sem_compra[['RazaoSocial', 'CPF_CNPJ', 'Vendedor', 'Cidade', 'Estado', 'ValorHistorico']]
    
    if vendedor_churn_filtro != 'Todos':
        clientes_sem_compra = clientes_sem_compra[clientes_sem_compra['Vendedor'] == vendedor_churn_filtro]
    if estado_churn_filtro != 'Todos':
        clientes_sem_compra = clientes_sem_compra[clientes_sem_compra['Estado'] == estado_churn_filtro]
    
    # Filtro de busca por nome do cliente
    if busca_cliente_churn and len(busca_cliente_churn) >= 2:
        clientes_sem_compra = clientes_sem_compra[
            clientes_sem_compra['RazaoSocial'].str.contains(busca_cliente_churn, case=False, na=False)
        ]
    
    if ordem == "Valor Histórico (Maior)":
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=False)
    elif ordem == "Valor Histórico (Menor)":
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=True)
    elif ordem == "Nome (A-Z)":
        clientes_sem_compra = clientes_sem_compra.sort_values('RazaoSocial')
    
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
            template='plotly_white',
            color='ValorHistorico',
            color_continuous_scale='Reds',
            title='Top 15 Clientes sem Compra por Valor Histórico'
        )
        st.plotly_chart(fig_churn, use_container_width=True)
    
    # Formatar para exibição
    clientes_sem_compra_display = formatar_dataframe_moeda(clientes_sem_compra, ['ValorHistorico'])
    
    # Renomear colunas para exibição
    clientes_sem_compra_display = clientes_sem_compra_display.rename(columns={
        'RazaoSocial': 'Razão Social',
        'CPF_CNPJ': 'CPF/CNPJ',
        'Vendedor': 'Vendedor',
        'Cidade': 'Cidade',
        'Estado': 'Estado',
        'ValorHistorico': 'Valor Histórico'
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
    st.header("📜 Histórico de Vendas")
    
    tab1, tab2 = st.tabs(["👤 Por Cliente", "🧑‍💼 Por Vendedor"])
    
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
                        template='plotly_white',
                        title='Evolução de Compras'
                    )
                    fig_hist.update_traces(line_color='#2ECC71', line_width=3)
                    st.plotly_chart(fig_hist, use_container_width=True)
                
                st.markdown("---")
                
                st.subheader("📋 Detalhamento de Produtos")
                
                historico_display = historico[['DataEmissao', 'TipoMov', 'Numero_NF', 'CodigoProduto', 'NomeProduto', 'Quantidade', 'PrecoUnit', 'TotalProduto']].copy()
                historico_display['DataEmissao'] = historico_display['DataEmissao'].dt.strftime('%d/%m/%Y')
                
                # Formatar valores monetários
                historico_display['PrecoUnit'] = historico_display['PrecoUnit'].apply(lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00")
                historico_display['TotalProduto'] = historico_display['TotalProduto'].apply(lambda x: formatar_moeda(x) if pd.notnull(x) else "R$ 0,00")
                
                historico_display = historico_display.rename(columns={
                    'DataEmissao': 'Data',
                    'TipoMov': 'Tipo',
                    'Numero_NF': 'Nota Fiscal',
                    'CodigoProduto': 'Código',
                    'NomeProduto': 'Produto',
                    'Quantidade': 'Qtd',
                    'PrecoUnit': 'Preço Unit.',
                    'TotalProduto': 'Total'
                })
                
                st.dataframe(historico_display, use_container_width=True, height=400)
                
                st.download_button(
                    "📥 Exportar Histórico",
                    to_excel(historico),
                    f"historico_{cpf_cnpj}.xlsx",
                    "application/vnd.ms-excel"
                )
            else:
                st.warning("Nenhum registro encontrado para este cliente")
        else:
            st.info("👆 Digite pelo menos 3 caracteres para buscar um cliente")
    
    # ========== ABA: POR VENDEDOR ==========
    with tab2:
        st.subheader("Histórico de Vendas por Vendedor")
        
        # Filtros
        col_f1, col_f2, col_f3 = st.columns(3)
        
        with col_f1:
            vendedores_hist = ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist())
            vendedor_hist_filtro = st.selectbox("Vendedor", vendedores_hist, key="vend_hist")
        
        with col_f2:
            data_inicial_hist = st.date_input(
                "Data Inicial", 
                value=None, 
                key="data_ini_hist",
                format="DD/MM/YYYY"
            )
        
        with col_f3:
            data_final_hist = st.date_input(
                "Data Final", 
                value=None, 
                key="data_fim_hist",
                format="DD/MM/YYYY"
            )
        
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
            historico_vendedor = notas_vendedor[['DataEmissao', 'RazaoSocial', 'Numero_NF', 'TotalProduto', 'Vendedor']].copy()
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
            historico_vendedor_display = historico_vendedor_display.rename(columns={
                'DataEmissao': 'Data',
                'RazaoSocial': 'Cliente',
                'Numero_NF': 'Nota Fiscal',
                'TotalProduto': 'Valor Total',
                'Vendedor': 'Vendedor'
            })
            
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

# ====================== PREÇO MÉDIO ======================
elif menu == "Preço Médio":
    st.header("💰 Análise de Preço Médio por Produto")
    
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
    
    # Verificar se as colunas necessárias existem
    colunas_vendas_necessarias = ['CODPRODUTO', 'TOTQTD', 'PRECOUNITMEDIO', 'TOTLIQUIDO']
    colunas_produtos_necessarias = ['ID_COD', 'GRUPO', 'DESCRIÇÃO', 'LINHA', 'GRAMATURA']
    
    # Verificar colunas alternativas
    if 'DESCRIÇÃO' not in df_produtos.columns and 'DESCRICAO' in df_produtos.columns:
        df_produtos = df_produtos.rename(columns={'DESCRICAO': 'DESCRIÇÃO'})
    
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
    df_produtos['NOMEPRODUTO'] = (
        df_produtos['GRUPO'].fillna('').astype(str) + ' ' +
        df_produtos['DESCRIÇÃO'].fillna('').astype(str) + ' ' +
        df_produtos['LINHA'].fillna('').astype(str)
    ).str.strip()
    
    # Renomear ID_COD para CODPRODUTO para facilitar o merge
    df_produtos = df_produtos.rename(columns={'ID_COD': 'CODPRODUTO'})
    
    # Adicionar coluna DATA com o mês/ano atual (já que não existe na planilha)
    data_atual = pd.Timestamp.now()
    df_vendas_produto['DATA'] = data_atual
    df_vendas_produto['Mes'] = data_atual.month
    df_vendas_produto['Ano'] = data_atual.year
    df_vendas_produto['MesAno'] = data_atual.strftime('%Y-%m')
    
    # Fazer o merge (PROCV) entre as planilhas
    df_preco_medio = pd.merge(
        df_vendas_produto,
        df_produtos[['CODPRODUTO', 'NOMEPRODUTO', 'GRAMATURA']],
        on='CODPRODUTO',
        how='left'
    )
    
    # Preencher produtos não encontrados
    df_preco_medio['NOMEPRODUTO'] = df_preco_medio['NOMEPRODUTO'].fillna('Produto não catalogado')
    df_preco_medio['GRAMATURA'] = df_preco_medio['GRAMATURA'].fillna(0)
    
    st.success(f"✅ Dados carregados: {len(df_preco_medio):,} registros de vendas")
    st.info(f"📊 Planilha de Vendas: **{planilhas_disponiveis['vendas_produto']['nome']}**")
    st.info(f"📦 Planilha de Produtos: **{planilhas_disponiveis['produtos_agrupados']['nome']}**")
    st.info(f"📅 Período de Referência: **{data_atual.strftime('%B/%Y')}** (mês atual)")
    
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
        st.metric("💰 Total Vendido", f"R$ {total_vendido:,.2f}")
    
    with col2:
        qtd_total = df_preco_filtrado['TOTQTD'].sum()
        st.metric("📦 Qtd Total Vendida", f"{qtd_total:,.0f}")
    
    with col3:
        preco_medio_geral = df_preco_filtrado['PRECOUNITMEDIO'].mean() if len(df_preco_filtrado) > 0 else 0
        st.metric("💵 Preço Médio Geral", f"R$ {preco_medio_geral:,.2f}")
    
    with col4:
        produtos_unicos = df_preco_filtrado['CODPRODUTO'].nunique()
        st.metric("🏷️ Produtos Únicos", f"{produtos_unicos:,}")
    
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
            template='plotly_white',
            color='TOTLIQUIDO',
            color_continuous_scale='Blues'
        )
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
            template='plotly_white',
            color='TOTQTD',
            color_continuous_scale='Greens'
        )
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
            template='plotly_white',
            title=f'Evolução do Preço Médio - {produto_selecionado}'
        )
        fig_evolucao.update_traces(line_color='#FF6B6B', line_width=3)
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

# ====================== RANKINGS ======================
elif menu == "Rankings":
    st.header("🏆 Rankings")
    
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
            template='plotly_white',
            color='Valor Total',
            color_continuous_scale='Purples',
            title='Top 15 Vendedores por Valor'
        )
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
            template='plotly_white',
            color='Valor Total',
            color_continuous_scale='Oranges',
            title=f'Top 15 Clientes por Valor'
        )
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

st.markdown("---")
st.caption("Dashboard BI Medtextil 2.0 | Desenvolvido com Streamlit 🚀")
