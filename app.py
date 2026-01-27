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
        
        planilhas = []
        for content in contents:
            if content.name.endswith(('.xlsx', '.xls')):
                planilhas.append({
                    'nome': content.name,
                    'url': content.download_url,
                    'path': content.path
                })
        
        if not planilhas:
            st.warning(f"⚠️ Nenhuma planilha Excel encontrada na pasta '{GITHUB_FOLDER}'")
        
        return planilhas
    except Exception as e:
        st.error(f"❌ Erro ao conectar ao GitHub: {str(e)}")
        st.info(f"💡 Verificando: {GITHUB_REPO}/{GITHUB_FOLDER}")
        return []

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
        if st.session_state["password"] == SENHA_CORRETA:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### 🔐 Login - Dashboard BI Medtextil")
        st.text_input("Senha", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### 🔐 Login - Dashboard BI Medtextil")
        st.text_input("Senha", type="password", on_change=password_entered, key="password")
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

# ====================== INÍCIO DO APP ======================
if not check_password():
    st.stop()

st.title("📊 Dashboard BI Medtextil 2.0")
st.markdown("---")

col_header1, col_header2 = st.columns([3, 1])

with col_header1:
    with st.spinner("🔄 Conectando ao GitHub..."):
        planilhas_disponiveis = listar_planilhas_github()
    
    if planilhas_disponiveis:
        planilha_selecionada = st.selectbox(
            f"📁 Selecione a planilha da pasta '{GITHUB_FOLDER}'",
            options=[p['nome'] for p in planilhas_disponiveis],
            index=0
        )
        url_planilha = next(p['url'] for p in planilhas_disponiveis if p['nome'] == planilha_selecionada)
    else:
        st.error(f"❌ Não foi possível carregar planilhas da pasta '{GITHUB_FOLDER}'")
        st.info("💡 Verifique se:")
        st.info(f"  • O repositório '{GITHUB_REPO}' existe e é público")
        st.info(f"  • A pasta '{GITHUB_FOLDER}' existe no repositório")
        st.info(f"  • Há arquivos .xlsx ou .xls dentro da pasta '{GITHUB_FOLDER}'")
        st.stop()

with col_header2:
    if st.button("🔄 Recarregar Dados", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

with st.spinner("📥 Carregando dados do GitHub..."):
    df = carregar_planilha_github(url_planilha)

if df is None:
    st.error("❌ Não foi possível carregar os dados")
    st.stop()

df = processar_dados(df)
st.success(f"✅ Planilha carregada: **{planilha_selecionada}** ({len(df):,} registros)")

# ====================== SIDEBAR - FILTROS GLOBAIS ======================
st.sidebar.header("🔍 Filtros Globais")

col1, col2 = st.sidebar.columns(2)
with col1:
    data_inicial = st.date_input("Data Inicial", value=None, key="data_ini")
with col2:
    data_final = st.date_input("Data Final", value=None, key="data_fim")

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
    ["Dashboard", "Positivação", "Clientes sem Compra", "Histórico", "Rankings"],
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
        
        st.dataframe(relatorio_positivacao, use_container_width=True)
        
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
            
            st.dataframe(clientes_vendedor, use_container_width=True)
            
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
        
        st.dataframe(relatorio_estado, use_container_width=True)
        
        st.download_button(
            "📥 Exportar Positivação por Estado",
            to_excel(relatorio_estado),
            "positivacao_estado.xlsx",
            "application/vnd.ms-excel"
        )

# ====================== CLIENTES SEM COMPRA ======================
elif menu == "Clientes sem Compra":
    st.header("⚠️ Clientes sem Compra no Período (Churn)")
    
    col_f1, col_f2, col_f3 = st.columns(3)
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
            ["Valor Histórico (Maior)", "Valor Histórico (Menor)", "Nome (A-Z)"],
            key="ordem_churn"
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
    
    if ordem == "Valor Histórico (Maior)":
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=False)
    elif ordem == "Valor Histórico (Menor)":
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=True)
    else:
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
    
    st.dataframe(clientes_sem_compra, use_container_width=True, height=400)
    
    st.download_button(
        "📥 Exportar Clientes sem Compra",
        to_excel(clientes_sem_compra),
        "clientes_sem_compra.xlsx",
        "application/vnd.ms-excel"
    )

# ====================== HISTÓRICO ======================
elif menu == "Histórico":
    st.header("📜 Histórico de Vendas por Cliente")
    
    # Buscar cliente por CPF/CNPJ ou Nome
    col_busca1, col_busca2 = st.columns(2)
    
    with col_busca1:
        busca_tipo = st.radio("Buscar por:", ["Nome", "CPF/CNPJ"], horizontal=True)
    
    with col_busca2:
        if busca_tipo == "Nome":
            busca_texto = st.text_input("Digite o nome do cliente", placeholder="Ex: Nome da Empresa")
        else:
            busca_texto = st.text_input("Digite o CPF/CNPJ", placeholder="Ex: 12345678901234")
    
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
        
        st.dataframe(ranking_vendedores, use_container_width=True)
        
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
        
        st.dataframe(ranking_clientes, use_container_width=True)
        
        st.download_button(
            "📥 Exportar Ranking Clientes",
            to_excel(ranking_clientes),
            f"ranking_top{top_n}_clientes.xlsx",
            "application/vnd.ms-excel"
        )

st.markdown("---")
st.caption("Dashboard BI Medtextil 2.0 | Desenvolvido com Streamlit 🚀")
