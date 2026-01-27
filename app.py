import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io

# Configuração da página
st.set_page_config(page_title="Dashboard BI - Vendas", layout="wide", initial_sidebar_state="expanded")

# Função de autenticação
def check_password():
    def password_entered():
        if st.session_state["password"] == "admin123":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Senha", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Senha", type="password", on_change=password_entered, key="password")
        st.error("😕 Senha incorreta")
        return False
    else:
        return True

# Função para processar dados
@st.cache_data
def processar_dados(df):
    """Aplica as regras de negócio nos dados"""
    
    # Criar coluna Valor_Real (Regra Principal)
    df['Valor_Real'] = df.apply(
        lambda row: row['TotalProduto'] if row['TipoMov'] == 'NF Venda' else -row['TotalProduto'],
        axis=1
    )
    
    # Converter DataEmissao para datetime
    df['DataEmissao'] = pd.to_datetime(df['DataEmissao'], errors='coerce')
    df['Mes'] = df['DataEmissao'].dt.month
    df['Ano'] = df['DataEmissao'].dt.year
    df['MesAno'] = df['DataEmissao'].dt.to_period('M').astype(str)
    
    return df

# Função para obter notas únicas (sem duplicatas)
def obter_notas_unicas(df):
    """Remove duplicatas de Numero_NF mantendo apenas primeira ocorrência"""
    return df.drop_duplicates(subset=['Numero_NF'], keep='first')

# Função para download
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
    return output.getvalue()

# INÍCIO DO APP
if not check_password():
    st.stop()

# Título
st.title("📊 Dashboard BI - Análise de Vendas")

# Upload de arquivo
uploaded_file = st.file_uploader("📁 Carregar planilha Excel (CONSULTA VENDEDORES.xlsx)", type=['xlsx', 'xls'])

if uploaded_file is not None:
    # Carregar dados
    df = pd.read_excel(uploaded_file)
    df = processar_dados(df)
    
    st.success(f"✅ Arquivo carregado: {len(df)} registros")
    
    # SIDEBAR - FILTROS GLOBAIS
    st.sidebar.header("🔍 Filtros Globais")
    
    # Filtro de Data
    col1, col2 = st.sidebar.columns(2)
    with col1:
        data_inicial = st.date_input("Data Inicial", value=None)
    with col2:
        data_final = st.date_input("Data Final", value=None)
    
    # Filtro de Vendedor
    vendedores = ['Todos'] + sorted(df['Vendedor'].dropna().unique().tolist())
    vendedor_filtro = st.sidebar.selectbox("Vendedor", vendedores)
    
    # Filtro de Estado
    estados = ['Todos'] + sorted(df['Estado'].dropna().unique().tolist())
    estado_filtro = st.sidebar.selectbox("Estado", estados)
    
    # Filtro de Mês e Ano
    col3, col4 = st.sidebar.columns(2)
    with col3:
        mes_filtro = st.selectbox("Mês", ['Todos'] + list(range(1, 13)))
    with col4:
        ano_filtro = st.selectbox("Ano", ['Todos'] + sorted(df['Ano'].dropna().unique().tolist()))
    
    # Aplicar filtros
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
    
    # Obter notas únicas para cálculos
    notas_unicas = obter_notas_unicas(df_filtrado)
    
    # MENU DE NAVEGAÇÃO
    menu = st.sidebar.radio(
        "📑 Menu",
        ["Dashboard", "Positivação", "Clientes sem Compra", "Histórico", "Rankings"]
    )
    
    # ====================== DASHBOARD ======================
    if menu == "Dashboard":
        # KPIs
        col1, col2 = st.columns(2)
        
        with col1:
            faturamento_total = notas_unicas['Valor_Real'].sum()
            st.metric("💰 Faturamento Total Líquido", f"R$ {faturamento_total:,.2f}")
        
        with col2:
            clientes_unicos = df_filtrado['CPF_CNPJ'].nunique()
            st.metric("👥 Clientes Únicos", f"{clientes_unicos:,}")
        
        st.markdown("---")
        
        # Gráficos
        col3, col4 = st.columns(2)
        
        with col3:
            st.subheader("📈 Evolução de Vendas")
            vendas_tempo = notas_unicas.groupby('MesAno')['Valor_Real'].sum().reset_index()
            vendas_tempo = vendas_tempo.sort_values('MesAno')
            
            fig_linha = px.line(vendas_tempo, x='MesAno', y='Valor_Real',
                               labels={'MesAno': 'Período', 'Valor_Real': 'Valor (R$)'},
                               template='plotly_white')
            st.plotly_chart(fig_linha, use_container_width=True)
        
        with col4:
            st.subheader("🗺️ Top 10 Estados")
            vendas_estado = notas_unicas.groupby('Estado')['Valor_Real'].sum().reset_index()
            vendas_estado = vendas_estado.sort_values('Valor_Real', ascending=False).head(10)
            
            fig_bar = px.bar(vendas_estado, x='Estado', y='Valor_Real',
                            labels={'Estado': 'Estado', 'Valor_Real': 'Valor (R$)'},
                            template='plotly_white', color='Valor_Real',
                            color_continuous_scale='Greens')
            st.plotly_chart(fig_bar, use_container_width=True)
    
    # ====================== POSITIVAÇÃO ======================
    elif menu == "Positivação":
        st.header("👥 Relatório de Positivação por Vendedor")
        
        # Calcular base total por vendedor (histórico completo)
        base_vendedor = df.groupby('Vendedor')['CPF_CNPJ'].nunique().reset_index()
        base_vendedor.columns = ['Vendedor', 'TotalBase']
        
        # Calcular clientes atendidos no período (apenas NF Venda)
        vendas_periodo = df_filtrado[df_filtrado['TipoMov'] == 'NF Venda']
        atendidos = vendas_periodo.groupby('Vendedor')['CPF_CNPJ'].nunique().reset_index()
        atendidos.columns = ['Vendedor', 'QtdAtendidos']
        
        # Merge
        relatorio_positivacao = pd.merge(base_vendedor, atendidos, on='Vendedor', how='left')
        relatorio_positivacao['QtdAtendidos'] = relatorio_positivacao['QtdAtendidos'].fillna(0).astype(int)
        relatorio_positivacao['Percentual'] = (relatorio_positivacao['QtdAtendidos'] / relatorio_positivacao['TotalBase'] * 100).round(1)
        relatorio_positivacao = relatorio_positivacao.sort_values('QtdAtendidos', ascending=False)
        
        # Tabela resumida
        st.dataframe(relatorio_positivacao, use_container_width=True)
        
        # Download
        st.download_button(
            "📥 Exportar Resumo",
            to_excel(relatorio_positivacao),
            "positivacao_resumo.xlsx",
            "application/vnd.ms-excel"
        )
        
        st.markdown("---")
        
        # Detalhamento por vendedor
        st.subheader("📋 Detalhamento de Clientes por Vendedor")
        vendedor_selecionado = st.selectbox("Selecione o vendedor", relatorio_positivacao['Vendedor'].tolist())
        
        if vendedor_selecionado:
            # Obter notas únicas do vendedor no período
            notas_vendedor = obter_notas_unicas(vendas_periodo[vendas_periodo['Vendedor'] == vendedor_selecionado])
            
            # Agrupar por cliente
            clientes_vendedor = notas_vendedor.groupby(['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado']).agg({
                'Valor_Real': 'sum'
            }).reset_index()
            clientes_vendedor.columns = ['CPF/CNPJ', 'Razão Social', 'Cidade', 'Estado', 'Valor Total']
            clientes_vendedor = clientes_vendedor.sort_values('Valor Total', ascending=False)
            
            # Métricas
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total de Clientes Atendidos", len(clientes_vendedor))
            with col2:
                st.metric("Valor Total Vendido", f"R$ {clientes_vendedor['Valor Total'].sum():,.2f}")
            
            # Tabela
            st.dataframe(clientes_vendedor, use_container_width=True)
            
            # Download
            st.download_button(
                f"📥 Exportar Clientes - {vendedor_selecionado}",
                to_excel(clientes_vendedor),
                f"clientes_{vendedor_selecionado}.xlsx",
                "application/vnd.ms-excel"
            )
    
    # ====================== CLIENTES SEM COMPRA ======================
    elif menu == "Clientes sem Compra":
        st.header("⚠️ Clientes sem Compra no Período (Churn)")
        
        # Clientes com venda no período
        clientes_com_venda = set(df_filtrado[df_filtrado['TipoMov'] == 'NF Venda']['CPF_CNPJ'].unique())
        
        # Todos os clientes históricos com último vendedor
        todos_clientes = df.sort_values('DataEmissao').groupby('CPF_CNPJ').last().reset_index()
        
        # Calcular valor histórico
        valor_historico = df[df['TipoMov'] == 'NF Venda'].groupby('CPF_CNPJ')['TotalProduto'].sum().reset_index()
        valor_historico.columns = ['CPF_CNPJ', 'ValorHistorico']
        
        # Merge
        todos_clientes = pd.merge(todos_clientes, valor_historico, on='CPF_CNPJ', how='left')
        todos_clientes['ValorHistorico'] = todos_clientes['ValorHistorico'].fillna(0)
        
        # Filtrar clientes sem compra
        clientes_sem_compra = todos_clientes[~todos_clientes['CPF_CNPJ'].isin(clientes_com_venda)]
        clientes_sem_compra = clientes_sem_compra[['RazaoSocial', 'CPF_CNPJ', 'Vendedor', 'Cidade', 'Estado', 'ValorHistorico']]
        clientes_sem_compra = clientes_sem_compra.sort_values('ValorHistorico', ascending=False)
        
        # Métricas
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total de Clientes sem Compra", len(clientes_sem_compra))
        with col2:
            st.metric("Valor Potencial Perdido", f"R$ {clientes_sem_compra['ValorHistorico'].sum():,.2f}")
        
        # Tabela
        st.dataframe(clientes_sem_compra.head(50), use_container_width=True)
        
        # Download
        st.download_button(
            "📥 Exportar Clientes sem Compra",
            to_excel(clientes_sem_compra),
            "clientes_sem_compra.xlsx",
            "application/vnd.ms-excel"
        )
    
    # ====================== HISTÓRICO ======================
    elif menu == "Histórico":
        st.header("📜 Histórico de Vendas por Cliente")
        
        # Busca
        busca = st.text_input("🔍 Buscar por Razão Social ou CPF/CNPJ")
        
        if busca:
            historico = df[
                (df['RazaoSocial'].str.contains(busca, case=False, na=False)) |
                (df['CPF_CNPJ'].str.contains(busca, case=False, na=False))
            ].sort_values('DataEmissao', ascending=False)
            
            if len(historico) > 0:
                # Info do cliente
                st.info(f"📊 **Cliente:** {historico.iloc[0]['RazaoSocial']} | **CPF/CNPJ:** {historico.iloc[0]['CPF_CNPJ']} | **Total de Registros:** {len(historico)}")
                
                # Preparar dados para exibição
                historico_display = historico[['DataEmissao', 'TipoMov', 'Numero_NF', 'CodigoProduto', 'NomeProduto', 'Quantidade', 'TotalProduto']].copy()
                historico_display['DataEmissao'] = historico_display['DataEmissao'].dt.strftime('%d/%m/%Y')
                
                st.dataframe(historico_display, use_container_width=True)
                
                # Download
                st.download_button(
                    "📥 Exportar Histórico",
                    to_excel(historico),
                    "historico_cliente.xlsx",
                    "application/vnd.ms-excel"
                )
            else:
                st.warning("Nenhum resultado encontrado")
    
    # ====================== RANKINGS ======================
    elif menu == "Rankings":
        st.header("🏆 Rankings")
        
        tab1, tab2 = st.tabs(["📊 Vendedores", "👥 Clientes"])
        
        with tab1:
            st.subheader("Ranking de Vendedores por Valor")
            
            ranking_vendedores = notas_unicas.groupby('Vendedor').agg({
                'Valor_Real': 'sum',
                'Numero_NF': 'count'
            }).reset_index()
            ranking_vendedores.columns = ['Vendedor', 'Valor Total', 'Qtd Notas']
            ranking_vendedores = ranking_vendedores.sort_values('Valor Total', ascending=False)
            ranking_vendedores.insert(0, 'Posição', range(1, len(ranking_vendedores) + 1))
            
            st.dataframe(ranking_vendedores, use_container_width=True)
            
            st.download_button(
                "📥 Exportar Ranking Vendedores",
                to_excel(ranking_vendedores),
                "ranking_vendedores.xlsx",
                "application/vnd.ms-excel"
            )
        
        with tab2:
            st.subheader("Ranking de Clientes por Valor")
            
            top_n = st.selectbox("Exibir Top:", [10, 20, 50])
            
            ranking_clientes = notas_unicas.groupby(['CPF_CNPJ', 'RazaoSocial', 'Cidade', 'Estado']).agg({
                'Valor_Real': 'sum',
                'Numero_NF': 'count'
            }).reset_index()
            ranking_clientes.columns = ['CPF/CNPJ', 'Razão Social', 'Cidade', 'Estado', 'Valor Total', 'Qtd Notas']
            ranking_clientes = ranking_clientes.sort_values('Valor Total', ascending=False).head(top_n)
            ranking_clientes.insert(0, 'Posição', range(1, len(ranking_clientes) + 1))
            
            st.dataframe(ranking_clientes, use_container_width=True)
            
            st.download_button(
                "📥 Exportar Ranking Clientes",
                to_excel(ranking_clientes),
                f"ranking_top{top_n}_clientes.xlsx",
                "application/vnd.ms-excel"
            )

else:
    st.info("👆 Por favor, faça upload da planilha Excel para começar a análise")
