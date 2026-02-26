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
            'pedidos_pendentes': None,
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
            "modulos": ["Dashboard", "Positivação", "Inadimplência", "Clientes sem Compra", "Histórico", "Preço Médio", "Pedidos Pendentes", "Rankings"]
        },
        "colaborador123": {  # ⬅️ MUDE ESTA SENHA
            "tipo": "colaborador",
            "nome": "Colaborador",
            "modulos": ["Inadimplência", "Histórico", "Pedidos Pendentes"]
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

    # Tela de login
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
    
    # CABEÇALHO COM LOGO DO GITHUB
    try:
        # Buscar logo do GitHub
        logo_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{GITHUB_FOLDER}/logo.png"
        response = requests.get(logo_url)
        if response.status_code == 200:
            logo_buffer = io.BytesIO(response.content)
            # Manter proporção do logo - altura fixa, largura proporcional
            logo_img = Image(logo_buffer, height=15*mm)  # Altura fixa de 15mm
            logo_img.hAlign = 'LEFT'  # Alinhar à esquerda como no modelo
            elements.append(logo_img)
            elements.append(Spacer(1, 5*mm))
        else:
            # Fallback para texto se logo não encontrado
            elements.append(Paragraph("<b>MEDTEXTIL PRODUTOS TEXTIL HOSPITALARES</b>", style_title))
            elements.append(Spacer(1, 3*mm))
    except:
        # Fallback para texto em caso de erro
        elements.append(Paragraph("<b>MEDTEXTIL PRODUTOS TEXTIL HOSPITALARES</b>", style_title))
        elements.append(Spacer(1, 3*mm))
    
    elements.append(Paragraph("CNPJ: 40.357.820/0001-50 | Inscrição Estadual: 16.390.286-0", style_small))
    elements.append(Spacer(1, 10*mm))
    
    # REPRESENTANTE
    data_repr = [[Paragraph("<b>Representante</b>", style_normal), dados_cliente.get('representante', '')]]
    table_repr = Table(data_repr, colWidths=[40*mm, 150*mm])
    table_repr.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
    ]))
    elements.append(table_repr)
    elements.append(Spacer(1, 3*mm))
    
    # INFORMAÇÕES DO CLIENTE
    data_cliente = [
        [Paragraph("<b>Informações sobre o Cliente</b>", style_normal), '', '', ''],
        ['Cliente', dados_cliente.get('razao_social', ''), 'Nome Fantasia:', dados_cliente.get('nome_fantasia', '')],
        ['CNPJ:', dados_cliente.get('cnpj', ''), 'Inscr. Estadual:', dados_cliente.get('ie', '')],
        ['Telefone:', dados_cliente.get('telefone', ''), 'Email NF-e:', dados_cliente.get('email', '')],
        ['Endereço:', dados_cliente.get('endereco', ''), '', ''],
        ['Observação:', dados_cliente.get('obs_cliente', ''), '', '']
    ]
    
    table_cliente = Table(data_cliente, colWidths=[30*mm, 70*mm, 30*mm, 60*mm])
    table_cliente.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('INNERGRID', (0, 1), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('SPAN', (0, 0), (-1, 0)),
        ('SPAN', (1, 4), (-1, 4)),
        ('SPAN', (1, 5), (-1, 5)),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ]))
    elements.append(table_cliente)
    elements.append(Spacer(1, 3*mm))
    
    # INFORMAÇÕES DO PEDIDO
    data_pedido_info = [
        [Paragraph("<b>Informações sobre Pedido Nº</b>", style_normal), dados_pedido.get('numero', ''), 'Tabela de Preço:', dados_pedido.get('tabela_preco', '')],
        ['Condições de Pagto:', dados_pedido.get('condicoes_pagto', ''), 'Tipo de Frete:', dados_pedido.get('tipo_frete', 'CIF')],
        ['Data da Venda:', dados_pedido.get('data_venda', ''), '', '']
    ]
    
    table_pedido = Table(data_pedido_info, colWidths=[50*mm, 50*mm, 40*mm, 50*mm])
    table_pedido.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
        ('SPAN', (0, 0), (0, 0)),
        ('SPAN', (1, 2), (-1, 2)),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),
    ]))
    elements.append(table_pedido)
    elements.append(Spacer(1, 5*mm))
    
    # DETALHE DO PEDIDO
    elements.append(Paragraph("<b>Detalhe do Pedido</b>", style_normal))
    elements.append(Spacer(1, 2*mm))
    
    # Cabeçalho da tabela de itens
    header_itens = ['COD.', 'PRODUTO', 'PESO', 'CAIXA EMB', 'QTDE', 'VALOR', 'TOTAL', 'COMISSÃO%']
    data_itens = [header_itens]
    
    # Adicionar itens
    for item in itens_pedido:
        data_itens.append([
            str(item.get('codigo', '')),
            str(item.get('descricao', ''))[:50],  # Limitar tamanho
            str(item.get('peso', '')),
            str(item.get('cx_embarque', '')),
            f"{item.get('quantidade', 0):.0f}",
            f"R$ {item.get('valor_unit', 0):.2f}",
            f"R$ {item.get('total', 0):.2f}",
            str(item.get('comissao', ''))
        ])
    
    # Calcular totais
    total_qtde = sum([item.get('quantidade', 0) for item in itens_pedido])
    total_valor = sum([item.get('total', 0) for item in itens_pedido])
    
    # Linha de total
    data_itens.append(['', '', '', '', f"{total_qtde:.0f}", '', f"R$ {total_valor:,.2f}", ''])
    
    col_widths = [15*mm, 65*mm, 15*mm, 20*mm, 15*mm, 20*mm, 25*mm, 20*mm]
    table_itens = Table(data_itens, colWidths=col_widths)
    
    # Estilo da tabela de itens
    style_itens = [
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f4788')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
    ]
    
    table_itens.setStyle(TableStyle(style_itens))
    elements.append(table_itens)
    elements.append(Spacer(1, 5*mm))
    
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

# ====================== INÍCIO DO APP ======================
if not check_password():
    st.stop()

# Obter informações do usuário logado
usuario = st.session_state.get("usuario", {})
tipo_usuario = usuario.get("tipo", "")
nome_usuario = usuario.get("nome", "Usuário")
modulos_permitidos = usuario.get("modulos", [])

# Header com informação do usuário
col_titulo, col_usuario = st.columns([3, 1])

with col_titulo:
    st.title("📊 Dashboard BI Medtextil 2.0")

with col_usuario:
    st.markdown(f"**👤 {nome_usuario}**")
    if st.button("🚪 Sair", use_container_width=True):
        # Limpar sessão e fazer logout
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

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

# Carregar planilha de produtos para cálculo de comissão
if planilhas_disponiveis.get('produtos_agrupados'):
    with st.spinner("📥 Carregando tabela de preços de referência..."):
        df_ref_preco = carregar_planilha_github(planilhas_disponiveis['produtos_agrupados']['url'])
    
    if df_ref_preco is not None:
        df_ref_preco.columns = df_ref_preco.columns.str.upper()
        
        # Verificar se as colunas necessárias existem
        if 'ID_COD' in df_ref_preco.columns and 'PRECO' in df_ref_preco.columns:
            # Manter apenas código e preço de referência
            df_ref_preco = df_ref_preco[['ID_COD', 'PRECO']].rename(
                columns={'ID_COD': 'CodigoProduto', 'PRECO': 'PrecoRef'}
            )
            # Normalizar código para inteiro em ambas as planilhas
            # Vendas tem '3.0' e Produtos tem '3' — converter os dois para int como string
            def normalizar_codigo(val):
                try:
                    return str(int(float(str(val).strip())))
                except:
                    return str(val).strip()
            
            df['CodigoProduto'] = df['CodigoProduto'].apply(normalizar_codigo)
            df_ref_preco['CodigoProduto'] = df_ref_preco['CodigoProduto'].apply(normalizar_codigo)
            
            # Remover duplicatas (manter primeiro preço por produto)
            df_ref_preco = df_ref_preco.drop_duplicates(subset=['CodigoProduto'], keep='first')
            # Fazer join com o df principal
            df = df.merge(df_ref_preco, on='CodigoProduto', how='left')
            # Calcular comissão para cada linha
            df['Comissao'] = df.apply(
                lambda row: calcular_comissao(row['PrecoUnit'], row['PrecoRef']),
                axis=1
            )
        else:
            df['PrecoRef'] = None
            df['Comissao'] = ''
            colunas_encontradas = df_ref_preco.columns.tolist()
            st.warning(f"⚠️ Coluna 'PRECO' ou 'ID_COD' não encontrada. Colunas disponíveis: {colunas_encontradas}")
else:
    df['PrecoRef'] = None
    df['Comissao'] = ''

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

# Mostrar apenas módulos permitidos para o usuário
modulos_visiveis = modulos_permitidos if modulos_permitidos else ["Dashboard", "Positivação", "Inadimplência", "Clientes sem Compra", "Histórico", "Preço Médio", "Pedidos Pendentes", "Rankings"]

# Definir índice padrão baseado nas permissões
indice_padrao = 0
if "Dashboard" not in modulos_visiveis and len(modulos_visiveis) > 0:
    indice_padrao = 0  # Sempre começa no primeiro disponível

menu = st.sidebar.radio(
    "📑 Navegação",
    modulos_visiveis,
    index=indice_padrao
)

# Verificar se o usuário tem permissão para acessar o módulo
if menu not in modulos_permitidos:
    st.error("🚫 Você não tem permissão para acessar este módulo")
    st.info(f"📋 Módulos disponíveis: {', '.join(modulos_permitidos)}")
    st.stop()

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
            
            # Verificar se as colunas necessárias existem
            if 'NumeroDoc' not in df_inad_filtrado.columns:
                # Tentar encontrar coluna alternativa
                possiveis_nomes = [col for col in df_inad_filtrado.columns if 'DOC' in col.upper() or 'NUMERO' in col.upper()]
                if possiveis_nomes:
                    df_inad_filtrado['NumeroDoc'] = df_inad_filtrado[possiveis_nomes[0]]
                else:
                    # Criar coluna fake apenas para não quebrar
                    df_inad_filtrado['NumeroDoc'] = 1
            
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
    
    tab1, tab2, tab3 = st.tabs(["👤 Por Cliente", "🧑‍💼 Por Vendedor", "📝 Pedidos"])
    
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
                    else:
                        produto_info['preco_sugerido'] = prod.get('PRECO', 0)
                else:
                    produto_info['preco_sugerido'] = prod.get('PRECO', 0)
        
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
                        'total': qtde_item * valor_item,
                        'comissao': comissao
                    }
                    st.session_state.itens_pedido.append(item)
                    st.success(f"✅ Item adicionado: {produto_info['descricao']}")
                    st.rerun()
        
        # Mostrar produtos adicionados
        if st.session_state.itens_pedido:
            st.markdown("---")
            st.markdown("### 📦 Itens do Pedido")
            
            # Criar DataFrame dos itens
            df_itens = pd.DataFrame(st.session_state.itens_pedido)
            
            # Formatar para exibição
            df_itens_display = df_itens.copy()
            df_itens_display['valor_unit'] = df_itens_display['valor_unit'].apply(lambda x: f"R$ {x:,.2f}")
            df_itens_display['total'] = df_itens_display['total'].apply(lambda x: f"R$ {x:,.2f}")
            
            df_itens_display = df_itens_display.rename(columns={
                'codigo': 'COD.',
                'descricao': 'PRODUTO',
                'peso': 'PESO',
                'cx_embarque': 'CAIXA EMBARQUE',
                'quantidade': 'QTDE',
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
    
    # IMPORTANTE: Se a planilha de vendas já tiver NOMEPRODUTO, remover para usar apenas o da planilha de produtos
    if 'NOMEPRODUTO' in df_vendas_produto.columns:
        df_vendas_produto = df_vendas_produto.drop(columns=['NOMEPRODUTO'])
        st.info("ℹ️ Coluna NOMEPRODUTO da planilha de vendas foi substituída pela descrição da planilha de produtos")
    
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
    
    st.success(f"✅ Dados carregados: {len(df_preco_medio):,} registros de vendas")
    st.info(f"📊 Planilha de Vendas: **{planilhas_disponiveis['vendas_produto']['nome']}**")
    st.info(f"📦 Planilha de Produtos: **{planilhas_disponiveis['produtos_agrupados']['nome']}**")
    st.info(f"📅 Período de Referência: **{data_atual.strftime('%B/%Y')}** (mês atual)")
    
    if produtos_nao_catalogados > 0:
        st.warning(f"⚠️ {produtos_nao_catalogados} produtos sem cadastro na planilha de produtos (verifique se os códigos coincidem)")
    else:
        st.success("✅ Todos os produtos foram encontrados no cadastro!")
    
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

# ====================== PEDIDOS PENDENTES ======================
elif menu == "Pedidos Pendentes":
    st.header("📦 Pedidos Pendentes de Faturamento")
    
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
            
            st.success(f"✅ {len(df_pendentes):,} itens pendentes carregados")
            
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
        st.metric("💰 Valor Total Pendente", f"R$ {total_pendente:,.2f}")
    
    with col2:
        qtd_pendente = df_pend_filtrado['QtdPendente'].sum()
        st.metric("📦 Qtd Total Pendente", f"{qtd_pendente:,.0f}")
    
    with col3:
        pedidos_unicos = df_pend_filtrado['NumeroPedido'].nunique()
        st.metric("📋 Pedidos Únicos", f"{pedidos_unicos:,}")
    
    with col4:
        perc_medio = df_pend_filtrado['PercEntregue'].mean() if len(df_pend_filtrado) > 0 else 0
        st.metric("📊 % Médio Entregue", f"{perc_medio:.1f}%")
    
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
            template='plotly_white',
            color='ValorPendente',
            color_continuous_scale='Reds'
        )
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
            template='plotly_white',
            color='ValorPendente',
            color_continuous_scale='Oranges'
        )
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
    
    # Botão de download
    st.download_button(
        "📥 Exportar Pedidos Pendentes (Separado por Tipo)",
        to_excel_pedidos_pendentes(df_pend_filtrado),
        "pedidos_pendentes.xlsx",
        "application/vnd.ms-excel",
        key="download_pendentes"
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
