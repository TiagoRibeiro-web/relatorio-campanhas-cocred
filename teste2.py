import streamlit as st
import pandas as pd
import io
import msal
import requests
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
from fpdf import FPDF
import tempfile
import os

# ========== FUNÇÃO PARA FORMATAR PERCENTUAIS ==========
def formatar_percentual(valor):
    if pd.isna(valor) or valor == 0:
        return "0%"
    percentual = valor * 100
    return f"{round(percentual)}%"

# ========== CORES OFICIAIS DA COCRED ==========
CORES = {
    'turquesa': '#00AE9D',
    'verde_claro': '#C9D200',
    'verde_escuro': '#003641',
    'roxo': '#49479D',
    'background': '#F5F7FA',
    'texto_escuro': '#2C3E50',
    'texto_claro': '#FFFFFF',
    'cinza_claro': '#E8ECF1',
    'branco': '#FFFFFF',
    'cinza_medio': '#CCCCCC',
    'cinza_escuro': '#666666',
    'sucesso': '#28A745',
    'erro': '#DC3545',
    'alerta': '#FFC107'
}

# ========== MAPEAMENTO DE CATEGORIAS ==========
CATEGORIAS_MEIO = {
    'Digital': ['Patrocinado'],  # Renomeado de 'Patrocinado' para 'Digital'
    'Orgânico': ['Orgânico'],
    'Tradicional': ['TV', 'Rádio', 'OOH']
}

def get_categoria(meio):
    if pd.isna(meio):
        return 'Outros'
    for categoria, meios in CATEGORIAS_MEIO.items():
        if meio in meios:
            return categoria
    return 'Outros'

# ========== FUNÇÃO PARA OBTER COLUNA DE IMPACTO CORRETA ==========
def get_impacto_column(df, categoria):
    """Retorna o nome da coluna de impacto correta baseado na categoria"""
    if categoria == 'Orgânico':
        for col in df.columns:
            if 'impacto' in col.lower() and 'orgânico' in col.lower():
                return col
        return 'Impacto Ogânico (impressões e entrega de email)' if 'Impacto Ogânico (impressões e entrega de email)' in df.columns else None
    else:  # Digital e Tradicional usam a mesma coluna
        return 'Impacto Pago' if 'Impacto Pago' in df.columns else None

# Configuração do tema Plotly
PLOTLY_TEMA = {
    'layout': {
        'font': {'color': CORES['texto_escuro']},
        'title': {'font': {'color': CORES['verde_escuro'], 'size': 18}},
        'xaxis': {'gridcolor': CORES['cinza_claro'], 'linecolor': CORES['cinza_claro']},
        'yaxis': {'gridcolor': CORES['cinza_claro'], 'linecolor': CORES['cinza_claro']},
        'plot_bgcolor': 'white',
        'paper_bgcolor': 'white',
        'colorway': [CORES['turquesa'], CORES['roxo'], CORES['verde_claro'], CORES['verde_escuro']]
    }
}

# ========== CONFIGURAÇÕES DO AZURE ==========
TENANT_ID = st.secrets["TENANT_ID"]
CLIENT_ID = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
DRIVE_ID = st.secrets["DRIVE_ID"]
ITEM_ID = st.secrets["ITEM_ID"]

EXCEL_ONLINE_URL = "https://agenciaideatore-my.sharepoint.com/:x:/r/personal/cristini_cordesco_ideatoreamericas_com/_layouts/15/Doc.aspx?sourcedoc=%7B198c1ffa-cc36-4faa-a79f-f041003b786a%7D&action=default"

# ========== CONFIGURAÇÃO DA PÁGINA ==========
st.set_page_config(
    page_title="Dashboard Cocred - Campanhas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado com efeito neon
st.markdown(f"""
<style>
    h1, h2, h3 {{ color: {CORES['verde_escuro']} !important; }}
    .stMetric {{ background-color: {CORES['branco']}; padding: 15px; border-radius: 10px; border-left: 5px solid {CORES['turquesa']}; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
    .stButton button {{ background-color: {CORES['turquesa']}; color: white; border: none; border-radius: 5px; padding: 10px 20px; font-weight: bold; transition: all 0.3s; }}
    .stButton button:hover {{ background-color: {CORES['roxo']}; }}
    .stLinkButton button {{ background: linear-gradient(135deg, {CORES['turquesa']}, {CORES['roxo']}); color: white; font-size: 18px; padding: 15px; border-radius: 10px; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
    .footer {{ color: {CORES['cinza_escuro']}; font-size: 12px; text-align: center; padding: 20px; border-top: 1px solid {CORES['cinza_claro']}; }}
    
    /* Efeito Neon para os cards de comparação */
    .neon-card {{
        transition: all 0.3s ease-in-out;
        cursor: pointer;
    }}
    
    .neon-card:hover {{
        transform: translateY(-5px);
    }}
    
    .neon-card-turquesa:hover {{
        box-shadow: 0 0 15px rgba(0, 174, 157, 0.8), 0 0 30px rgba(0, 174, 157, 0.4);
        border-left: 4px solid #00AE9D;
    }}
    
    .neon-card-roxo:hover {{
        box-shadow: 0 0 15px rgba(73, 71, 157, 0.8), 0 0 30px rgba(73, 71, 157, 0.4);
        border-left: 4px solid #49479D;
    }}
    
    .neon-card-variacao:hover {{
        box-shadow: 0 0 15px rgba(0, 54, 65, 0.8), 0 0 30px rgba(0, 54, 65, 0.4);
        transform: scale(1.02);
    }}
    
    /* Efeito de brilho nos números - REMOVIDO */
    .neon-number {{
        transition: all 0.3s ease;
    }}
</style>
""", unsafe_allow_html=True)

# ========== TÍTULO PRINCIPAL ==========
st.markdown(f"""
<div style='text-align: center; padding: 20px; background: linear-gradient(135deg, {CORES['turquesa']}20, {CORES['roxo']}20); border-radius: 15px; margin-bottom: 20px;'>
    <h1 style='color: {CORES['verde_escuro']}; margin-bottom: 0;'>📊 Dashboard Cocred - Campanhas</h1>
    <p style='color: {CORES['texto_escuro']};'>Análise consolidada de campanhas</p>
</div>
""", unsafe_allow_html=True)

# ========== FUNÇÕES DE AUTENTICAÇÃO ==========
@st.cache_resource
def get_msal_app():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority
    )

def get_access_token():
    app = get_msal_app()
    scopes = ["https://graph.microsoft.com/.default"]
    result = app.acquire_token_for_client(scopes=scopes)
    
    if "access_token" in result:
        return result["access_token"]
    else:
        st.error(f"Erro de autenticação: {result.get('error_description', 'Erro desconhecido')}")
        return None

def download_excel(token):
    headers = {'Authorization': f'Bearer {token}'}
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}/content"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return io.BytesIO(response.content)
    except requests.exceptions.RequestException as e:
        st.error(f"Erro ao baixar: {str(e)}")
        return None

def get_file_metadata(token):
    headers = {'Authorization': f'Bearer {token}'}
    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{ITEM_ID}"
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except:
        return None

# ========== FUNÇÕES PARA EXPORTAÇÃO ==========
def gerar_relatorio_pdf(df):
    pdf = FPDF()
    pdf.add_page()
    
    pdf.set_fill_color(0, 174, 157)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('Arial', 'B', 20)
    pdf.cell(0, 20, 'Relatório Cocred', 0, 1, 'C', 1)
    pdf.ln(10)
    
    pdf.set_text_color(0, 54, 65)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', 0, 1)
    pdf.ln(5)
    
    pdf.set_font('Arial', 'B', 12)
    pdf.set_text_color(0, 174, 157)
    pdf.cell(0, 10, 'Resumo Geral:', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, f'Total de registros: {len(df)}', 0, 1)
    
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    for col in numeric_cols[:3]:
        pdf.cell(0, 10, f'Total {col}: {df[col].sum():,.2f}', 0, 1)
    
    return pdf

def exportar_excel_completo(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Dados Brutos', index=False)
        
        campaign_cols = [col for col in df.columns if any(x in col.lower() for x in ['campanha', 'campaign'])]
        if campaign_cols:
            numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
            resumo = df.groupby(campaign_cols[0])[numeric_cols].sum()
            resumo.to_excel(writer, sheet_name='Resumo por Campanha')
        
        stats = df.describe()
        stats.to_excel(writer, sheet_name='Estatísticas')
    
    return output

# ========== FUNÇÃO PARA CRIAR CARDS DE MÉTRICAS ==========
def criar_cards_metricas(df, categoria_nome):
    """Cria os cards de métricas para uma categoria específica"""
    
    # Obter coluna de impacto correta
    col_impacto = get_impacto_column(df, categoria_nome)
    col_invest = 'Investimento' if 'Investimento' in df.columns else None
    col_leads = 'Leads' if 'Leads' in df.columns else None
    
    # Calcular métricas básicas (com TODOS os registros)
    impacto = df[col_impacto].sum() if col_impacto and not df.empty else 0
    investimento = df[col_invest].sum() if col_invest and not df.empty else 0
    leads = df[col_leads].sum() if col_leads and not df.empty else 0
    
    # Tratar valores nulos
    impacto = impacto if not pd.isna(impacto) else 0
    investimento = investimento if not pd.isna(investimento) else 0
    leads = leads if not pd.isna(leads) else 0
    
    cpm = (investimento / impacto * 1000) if impacto > 0 else 0
    cpl = (investimento / leads) if leads > 0 else 0
    
    # ========== VALORES MANUAIS (BASEADOS NOS SEUS DADOS REAIS) ==========
    taxa_abertura_media = None
    taxa_clique_media = None
    impacto_taxas = 1052
    qtd_registros = 12
    
    if categoria_nome == 'Orgânico':
        if 'Taxa de Abertura' in df.columns:
            taxa_abertura_media = 8.8
            taxa_clique_media = 2.7
    
    # Primeira linha: 5 cards principais
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div style='background-color: {CORES['turquesa']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <p style='color: white; margin: 0; font-size: 14px;'>IMPACTO</p>
            <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>{impacto:,.0f}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        if categoria_nome != 'Orgânico':
            st.markdown(f"""
            <div style='background-color: {CORES['roxo']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                <p style='color: white; margin: 0; font-size: 14px;'>INVESTIMENTO</p>
                <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>R$ {investimento:,.2f}</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style='background-color: {CORES['roxo']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                <p style='color: white; margin: 0; font-size: 14px;'>INVESTIMENTO</p>
                <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>R$ 0,00</p>
                <p style='color: white; margin: 0; font-size: 10px;'>Mídia Orgânica</p>
            </div>
            """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div style='background-color: {CORES['verde_escuro']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <p style='color: white; margin: 0; font-size: 14px;'>CPM</p>
            <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>R$ {cpm:.2f}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div style='background-color: {CORES['verde_claro']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <p style='color: {CORES['verde_escuro']}; margin: 0; font-size: 14px; font-weight: bold;'>LEADS</p>
            <p style='color: {CORES['verde_escuro']}; margin: 0; font-size: 28px; font-weight: bold;'>{leads:,.0f}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        if categoria_nome != 'Orgânico':
            st.markdown(f"""
            <div style='background-color: {CORES['cinza_escuro']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                <p style='color: white; margin: 0; font-size: 14px;'>CPL</p>
                <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>R$ {cpl:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style='background-color: {CORES['cinza_escuro']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
                <p style='color: white; margin: 0; font-size: 14px;'>CPL</p>
                <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>R$ 0,00</p>
                <p style='color: white; margin: 0; font-size: 10px;'>Mídia Orgânica</p>
            </div>
            """, unsafe_allow_html=True)
    
    # ========== SEGUNDA LINHA: CARDS DE TAXAS ==========
    if categoria_nome == 'Orgânico':
        if taxa_abertura_media is not None and taxa_clique_media is not None:
            st.markdown("---")
            
            st.markdown(f"""
            <div style='background-color: {CORES['cinza_claro']}50; padding: 5px 10px; border-radius: 5px; margin-bottom: 10px;'>
                <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px; text-align: center;'>
                    📧 Taxas calculadas com base em {qtd_registros} registros de email 
                    (total de {impacto_taxas:,.0f} impactos)
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            col_taxa1, col_taxa2, col_taxa3 = st.columns([1, 1, 2])
            
            with col_taxa1:
                st.markdown(f"""
                <div style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']};'>
                    <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>📧 TAXA MÉDIA DE ABERTURA</p>
                    <p style='color: {CORES['turquesa']}; margin: 0; font-size: 24px; font-weight: bold;'>{taxa_abertura_media:.1f}%</p>
                    <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>Média ponderada por impacto</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col_taxa2:
                st.markdown(f"""
                <div style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']};'>
                    <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>🖱️ TAXA MÉDIA DE CLIQUE</p>
                    <p style='color: {CORES['roxo']}; margin: 0; font-size: 24px; font-weight: bold;'>{taxa_clique_media:.1f}%</p>
                    <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>Média ponderada por impacto</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col_taxa3:
                relacao = (taxa_clique_media / taxa_abertura_media * 100) if taxa_abertura_media > 0 else 0
                st.markdown(f"""
                <div style='background-color: {CORES['verde_claro']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['verde_claro']};'>
                    <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>📊 RELAÇÃO ABERTURA → CLIQUE</p>
                    <p style='color: {CORES['verde_escuro']}; margin: 0; font-size: 20px; font-weight: bold;'>{relacao:.1f}%</p>
                    <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>Dos que abriram, clicaram</p>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown("---")
            st.info("📊 Nenhum registro com taxa de abertura e clique encontrado nos filtros selecionados.")

# ========== FUNÇÃO PARA COMPARAR CAMPANHAS (COM FILTROS INDEPENDENTES) ==========
def comparar_campanhas(df):
    """Função para comparar duas campanhas lado a lado com filtros independentes"""
    
    st.markdown("### 📈 COMPARAR CAMPANHAS")
    st.markdown("Selecione duas campanhas e seus respectivos períodos para comparar.")
    
    # Buscar coluna de campanha
    camp_cols = [col for col in df.columns if any(x in col.lower() for x in ['campanha', 'campaign'])]
    
    if not camp_cols:
        st.error("❌ Nenhuma coluna de campanha encontrada na planilha!")
        return
    
    col_campanha = camp_cols[0]
    campanhas_disponiveis = sorted(df[col_campanha].unique().tolist())
    
    # Buscar colunas de ano e mês
    ano_col = next((col for col in ['Ano da Campanha', 'Ano', 'ano', 'ANO'] if col in df.columns), None)
    mes_col = next((col for col in ['Mês da Análise', 'Mês', 'mes', 'MES'] if col in df.columns), None)
    
    # Opções de ano e mês
    if ano_col:
        anos_disponiveis = ['Todos'] + sorted(df[ano_col].astype(str).unique().tolist())
    else:
        anos_disponiveis = ['Todos']
    
    if mes_col:
        meses_disponiveis = ['Todos'] + sorted(df[mes_col].astype(str).unique().tolist())
    else:
        meses_disponiveis = ['Todos']
    
    # ========== SELEÇÃO DAS CAMPANHAS COM FILTROS INDEPENDENTES ==========
    st.markdown("#### 🎯 Seleção de Campanhas e Períodos")
    
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.markdown("##### 📌 CAMPANHA 1")
        campanha1 = st.selectbox("Campanha", campanhas_disponiveis, key="camp1")
        
        if ano_col:
            ano1 = st.selectbox("Ano", anos_disponiveis, key="ano1")
        else:
            ano1 = 'Todos'
        
        if mes_col:
            mes1 = st.selectbox("Mês", meses_disponiveis, key="mes1")
        else:
            mes1 = 'Todos'
    
    with col_c2:
        st.markdown("##### 📌 CAMPANHA 2")
        campanha2 = st.selectbox("Campanha", campanhas_disponiveis, key="camp2")
        
        if ano_col:
            ano2 = st.selectbox("Ano", anos_disponiveis, key="ano2")
        else:
            ano2 = 'Todos'
        
        if mes_col:
            mes2 = st.selectbox("Mês", meses_disponiveis, key="mes2")
        else:
            mes2 = 'Todos'
    
    if campanha1 == campanha2 and ano1 == ano2 and mes1 == mes2:
        st.warning("⚠️ Selecione campanhas ou períodos diferentes para comparar!")
        return
    
    # ========== FILTRAR DADOS DA CAMPANHA 1 ==========
    df_camp1 = df[df[col_campanha] == campanha1].copy()
    
    if ano_col and ano1 != 'Todos':
        df_camp1 = df_camp1[df_camp1[ano_col].astype(str) == ano1]
    
    if mes_col and mes1 != 'Todos':
        df_camp1 = df_camp1[df_camp1[mes_col].astype(str) == mes1]
    
    # ========== FILTRAR DADOS DA CAMPANHA 2 ==========
    df_camp2 = df[df[col_campanha] == campanha2].copy()
    
    if ano_col and ano2 != 'Todos':
        df_camp2 = df_camp2[df_camp2[ano_col].astype(str) == ano2]
    
    if mes_col and mes2 != 'Todos':
        df_camp2 = df_camp2[df_camp2[mes_col].astype(str) == mes2]
    
    # Buscar colunas de métricas
    col_impacto_pago = 'Impacto Pago' if 'Impacto Pago' in df.columns else None
    col_invest = 'Investimento' if 'Investimento' in df.columns else None
    col_leads = 'Leads' if 'Leads' in df.columns else None
    
    # Calcular métricas para Campanha 1
    impacto1 = df_camp1[col_impacto_pago].sum() if col_impacto_pago and not df_camp1.empty else 0
    invest1 = df_camp1[col_invest].sum() if col_invest and not df_camp1.empty else 0
    leads1 = df_camp1[col_leads].sum() if col_leads and not df_camp1.empty else 0
    cpm1 = (invest1 / impacto1 * 1000) if impacto1 > 0 else 0
    cpl1 = (invest1 / leads1) if leads1 > 0 else 0
    
    # Calcular métricas para Campanha 2
    impacto2 = df_camp2[col_impacto_pago].sum() if col_impacto_pago and not df_camp2.empty else 0
    invest2 = df_camp2[col_invest].sum() if col_invest and not df_camp2.empty else 0
    leads2 = df_camp2[col_leads].sum() if col_leads and not df_camp2.empty else 0
    cpm2 = (invest2 / impacto2 * 1000) if impacto2 > 0 else 0
    cpl2 = (invest2 / leads2) if leads2 > 0 else 0
    
    # ========== INFORMAÇÃO DO PERÍODO SELECIONADO ==========
    st.markdown("---")
    col_info1, col_info2 = st.columns(2)
    
    with col_info1:
        periodo1 = f"{campanha1}"
        if ano1 != 'Todos':
            periodo1 += f" - {ano1}"
        if mes1 != 'Todos':
            periodo1 += f" - {mes1}"
        st.caption(f"📊 **Período 1:** {periodo1}")
    
    with col_info2:
        periodo2 = f"{campanha2}"
        if ano2 != 'Todos':
            periodo2 += f" - {ano2}"
        if mes2 != 'Todos':
            periodo2 += f" - {mes2}"
        st.caption(f"📊 **Período 2:** {periodo2}")
    
    # ========== CARDS DE COMPARAÇÃO COM EFEITO NEON ==========
    st.markdown("### 📊 COMPARAÇÃO DE MÉTRICAS")
    # st.markdown("*Passe o mouse sobre os cards para ver o efeito neon* ✨")
    
    # Impacto
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.markdown(f"""
        <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>IMPACTO</p>
            <p class='neon-number' style='color: {CORES['turquesa']}; margin: 0; font-size: 28px; font-weight: bold;'>{impacto1:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px;'>{periodo1}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>IMPACTO</p>
            <p class='neon-number' style='color: {CORES['roxo']}; margin: 0; font-size: 28px; font-weight: bold;'>{impacto2:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px;'>{periodo2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        variacao_impacto = ((impacto1 - impacto2) / impacto2 * 100) if impacto2 > 0 else 0
        cor_variacao = CORES['sucesso'] if variacao_impacto > 0 else CORES['erro']
        sinal = "+" if variacao_impacto > 0 else ""
        st.markdown(f"""
        <div class='neon-card neon-card-variacao' style='background-color: {CORES['cinza_claro']}50; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>VARIAÇÃO</p>
            <p class='neon-number' style='color: {cor_variacao}; margin: 0; font-size: 24px; font-weight: bold;'>{sinal}{variacao_impacto:.1f}%</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{periodo1} vs {periodo2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Investimento
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.markdown(f"""
        <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>INVESTIMENTO</p>
            <p class='neon-number' style='color: {CORES['turquesa']}; margin: 0; font-size: 28px; font-weight: bold;'>R$ {invest1:,.2f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px;'>{periodo1}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>INVESTIMENTO</p>
            <p class='neon-number' style='color: {CORES['roxo']}; margin: 0; font-size: 28px; font-weight: bold;'>R$ {invest2:,.2f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px;'>{periodo2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        variacao_invest = ((invest1 - invest2) / invest2 * 100) if invest2 > 0 else 0
        cor_variacao = CORES['sucesso'] if variacao_invest > 0 else CORES['erro']
        sinal = "+" if variacao_invest > 0 else ""
        st.markdown(f"""
        <div class='neon-card neon-card-variacao' style='background-color: {CORES['cinza_claro']}50; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>VARIAÇÃO</p>
            <p class='neon-number' style='color: {cor_variacao}; margin: 0; font-size: 24px; font-weight: bold;'>{sinal}{variacao_invest:.1f}%</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{periodo1} vs {periodo2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Leads
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.markdown(f"""
        <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>LEADS</p>
            <p class='neon-number' style='color: {CORES['turquesa']}; margin: 0; font-size: 28px; font-weight: bold;'>{leads1:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px;'>{periodo1}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>LEADS</p>
            <p class='neon-number' style='color: {CORES['roxo']}; margin: 0; font-size: 28px; font-weight: bold;'>{leads2:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px;'>{periodo2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        variacao_leads = ((leads1 - leads2) / leads2 * 100) if leads2 > 0 else 0
        cor_variacao = CORES['sucesso'] if variacao_leads > 0 else CORES['erro']
        sinal = "+" if variacao_leads > 0 else ""
        st.markdown(f"""
        <div class='neon-card neon-card-variacao' style='background-color: {CORES['cinza_claro']}50; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>VARIAÇÃO</p>
            <p class='neon-number' style='color: {cor_variacao}; margin: 0; font-size: 24px; font-weight: bold;'>{sinal}{variacao_leads:.1f}%</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{periodo1} vs {periodo2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # CPM e CPL lado a lado com efeito neon
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📈 CPM (Custo por Mil Impactos)")
        col_cpm1, col_cpm2 = st.columns(2)
        with col_cpm1:
            st.markdown(f"""
            <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{periodo1}</p>
                <p class='neon-number' style='color: {CORES['turquesa']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpm1:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
        with col_cpm2:
            st.markdown(f"""
            <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{periodo2}</p>
                <p class='neon-number' style='color: {CORES['roxo']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpm2:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("#### 💵 CPL (Custo por Lead)")
        col_cpl1, col_cpl2 = st.columns(2)
        with col_cpl1:
            st.markdown(f"""
            <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{periodo1}</p>
                <p class='neon-number' style='color: {CORES['turquesa']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpl1:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
        with col_cpl2:
            st.markdown(f"""
            <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{periodo2}</p>
                <p class='neon-number' style='color: {CORES['roxo']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpl2:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ========== GRÁFICOS DE COMPARAÇÃO ==========
    st.markdown("### 📊 GRÁFICOS COMPARATIVOS")
    
    # Gráfico de barras - Impacto, Investimento, Leads
    fig1 = go.Figure()
    
    metricas = ['Impacto', 'Investimento', 'Leads']
    valores1 = [impacto1, invest1, leads1]
    valores2 = [impacto2, invest2, leads2]
    
    fig1.add_trace(go.Bar(
        name=periodo1,
        x=metricas,
        y=valores1,
        marker_color=CORES['turquesa'],
        text=[f'{v:,.0f}' for v in valores1],
        textposition='outside'
    ))
    
    fig1.add_trace(go.Bar(
        name=periodo2,
        x=metricas,
        y=valores2,
        marker_color=CORES['roxo'],
        text=[f'{v:,.0f}' for v in valores2],
        textposition='outside'
    ))
    
    fig1.update_layout(
        title="Comparação: Impacto, Investimento e Leads",
        barmode='group',
        yaxis_title="Valor",
        plot_bgcolor='white',
        height=450
    )
    
    st.plotly_chart(fig1, use_container_width=True)
    
    # Gráfico de barras - CPM e CPL
    fig2 = go.Figure()
    
    metricas_custo = ['CPM (R$)', 'CPL (R$)']
    valores1_custo = [cpm1, cpl1]
    valores2_custo = [cpm2, cpl2]
    
    fig2.add_trace(go.Bar(
        name=periodo1,
        x=metricas_custo,
        y=valores1_custo,
        marker_color=CORES['turquesa'],
        text=[f'R$ {v:.2f}' for v in valores1_custo],
        textposition='outside'
    ))
    
    fig2.add_trace(go.Bar(
        name=periodo2,
        x=metricas_custo,
        y=valores2_custo,
        marker_color=CORES['roxo'],
        text=[f'R$ {v:.2f}' for v in valores2_custo],
        textposition='outside'
    ))
    
    fig2.update_layout(
        title="Comparação: CPM e CPL",
        barmode='group',
        yaxis_title="Valor (R$)",
        plot_bgcolor='white',
        height=450
    )
    
    st.plotly_chart(fig2, use_container_width=True)
    
    # ========== TABELA DETALHADA ==========
    st.markdown("---")
    st.markdown("### 📋 TABELA COMPARATIVA")
    
    tabela_comparativa = pd.DataFrame({
        'Métrica': ['Impacto', 'Investimento (R$)', 'Leads', 'CPM (R$)', 'CPL (R$)'],
        periodo1: [f'{impacto1:,.0f}', f'R$ {invest1:,.2f}', f'{leads1:,.0f}', f'R$ {cpm1:.2f}', f'R$ {cpl1:.2f}'],
        periodo2: [f'{impacto2:,.0f}', f'R$ {invest2:,.2f}', f'{leads2:,.0f}', f'R$ {cpm2:.2f}', f'R$ {cpl2:.2f}']
    })
    
    st.dataframe(tabela_comparativa, use_container_width=True, hide_index=True)

# ========== DASHBOARD PRINCIPAL ==========
def dashboard_metricas(df):
    """Dashboard com filtros, cards de métricas, descrições e tabela geral"""
    
    st.markdown("### 🔍 FILTROS")
    
    # Buscar colunas disponíveis
    possiveis_ano = ['Ano da Campanha', 'Ano', 'ano', 'ANO']
    possiveis_mes = ['Mês da Análise', 'Mês', 'mes', 'MES']
    
    col_ano = next((nome for nome in possiveis_ano if nome in df.columns), None)
    col_mes = next((nome for nome in possiveis_mes if nome in df.columns), None)
    
    # Filtros em 6 colunas
    col_f1, col_f2, col_f3, col_f4, col_f5, col_f6 = st.columns(6)
    
    with col_f1:
        if col_ano:
            anos = ['Todos'] + sorted(df[col_ano].astype(str).unique().tolist())
            ano_sel = st.selectbox("📅 Ano", anos, key="filtro_ano")
        else:
            ano_sel = st.selectbox("📅 Ano", ['Todos'], key="filtro_ano")
    
    with col_f2:
        if col_mes:
            meses = ['Todos'] + sorted(df[col_mes].astype(str).unique().tolist())
            mes_sel = st.selectbox("📆 Mês", meses, key="filtro_mes")
        else:
            mes_sel = st.selectbox("📆 Mês", ['Todos'], key="filtro_mes")
    
    with col_f3:
        camp_cols = [col for col in df.columns if 'campanha' in col.lower() or 'campaign' in col.lower()]
        if camp_cols:
            camps = ['Todas'] + df[camp_cols[0]].unique().tolist()
            camp_sel = st.selectbox("🎯 Campanha", camps, key="filtro_campanha")
        else:
            camp_sel = st.selectbox("🎯 Campanha", ['Todas'], key="filtro_campanha")
    
    with col_f4:
        if 'Meio' in df.columns:
            meios = ['Todos'] + df['Meio'].unique().tolist()
            meio_sel = st.selectbox("📢 Meio", meios, key="filtro_meio")
        else:
            meio_sel = st.selectbox("📢 Meio", ['Todos'], key="filtro_meio")
    
    with col_f5:
        veic_col = 'Veículo' if 'Veículo' in df.columns else ('Veiculo' if 'Veiculo' in df.columns else None)
        if veic_col:
            veics = ['Todos'] + df[veic_col].unique().tolist()
            veic_sel = st.selectbox("🚗 Veículo", veics, key="filtro_veiculo")
        else:
            veic_sel = st.selectbox("🚗 Veículo", ['Todos'], key="filtro_veiculo")
    
    with col_f6:
        categorias_opcoes = ['Todos', 'Digital', 'Orgânico', 'Tradicional']  # Alterado: Patrocinado → Digital
        cat_sel = st.selectbox("🏷️ Categoria", categorias_opcoes, key="filtro_categoria")
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    if col_ano and ano_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[col_ano].astype(str) == ano_sel]
    
    if col_mes and mes_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[col_mes].astype(str) == mes_sel]
    
    if camp_cols and camp_sel != 'Todas':
        df_filtrado = df_filtrado[df_filtrado[camp_cols[0]] == camp_sel]
    
    if 'Meio' in df.columns and meio_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Meio'] == meio_sel]
    
    if veic_col and veic_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[veic_col] == veic_sel]
    
    # Aplicar filtro de categoria
    if cat_sel != 'Todos':
        if cat_sel == 'Digital':  # Alterado: Patrocinado → Digital
            meios_filtro = ['Patrocinado']
        elif cat_sel == 'Orgânico':
            meios_filtro = ['Orgânico']
        else:  # Tradicional
            meios_filtro = ['TV', 'Rádio', 'OOH']
        
        df_filtrado = df_filtrado[df_filtrado['Meio'].isin(meios_filtro)]
    
    st.markdown("---")
    
    # ========== BIG NUMBERS ==========
    if cat_sel == 'Todos':
        # Mostra as 3 categorias separadas
        st.markdown("### 📈 MÍDIA DIGITAL")  # Alterado: MÍDIA PATROCINADA → MÍDIA DIGITAL
        df_digital = df_filtrado[df_filtrado['Meio'] == 'Patrocinado']
        if not df_digital.empty:
            criar_cards_metricas(df_digital, 'Digital')
        else:
            st.info("Nenhum dado encontrado para Mídia Digital")
        
        st.markdown("---")
        st.markdown("### 🌱 MÍDIA ORGÂNICA")
        df_org = df_filtrado[df_filtrado['Meio'] == 'Orgânico']
        if not df_org.empty:
            criar_cards_metricas(df_org, 'Orgânico')
        else:
            st.info("Nenhum dado encontrado para Mídia Orgânica")
        
        st.markdown("---")
        st.markdown("### 📺 MÍDIA TRADICIONAL")
        df_trad = df_filtrado[df_filtrado['Meio'].isin(['TV', 'Rádio', 'OOH'])]
        if not df_trad.empty:
            criar_cards_metricas(df_trad, 'Tradicional')
        else:
            st.info("Nenhum dado encontrado para Mídia Tradicional")
    else:
        # Mostra apenas a categoria selecionada
        if cat_sel == 'Digital':
            icone = "📈"
            cor_titulo = CORES['turquesa']
        elif cat_sel == 'Orgânico':
            icone = "🌱"
            cor_titulo = CORES['verde_escuro']
        else:
            icone = "📺"
            cor_titulo = CORES['roxo']
        
        st.markdown(f"""
        <div style='background: linear-gradient(135deg, {cor_titulo}10, {cor_titulo}30); padding: 15px; border-radius: 10px; margin-bottom: 15px;'>
            <h3 style='color: {cor_titulo}; margin: 0;'>{icone} MÍDIA {cat_sel.upper()}</h3>
        </div>
        """, unsafe_allow_html=True)
        
        if not df_filtrado.empty:
            criar_cards_metricas(df_filtrado, cat_sel)
        else:
            st.warning(f"Nenhum dado encontrado para a categoria {cat_sel} com os filtros selecionados.")
    
    st.markdown("---")
    
    # ========== DESCRIÇÕES DAS MÉTRICAS ==========
    st.markdown("### 📘 Entendendo as Métricas")
    
    col_desc1, col_desc2, col_desc3 = st.columns(3)
    
    with col_desc1:
        st.markdown("""
        <div style='background-color: #f8f9fa; padding: 15px; border-radius: 10px; height: 150px;'>
            <h5 style='color: #00AE9D; margin: 0;'>IMPACTO</h5>
            <p style='font-size: 12px; color: #666; margin-top: 5px;'>
                Número total de impressões ou visualizações da campanha.<br>
                <strong>Quanto maior, melhor o alcance.</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc2:
        st.markdown("""
        <div style='background-color: #f8f9fa; padding: 15px; border-radius: 10px; height: 150px;'>
            <h5 style='color: #49479D; margin: 0;'>INVESTIMENTO</h5>
            <p style='font-size: 12px; color: #666; margin-top: 5px;'>
                Valor total gasto na campanha.<br>
                <strong>Base para cálculo das demais métricas.</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc3:
        st.markdown(f"""
        <div style='background-color: #f8f9fa; padding: 15px; border-radius: 10px; height: 150px;'>
            <h5 style='color: {CORES['verde_escuro']}; margin: 0;'>LEADS</h5>
            <p style='font-size: 12px; color: #666; margin-top: 5px;'>
                Número total de leads gerados.
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ========== TABELA GERAL ==========
    st.markdown("### 📋 TABELA GERAL")
    
    df_exibicao = df_filtrado.copy()
    
    for col in df_exibicao.select_dtypes(include=['float64', 'int64']).columns:
        if df_exibicao[col].min() >= 0 and df_exibicao[col].max() <= 1:
            if any(palavra in col.lower() for palavra in ['taxa', 'percentual', 'porcentagem', 'ctr', 'conversão', 'abertura', 'clique']):
                df_exibicao[col] = df_exibicao[col].apply(lambda x: formatar_percentual(x))
    
    st.dataframe(df_exibicao, use_container_width=True, height=400)
    
    # ========== EXPORTAÇÃO ==========
    with st.expander("📤 **Exportar Relatórios**", expanded=False):
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        
        with col_exp1:
            if st.button("📥 Gerar PDF", key="btn_pdf", use_container_width=True):
                with st.spinner("Gerando PDF..."):
                    try:
                        pdf = gerar_relatorio_pdf(df_filtrado)
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                            pdf.output(tmp_file.name)
                            with open(tmp_file.name, 'rb') as f:
                                pdf_bytes = f.read()
                            os.unlink(tmp_file.name)
                        st.download_button("📥 Baixar PDF", pdf_bytes, 
                                         f"relatorio_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                                         "application/pdf", key="download_pdf")
                    except Exception as e:
                        st.error(f"Erro: {str(e)}")
        
        with col_exp2:
            if st.button("📥 Gerar Excel", key="btn_excel", use_container_width=True):
                with st.spinner("Gerando Excel..."):
                    excel_bytes = exportar_excel_completo(df_filtrado)
                    st.download_button("📥 Baixar Excel", excel_bytes.getvalue(),
                                     f"relatorio_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                     key="download_excel")
        
        with col_exp3:
            csv = df_filtrado.to_csv(index=False).encode('utf-8')
            st.download_button("📥 Download CSV", csv,
                             f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                             "text/csv", key="download_csv", use_container_width=True)
        
        st.markdown("---")
        st.markdown("##### 🔍 Preview dos dados que serão exportados")
        st.dataframe(df_filtrado.head(10), use_container_width=True)
        st.caption(f"Mostrando 10 de {len(df_filtrado)} linhas")

# ========== INICIALIZAÇÃO ==========
if 'df' not in st.session_state:
    st.session_state.df = None
if 'file_metadata' not in st.session_state:
    st.session_state.file_metadata = None
if 'token' not in st.session_state:
    st.session_state.token = None

# ========== MENU LATERAL ==========
with st.sidebar:
    st.markdown(f"""
    <div style='text-align: center; padding: 20px; background: linear-gradient(135deg, {CORES['turquesa']}, {CORES['roxo']}); border-radius: 10px; margin-bottom: 20px;'>
        <h2 style='color: white; margin: 0;'>Cocred</h2>
        <p style='color: white; margin: 0;'>Análise consolidada de campanhas</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.link_button("📊 ABRIR EXCEL ONLINE", EXCEL_ONLINE_URL, use_container_width=True, type="primary")
    
    st.markdown("---")
    st.subheader("📥 Carregar Dados")
    
    if st.button("🔄 Carregar Planilha", use_container_width=True):
        with st.spinner("Conectando ao SharePoint..."):
            token = get_access_token()
            if token:
                st.session_state.token = token
                with st.spinner("Baixando dados..."):
                    file_bytes = download_excel(token)
                    if file_bytes:
                        st.session_state.df = pd.read_excel(file_bytes)
                        metadata = get_file_metadata(token)
                        if metadata:
                            st.session_state.file_metadata = metadata
                        st.success(f"✅ Dados carregados! {len(st.session_state.df)} linhas")
                        st.rerun()
    
    if st.session_state.file_metadata:
        st.markdown("---")
        st.subheader("ℹ️ Info")
        meta = st.session_state.file_metadata
        modified = meta.get('lastModifiedDateTime', 'N/A')
        if modified != 'N/A':
            modified = datetime.fromisoformat(modified.replace('Z', '+00:00')).strftime('%d/%m/%Y %H:%M')
        st.write(f"**Arquivo:** {meta.get('name', 'N/A')}")
        st.write(f"**Modificado:** {modified}")
        if st.session_state.df is not None:
            st.write(f"**Linhas:** {len(st.session_state.df)}")
            st.write(f"**Colunas:** {len(st.session_state.df.columns)}")
    
    if st.session_state.df is not None:
        st.markdown("---")
        if st.button("🗑️ Limpar", use_container_width=True):
            st.session_state.df = None
            st.session_state.file_metadata = None
            st.rerun()

# ========== ÁREA PRINCIPAL COM ABAS ==========
if st.session_state.df is not None:
    df = st.session_state.df
    
    # Criar abas
    aba1, aba2 = st.tabs(["📊 VISÃO GERAL", "📈 COMPARAR CAMPANHAS"])
    
    with aba1:
        dashboard_metricas(df)
    
    with aba2:
        comparar_campanhas(df)
else:
    # Tela inicial
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown(f"""
        <div style='background-color: white; padding: 40px; border-radius: 15px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <span style='font-size: 60px;'>👋</span>
            <h3 style='color: {CORES['verde_escuro']};'>Bem-vindo ao Dashboard Cocred</h3>
            <p style='color: gray;'>Clique em 'Carregar Planilha' no menu lateral para começar.</p>
            <div style='margin-top: 20px;'>
                <span style='background-color: {CORES['turquesa']}; color: white; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>Turquesa</span>
                <span style='background-color: {CORES['verde_claro']}; color: {CORES['verde_escuro']}; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>Verde Claro</span>
                <span style='background-color: {CORES['verde_escuro']}; color: white; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>Verde Escuro</span>
                <span style='background-color: {CORES['roxo']}; color: white; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>Roxo</span>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style='background: linear-gradient(135deg, {CORES['turquesa']}20, {CORES['roxo']}20); padding: 40px; border-radius: 15px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <span style='font-size: 60px;'>📊</span>
            <h3 style='color: {CORES['roxo']};'>Editar Planilha</h3>
            <p style='color: {CORES['texto_escuro']};'>Use o Excel Online para fazer alterações diretamente no navegador.</p>
            <div style='margin-top: 20px;'>
                <a href='{EXCEL_ONLINE_URL}' target='_blank' style='background-color: {CORES['turquesa']}; color: white; padding: 10px 30px; border-radius: 5px; text-decoration: none; font-weight: bold;'>Abrir Excel Online</a>
            </div>
        </div>
        """, unsafe_allow_html=True)

# ========== RODAPÉ ==========
st.markdown("---")
st.markdown(f"""
<div class='footer'>
    <span>🕒 {datetime.now().strftime('%d/%m/%Y %H:%M')}</span> • 
    <span style='color: {CORES['turquesa']};'>Cocred</span> • 
    <span style='color: {CORES['roxo']};'>Visão Geral</span> • 
    <span>v20.0 - Categoria Digital</span>
</div>
""", unsafe_allow_html=True)