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
    'azul': '#1E88E5',
    'background': '#F5F7FA',
    'texto_escuro': '#2C3E50',
    'texto_claro': '#FFFFFF',
    'cinza_claro': '#E8ECF1',
    'branco': '#FFFFFF',
    'cinza_medio': '#CCCCCC',
    'cinza_escuro': '#666666',
    'sucesso': '#28A745',
    'erro': '#DC3545',
    'alerta': '#FFC107',
    'laranja': '#FF6B35'
}

# ========== MAPEAMENTO DE CATEGORIAS ==========
CATEGORIAS_MEIO = {
    'Digital': ['Patrocinado', 'Portal'],
    'Orgânico': ['Orgânico'],
    'Tradicional': ['TV', 'Rádio', 'OOH', 'Revista']
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
    else:
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
    
    /* Efeito Neon para TODOS os cards */
    .neon-card {{
        transition: all 0.3s ease-in-out;
        cursor: pointer;
    }}
    
    .neon-card:hover {{
        transform: translateY(-5px);
    }}
    
    .neon-card-turquesa:hover {{
        box-shadow: 0 0 15px rgba(0, 174, 157, 0.8), 0 0 30px rgba(0, 174, 157, 0.4);
    }}
    
    .neon-card-roxo:hover {{
        box-shadow: 0 0 15px rgba(73, 71, 157, 0.8), 0 0 30px rgba(73, 71, 157, 0.4);
    }}
    
    .neon-card-verde:hover {{
        box-shadow: 0 0 15px rgba(0, 54, 65, 0.8), 0 0 30px rgba(0, 54, 65, 0.4);
    }}
    
    .neon-card-laranja:hover {{
        box-shadow: 0 0 15px rgba(255, 107, 53, 0.8), 0 0 30px rgba(255, 107, 53, 0.4);
    }}
    
    .neon-card-claro:hover {{
        box-shadow: 0 0 15px rgba(201, 210, 0, 0.8), 0 0 30px rgba(201, 210, 0, 0.4);
    }}
    
    .neon-card-cinza:hover {{
        box-shadow: 0 0 15px rgba(102, 102, 102, 0.8), 0 0 30px rgba(102, 102, 102, 0.4);
    }}
    
    .neon-card-azul:hover {{
        box-shadow: 0 0 15px rgba(30, 136, 229, 0.8), 0 0 30px rgba(30, 136, 229, 0.4);
    }}
</style>
""", unsafe_allow_html=True)

# ========== TÍTULO PRINCIPAL ==========
st.markdown(f"""
<div style='text-align: center; padding: 20px; background: linear-gradient(135deg, {CORES['turquesa']}20, {CORES['roxo']}20); border-radius: 15px; margin-bottom: 20px;'>
    <h1 style='color: {CORES['verde_escuro']}; margin-bottom: 0;'>📊 Dashboard Cocred - Campanhas</h1>
    <p style='color: {CORES['texto_escuro']};'>Análise consolidada de campanhas | Mídia ON, OFF e Orgânica</p>
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

# ========== FUNÇÃO: CARDS CONSOLIDADOS ON/OFF ==========
def criar_cards_consolidados(df):
    """Cria cards de KPIs consolidados no formato Mídia ON e OFF"""
    
    # Identificar colunas
    col_impacto_pago = 'Impacto Pago' if 'Impacto Pago' in df.columns else None
    col_impacto_org = None
    for col in df.columns:
        if 'impacto' in col.lower() and 'orgânico' in col.lower():
            col_impacto_org = col
            break
    if not col_impacto_org:
        col_impacto_org = 'Impacto Ogânico (impressões e entrega de email)' if 'Impacto Ogânico (impressões e entrega de email)' in df.columns else None
    
    col_invest = 'Investimento' if 'Investimento' in df.columns else None
    col_leads = 'Leads' if 'Leads' in df.columns else None
    
    # Separar dados por Mídia ON e OFF
    midia_on_meios = ['Patrocinado', 'Portal']
    midia_off_meios = ['TV', 'Rádio', 'OOH', 'Revista']
    
    df_on = df[df['Meio'].isin(midia_on_meios)] if 'Meio' in df.columns else pd.DataFrame()
    df_off = df[df['Meio'].isin(midia_off_meios)] if 'Meio' in df.columns else pd.DataFrame()
    
    # Impacto Total (ON + OFF + Orgânico)
    impacto_on = df_on[col_impacto_pago].sum() if col_impacto_pago and not df_on.empty else 0
    impacto_off = df_off[col_impacto_pago].sum() if col_impacto_pago and not df_off.empty else 0
    impacto_org = df[col_impacto_org].sum() if col_impacto_org and not df.empty else 0
    
    impacto_total = impacto_on + impacto_off + impacto_org
    impacto_total = impacto_total if not pd.isna(impacto_total) else 0
    
    # Investimento Mídia OFF
    invest_off = df_off[col_invest].sum() if col_invest and not df_off.empty else 0
    invest_off = invest_off if not pd.isna(invest_off) else 0
    
    # Investimento Mídia ON
    invest_on = df_on[col_invest].sum() if col_invest and not df_on.empty else 0
    invest_on = invest_on if not pd.isna(invest_on) else 0
    
    # CPM Mídia ON
    cpm_on = (invest_on / impacto_on * 1000) if impacto_on > 0 else 0
    
    # CPM Mídia OFF
    cpm_off = (invest_off / impacto_off * 1000) if impacto_off > 0 else 0
    
    # CPM TOTAL (ON + OFF unificado)
    investimento_total_on_off = invest_on + invest_off
    impacto_total_on_off = impacto_on + impacto_off
    cpm_total = (investimento_total_on_off / impacto_total_on_off * 1000) if impacto_total_on_off > 0 else 0
    
    # Leads Total
    leads_total = df[col_leads].sum() if col_leads and not df.empty else 0
    leads_total = leads_total if not pd.isna(leads_total) else 0
    
    # CPL Médio (consolidado)
    cpl_medio = (investimento_total_on_off / leads_total) if leads_total > 0 else 0
    
    # Custo de Produção
    df_producao = df[df['Meio'] == 'Produção'] if 'Meio' in df.columns else pd.DataFrame()
    custo_producao = df_producao[col_invest].sum() if col_invest and not df_producao.empty else 0
    custo_producao = custo_producao if not pd.isna(custo_producao) else 0
    
    # ========== MÉTRICAS CONSOLIDADAS ==========
    st.markdown("#### 📊 MÉTRICAS CONSOLIDADAS")
    
    # ========== LINHA ÚNICA COM 6 CARDS AGRUPADOS ==========
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    
    with col1:
        st.markdown(f"""
        <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 120px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>📊 IMPACTO TOTAL</p>
            <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>{impacto_total:,.0f}</p>
            <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>ON + OFF + Orgânico</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='neon-card neon-card-azul' style='background-color: {CORES['azul']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 120px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>📈 CPM TOTAL</p>
            <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>R$ {cpm_total:.2f}</p>
            <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>ON + OFF Unificado</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div class='neon-card neon-card-laranja' style='background-color: {CORES['laranja']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 120px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>📺 INVEST. MÍDIA OFF</p>
            <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>R$ {invest_off:,.2f}</p>
            <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>TV, Rádio, OOH, Revista</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div class='neon-card neon-card-laranja' style='background-color: {CORES['laranja']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 120px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>📺 CPM MÍDIA OFF</p>
            <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>R$ {cpm_off:.2f}</p>
            <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>Custo por Mil Impactos</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 120px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>💻 INVEST. MÍDIA ON</p>
            <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>R$ {invest_on:,.2f}</p>
            <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>Patrocinado, Portal (Digital)</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col6:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 120px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>💻 CPM MÍDIA ON</p>
            <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>R$ {cpm_on:.2f}</p>
            <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>Custo por Mil Impactos</p>
        </div>
        """, unsafe_allow_html=True)
    
    # ========== LINHA 2: LEADS + CPL + PRODUÇÃO (3 cards centralizados) ==========
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_center1, col_center2, col_center3 = st.columns([0.5, 3, 0.5])
    
    with col_center2:
        col_a, col_b, col_c = st.columns(3)
        
        with col_a:
            st.markdown(f"""
            <div class='neon-card neon-card-verde' style='background-color: {CORES['verde_escuro']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 110px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>🎯 LEADS</p>
                <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>{leads_total:,.0f}</p>
                <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>Total de Leads</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col_b:
            st.markdown(f"""
            <div class='neon-card neon-card-cinza' style='background-color: {CORES['cinza_escuro']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 110px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: white; margin: 0; font-size: 11px; opacity: 0.9;'>💵 CPL MÉDIO</p>
                <p style='color: white; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>R$ {cpl_medio:.2f}</p>
                <p style='color: white; margin: 0; font-size: 9px; opacity: 0.7;'>Custo por Lead</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col_c:
            st.markdown(f"""
            <div class='neon-card neon-card-claro' style='background-color: {CORES['verde_claro']}; padding: 18px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1); height: 110px; display: flex; flex-direction: column; justify-content: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['verde_escuro']}; margin: 0; font-size: 11px; opacity: 0.9; font-weight: bold;'>🎬 CUSTO DE PRODUÇÃO</p>
                <p style='color: {CORES['verde_escuro']}; margin: 5px 0 0 0; font-size: 22px; font-weight: bold;'>R$ {custo_producao:,.2f}</p>
                <p style='color: {CORES['verde_escuro']}; margin: 0; font-size: 9px; opacity: 0.7;'>Meio = Produção</p>
            </div>
            """, unsafe_allow_html=True)
    
    # ========== LINHA 3: TAXAS DE EMAIL (quando houver dados) ==========
    if 'Taxa de Abertura' in df.columns and 'Taxa de Clique' in df.columns:
        df_email = df[df['Taxa de Abertura'].notna() & df['Taxa de Clique'].notna()]
        
        if not df_email.empty:
            # Média simples das taxas (valores já em percentual)
            taxa_abertura_media = df_email['Taxa de Abertura'].mean()*100
            taxa_clique_media = df_email['Taxa de Clique'].mean()*100
            
            st.markdown("---")
            st.markdown(f"""
            <div style='background-color: {CORES['cinza_claro']}50; padding: 5px 10px; border-radius: 5px; margin-bottom: 10px;'>
                <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 11px; text-align: center;'>
                    📧 Taxas de Email Marketing - Média simples com base em {len(df_email)} registros
                </p>
            </div>
            """, unsafe_allow_html=True)
            
            col_center1, col_center2, col_center3 = st.columns([1.5, 2, 1.5])
            
            with col_center2:
                col_taxa1, col_taxa2 = st.columns(2)
                
                with col_taxa1:
                    st.markdown(f"""
                    <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 18px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']}; transition: all 0.3s ease-in-out; cursor: pointer; height: 120px; display: flex; flex-direction: column; justify-content: center;'>
                        <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>📧 TAXA MÉDIA DE ABERTURA</p>
                        <p style='color: {CORES['turquesa']}; margin: 5px 0 0 0; font-size: 24px; font-weight: bold;'>{taxa_abertura_media:.1f}%</p>
                        <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 9px;'>Média simples</p>
                        <p style='color: {CORES['cinza_escuro']}; margin: 3px 0 0 0; font-size: 9px;'>🎯 Benchmark: 30%</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col_taxa2:
                    st.markdown(f"""
                    <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 18px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']}; transition: all 0.3s ease-in-out; cursor: pointer; height: 120px; display: flex; flex-direction: column; justify-content: center;'>
                        <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>🖱️ TAXA MÉDIA DE CLIQUE</p>
                        <p style='color: {CORES['roxo']}; margin: 5px 0 0 0; font-size: 24px; font-weight: bold;'>{taxa_clique_media:.1f}%</p>
                        <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 9px;'>Média simples</p>
                        <p style='color: {CORES['cinza_escuro']}; margin: 3px 0 0 0; font-size: 9px;'>🎯 Benchmark: 2% a 5%</p>
                    </div>
                    """, unsafe_allow_html=True)

# ========== FUNÇÃO PARA COMPARAR CAMPANHAS (COM FILTROS DE MEIO E VEÍCULO) ==========
def comparar_campanhas(df):
    """Função para comparar duas campanhas lado a lado com filtros independentes"""
    
    st.markdown("### 📈 COMPARAR CAMPANHAS")
    st.markdown("Selecione duas campanhas e seus respectivos filtros para comparar.")
    
    # Buscar colunas
    camp_cols = [col for col in df.columns if any(x in col.lower() for x in ['campanha', 'campaign'])]
    
    if not camp_cols:
        st.error("❌ Nenhuma coluna de campanha encontrada na planilha!")
        return
    
    col_campanha = camp_cols[0]
    campanhas_disponiveis = sorted(df[col_campanha].dropna().astype(str).unique().tolist())
    
    ano_col = next((col for col in ['Ano da Campanha', 'Ano', 'ano', 'ANO'] if col in df.columns), None)
    mes_col = next((col for col in ['Mês da Análise', 'Mês', 'mes', 'MES'] if col in df.columns), None)
    veic_col = 'Veículo' if 'Veículo' in df.columns else ('Veiculo' if 'Veiculo' in df.columns else None)
    
    # Opções de filtros
    if ano_col:
        anos_disponiveis = ['Todos'] + sorted(df[ano_col].dropna().astype(str).unique().tolist())
    else:
        anos_disponiveis = ['Todos']
    
    if mes_col:
        meses_disponiveis = ['Todos'] + sorted(df[mes_col].dropna().astype(str).unique().tolist())
    else:
        meses_disponiveis = ['Todos']
    
    if 'Meio' in df.columns:
        meios_disponiveis = ['Todos'] + sorted(df['Meio'].dropna().astype(str).unique().tolist())
    else:
        meios_disponiveis = ['Todos']
    
    if veic_col:
        veiculos_disponiveis = ['Todos'] + sorted(df[veic_col].dropna().astype(str).unique().tolist())
    else:
        veiculos_disponiveis = ['Todos']
    
    # ========== SELEÇÃO DAS CAMPANHAS COM FILTROS INDEPENDENTES ==========
    st.markdown("#### 🎯 Seleção de Campanhas e Filtros")
    
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        st.markdown("##### 📌 CAMPANHA 1")
        campanha1 = st.selectbox("Campanha", campanhas_disponiveis, key="camp1")
        
        col_f1a, col_f1b = st.columns(2)
        with col_f1a:
            if ano_col:
                ano1 = st.selectbox("Ano", anos_disponiveis, key="ano1")
            else:
                ano1 = 'Todos'
        with col_f1b:
            if mes_col:
                mes1 = st.selectbox("Mês", meses_disponiveis, key="mes1")
            else:
                mes1 = 'Todos'
        
        col_f1c, col_f1d = st.columns(2)
        with col_f1c:
            meio1 = st.selectbox("Meio", meios_disponiveis, key="meio1")
        with col_f1d:
            veiculo1 = st.selectbox("Veículo", veiculos_disponiveis, key="veiculo1")
    
    with col_c2:
        st.markdown("##### 📌 CAMPANHA 2")
        campanha2 = st.selectbox("Campanha", campanhas_disponiveis, key="camp2")
        
        col_f2a, col_f2b = st.columns(2)
        with col_f2a:
            if ano_col:
                ano2 = st.selectbox("Ano", anos_disponiveis, key="ano2")
            else:
                ano2 = 'Todos'
        with col_f2b:
            if mes_col:
                mes2 = st.selectbox("Mês", meses_disponiveis, key="mes2")
            else:
                mes2 = 'Todos'
        
        col_f2c, col_f2d = st.columns(2)
        with col_f2c:
            meio2 = st.selectbox("Meio", meios_disponiveis, key="meio2")
        with col_f2d:
            veiculo2 = st.selectbox("Veículo", veiculos_disponiveis, key="veiculo2")
    
    # Verificar se selecionou a mesma campanha com os mesmos filtros
    if campanha1 == campanha2 and ano1 == ano2 and mes1 == mes2 and meio1 == meio2 and veiculo1 == veiculo2:
        st.warning("⚠️ Selecione campanhas ou filtros diferentes para comparar!")
        return
    
    # ========== FILTRAR DADOS DA CAMPANHA 1 ==========
    df_camp1 = df[df[col_campanha] == campanha1].copy()
    
    if ano_col and ano1 != 'Todos':
        df_camp1 = df_camp1[df_camp1[ano_col].astype(str).str.strip() == ano1]
    
    if mes_col and mes1 != 'Todos':
        df_camp1 = df_camp1[df_camp1[mes_col].astype(str).str.strip() == mes1]
    
    if 'Meio' in df.columns and meio1 != 'Todos':
        df_camp1 = df_camp1[df_camp1['Meio'] == meio1]
    
    if veic_col and veiculo1 != 'Todos':
        df_camp1 = df_camp1[df_camp1[veic_col] == veiculo1]
    
    # ========== FILTRAR DADOS DA CAMPANHA 2 ==========
    df_camp2 = df[df[col_campanha] == campanha2].copy()
    
    if ano_col and ano2 != 'Todos':
        df_camp2 = df_camp2[df_camp2[ano_col].astype(str).str.strip() == ano2]
    
    if mes_col and mes2 != 'Todos':
        df_camp2 = df_camp2[df_camp2[mes_col].astype(str).str.strip() == mes2]
    
    if 'Meio' in df.columns and meio2 != 'Todos':
        df_camp2 = df_camp2[df_camp2['Meio'] == meio2]
    
    if veic_col and veiculo2 != 'Todos':
        df_camp2 = df_camp2[df_camp2[veic_col] == veiculo2]
    
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
    
    # ========== MONTAR DESCRIÇÃO DOS PERÍODOS ==========
    st.markdown("---")
    col_info1, col_info2 = st.columns(2)
    
    with col_info1:
        periodo1 = f"{campanha1}"
        if ano1 != 'Todos':
            periodo1 += f" | {ano1}"
        if mes1 != 'Todos':
            periodo1 += f" | {mes1}"
        if meio1 != 'Todos':
            periodo1 += f" | {meio1}"
        if veiculo1 != 'Todos':
            periodo1 += f" | {veiculo1}"
        st.caption(f"📊 **Campanha 1:** {periodo1}")
    
    with col_info2:
        periodo2 = f"{campanha2}"
        if ano2 != 'Todos':
            periodo2 += f" | {ano2}"
        if mes2 != 'Todos':
            periodo2 += f" | {mes2}"
        if meio2 != 'Todos':
            periodo2 += f" | {meio2}"
        if veiculo2 != 'Todos':
            periodo2 += f" | {veiculo2}"
        st.caption(f"📊 **Campanha 2:** {periodo2}")
    
    # ========== CARDS DE COMPARAÇÃO COM EFEITO NEON ==========
    st.markdown("### 📊 COMPARAÇÃO DE MÉTRICAS")
    
    # Impacto
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.markdown(f"""
        <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>IMPACTO</p>
            <p style='color: {CORES['turquesa']}; margin: 0; font-size: 28px; font-weight: bold;'>{impacto1:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{campanha1}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>IMPACTO</p>
            <p style='color: {CORES['roxo']}; margin: 0; font-size: 28px; font-weight: bold;'>{impacto2:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{campanha2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        variacao_impacto = ((impacto1 - impacto2) / impacto2 * 100) if impacto2 > 0 else 0
        cor_variacao = CORES['sucesso'] if variacao_impacto > 0 else CORES['erro']
        sinal = "+" if variacao_impacto > 0 else ""
        st.markdown(f"""
        <div class='neon-card' style='background-color: {CORES['cinza_claro']}50; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>VARIAÇÃO</p>
            <p style='color: {cor_variacao}; margin: 0; font-size: 24px; font-weight: bold;'>{sinal}{variacao_impacto:.1f}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Investimento
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.markdown(f"""
        <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>INVESTIMENTO</p>
            <p style='color: {CORES['turquesa']}; margin: 0; font-size: 28px; font-weight: bold;'>R$ {invest1:,.2f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{campanha1}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>INVESTIMENTO</p>
            <p style='color: {CORES['roxo']}; margin: 0; font-size: 28px; font-weight: bold;'>R$ {invest2:,.2f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{campanha2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        variacao_invest = ((invest1 - invest2) / invest2 * 100) if invest2 > 0 else 0
        cor_variacao = CORES['sucesso'] if variacao_invest > 0 else CORES['erro']
        sinal = "+" if variacao_invest > 0 else ""
        st.markdown(f"""
        <div class='neon-card' style='background-color: {CORES['cinza_claro']}50; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>VARIAÇÃO</p>
            <p style='color: {cor_variacao}; margin: 0; font-size: 24px; font-weight: bold;'>{sinal}{variacao_invest:.1f}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Leads
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.markdown(f"""
        <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['turquesa']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>LEADS</p>
            <p style='color: {CORES['turquesa']}; margin: 0; font-size: 28px; font-weight: bold;'>{leads1:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{campanha1}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; border-left: 4px solid {CORES['roxo']}; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>LEADS</p>
            <p style='color: {CORES['roxo']}; margin: 0; font-size: 28px; font-weight: bold;'>{leads2:,.0f}</p>
            <p style='color: {CORES['cinza_escuro']}; margin: 0; font-size: 10px;'>{campanha2}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        variacao_leads = ((leads1 - leads2) / leads2 * 100) if leads2 > 0 else 0
        cor_variacao = CORES['sucesso'] if variacao_leads > 0 else CORES['erro']
        sinal = "+" if variacao_leads > 0 else ""
        st.markdown(f"""
        <div class='neon-card' style='background-color: {CORES['cinza_claro']}50; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
            <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 12px;'>VARIAÇÃO</p>
            <p style='color: {cor_variacao}; margin: 0; font-size: 24px; font-weight: bold;'>{sinal}{variacao_leads:.1f}%</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # CPM e CPL lado a lado
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 📈 CPM (Custo por Mil Impactos)")
        col_cpm1, col_cpm2 = st.columns(2)
        with col_cpm1:
            st.markdown(f"""
            <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{campanha1}</p>
                <p style='color: {CORES['turquesa']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpm1:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
        with col_cpm2:
            st.markdown(f"""
            <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{campanha2}</p>
                <p style='color: {CORES['roxo']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpm2:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("#### 💵 CPL (Custo por Lead)")
        col_cpl1, col_cpl2 = st.columns(2)
        with col_cpl1:
            st.markdown(f"""
            <div class='neon-card neon-card-turquesa' style='background-color: {CORES['turquesa']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{campanha1}</p>
                <p style='color: {CORES['turquesa']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpl1:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
        with col_cpl2:
            st.markdown(f"""
            <div class='neon-card neon-card-roxo' style='background-color: {CORES['roxo']}20; padding: 15px; border-radius: 10px; text-align: center; transition: all 0.3s ease-in-out; cursor: pointer;'>
                <p style='color: {CORES['texto_escuro']}; margin: 0; font-size: 11px;'>{campanha2}</p>
                <p style='color: {CORES['roxo']}; margin: 0; font-size: 24px; font-weight: bold;'>R$ {cpl2:.2f}</p>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ========== GRÁFICOS EM EXPANDERS ==========
    st.markdown("### 📊 GRÁFICOS COMPARATIVOS")
    st.markdown("*Clique em cada métrica para expandir o gráfico*")
    
    # Expander 1: Impacto
    with st.expander("📊 IMPACTO", expanded=False):
        fig_impacto = go.Figure()
        
        fig_impacto.add_trace(go.Bar(
            name=campanha1,
            x=['Impacto'],
            y=[impacto1],
            marker_color=CORES['turquesa'],
            text=[f'{impacto1:,.0f}'],
            textposition='outside'
        ))
        
        fig_impacto.add_trace(go.Bar(
            name=campanha2,
            x=['Impacto'],
            y=[impacto2],
            marker_color=CORES['roxo'],
            text=[f'{impacto2:,.0f}'],
            textposition='outside'
        ))
        
        fig_impacto.update_layout(
            title="Comparação de Impacto",
            barmode='group',
            yaxis_title="Impactos",
            plot_bgcolor='white',
            height=300,
            showlegend=True
        )
        
        st.plotly_chart(fig_impacto, use_container_width=True)
    
    # Expander 2: Investimento
    with st.expander("💰 INVESTIMENTO", expanded=False):
        fig_invest = go.Figure()
        
        fig_invest.add_trace(go.Bar(
            name=campanha1,
            x=['Investimento'],
            y=[invest1],
            marker_color=CORES['turquesa'],
            text=[f'R$ {invest1:,.2f}'],
            textposition='outside'
        ))
        
        fig_invest.add_trace(go.Bar(
            name=campanha2,
            x=['Investimento'],
            y=[invest2],
            marker_color=CORES['roxo'],
            text=[f'R$ {invest2:,.2f}'],
            textposition='outside'
        ))
        
        fig_invest.update_layout(
            title="Comparação de Investimento",
            barmode='group',
            yaxis_title="Investimento (R$)",
            plot_bgcolor='white',
            height=300,
            showlegend=True
        )
        
        st.plotly_chart(fig_invest, use_container_width=True)
    
    # Expander 3: Leads
    with st.expander("🎯 LEADS", expanded=False):
        fig_leads = go.Figure()
        
        fig_leads.add_trace(go.Bar(
            name=campanha1,
            x=['Leads'],
            y=[leads1],
            marker_color=CORES['turquesa'],
            text=[f'{leads1:,.0f}'],
            textposition='outside'
        ))
        
        fig_leads.add_trace(go.Bar(
            name=campanha2,
            x=['Leads'],
            y=[leads2],
            marker_color=CORES['roxo'],
            text=[f'{leads2:,.0f}'],
            textposition='outside'
        ))
        
        fig_leads.update_layout(
            title="Comparação de Leads",
            barmode='group',
            yaxis_title="Leads",
            plot_bgcolor='white',
            height=300,
            showlegend=True
        )
        
        st.plotly_chart(fig_leads, use_container_width=True)
    
    # Expander 4: CPM
    with st.expander("📈 CPM (CUSTO POR MIL IMPACTOS)", expanded=False):
        fig_cpm = go.Figure()
        
        fig_cpm.add_trace(go.Bar(
            name=campanha1,
            x=['CPM'],
            y=[cpm1],
            marker_color=CORES['turquesa'],
            text=[f'R$ {cpm1:.2f}'],
            textposition='outside'
        ))
        
        fig_cpm.add_trace(go.Bar(
            name=campanha2,
            x=['CPM'],
            y=[cpm2],
            marker_color=CORES['roxo'],
            text=[f'R$ {cpm2:.2f}'],
            textposition='outside'
        ))
        
        fig_cpm.update_layout(
            title="Comparação de CPM",
            barmode='group',
            yaxis_title="CPM (R$)",
            plot_bgcolor='white',
            height=300,
            showlegend=True
        )
        
        st.plotly_chart(fig_cpm, use_container_width=True)
    
    # Expander 5: CPL
    with st.expander("💵 CPL (CUSTO POR LEAD)", expanded=False):
        fig_cpl = go.Figure()
        
        fig_cpl.add_trace(go.Bar(
            name=campanha1,
            x=['CPL'],
            y=[cpl1],
            marker_color=CORES['turquesa'],
            text=[f'R$ {cpl1:.2f}'],
            textposition='outside'
        ))
        
        fig_cpl.add_trace(go.Bar(
            name=campanha2,
            x=['CPL'],
            y=[cpl2],
            marker_color=CORES['roxo'],
            text=[f'R$ {cpl2:.2f}'],
            textposition='outside'
        ))
        
        fig_cpl.update_layout(
            title="Comparação de CPL",
            barmode='group',
            yaxis_title="CPL (R$)",
            plot_bgcolor='white',
            height=300,
            showlegend=True
        )
        
        st.plotly_chart(fig_cpl, use_container_width=True)
    
    # Tabela comparativa
    st.markdown("---")
    st.markdown("### 📋 TABELA COMPARATIVA")
    
    tabela_comparativa = pd.DataFrame({
        'Métrica': ['Impacto', 'Investimento (R$)', 'Leads', 'CPM (R$)', 'CPL (R$)'],
        periodo1: [f'{impacto1:,.0f}', f'R$ {invest1:,.2f}', f'{leads1:,.0f}', f'R$ {cpm1:.2f}', f'R$ {cpl1:.2f}'],
        periodo2: [f'{impacto2:,.0f}', f'R$ {invest2:,.2f}', f'{leads2:,.0f}', f'R$ {cpm2:.2f}', f'R$ {cpl2:.2f}']
    })
    
    st.dataframe(tabela_comparativa, use_container_width=True, hide_index=True)

# ========== ABA: VISÃO GERAL (COM CARDS CONSOLIDADOS ON/OFF) ==========
def dashboard_visao_geral(df):
    """Dashboard com filtros, cards consolidados ON/OFF, descrições e tabela geral"""
    
    st.markdown("### 🔍 FILTROS")
    
    possiveis_ano = ['Ano da Campanha', 'Ano', 'ano', 'ANO']
    possiveis_mes = ['Mês da Análise', 'Mês', 'mes', 'MES']
    
    col_ano = next((nome for nome in possiveis_ano if nome in df.columns), None)
    col_mes = next((nome for nome in possiveis_mes if nome in df.columns), None)
    
    col_f1, col_f2, col_f3, col_f4, col_f5, col_f6 = st.columns(6)
    
    with col_f1:
        if col_ano:
            anos = ['Todos'] + sorted(df[col_ano].dropna().astype(str).str.strip().unique().tolist())
            ano_sel = st.selectbox("📅 Ano", anos, key="filtro_ano")
        else:
            ano_sel = st.selectbox("📅 Ano", ['Todos'], key="filtro_ano")
    
    with col_f2:
        if col_mes:
            # Limpar, remover nulos e pegar valores únicos
            meses_limpos = df[col_mes].dropna().astype(str).str.strip().unique().tolist()
            
            # Ordenação correta dos meses
            ordem_meses = {
                'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4,
                'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
                'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12
            }
            
            meses_ordenados = sorted(meses_limpos, key=lambda x: ordem_meses.get(x, 99))
            meses = ['Todos'] + meses_ordenados
            mes_sel = st.selectbox("📆 Mês", meses, key="filtro_mes")
        else:
            mes_sel = st.selectbox("📆 Mês", ['Todos'], key="filtro_mes")
    
    with col_f3:
        camp_cols = [col for col in df.columns if 'campanha' in col.lower() or 'campaign' in col.lower()]
        if camp_cols:
            camps = ['Todas'] + df[camp_cols[0]].dropna().astype(str).str.strip().unique().tolist()
            camp_sel = st.selectbox("🎯 Campanha", camps, key="filtro_campanha")
        else:
            camp_sel = st.selectbox("🎯 Campanha", ['Todas'], key="filtro_campanha")
    
    with col_f4:
        if 'Meio' in df.columns:
            meios = ['Todos'] + df['Meio'].dropna().astype(str).str.strip().unique().tolist()
            meio_sel = st.selectbox("📢 Meio", meios, key="filtro_meio")
        else:
            meio_sel = st.selectbox("📢 Meio", ['Todos'], key="filtro_meio")
    
    with col_f5:
        veic_col = 'Veículo' if 'Veículo' in df.columns else ('Veiculo' if 'Veiculo' in df.columns else None)
        if veic_col:
            veics = ['Todos'] + df[veic_col].dropna().astype(str).str.strip().unique().tolist()
            veic_sel = st.selectbox("🚗 Veículo", veics, key="filtro_veiculo")
        else:
            veic_sel = st.selectbox("🚗 Veículo", ['Todos'], key="filtro_veiculo")
    
    with col_f6:
        categorias_opcoes = ['Todos', 'Mídia ON', 'Mídia OFF', 'Orgânico']
        cat_sel = st.selectbox("🏷️ Categoria", categorias_opcoes, key="filtro_categoria")
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    if col_ano and ano_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[col_ano].astype(str).str.strip() == ano_sel]
    
    if col_mes and mes_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[col_mes].astype(str).str.strip() == mes_sel]
    
    if camp_cols and camp_sel != 'Todas':
        df_filtrado = df_filtrado[df_filtrado[camp_cols[0]].astype(str).str.strip() == camp_sel]
    
    if 'Meio' in df.columns and meio_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Meio'].astype(str).str.strip() == meio_sel]
    
    if veic_col and veic_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[veic_col].astype(str).str.strip() == veic_sel]
    
    if cat_sel != 'Todos':
        if cat_sel == 'Mídia ON':
            meios_filtro = ['Patrocinado', 'Portal']
        elif cat_sel == 'Orgânico':
            meios_filtro = ['Orgânico']
        elif cat_sel == 'Mídia OFF':
            meios_filtro = ['TV', 'Rádio', 'OOH', 'Revista']
        
        df_filtrado = df_filtrado[df_filtrado['Meio'].isin(meios_filtro)]
    
    st.markdown("---")
    
    # ========== CARDS CONSOLIDADOS ON/OFF ==========
    if not df_filtrado.empty:
        criar_cards_consolidados(df_filtrado)
    else:
        st.warning("Nenhum dado encontrado com os filtros selecionados.")
    
    st.markdown("---")
    
    # ========== DESCRIÇÕES DAS MÉTRICAS ==========
    st.markdown("### 📘 Entendendo as Métricas")
    
    col_desc1, col_desc2, col_desc3 = st.columns(3)
    
    with col_desc1:
        st.markdown(f"""
        <div style='background-color: #f8f9fa; padding: 15px; border-radius: 10px; height: 150px;'>
            <h5 style='color: {CORES['turquesa']}; margin: 0;'>📊 IMPACTO TOTAL</h5>
            <p style='font-size: 12px; color: #666; margin-top: 5px;'>
                Soma de todos os impactos (ON + OFF + Orgânico).<br>
                <strong>Quanto maior, melhor o alcance total.</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc2:
        st.markdown(f"""
        <div style='background-color: #f8f9fa; padding: 15px; border-radius: 10px; height: 150px;'>
            <h5 style='color: {CORES['laranja']}; margin: 0;'>📺 MÍDIA OFF</h5>
            <p style='font-size: 12px; color: #666; margin-top: 5px;'>
                <strong>Investimento:</strong> TV, Rádio, OOH e Revista.<br>
                <strong>CPM OFF:</strong> Custo por mil impactos da mídia OFF.
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc3:
        st.markdown(f"""
        <div style='background-color: #f8f9fa; padding: 15px; border-radius: 10px; height: 150px;'>
            <h5 style='color: {CORES['roxo']}; margin: 0;'>💻 MÍDIA ON</h5>
            <p style='font-size: 12px; color: #666; margin-top: 5px;'>
                <strong>Investimento:</strong> Patrocinado e Portal (Digital).<br>
                <strong>CPM ON:</strong> Custo por mil impactos da mídia ON.
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
        <p style='color: white; margin: 0;'>Análise de Campanhas</p>
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

# ========== ÁREA PRINCIPAL COM 2 ABAS ==========
if st.session_state.df is not None:
    df = st.session_state.df
    
    # Criar 2 abas
    aba1, aba2 = st.tabs(["📊 VISÃO GERAL", "📈 COMPARAR CAMPANHAS"])
    
    with aba1:
        dashboard_visao_geral(df)
    
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
                <span style='background-color: {CORES['turquesa']}; color: white; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>ON</span>
                <span style='background-color: {CORES['laranja']}; color: white; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>OFF</span>
                <span style='background-color: {CORES['azul']}; color: white; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>TOTAL</span>
                <span style='background-color: {CORES['verde_escuro']}; color: white; padding: 5px 15px; border-radius: 20px; margin: 0 5px;'>Leads</span>
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
    <span style='color: {CORES['azul']};'>CPM Total</span> • 
    <span>v24.0 - Build Corrigida</span>
</div>
""", unsafe_allow_html=True)