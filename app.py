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
    """Formata qualquer valor como percentual arredondado"""
    if pd.isna(valor) or valor == 0:
        return "0%"
    # Converte para percentual (0.15 → 15)
    percentual = valor * 100
    # Arredonda para inteiro
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

# Configuração do tema Plotly com as cores da Cocred
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

# Link direto para o Excel Online
EXCEL_ONLINE_URL = "https://agenciaideatore-my.sharepoint.com/:x:/r/personal/cristini_cordesco_ideatoreamericas_com/_layouts/15/Doc.aspx?sourcedoc=%7B198c1ffa-cc36-4faa-a79f-f041003b786a%7D&action=default"
# ========================================

# ========== CONFIGURAÇÃO DA PÁGINA ==========
st.set_page_config(
    page_title="Dashboard Cocred - Campanhas",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personalizado
st.markdown(f"""
<style>
    h1, h2, h3 {{ color: {CORES['verde_escuro']} !important; }}
    .stMetric {{ background-color: {CORES['branco']}; padding: 15px; border-radius: 10px; border-left: 5px solid {CORES['turquesa']}; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
    .stButton button {{ background-color: {CORES['turquesa']}; color: white; border: none; border-radius: 5px; padding: 10px 20px; font-weight: bold; transition: all 0.3s; }}
    .stButton button:hover {{ background-color: {CORES['roxo']}; }}
    .stLinkButton button {{ background: linear-gradient(135deg, {CORES['turquesa']}, {CORES['roxo']}); color: white; font-size: 18px; padding: 15px; border-radius: 10px; border: none; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
    .footer {{ color: {CORES['cinza_escuro']}; font-size: 12px; text-align: center; padding: 20px; border-top: 1px solid {CORES['cinza_claro']}; }}
    .tooltip {{ position: relative; display: inline-block; cursor: help; }}
    .tooltip .tooltiptext {{ visibility: hidden; width: 200px; background-color: {CORES['verde_escuro']}; color: white; text-align: center; border-radius: 6px; padding: 5px; position: absolute; z-index: 1; bottom: 125%; left: 50%; margin-left: -100px; opacity: 0; transition: opacity 0.3s; }}
    .tooltip:hover .tooltiptext {{ visibility: visible; opacity: 1; }}
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

# ========== FUNÇÕES PARA EXPORTAÇÃO DE RELATÓRIOS ==========
def gerar_relatorio_pdf(df):
    """Gera um relatório PDF com análises"""
    pdf = FPDF()
    pdf.add_page()
    
    # Título
    pdf.set_fill_color(0, 174, 157)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font('Arial', 'B', 20)
    pdf.cell(0, 20, 'Relatório Cocred', 0, 1, 'C', 1)
    pdf.ln(10)
    
    # Data
    pdf.set_text_color(0, 54, 65)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M")}', 0, 1)
    pdf.ln(5)
    
    # Estatísticas gerais
    pdf.set_font('Arial', 'B', 12)
    pdf.set_text_color(0, 174, 157)
    pdf.cell(0, 10, 'Resumo Geral:', 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, f'Total de registros: {len(df)}', 0, 1)
    
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns
    for col in numeric_cols[:3]:
        pdf.cell(0, 10, f'Total {col}: {df[col].sum():,.2f}', 0, 1)
        pdf.cell(0, 10, f'Média {col}: {df[col].mean():,.2f}', 0, 1)
    
    return pdf

def exportar_excel_completo(df):
    """Exporta todos os dados e análises para Excel"""
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

# ========== DASHBOARD DE MÉTRICAS ==========
def dashboard_metricas(df):
    """Dashboard com filtros, cards de métricas, descrições e tabela geral"""
    
    st.markdown("### 🔍 FILTROS")
    
    # Filtros em linha
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        # Busca por "Ano da Campanha"
        possiveis_ano = [
            'Ano da Campanha',
            'Ano', 'ano', 'ANO',
            'Ano da campanha', 'ano da campanha'
        ]
        
        col_ano = None
        for nome in possiveis_ano:
            if nome in df.columns:
                col_ano = nome
                break
        
        if col_ano:
            anos = ['Todos'] + sorted(df[col_ano].astype(str).unique().tolist())
            ano_sel = st.selectbox("Ano", anos, key="filtro_ano")
        else:
            ano_sel = st.selectbox("Ano", ['Todos'], key="filtro_ano")
            st.caption("⚠️ Coluna 'Ano da Campanha' não encontrada")
    
    with col_f2:
        camp_cols = [col for col in df.columns if any(x in col.lower() for x in ['campanha', 'campaign'])]
        if camp_cols:
            camps = ['Todas'] + df[camp_cols[0]].unique().tolist()
            camp_sel = st.selectbox("Campanha", camps, key="filtro_campanha")
        else:
            camp_sel = st.selectbox("Campanha", ['Todas'], key="filtro_campanha")
    
    with col_f3:
        if 'Meio' in df.columns:
            meios = ['Todos'] + df['Meio'].unique().tolist()
            meio_sel = st.selectbox("Meio", meios, key="filtro_meio")
        else:
            meio_sel = st.selectbox("Meio", ['Todos'], key="filtro_meio")
    
    with col_f4:
        veic_col = None
        if 'Veículo' in df.columns:
            veic_col = 'Veículo'
        elif 'Veiculo' in df.columns:
            veic_col = 'Veiculo'
        
        if veic_col:
            veics = ['Todos'] + df[veic_col].unique().tolist()
            veic_sel = st.selectbox("Veículo", veics, key="filtro_veiculo")
        else:
            veic_sel = st.selectbox("Veículo", ['Todos'], key="filtro_veiculo")
    
    # Aplicar filtros
    df_filtrado = df.copy()
    
    if col_ano and ano_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[col_ano].astype(str) == ano_sel]
    
    if camp_cols and camp_sel != 'Todas':
        df_filtrado = df_filtrado[df_filtrado[camp_cols[0]] == camp_sel]
    
    if 'Meio' in df.columns and meio_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Meio'] == meio_sel]
    
    if veic_col and veic_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[veic_col] == veic_sel]
    
    st.markdown("---")
    
    # ========== BIG NUMBERS ==========
    st.markdown("### 📊 BIG NUMBERS")
    
    # Busca por IMPACTO
    possiveis_impacto = [
        'Impacto (impressões e entrega de email)',
        'Impacto', 'impacto', 'IMPACTO',
        'Impressões', 'impressões', 'IMPRESSÕES',
        'Impressoes', 'impressoes', 'IMPRESSOES',
        'Visualizações', 'visualizações', 'VISUALIZAÇÕES',
        'Visualizacoes', 'visualizacoes', 'VISUALIZACOES',
        'views', 'Views', 'VIEWS',
        'alcance', 'Alcance', 'ALCANCE'
    ]
    
    col_impacto = None
    for nome in possiveis_impacto:
        if nome in df_filtrado.columns:
            col_impacto = nome
            break
    
    col_invest = next((col for col in ['Investimento', 'investimento', 'INVESTIMENTO', 'gasto', 'custo'] if col in df_filtrado.columns), None)
    col_leads = next((col for col in ['Leads', 'leads', 'LEADS', 'conversoes', 'conversões'] if col in df_filtrado.columns), None)
    
    impacto = df_filtrado[col_impacto].sum() if col_impacto else 0
    investimento = df_filtrado[col_invest].sum() if col_invest else 0
    leads = df_filtrado[col_leads].sum() if col_leads else 0
    
    cpm = (investimento / impacto * 1000) if impacto > 0 else 0
    cpl = (investimento / leads) if leads > 0 else 0
    
    # Cards
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown(f"""
        <div style='background-color: {CORES['turquesa']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <p style='color: white; margin: 0; font-size: 14px;'>IMPACTO</p>
            <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>{impacto:,.0f}</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style='background-color: {CORES['roxo']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <p style='color: white; margin: 0; font-size: 14px;'>INVESTIMENTO</p>
            <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>R$ {investimento:,.2f}</p>
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
        st.markdown(f"""
        <div style='background-color: {CORES['cinza_escuro']}; padding: 20px; border-radius: 10px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <p style='color: white; margin: 0; font-size: 14px;'>CPL</p>
            <p style='color: white; margin: 0; font-size: 28px; font-weight: bold;'>R$ {cpl:.2f}</p>
        </div>
        """, unsafe_allow_html=True)
    
    # ========== DESCRIÇÕES DAS MÉTRICAS ==========
    st.markdown("---")
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
                Número total de leads gerados.<br>
                <strong>Total: {leads:,.0f} leads</strong><br>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ========== TABELA GERAL ==========
    st.markdown("### 📋 TABELA GERAL")
    
    # Formata colunas de porcentagem na tabela
    df_exibicao = df_filtrado.copy()
    
    # Detecta colunas que parecem ser taxas/percentuais
    for col in df_exibicao.select_dtypes(include=['float64', 'int64']).columns:
        # Se a coluna tem valores entre 0 e 1 (provável percentual)
        if df_exibicao[col].min() >= 0 and df_exibicao[col].max() <= 1:
            # Verifica se o nome da coluna sugere taxa
            if any(palavra in col.lower() for palavra in ['taxa', 'percentual', 'porcentagem', 'ctr', 'conversão', 'abertura', 'clique']):
                df_exibicao[col] = df_exibicao[col].apply(lambda x: formatar_percentual(x))
    
    st.dataframe(df_exibicao, use_container_width=True, height=400)
    
    # ========== EXPORTAÇÃO DE RELATÓRIOS (EM EXPANDER) ==========
    with st.expander("📤 **Exportar Relatórios**", expanded=False):
        st.markdown(f"""
        <div style='background-color: {CORES['roxo']}10; padding: 15px; border-radius: 10px; margin-bottom: 20px; border-left: 5px solid {CORES['roxo']};'>
            <p style='margin: 0; color: {CORES['texto_escuro']};'>Escolha o formato desejado para exportar os dados filtrados:</p>
        </div>
        """, unsafe_allow_html=True)
        
        col_exp1, col_exp2, col_exp3 = st.columns(3)
        
        with col_exp1:
            st.markdown(f"""
            <div style='background-color: white; padding: 15px; border-radius: 10px; text-align: center; border: 1px solid {CORES['cinza_claro']}; margin-bottom: 10px;'>
                <span style='font-size: 30px;'>📄</span>
                <h5 style='color: {CORES['verde_escuro']}; margin: 5px 0;'>PDF</h5>
                <p style='color: gray; font-size: 12px; margin: 0;'>Relatório executivo</p>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("📥 Gerar PDF", key="btn_pdf", use_container_width=True):
                with st.spinner("Gerando PDF..."):
                    try:
                        pdf = gerar_relatorio_pdf(df_filtrado)
                        
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                            pdf.output(tmp_file.name)
                            tmp_file_path = tmp_file.name
                        
                        with open(tmp_file_path, 'rb') as f:
                            pdf_bytes = f.read()
                        
                        os.unlink(tmp_file_path)
                        
                        st.download_button(
                            label="📥 Clique para baixar PDF",
                            data=pdf_bytes,
                            file_name=f"relatorio_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf",
                            key="download_pdf"
                        )
                    except Exception as e:
                        st.error(f"Erro ao gerar PDF: {str(e)}")
        
        with col_exp2:
            st.markdown(f"""
            <div style='background-color: white; padding: 15px; border-radius: 10px; text-align: center; border: 1px solid {CORES['cinza_claro']}; margin-bottom: 10px;'>
                <span style='font-size: 30px;'>📊</span>
                <h5 style='color: {CORES['verde_escuro']}; margin: 5px 0;'>Excel</h5>
                <p style='color: gray; font-size: 12px; margin: 0;'>Planilha completa</p>
            </div>
            """, unsafe_allow_html=True)
            
            if st.button("📥 Gerar Excel", key="btn_excel", use_container_width=True):
                with st.spinner("Gerando Excel..."):
                    excel_bytes = exportar_excel_completo(df_filtrado)
                    
                    st.download_button(
                        label="📥 Clique para baixar Excel",
                        data=excel_bytes.getvalue(),
                        file_name=f"relatorio_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_excel"
                    )
        
        with col_exp3:
            st.markdown(f"""
            <div style='background-color: white; padding: 15px; border-radius: 10px; text-align: center; border: 1px solid {CORES['cinza_claro']}; margin-bottom: 10px;'>
                <span style='font-size: 30px;'>📈</span>
                <h5 style='color: {CORES['verde_escuro']}; margin: 5px 0;'>CSV</h5>
                <p style='color: gray; font-size: 12px; margin: 0;'>Dados brutos</p>
            </div>
            """, unsafe_allow_html=True)
            
            csv = df_filtrado.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Download CSV",
                data=csv,
                file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                key="download_csv",
                use_container_width=True
            )
        
        # Preview dos dados
        with st.expander("🔍 Preview dos dados que serão exportados", expanded=False):
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

# ========== ÁREA PRINCIPAL ==========
if st.session_state.df is not None:
    df = st.session_state.df
    
    # Agora apenas o dashboard de métricas, sem abas
    dashboard_metricas(df)

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
    <span>v7.1 - Com Exportação</span>
</div>
""", unsafe_allow_html=True)