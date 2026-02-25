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

# ========== CONFIGURAÇÕES DO AZURE (via Streamlit Secrets) ==========
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
    page_title="Dashboard Cocred",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ========== CSS PERSONALIZADO COM SUPORTE A DARK THEME ==========
st.markdown(f"""
<style>
    /* Cabeçalhos - se adaptam ao tema */
    h1, h2, h3 {{
        color: {CORES['verde_escuro']} !important;
    }}
    /* No dark theme, os cabeçalhos ficam mais claros */
    @media (prefers-color-scheme: dark) {{
        h1, h2, h3 {{
            color: {CORES['verde_claro']} !important;
        }}
    }}
    
    /* Métricas - funcionam em ambos os temas */
    .stMetric {{
        background-color: rgba(255, 255, 255, 0.1);
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid {CORES['turquesa']};
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        backdrop-filter: blur(5px);
    }}
    
    /* Botões */
    .stButton button {{
        background-color: {CORES['turquesa']};
        color: white;
        border: none;
        border-radius: 5px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s;
    }}
    .stButton button:hover {{
        background-color: {CORES['roxo']};
        color: white;
    }}
    
    /* Botão primário (Excel Online) */
    .stLinkButton button {{
        background: linear-gradient(135deg, {CORES['turquesa']}, {CORES['roxo']});
        color: white;
        font-size: 18px;
        padding: 15px;
        border-radius: 10px;
        border: none;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
    
    /* Cards de informação - adaptáveis */
    .info-card {{
        background-color: rgba(255, 255, 255, 0.05);
        padding: 20px;
        border-radius: 10px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin: 10px 0;
        color: inherit;
    }}
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
    }}
    .stTabs [data-baseweb="tab"] {{
        background-color: rgba(128, 128, 128, 0.1);
        border-radius: 5px 5px 0 0;
        padding: 10px 20px;
        color: inherit;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {CORES['turquesa']} !important;
        color: white !important;
    }}
    
    /* Sidebar */
    .css-1d391kg {{
        background-color: rgba(0, 0, 0, 0.05);
    }}
    
    /* Alertas e mensagens */
    .stAlert {{
        background-color: {CORES['verde_claro']}20;
        border-left-color: {CORES['verde_claro']};
        color: inherit;
    }}
    
    /* Links */
    a {{
        color: {CORES['turquesa']};
        text-decoration: none;
    }}
    a:hover {{
        color: {CORES['roxo']};
    }}
    
    /* Rodapé */
    .footer {{
        color: rgba(128, 128, 128, 0.8);
        font-size: 12px;
        text-align: center;
        padding: 20px;
        border-top: 1px solid rgba(128, 128, 128, 0.2);
    }}
    
    /* Cards das métricas na seção "Entendendo as Métricas" */
    .metric-card {{
        background-color: rgba(255, 255, 255, 0.05);
        padding: 15px;
        border-radius: 10px;
        border: 1px solid rgba(255, 255, 255, 0.1);
        margin: 5px;
        color: inherit;
        height: 150px;
    }}
    .metric-card h5 {{
        margin: 0;
        font-size: 16px;
    }}
    .metric-card p {{
        font-size: 12px;
        margin-top: 5px;
        color: inherit;
    }}
    
    /* Tabelas */
    .stDataFrame {{
        background-color: transparent;
    }}
    
    /* No dark theme, ajustes adicionais */
    @media (prefers-color-scheme: dark) {{
        body {{
            color: #ffffff;
        }}
        .stMarkdown p {{
            color: #cccccc;
        }}
        .metric-card p {{
            color: #cccccc;
        }}
        .st-bb {{
            background-color: #1e1e1e;
        }}
    }}
</style>
""", unsafe_allow_html=True)

# ========== TÍTULO PRINCIPAL ==========
st.markdown(f"""
<div style='text-align: center; padding: 20px; background: linear-gradient(135deg, {CORES['turquesa']}20, {CORES['roxo']}20); border-radius: 15px; margin-bottom: 20px;'>
    <h1 style='color: {CORES['verde_escuro']}; margin-bottom: 0;'>📊 Dashboard Cocred</h1>
    <p style='color: {CORES['texto_escuro']};'>Análise de Campanhas</p>
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

# ========== DASHBOARD DE MÉTRICAS (COM DESCRIÇÕES E DARK THEME) ==========
def dashboard_metricas(df):
    """Dashboard com filtros, cards de métricas, descrições e tabela geral"""
    
    st.markdown("### 🔍 FILTROS")
    
    # Filtros em linha
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    
    with col_f1:
        if 'Ano' in df.columns:
            anos = ['Todos'] + sorted(df['Ano'].astype(str).unique().tolist())
            ano_sel = st.selectbox("Ano", anos, key="filtro_ano")
        else:
            ano_sel = st.selectbox("Ano", ['Todos'], key="filtro_ano")
    
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
    
    if 'Ano' in df.columns and ano_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Ano'].astype(str) == ano_sel]
    
    if camp_cols and camp_sel != 'Todas':
        df_filtrado = df_filtrado[df_filtrado[camp_cols[0]] == camp_sel]
    
    if 'Meio' in df.columns and meio_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado['Meio'] == meio_sel]
    
    if veic_col and veic_sel != 'Todos':
        df_filtrado = df_filtrado[df_filtrado[veic_col] == veic_sel]
    
    st.markdown("---")
    
    # ========== BIG NUMBERS ==========
    st.markdown("### 📊 BIG NUMBERS")
    
    # Calcula métricas
    col_impacto = next((col for col in ['Impacto', 'impressoes', 'visualizacoes'] if col in df_filtrado.columns), None)
    col_invest = next((col for col in ['Investimento', 'investimento', 'gasto', 'custo'] if col in df_filtrado.columns), None)
    col_leads = next((col for col in ['Leads', 'leads', 'conversoes'] if col in df_filtrado.columns), None)
    
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
    
    # ========== DESCRIÇÕES DAS MÉTRICAS (ADAPTÁVEIS AO TEMA) ==========
    st.markdown("---")
    st.markdown("### 📘 Entendendo as Métricas")
    
    col_desc1, col_desc2, col_desc3, col_desc4, col_desc5 = st.columns(5)
    
    with col_desc1:
        st.markdown(f"""
        <div class='metric-card'>
            <h5 style='color: {CORES['turquesa']};'>IMPACTO</h5>
            <p style='margin-top: 5px;'>
                Número total de impressões ou visualizações da campanha.<br>
                <strong>Quanto maior, melhor o alcance.</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc2:
        st.markdown(f"""
        <div class='metric-card'>
            <h5 style='color: {CORES['roxo']};'>INVESTIMENTO</h5>
            <p style='margin-top: 5px;'>
                Valor total gasto na campanha.<br>
                <strong>Base para cálculo das demais métricas.</strong>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc3:
        st.markdown(f"""
        <div class='metric-card'>
            <h5 style='color: {CORES['verde_escuro']}80;'>CPM</h5>
            <p style='margin-top: 5px;'>
                <strong>Custo por Mil Impressões</strong><br>
                Investimento ÷ Impacto × 1000<br>
                <strong>R$ {cpm:.2f}</strong> por mil visualizações.<br>
                <span style='color: {"#00ff00" if cpm < 10 else "#ffff00" if cpm < 20 else "#ff0000"};'>
                    {'✅ Excelente' if cpm < 10 else '⚠️ Médio' if cpm < 20 else '🔴 Alto'}
                </span>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc4:
        st.markdown(f"""
        <div class='metric-card'>
            <h5 style='color: {CORES['verde_claro']};'>LEADS</h5>
            <p style='margin-top: 5px;'>
                Número total de leads gerados.<br>
                <strong>Total: {leads:,.0f} leads</strong><br>
                Taxa de conversão: {(leads/impacto*100) if impacto>0 else 0:.2f}%
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    with col_desc5:
        st.markdown(f"""
        <div class='metric-card'>
            <h5 style='color: {CORES['cinza_escuro']}80;'>CPL</h5>
            <p style='margin-top: 5px;'>
                <strong>Custo por Lead</strong><br>
                Investimento ÷ Leads<br>
                <strong>R$ {cpl:.2f}</strong> por lead.<br>
                <span style='color: {"#00ff00" if cpl < 5 else "#ffff00" if cpl < 10 else "#ff0000"};'>
                    {'✅ Bom' if cpl < 5 else '⚠️ Médio' if cpl < 10 else '🔴 Alto'}
                </span>
            </p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # ========== TABELA GERAL ==========
    st.markdown("### 📋 TABELA GERAL")
    st.dataframe(df_filtrado, use_container_width=True, height=400)
    
    csv = df_filtrado.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download CSV (filtrado)",
        data=csv,
        file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )

# ========== ANÁLISE TEMPORAL ==========
def analise_temporal(df):
    """Análise ao longo do tempo com suporte a 'mês da análise'"""
    st.subheader("📈 Análise Temporal")
    
    with st.expander("📋 Ver colunas disponíveis"):
        st.write("Colunas no DataFrame:", df.columns.tolist())
    
    # Identifica colunas de data
    date_cols = []
    
    if 'mês da análise' in df.columns:
        date_cols.append('mês da análise')
        st.success("✅ Coluna 'mês da análise' encontrada!")
    
    for col in df.columns:
        if col not in date_cols:
            if any(x in col.lower() for x in ['data', 'date', 'mês', 'mes', 'ano', 'year']):
                date_cols.append(col)
            else:
                try:
                    pd.to_datetime(df[col])
                    date_cols.append(col)
                except:
                    pass
    
    if not date_cols:
        st.error("Nenhuma coluna de data encontrada!")
        col_manual = st.selectbox("Selecione manualmente:", df.columns.tolist())
        if col_manual:
            date_cols = [col_manual]
    
    if not date_cols:
        return
    
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    
    if not numeric_cols:
        st.warning("Não há colunas numéricas para análise temporal.")
        return
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        default_data_col = 'mês da análise' if 'mês da análise' in date_cols else date_cols[0]
        data_col = st.selectbox("Coluna de data/mês:", date_cols, index=date_cols.index(default_data_col))
    
    with col2:
        metrica = st.selectbox("Métrica a analisar:", numeric_cols, key="temp_metrica")
    
    with col3:
        periodo = st.selectbox("Agrupar por:", ['Mês', 'Trimestre', 'Semestre', 'Ano'])
    
    df_temp = df.copy()
    
    try:
        if data_col == 'mês da análise':
            df_temp['data_analise'] = pd.to_datetime(df_temp[data_col], errors='coerce')
            
            if df_temp['data_analise'].isna().all():
                df_temp['mes_extraido'] = df_temp[data_col].str.extract(r'([A-Za-zç]+)')
                df_temp['ano_extraido'] = df_temp[data_col].str.extract(r'(\d{4})')
                
                meses_map = {
                    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'abril': 4,
                    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8,
                    'setembro': 9, 'outubro': 10, 'novembro': 11, 'dezembro': 12
                }
                
                df_temp['mes_num'] = df_temp['mes_extraido'].str.lower().map(meses_map)
                df_temp['data_analise'] = pd.to_datetime(
                    df_temp['ano_extraido'].astype(str) + '-' + 
                    df_temp['mes_num'].astype(str) + '-01', 
                    errors='coerce'
                )
        else:
            df_temp['data_analise'] = pd.to_datetime(df_temp[data_col], errors='coerce')
        
        df_temp = df_temp.dropna(subset=['data_analise'])
        
        if len(df_temp) == 0:
            st.error("Não foi possível converter a coluna selecionada para data.")
            return
            
    except Exception as e:
        st.error(f"Erro ao processar datas: {str(e)}")
        return
    
    if periodo == 'Mês':
        df_temp['periodo'] = df_temp['data_analise'].dt.to_period('M').astype(str)
        titulo = f"Evolução Mensal de {metrica}"
    elif periodo == 'Trimestre':
        df_temp['periodo'] = df_temp['data_analise'].dt.to_period('Q').astype(str)
        titulo = f"Evolução Trimestral de {metrica}"
    elif periodo == 'Semestre':
        df_temp['periodo'] = df_temp['data_analise'].dt.to_period('2Q').astype(str)
        titulo = f"Evolução Semestral de {metrica}"
    else:
        df_temp['periodo'] = df_temp['data_analise'].dt.year
        titulo = f"Evolução Anual de {metrica}"
    
    temporal = df_temp.groupby('periodo')[metrica].sum().reset_index()
    temporal = temporal.sort_values('periodo')
    
    fig = px.line(
        temporal,
        x='periodo',
        y=metrica,
        title=titulo,
        markers=True,
        color_discrete_sequence=[CORES['turquesa']]
    )
    fig.update_layout(**PLOTLY_TEMA['layout'])
    st.plotly_chart(fig, use_container_width=True)
    
    if periodo == 'Mês':
        st.markdown("---")
        st.subheader("📅 Análise Mensal Detalhada")
        
        fig_mensal = px.bar(
            temporal,
            x='periodo',
            y=metrica,
            title=f"Comparativo Mensal de {metrica}",
            color_discrete_sequence=[CORES['roxo']],
            text_auto=True
        )
        fig_mensal.update_layout(**PLOTLY_TEMA['layout'])
        st.plotly_chart(fig_mensal, use_container_width=True)
        
        st.dataframe(temporal, use_container_width=True)
        
        st.markdown("### 📊 Estatísticas Mensais")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        
        with col_m1:
            st.metric("Média Mensal", f"{temporal[metrica].mean():,.0f}")
        with col_m2:
            st.metric("Total do Período", f"{temporal[metrica].sum():,.0f}")
        with col_m3:
            melhor_mes = temporal.loc[temporal[metrica].idxmax(), 'periodo']
            st.metric("Melhor Mês", melhor_mes)
        with col_m4:
            pior_mes = temporal.loc[temporal[metrica].idxmin(), 'periodo']
            st.metric("Pior Mês", pior_mes)
        
        st.markdown("### 📈 Crescimento Mês a Mês")
        temporal['crescimento'] = temporal[metrica].pct_change() * 100
        temporal['crescimento_formatado'] = temporal['crescimento'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
        
        crescimento_df = temporal[['periodo', metrica, 'crescimento_formatado']].rename(columns={
            'periodo': 'Mês',
            metrica: 'Valor',
            'crescimento_formatado': 'Crescimento'
        })
        st.dataframe(crescimento_df, use_container_width=True)
    
    st.markdown("---")
    st.markdown("### 📊 Estatísticas Gerais")
    
    col_est1, col_est2, col_est3, col_est4 = st.columns(4)
    
    with col_est1:
        st.metric("Total do Período", f"{temporal[metrica].sum():,.0f}")
    with col_est2:
        st.metric("Média por Período", f"{temporal[metrica].mean():,.0f}")
    with col_est3:
        st.metric("Melhor Período", temporal.loc[temporal[metrica].idxmax(), 'periodo'])
    with col_est4:
        st.metric("Pior Período", temporal.loc[temporal[metrica].idxmin(), 'periodo'])
    
    csv_temporal = temporal.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Dados Temporais (CSV)",
        data=csv_temporal,
        file_name=f"analise_temporal_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv"
    )

# ========== DEMAIS FUNÇÕES DE ANÁLISE ==========

def analise_comparativa_campanhas(df):
    """Comparativo entre campanhas"""
    st.subheader("📊 Comparativo entre Campanhas")
    
    campaign_cols = [col for col in df.columns if any(x in col.lower() for x in ['campanha', 'campaign', 'nome', 'name'])]
    campaign_col = campaign_cols[0] if campaign_cols else df.columns[0]
    
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    
    if not numeric_cols:
        st.warning("Não há colunas numéricas para análise comparativa.")
        return
    
    st.markdown("### Configure a comparação")
    col1, col2 = st.columns(2)
    
    with col1:
        metrica_principal = st.selectbox("Métrica principal:", numeric_cols, index=0, key="comp_metrica")
    
    with col2:
        top_n = st.slider("Mostrar top N campanhas:", 5, 20, 10)
    
    comparativo = df.groupby(campaign_col).agg({
        metrica_principal: ['sum', 'mean', 'count']
    }).round(2)
    
    comparativo.columns = ['Total', 'Média', 'Contagem']
    comparativo = comparativo.sort_values('Total', ascending=False).head(top_n)
    
    st.markdown(f"### Top {top_n} Campanhas por {metrica_principal}")
    st.dataframe(comparativo, use_container_width=True)
    
    fig = px.bar(
        comparativo.reset_index(),
        x=campaign_col,
        y='Total',
        title=f"Comparativo de {metrica_principal} por Campanha",
        labels={'Total': metrica_principal, campaign_col: 'Campanha'},
        color_discrete_sequence=[CORES['turquesa']]
    )
    fig.update_layout(**PLOTLY_TEMA['layout'])
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("### 🏆 Ranking de Performance")
    ranking = comparativo.reset_index()[[campaign_col, 'Total']].head(5)
    
    cores_ranking = [CORES['turquesa'], CORES['roxo'], CORES['verde_claro'], CORES['verde_escuro'], CORES['cinza_escuro']]
    for idx, (_, row) in enumerate(ranking.iterrows()):
        st.markdown(f"""
        <div style='background-color: {cores_ranking[idx]}20; padding: 10px; border-radius: 5px; margin: 5px 0; border-left: 5px solid {cores_ranking[idx]};'>
            <span style='font-size: 18px; font-weight: bold; color: inherit;'>{idx+1}º {row[campaign_col]}</span>
            <span style='float: right; font-size: 18px; font-weight: bold; color: {cores_ranking[idx]};'>{row['Total']:,.2f}</span>
        </div>
        """, unsafe_allow_html=True)

def tabela_dinamica_interativa(df):
    """Tabela dinâmica configurável"""
    st.subheader("🔄 Tabela Dinâmica Interativa")
    
    st.markdown(f"""
    <div style='background-color: {CORES['turquesa']}10; padding: 15px; border-radius: 10px; margin-bottom: 20px; border-left: 5px solid {CORES['turquesa']};'>
        <p style='margin: 0; color: inherit;'><strong>💡 Como usar:</strong> Selecione as colunas para linhas, colunas e valores.</p>
    </div>
    """, unsafe_allow_html=True)
    
    categorical_cols = df.select_dtypes(include=['object']).columns.tolist()
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    
    if not categorical_cols or not numeric_cols:
        st.warning("Precisa de colunas categóricas e numéricas para criar tabela dinâmica.")
        return
    
    col_conf1, col_conf2, col_conf3 = st.columns(3)
    
    with col_conf1:
        linhas = st.multiselect("Linhas (agrupar por):", categorical_cols, default=[categorical_cols[0]] if categorical_cols else [])
    
    with col_conf2:
        colunas = st.multiselect("Colunas (opcional):", categorical_cols, default=[])
    
    with col_conf3:
        valores = st.selectbox("Valores (métrica):", numeric_cols, index=0 if numeric_cols else None)
        agg_func = st.selectbox("Função de agregação:", ['Soma', 'Média', 'Contagem', 'Máximo', 'Mínimo'])
    
    if linhas and valores:
        agg_map = {'Soma': 'sum', 'Média': 'mean', 'Contagem': 'count', 'Máximo': 'max', 'Mínimo': 'min'}
        
        if colunas:
            pivot = pd.pivot_table(df, values=valores, index=linhas, columns=colunas, aggfunc=agg_map[agg_func], fill_value=0)
        else:
            pivot = df.groupby(linhas)[valores].agg(agg_map[agg_func]).reset_index()
            pivot = pivot.sort_values(valores, ascending=False)
        
        st.markdown("### Resultado")
        st.dataframe(pivot, use_container_width=True, height=400)
        
        csv_pivot = pivot.to_csv().encode('utf-8')
        st.download_button(
            label="📥 Download Tabela Dinâmica (CSV)",
            data=csv_pivot,
            file_name=f"tabela_dinamica_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

def exportar_relatorios(df):
    """Aba para exportação de relatórios"""
    st.subheader("📤 Exportar Relatórios")
    
    st.markdown(f"""
    <div style='background-color: {CORES['roxo']}10; padding: 20px; border-radius: 10px; margin-bottom: 20px; border-left: 5px solid {CORES['roxo']};'>
        <h4 style='margin-top: 0; color: {CORES['roxo']};'>Escolha o formato desejado:</h4>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
        <div style='background-color: rgba(255,255,255,0.05); padding: 20px; border-radius: 10px; text-align: center; border: 1px solid rgba(255,255,255,0.1);'>
            <span style='font-size: 40px;'>📄</span>
            <h4 style='color: {CORES['verde_escuro']}'>PDF</h4>
            <p style='color: gray;'>Relatório executivo</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("📥 Gerar PDF", use_container_width=True):
            with st.spinner("Gerando PDF..."):
                try:
                    pdf = gerar_relatorio_pdf(df)
                    
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
                        mime="application/pdf"
                    )
                except Exception as e:
                    st.error(f"Erro ao gerar PDF: {str(e)}")
    
    with col2:
        st.markdown(f"""
        <div style='background-color: rgba(255,255,255,0.05); padding: 20px; border-radius: 10px; text-align: center; border: 1px solid rgba(255,255,255,0.1);'>
            <span style='font-size: 40px;'>📊</span>
            <h4 style='color: {CORES['verde_escuro']}'>Excel</h4>
            <p style='color: gray;'>Planilha completa</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("📥 Gerar Excel", use_container_width=True):
            with st.spinner("Gerando Excel..."):
                excel_bytes = exportar_excel_completo(df)
                
                st.download_button(
                    label="📥 Clique para baixar Excel",
                    data=excel_bytes.getvalue(),
                    file_name=f"relatorio_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    with col3:
        st.markdown(f"""
        <div style='background-color: rgba(255,255,255,0.05); padding: 20px; border-radius: 10px; text-align: center; border: 1px solid rgba(255,255,255,0.1);'>
            <span style='font-size: 40px;'>📈</span>
            <h4 style='color: {CORES['verde_escuro']}'>CSV</h4>
            <p style='color: gray;'>Dados brutos</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("📥 Gerar CSV", use_container_width=True):
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Clique para baixar CSV",
                data=csv,
                file_name=f"dados_cocred_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv"
            )
    
    with st.expander("🔍 Preview dos dados que serão exportados"):
        st.dataframe(df.head(10), use_container_width=True)

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
        <p style='color: white; margin: 0;'>Relatório de Campanhas</p>
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
    
    # TABS PRINCIPAIS
    tab1, tab2, tab3 = st.tabs([
        "📊 Dashboard de Métricas",
        "📈 Análises Avançadas",
        "ℹ️ Sobre"
    ])
    
    with tab1:
        dashboard_metricas(df)
    
    with tab2:
        sub_tab1, sub_tab2, sub_tab3, sub_tab4 = st.tabs([
            "📊 Comparativo Campanhas",
            "📈 Análise Temporal",
            "🔄 Tabela Dinâmica",
            "📤 Exportar Relatórios"
        ])
        
        with sub_tab1:
            analise_comparativa_campanhas(df)
        
        with sub_tab2:
            analise_temporal(df)
        
        with sub_tab3:
            tabela_dinamica_interativa(df)
        
        with sub_tab4:
            exportar_relatorios(df)
    
    with tab3:
        st.subheader("ℹ️ Sobre o Dashboard")
        
        st.markdown(f"""
        <div style='background: linear-gradient(135deg, {CORES['turquesa']}20, {CORES['roxo']}20); padding: 25px; border-radius: 15px; margin-bottom: 20px;'>
            <h2 style='color: {CORES['verde_escuro']}; margin-top: 0;'>Dashboard Cocred</h2>
            <p style='font-size: 16px; color: inherit;'>Visualização e análise dos dados de campanhas da Cocred, integrado com SharePoint via Microsoft Graph API.</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 📌 Funcionalidades")
        
        col_func1, col_func2 = st.columns(2)
        
        with col_func1:
            st.markdown(f"""
            <div class='info-card'>
                <h4 style='color: {CORES['turquesa']}; margin: 0;'>📊 Dashboard de Métricas</h4>
                <p style='margin: 5px 0 0 0;'>Filtros interativos, cards com KPIs e explicações detalhadas de CPM e CPL.</p>
            </div>
            <div class='info-card'>
                <h4 style='color: {CORES['roxo']}; margin: 0;'>📈 Comparativo entre Campanhas</h4>
                <p style='margin: 5px 0 0 0;'>Ranking de performance, gráficos comparativos e top N campanhas.</p>
            </div>
            <div class='info-card'>
                <h4 style='color: {CORES['verde_escuro']}; margin: 0;'>📅 Análise Temporal</h4>
                <p style='margin: 5px 0 0 0;'>Evolução por mês, trimestre, semestre e ano com estatísticas detalhadas.</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col_func2:
            st.markdown(f"""
            <div class='info-card'>
                <h4 style='color: {CORES['verde_claro']}; margin: 0;'>🔄 Tabela Dinâmica</h4>
                <p style='margin: 5px 0 0 0;'>Configure suas próprias visões com linhas, colunas e funções de agregação.</p>
            </div>
            <div class='info-card'>
                <h4 style='color: {CORES['cinza_escuro']}; margin: 0;'>📤 Exportação</h4>
                <p style='margin: 5px 0 0 0;'>Relatórios em PDF, Excel e CSV com preview dos dados.</p>
            </div>
            <div class='info-card'>
                <h4 style='color: {CORES['turquesa']}; margin: 0;'>🔗 Excel Online</h4>
                <p style='margin: 5px 0 0 0;'>Edição direta no navegador com todas as funcionalidades do Excel.</p>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("### ⚙️ Informações Técnicas")
        
        col_tech1, col_tech2 = st.columns(2)
        
        with col_tech1:
            st.markdown(f"""
            <div class='info-card'>
                <h4 style='color: {CORES['verde_escuro']}; margin-top: 0;'>Tecnologias Utilizadas</h4>
                <ul>
                    <li>🐍 Python 3.12</li>
                    <li>📊 Streamlit</li>
                    <li>🔷 Microsoft Graph API</li>
                    <li>📈 Pandas & Plotly</li>
                    <li>🔐 MSAL (Autenticação)</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        with col_tech2:
            st.markdown(f"""
            <div class='info-card'>
                <h4 style='color: {CORES['roxo']}; margin-top: 0;'>Cores Institucionais</h4>
                <div style='display: flex; gap: 15px; flex-wrap: wrap;'>
                    <div><span style='background-color: {CORES['turquesa']}; width: 20px; height: 20px; display: inline-block; border-radius: 3px;'></span> Turquesa</div>
                    <div><span style='background-color: {CORES['verde_claro']}; width: 20px; height: 20px; display: inline-block; border-radius: 3px;'></span> Verde Claro</div>
                    <div><span style='background-color: {CORES['verde_escuro']}; width: 20px; height: 20px; display: inline-block; border-radius: 3px;'></span> Verde Escuro</div>
                    <div><span style='background-color: {CORES['roxo']}; width: 20px; height: 20px; display: inline-block; border-radius: 3px;'></span> Roxo</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        if st.session_state.file_metadata:
            st.markdown("### 📁 Arquivo Atual")
            
            meta = st.session_state.file_metadata
            modified = meta.get('lastModifiedDateTime', 'N/A')
            if modified != 'N/A':
                modified = datetime.fromisoformat(modified.replace('Z', '+00:00')).strftime('%d/%m/%Y %H:%M')
            
            col_file1, col_file2, col_file3 = st.columns(3)
            
            with col_file1:
                st.markdown(f"""
                <div class='info-card' style='text-align: center;'>
                    <p style='margin: 0; font-size: 14px; color: #666;'>Arquivo</p>
                    <p style='margin: 5px 0 0 0; font-weight: bold;'>{meta.get('name', 'N/A')}</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col_file2:
                st.markdown(f"""
                <div class='info-card' style='text-align: center;'>
                    <p style='margin: 0; font-size: 14px; color: #666;'>Última modificação</p>
                    <p style='margin: 5px 0 0 0; font-weight: bold;'>{modified}</p>
                </div>
                """, unsafe_allow_html=True)
            
            with col_file3:
                st.markdown(f"""
                <div class='info-card' style='text-align: center;'>
                    <p style='margin: 0; font-size: 14px; color: #666;'>Tamanho</p>
                    <p style='margin: 5px 0 0 0; font-weight: bold;'>{int(meta.get('size', 0))/1024:.1f} KB</p>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown(f"""
        <div style='text-align: center; padding: 20px; background-color: rgba(128,128,128,0.05); border-radius: 10px;'>
            <p style='margin: 0; color: {CORES['turquesa']}; font-weight: bold;'>Versão 6.4 - Dark Theme</p>
            <p style='margin: 5px 0 0 0; color: #666; font-size: 14px;'>Desenvolvido para a Cocred • {datetime.now().strftime('%Y')}</p>
            <p style='margin: 5px 0 0 0; color: #999; font-size: 12px;'>Integração com SharePoint via Microsoft Graph API</p>
        </div>
        """, unsafe_allow_html=True)

else:
    # Tela inicial
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown(f"""
        <div style='background-color: rgba(255,255,255,0.05); padding: 40px; border-radius: 15px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
            <span style='font-size: 60px;'>👋</span>
            <h3 style='color: {CORES['verde_escuro']}'>Bem-vindo ao Dashboard Cocred</h3>
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
            <h3 style='color: {CORES['roxo']}'>Editar Planilha</h3>
            <p style='color: inherit;'>Use o Excel Online para fazer alterações diretamente no navegador.</p>
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
    <span style='color: {CORES['turquesa']}'>Cocred</span> • 
    <span style='color: {CORES['roxo']}'>Relatório de Campanhas</span> • 
    <span>v6.4 - Dark Theme</span>
</div>
""", unsafe_allow_html=True)