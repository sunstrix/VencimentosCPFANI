import pandas as pd
import streamlit as st
import plotly.express as px
import streamlit.components.v1 as components
import os
import io
import requests
from urllib.parse import urlparse, parse_qs, urlencode, urlunparse

# Configuração inicial da página do Streamlit
st.set_page_config(
    page_title="Dashboard de Controle de Vencimentos",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização customizada e Regras Corrigidas de Impressão (PDF A4)
st.markdown("""
<style>
.main .block-container { padding-top: 2rem; }
div[data-testid="stMetricValue"] { font-size: 28px; font-weight: bold; color: #1E3A8A; }

/* Regras de Otimização para salvar em PDF (A4) */
@media print {
    @page {
        size: A4 portrait;
        margin: 0.8cm;
    }

    /* Esconde barra lateral, cabeçalho e botões */
    [data-testid="stSidebar"], 
    [data-testid="stHeader"], 
    [data-testid="stToolbar"],
    header, 
    footer, 
    .stButton,
    div.stActionButton {
        display: none !important;
    }
    
    /* CORREÇÃO: Zoom moderado (0.75) para melhor legibilidade + controle de alturas */
    html, body, [data-testid="stAppViewContainer"], .main, .block-container {
        zoom: 0.75 !important;
        height: auto !important;
        width: 100% !important;
        overflow: visible !important;
        position: static !important;
    }
    
    .main .block-container {
        max-width: 100% !important;
        padding: 0.3cm !important;
    }

    /* Reduz títulos e métricas proporcionalmente */
    h1 { font-size: 20px !important; margin: 5px 0 !important; }
    h2, h3, h4 { font-size: 14px !important; margin: 3px 0 !important; }
    
    div[data-testid="stMetricValue"] { 
        font-size: 16px !important; 
        font-weight: bold !important;
    }
    
    div[data-testid="stMetricLabel"] {
        font-size: 10px !important;
    }

    /* Força as cores de fundo */
    body, .stApp {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
    }
    
    /* Controla altura dos gráficos para caber na página */
    .stPlotlyChart {
        page-break-inside: avoid !important;
        max-height: 180px !important;
    }
    
    /* Controla altura das tabelas */
    div[data-testid="stDataFrame"] {
        page-break-inside: avoid !important;
        max-height: 200px !important;
        font-size: 8px !important;
    }
    
    /* Ajusta colunas para ficarem mais compactas */
    .stColumns {
        margin-bottom: 5px !important;
    }
}
</style>
""", unsafe_allow_html=True)

def formatar_vencimento_pt(val):
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ['nan', 'none', 'não informado', 'nat'] or val_str == '00:00:00':
        return "Não Informado"
    
    meses_pt = {
        1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
        7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
    }
    try:
        dt = pd.to_datetime(val, errors='coerce')
        if pd.notnull(dt):
            return f"{meses_pt[dt.month]}/{str(dt.year)[-2:]}"
    except:
        pass

    if '/' in val_str and not val_str.replace('/', '').isdigit():
        partes = val_str.split('/')
        if len(partes) == 2:
            return f"{partes[0].capitalize()}/{partes[1]}"
    return val_str

def ordenar_meses_cronologicamente(lista_meses):
    meses_map = {'Jan':1, 'Fev':2, 'Mar':3, 'Abr':4, 'Mai':5, 'Jun':6, 'Jul':7, 'Ago':8, 'Set':9, 'Out':10, 'Nov':11, 'Dez':12}
    def obter_chave(m):
        if '/' in m:
            p, a = m.split('/')
            return (int(a), meses_map.get(p.capitalize(), 0))
        return (99, 0)
    return sorted(lista_meses, key=obter_chave)

# CORREÇÃO: Função reescrita para apenas adicionar download=1 à URL original
# em vez de tentar extrair UniqueId (que é apenas um token de convidado)
def extrair_info_sharepoint(url):
    """
    CORREÇÃO: Converte link de compartilhamento em URL de download direto
    adicionando o parâmetro download=1, sem tentar extrair UniqueId.
    """
    try:
        if not url:
            return None
        
        # Parse da URL
        parsed = urlparse(url)
        
        # Se já é URL de download direto, retornar como está
        if 'download.aspx' in url or 'download=1' in url:
            return url
        
        # Extrair query string existente
        query_params = parse_qs(parsed.query, keep_blank_values=True)
        
        # Adicionar parâmetro download=1
        query_params['download'] = ['1']
        
        # Reconstruir URL com novo parâmetro
        new_query = urlencode(query_params, doseq=True)
        new_url = urlunparse((
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            new_query,
            parsed.fragment
        ))
        
        return new_url
        
    except Exception as e:
        st.error(f"Erro ao converter URL: {str(e)}")
        return None

@st.cache_data(ttl=1800)
def carregar_e_consolidar_dados():
    """
    Download direto via URL pública do SharePoint
    """
    try:
        sharepoint_url = st.secrets["sharepoint"]["sharepoint_url"]
        
        if not sharepoint_url:
            return None, "❌ URL do SharePoint não configurada"
        
        # CORREÇÃO: Converter URL adicionando download=1
        download_url = extrair_info_sharepoint(sharepoint_url)
        
        if not download_url:
            download_url = sharepoint_url
        
        # Headers para simular navegador
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*',
        }
        
        with st.spinner("Baixando arquivo do SharePoint..."):
            try:
                # Download direto com URL corrigida
                response = requests.get(download_url, headers=headers, timeout=30, allow_redirects=True)
                
            except requests.exceptions.RequestException as e:
                return None, f"❌ Erro de conexão: {str(e)}"
            
            if response.status_code != 200:
                return None, f"❌ Erro HTTP {response.status_code}\n\nURL: {download_url}"
            
            content_type = response.headers.get('content-type', '').lower()
            content_length = len(response.content)
            
            # Verificar se é HTML (página de login ou erro)
            if 'text/html' in content_type:
                return None, f"❌ Resposta HTML recebida. O arquivo pode exigir autenticação.\n\nContent-Type: {content_type}\n\n**Solução**: Verifique se o arquivo está compartilhado como 'Qualquer pessoa com o link pode visualizar'."
            
            if content_length < 1000:
                return None, f"❌ Arquivo muito pequeno ({content_length} bytes)"
        
        # Ler Excel
        try:
            file_object = io.BytesIO(response.content)
            xl = pd.ExcelFile(file_object, engine='openpyxl')
        except Exception as e:
            return None, f"❌ Erro ao ler Excel: {str(e)}"
        
        # Obter abas
        abas_lojas = [aba for aba in xl.sheet_names if aba.isdigit()]
        
        if not abas_lojas:
            return None, f"Nenhuma aba numérica. Abas: {', '.join(xl.sheet_names[:10])}"
        
        st.success(f"✅ Arquivo carregado! {len(abas_lojas)} lojas.")
        
        dados_consolidados = []
        for loja in abas_lojas:
            df = xl.parse(loja)
            df.columns = [str(col).strip() for col in df.columns]
            
            if 'Vencimento' in df.columns and 'Codigo' in df.columns:
                df_limpo = df[['Codigo', 'Descrição', 'Qtde.', 'Vencimento']].dropna(subset=['Codigo'])
                df_limpo['Loja'] = str(loja).strip()
                dados_consolidados.append(df_limpo)
            else:
                df_raw = xl.parse(loja, header=None)
                mes_atual = "Não Informado"
                linhas_finais = []
                for idx, row in df_raw.iterrows():
                    if row.isnull().all(): continue
                    val_a = str(row[0]).strip() if pd.notna(row[0]) else ""
                    val_b = str(row[1]).strip() if pd.notna(row[1]) else ""
                    if "/" in val_a and len(val_a) <= 7 and not val_a.isdigit():
                        mes_atual = val_a
                    elif "/" in val_b and len(val_b) <= 7 and val_a == "":
                        mes_atual = val_b
                    if val_a.isdigit():
                        linhas_finais.append({
                            'Codigo': val_a, 'Descrição': val_b,
                            'Qtde.': row[2] if pd.notna(row[2]) else 0,
                            'Vencimento': mes_atual, 'Loja': str(loja).strip()
                        })
                if linhas_finais:
                    dados_consolidados.append(pd.DataFrame(linhas_finais))
        
        if not dados_consolidados:
            return None, "Nenhum dado extraído"
        
        df_final = pd.concat(dados_consolidados, ignore_index=True)
        df_final['Loja'] = df_final['Loja'].astype(str).str.strip()
        df_final['Qtde.'] = pd.to_numeric(df_final['Qtde.'], errors='coerce').fillna(0).astype(int)
        df_final['Codigo'] = df_final['Codigo'].astype(str).str.replace(r'\.0$', '', regex=True)
        df_final['Descrição'] = df_final['Descrição'].astype(str).str.upper().str.strip()
        df_final['Vencimento'] = df_final['Vencimento'].apply(formatar_vencimento_pt)
        df_final = df_final[(df_final['Codigo'] != 'nan') & (df_final['Vencimento'] != 'Não Informado')]
        
        return df_final, None
        
    except Exception as e:
        return None, f"❌ Erro: {type(e).__name__}: {str(e)}"

# Interface Principal
st.title("📊 Dashboard de Controle de Vencimentos — Matriz")
st.markdown("Consolidação automática de dados de todas as lojas para análise de vencimentos.")

df, erro = carregar_e_consolidar_dados()

if erro:
    st.error(erro)
    st.info("💡 **Como obter link de download direto:**\n\n"
            "1. Acesse o arquivo no SharePoint\n"
            "2. Clique em **Compartilhar**\n"
            "3. Selecione **'Qualquer pessoa com o link pode visualizar'**\n"
            "4. Copie o link gerado\n"
            "5. Cole o link nos secrets como `sharepoint_url`")
elif df is None or df.empty:
    st.warning("⚠️ Nenhum dado encontrado.")
else:
    st.sidebar.header("🎯 Filtros de Análise")
    todas_lojas = sorted(df['Loja'].unique(), key=int)
    lojas_selecionadas = st.sidebar.multiselect("Selecione as Lojas:", todas_lojas, default=todas_lojas)
    todos_meses_ordenados = ordenar_meses_cronologicamente(df['Vencimento'].unique())
    meses_selecionados = st.sidebar.multiselect("Selecione os Meses:", todos_meses_ordenados, default=todos_meses_ordenados)
    busca_produto = st.sidebar.text_input("Buscar por Produto:").upper()
    st.sidebar.markdown("---")
    st.sidebar.subheader("🖨️ Exportar Relatório")
    if st.sidebar.button("💾 Salvar Tela em PDF (A4)"):
        components.html("<script>setTimeout(function(){ window.parent.print(); }, 500);</script>", height=0)

    df_filtrado = df[df['Loja'].isin(lojas_selecionadas) & df['Vencimento'].isin(meses_selecionados)]
    if busca_produto:
        df_filtrado = df_filtrado[df_filtrado['Descrição'].str.contains(busca_produto) | df_filtrado['Codigo'].str.contains(busca_produto)]

    df_produto_agrupado = df_filtrado.groupby(['Codigo', 'Descrição'])['Qtde.'].sum().reset_index().sort_values(by='Qtde.', ascending=False)

    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("Total de Itens a Vencer", f"{df_filtrado['Qtde.'].sum():,}")
    with col2: st.metric("Lojas Analisadas", len(df_filtrado['Loja'].unique()))
    with col3: st.metric("Produtos Únicos", len(df_filtrado['Codigo'].unique()))
    with col4: st.metric("Meses com Alertas", len(df_filtrado['Vencimento'].unique()))
    st.markdown("---")

    col_graf1, col_graf2 = st.columns(2)
    with col_graf1:
        st.subheader("🗓️ Volume por Mês")
        df_mes = df_filtrado.groupby('Vencimento')['Qtde.'].sum().reset_index()
        if not df_mes.empty:
            df_mes['Vencimento'] = pd.Categorical(df_mes['Vencimento'], categories=todos_meses_ordenados, ordered=True)
            df_mes = df_mes.sort_values('Vencimento')
            fig_mes = px.bar(df_mes, x='Vencimento', y='Qtde.', text='Qtde.', color_discrete_sequence=['#1E3A8A'], height=200)
            fig_mes.update_traces(textposition='outside', marker=dict(line=dict(width=0)))
            fig_mes.update_layout(plot_bgcolor='rgba(0,0,0,0)', yaxis_title="", xaxis_title="", margin=dict(l=10, r=10, t=20, b=10))
            fig_mes.update_xaxes(type='category')
            st.plotly_chart(fig_mes, use_container_width=True)

    with col_graf2:
        st.subheader("🏪 Top Lojas a Vencer")
        df_loja = df_filtrado.groupby('Loja')['Qtde.'].sum().reset_index()
        if not df_loja.empty:
            df_loja = df_loja.sort_values(by='Qtde.', ascending=False).head(10)
            fig_loja = px.bar(df_loja, x='Loja', y='Qtde.', text='Qtde.', color_discrete_sequence=['#10B981'], height=200)
            fig_loja.update_traces(textposition='outside', marker=dict(line=dict(width=0)))
            fig_loja.update_layout(plot_bgcolor='rgba(0,0,0,0)', yaxis_title="", xaxis_title="", margin=dict(l=10, r=10, t=20, b=10))
            fig_loja.update_xaxes(type='category')
            st.plotly_chart(fig_loja, use_container_width=True)
    st.markdown("---")

    st.subheader("📦 Top 15 Produtos a Vencer")
    if not df_produto_agrupado.empty:
        df_top_produtos = df_produto_agrupado.head(15).copy()
        fig_prod = px.bar(df_top_produtos, x='Qtde.', y='Descrição', orientation='h', text='Qtde.', color_discrete_sequence=['#F59E0B'], height=350)
        fig_prod.update_traces(textposition='outside', marker=dict(line=dict(width=0)))
        fig_prod.update_layout(plot_bgcolor='rgba(0,0,0,0)', yaxis={'categoryorder':'total ascending'}, yaxis_title="", xaxis_title="Quantidade", margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig_prod, use_container_width=True)
    st.markdown("---")

    col_tab1, col_tab2 = st.columns(2)
    with col_tab1:
        st.subheader("🛒 Resumo por Produto")
        if not df_produto_agrupado.empty:
            st.dataframe(df_produto_agrupado, use_container_width=True, hide_index=True, height=250)
    with col_tab2:
        st.subheader("📋 Lista (Loja a Loja)")
        df_exibicao = df_filtrado.copy()
        if not df_exibicao.empty:
            df_exibicao['Vencimento'] = pd.Categorical(df_exibicao['Vencimento'], categories=todos_meses_ordenados, ordered=True)
            df_exibicao = df_exibicao[['Loja', 'Vencimento', 'Codigo', 'Descrição', 'Qtde.']].sort_values(by=['Vencimento', 'Loja'])
            st.dataframe(df_exibicao, use_container_width=True, hide_index=True, height=250)