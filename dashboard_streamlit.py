import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ─── Configuração da Página ──────────────────────────────────────────────────
st.set_page_config(
    page_title="Mapa Antifraude | Dashboard UX Premium",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CSS Customizado (Premium Dark Mode Aesthetic) ──────────────────────────
st.markdown("""
<style>
    /* Mudança de Fonte para Montserrat (Premium, Clean) */
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Montserrat', sans-serif !important;
    }

    /* Esconder elementos padrões do Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Top padding adjustment */
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 2rem !important;
    }

    /* Custom KPI Cards HTML */
    .kpi-container {
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        margin-bottom: 2.5rem;
    }
    
    .kpi-card {
        flex: 1;
        background: linear-gradient(145deg, rgba(30,41,59,0.7) 0%, rgba(15,23,42,0.9) 100%);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.3), 0 4px 6px -2px rgba(0, 0, 0, 0.15);
        backdrop-filter: blur(10px);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .kpi-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.4), 0 10px 10px -5px rgba(0, 0, 0, 0.2);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    
    .kpi-title {
        color: #94A3B8;
        font-size: 0.85rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 0.5rem;
    }
    
    .kpi-value {
        color: #F8FAFC;
        font-size: 2.2rem;
        font-weight: 800;
        line-height: 1.2;
    }
    
    .kpi-subtitle {
        color: #64748B;
        font-size: 0.8rem;
        margin-top: 0.5rem;
        font-weight: 500;
    }

    /* Highlight classes */
    .text-pink { color: #EC4899; }
    .text-indigo { color: #818cf8; }
    .text-emerald { color: #10B981; }

    /* CSS Tooltip Customizado para KPIs */
    .kpi-tooltip {
      position: relative;
      display: inline-block;
      cursor: help;
      margin-left: 6px;
      font-size: 0.9rem;
    }
    
    .kpi-tooltip .kpi-tooltiptext {
      visibility: hidden;
      width: 220px;
      background-color: #0F172A;
      color: #F8FAFC;
      text-align: center;
      border-radius: 8px;
      padding: 10px;
      position: absolute;
      z-index: 100;
      bottom: 130%; 
      left: 50%;
      margin-left: -110px;
      opacity: 0;
      transition: opacity 0.2s, transform 0.2s;
      transform: translateY(5px);
      font-size: 0.75rem;
      border: 1px solid #334155;
      font-weight: 500;
      text-transform: none;
      line-height: 1.4;
      box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.5);
    }
    
    .kpi-tooltip:hover .kpi-tooltiptext {
      visibility: visible;
      opacity: 1;
      transform: translateY(0);
    }
</style>
""", unsafe_allow_html=True)

# ─── Funções de Carregamento e Tratamento de Dados ───────────────────────────

@st.cache_data
def load_data():
    try:
        df = pd.read_excel('relatorio_chargeback_consolidado_novo.xlsx', sheet_name='Consolidado Completo')
        
        # Filtro de Data Preciso
        if 'Data de lançamento' in df.columns:
            df['Data_Real'] = pd.to_datetime(df['Data de lançamento'], errors='coerce')
            df['Data_MesAno'] = df['Data_Real'].dt.to_period('M').astype(str)
        else:
            df['Data_Real'] = pd.to_datetime('2026-01-01')
            df['Data_MesAno'] = 'Desconhecido'

        # Calcular a Data da Venda Original (Mês de Referência)
        col_doc = 'Data do documento' if 'Data do documento' in df.columns else 'Data de lançamento'
        df['Data_Venda'] = pd.to_datetime(df.get('adyen_payment_date'), errors='coerce').combine_first(
            pd.to_datetime(df.get('cs_data_finalizacao'), errors='coerce')
        ).combine_first(
            pd.to_datetime(df.get(col_doc), errors='coerce')
        )
        df['Data_Venda_MesAno'] = df['Data_Venda'].dt.to_period('M').astype(str)
            
        df['Valor_Float'] = pd.to_numeric(df['Valor/MR'].fillna(0), errors='coerce').fillna(0)
        
        # Tipagem de Chargeback
        def classificar_tipo(motivo):
            if pd.isna(motivo):
                return 'Não Classificado'
            motivo = str(motivo).lower()
            fraude_keywords = ['fraud', 'no cardholder authorisation', 'card absent']
            if any(k in motivo for k in fraude_keywords):
                return 'Fraude Confirmada'
            return 'Desacordo Comercial'
            
        df['Tipo de Chargeback'] = df['adyen_dispute_reason'].apply(classificar_tipo)
        
        # Categorização SAP Avançada por Denominação
        col_denom = next((c for c in df.columns if 'Denomin' in c and 'objeto' not in c), None)
        if col_denom:
            def categorizar_sap(texto):
                if pd.isna(texto): return 'Não Identificado'
                t = str(texto).upper().strip()
                
                # Regras baseadas no padrão SAP mencionado
                if 'RE DISPUTA DE CHARGEBACK' in t:
                    return 'Re-disputa de Chargeback'
                elif 'REF CREDITO CHARGEBACK' in t or 'REF CRÉDITO CHARGEBACK' in t:
                    return 'Ref Crédito Chargeback'
                elif 'ESTORNO DE CHARGEBACK' in t or 'ESTORNO CHARGEBACK' in t:
                    return 'Estorno de Chargeback'
                elif 'REF CHARGEBACK' in t:
                    return 'Ref Chargeback'
                elif 'CASHBACK' in t:
                    return 'Cashback'
                elif 'CHARGEBACK' in t:
                    return 'Chargeback'
                
                # Fallback genérico: pegar texto antes de hífen ou números (ex: "Chargeback-123")
                import re
                parts = re.split(r'[-0-9]', t)
                if parts and parts[0].strip():
                    return parts[0].strip().title()
                return 'Outros Lançamentos'
                
            df['Categoria SAP'] = df[col_denom].apply(categorizar_sap)
        else:
            df['Categoria SAP'] = 'Não Identificado'
        
        # Fill NA para filtros e gráficos
        df['vtex_forma_entrega'] = df['vtex_forma_entrega'].fillna('Sem Registro')
        df['vtex_cidade'] = df['vtex_cidade'].fillna('Sem Registro')
        df['vtex_uf'] = df['vtex_uf'].fillna('NC')
        df['vtex_tipo_produto'] = df['vtex_tipo_produto'].fillna('Sem Registro')
        df['adyen_record_type'] = df['adyen_record_type'].fillna('Indefinido')
        df['adyen_dispute_reason'] = df['adyen_dispute_reason'].fillna('Sem Motivo Adyen')
        df['cs_status_chargeback'] = df['cs_status_chargeback'].fillna('Não Analisado')
        
        return df
    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame()

df_raw = load_data()

if df_raw.empty:
    st.warning("O arquivo `relatorio_chargeback_consolidado_novo.xlsx` não foi encontrado ou está vazio.")
    st.stop()

# ─── Header e Título ──────────────────────────────────────────────
st.markdown("<h1 style='font-weight: 800; color: #F8FAFC; margin-bottom: 0;'>🛡️ Painel de Controle de Chargebacks</h1>", unsafe_allow_html=True)
st.markdown("<p style='color: #94A3B8; font-size: 1.05rem; margin-bottom: 1.5rem; font-weight: 400;'>Análise cruzada de Fraude e Desacordos Comerciais.</p>", unsafe_allow_html=True)

# ─── Filtros (Topo da Página) ────────────────────────────────────────────────
col_header1, col_header2 = st.columns([8, 2])
with col_header1:
    st.markdown("<h4 style='color: #EC4899; font-weight: 600; margin-bottom: 0.5rem;'>Filtros de Análise</h4>", unsafe_allow_html=True)
with col_header2:
    if st.button("🧹 Limpar Filtros", use_container_width=True):
        for key in ['filtro_data', 'filtro_adyen', 'filtro_cat_adyen', 'filtro_sap']:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

col_f1, col_f2, col_f3, col_f4 = st.columns(4)

with col_f1:
    min_date = df_raw['Data_Real'].min() if not df_raw['Data_Real'].isna().all() else datetime(2025, 1, 1)
    max_date = df_raw['Data_Real'].max() if not df_raw['Data_Real'].isna().all() else datetime.today()
    
    date_range = st.date_input(
        "📅 Período (Lançamento):",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
        key='filtro_data'
    )

with col_f2:
    tipos_adyen = sorted(df_raw['adyen_record_type'].unique())
    tipo_selecionado = st.multiselect(
        "🏷️ Tipo Adyen:", 
        options=tipos_adyen,
        placeholder="Todos (Clique para filtrar)",
        key='filtro_adyen'
    )

with col_f3:
    tipos_cb = sorted(df_raw['Tipo de Chargeback'].unique())
    tipo_cb_selecionado = st.multiselect(
        "🚨 Categoria Adyen:", 
        options=tipos_cb,
        placeholder="Todas (Clique para filtrar)",
        key='filtro_cat_adyen'
    )

with col_f4:
    categorias_sap = sorted(df_raw['Categoria SAP'].unique())
    default_sap = ['Chargeback'] if 'Chargeback' in categorias_sap else None
    cat_sap_selecionada = st.multiselect(
        "📑 Denominação SAP:", 
        options=categorias_sap,
        default=default_sap,
        placeholder="Todas (Clique para filtrar)",
        key='filtro_sap'
    )

# Validação do date_range
if len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date.date(), max_date.date()

# Aplicação dos Filtros
mask_date = (df_raw['Data_Real'].dt.date >= start_date) & (df_raw['Data_Real'].dt.date <= end_date) if not df_raw['Data_Real'].isna().all() else True

df = df_raw[
    mask_date &
    (df_raw['adyen_record_type'].isin(tipo_selecionado if tipo_selecionado else tipos_adyen)) &
    (df_raw['Tipo de Chargeback'].isin(tipo_cb_selecionado if tipo_cb_selecionado else tipos_cb)) &
    (df_raw['Categoria SAP'].isin(cat_sap_selecionada if cat_sap_selecionada else categorias_sap))
]

# Helpers de Formatação
def fmt_currency(val):
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_number(val):
    return f"{val:,}".replace(",", ".")

# ─── KPIs Customizados ──────────────────────────────────────────────
total_valor = df['Valor_Float'].sum()
total_registros = len(df)
fraude_valor = df[df['Tipo de Chargeback'] == 'Fraude Confirmada']['Valor_Float'].sum()
comercial_valor = df[df['Tipo de Chargeback'] == 'Desacordo Comercial']['Valor_Float'].sum()
ticket_medio = total_valor / total_registros if total_registros > 0 else 0

# HTML dos KPIs Premium
kpi_html = f"""
<div class="kpi-container">
    <div class="kpi-card">
        <div class="kpi-title">
            Volume Bruto Afetado
            <div class="kpi-tooltip">❔<span class="kpi-tooltiptext">Valor total bruto das transações que sofreram alguma disputa no período selecionado.</span></div>
        </div>
        <div class="kpi-value text-indigo">{fmt_currency(total_valor)}</div>
        <div class="kpi-subtitle">Total de {fmt_number(total_registros)} transações</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">
            Prejuízo por Fraude
            <div class="kpi-tooltip">❔<span class="kpi-tooltiptext">Soma financeira perdida confirmada por transações não reconhecidas (cartão clonado, fraude, etc).</span></div>
        </div>
        <div class="kpi-value text-pink">{fmt_currency(fraude_valor)}</div>
        <div class="kpi-subtitle">Perda por cartões roubados</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">
            Desacordo Comercial
            <div class="kpi-tooltip">❔<span class="kpi-tooltiptext">Valores em disputa devido a devoluções, problemas operacionais ou mercadorias danificadas.</span></div>
        </div>
        <div class="kpi-value text-emerald">{fmt_currency(comercial_valor)}</div>
        <div class="kpi-subtitle">Processos adm. e devoluções</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">
            Ticket Médio das Disputas
            <div class="kpi-tooltip">❔<span class="kpi-tooltiptext">O valor financeiro médio de cada chargeback disparado no sistema neste período.</span></div>
        </div>
        <div class="kpi-value">{fmt_currency(ticket_medio)}</div>
        <div class="kpi-subtitle">Média por chargeback</div>
    </div>
</div>
"""
st.markdown(kpi_html, unsafe_allow_html=True)

# Cores do Gráfico
COLORS = {
    'Fraude Confirmada': '#EC4899', # Pink
    'Desacordo Comercial': '#818CF8', # Indigo Light
    'Não Classificado': '#64748B' # Slate
}

# Configuração global de layout dos gráficos Plotly para fonte Montserrat
def apply_premium_layout(fig):
    fig.update_layout(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(color='#94A3B8', family='Montserrat, sans-serif'),
        margin=dict(t=25, l=0, r=0, b=0),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1),
        hoverlabel=dict(bgcolor="#1E293B", font_size=13, font_family="Montserrat, sans-serif")
    )
    fig.update_xaxes(showgrid=False, zeroline=False, showline=True, linecolor='#334155')
    fig.update_yaxes(showgrid=True, gridcolor='#1E293B', zeroline=False)
    return fig

# ─── Linha 1: Gráficos de Evolução (Lançamento vs Venda) ─────────────────────────────
col_evol1, col_evol2 = st.columns(2)

with col_evol1:
    st.subheader("📈 Evolução por Mês de Lançamento (SAP)", help="Mostra o impacto financeiro no mês em que o chargeback ou estorno foi contabilizado no SAP.")
    df_evol = df.groupby(['Data_MesAno', 'Tipo de Chargeback'])['Valor_Float'].sum().reset_index()
    if not df_evol.empty:
        fig_evol = px.bar(df_evol, x='Data_MesAno', y='Valor_Float', color='Tipo de Chargeback',
                          color_discrete_map=COLORS, barmode='group')
        fig_evol.update_traces(hovertemplate='Lançamento: %{x}<br>Valor: R$ %{y:,.2f}<extra></extra>')
        fig_evol = apply_premium_layout(fig_evol)
        fig_evol.update_xaxes(type='category')
        fig_evol.update_layout(xaxis_title="", yaxis_title="")
        st.plotly_chart(fig_evol, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados para o período selecionado.")

with col_evol2:
    st.subheader("🛒 Origem: Mês de Referência da Venda", help="Mostra os mesmos chargebacks do período filtrado, mas reposicionados no mês em que a Venda Original (Pedido) ocorreu.")
    df_venda = df.groupby(['Data_Venda_MesAno', 'Tipo de Chargeback'])['Valor_Float'].sum().reset_index()
    if not df_venda.empty:
        df_venda = df_venda.sort_values('Data_Venda_MesAno')
        fig_venda = px.bar(df_venda, x='Data_Venda_MesAno', y='Valor_Float', color='Tipo de Chargeback',
                           color_discrete_map=COLORS, barmode='group')
        fig_venda.update_traces(hovertemplate='Venda Origem: %{x}<br>Valor: R$ %{y:,.2f}<extra></extra>')
        fig_venda = apply_premium_layout(fig_venda)
        fig_venda.update_xaxes(type='category')
        fig_venda.update_layout(xaxis_title="", yaxis_title="")
        st.plotly_chart(fig_venda, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados de venda na origem.")

st.markdown("<br>", unsafe_allow_html=True)

# ─── Linha 2: Gráficos de Composição (Roscas) ─────────────────────────────────
col_comp1, col_comp2 = st.columns(2)

with col_comp1:
    st.subheader("🎯 Proporção Fraude vs Desacordo", help="Demonstra o percentual financeiro de perdas causadas por fraude em cartões versus devoluções operacionais.")
    df_tipo_comp = df.groupby('Tipo de Chargeback')['Valor_Float'].sum().reset_index()
    if not df_tipo_comp.empty:
        fig_tipo = px.pie(df_tipo_comp, values='Valor_Float', names='Tipo de Chargeback', hole=0.7,
                          color='Tipo de Chargeback', color_discrete_map=COLORS)
        fig_tipo.update_traces(
            textposition='inside', 
            textinfo='percent',
            hovertemplate='%{label}<br>R$ %{value:,.2f}<extra></extra>',
            marker=dict(line=dict(color='#0F172A', width=3))
        )
        fig_tipo = apply_premium_layout(fig_tipo)
        fig_tipo.update_layout(showlegend=False, margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_tipo, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados.")

with col_comp2:
    st.subheader("📑 Distribuição Contábil SAP", help="Agrupa as denominações fiscais cadastradas no SAP, mostrando o equilíbrio entre lançamentos de débito (Chargebacks) e créditos (Estornos).")
    df_sap_comp = df.groupby('Categoria SAP')['Valor_Float'].sum().reset_index()
    if not df_sap_comp.empty:
        df_sap_comp['Valor_Abs'] = df_sap_comp['Valor_Float'].abs()
        fig_sap = px.pie(df_sap_comp, values='Valor_Abs', names='Categoria SAP', hole=0.7, custom_data=['Valor_Float'])
        fig_sap.update_traces(
            textposition='inside', 
            textinfo='percent',
            hovertemplate='%{label}<br>R$ %{customdata[0]:,.2f}<extra></extra>',
            marker=dict(line=dict(color='#0F172A', width=3))
        )
        fig_sap = apply_premium_layout(fig_sap)
        fig_sap.update_layout(showlegend=False, margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_sap, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados.")

st.markdown("<br>", unsafe_allow_html=True)

# ─── Linha 3: Top Produtos e Top Motivos Adyen ──────────────────────────────
col_prod, col_motivo = st.columns(2)

with col_prod:
    st.subheader("📦 Top 10 Produtos com Maior Índice", help="Lista os produtos da VTEX que mais sofrem contestações de pagamento, indicando potenciais alvos prediletos de fraudadores.")
    st.markdown("<span style='color: #64748B; font-weight: 500; font-size: 0.85rem;'>Produtos extraídos via API VTEX</span>", unsafe_allow_html=True)
    
    df_prod = df.groupby(['vtex_tipo_produto', 'Tipo de Chargeback'])['Valor_Float'].sum().reset_index()
    df_prod = df_prod[df_prod['vtex_tipo_produto'] != 'Sem Registro']
    
    if not df_prod.empty:
        prod_totals = df_prod.groupby('vtex_tipo_produto')['Valor_Float'].sum().nlargest(10).index
        df_prod = df_prod[df_prod['vtex_tipo_produto'].isin(prod_totals)]
        
        # Limpar o nome do produto para não ficar gigante no eixo Y
        df_prod['Produto Curto'] = df_prod['vtex_tipo_produto'].apply(lambda x: (str(x)[:45] + '...') if len(str(x)) > 45 else x)
        
        fig_prod = px.bar(df_prod, x='Valor_Float', y='Produto Curto', color='Tipo de Chargeback',
                          color_discrete_map=COLORS, orientation='h')
        
        fig_prod.update_traces(hovertemplate='Valor: R$ %{x:,.2f}<br>Produto: %{customdata}<extra></extra>', customdata=df_prod['vtex_tipo_produto'])
        fig_prod = apply_premium_layout(fig_prod)
        fig_prod.update_layout(xaxis_title="", yaxis_title="", yaxis={'categoryorder':'total ascending'})
        fig_prod.update_yaxes(showgrid=False)
        fig_prod.update_xaxes(showgrid=True, gridcolor='#1E293B')
        st.plotly_chart(fig_prod, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Nenhum produto rastreado no período selecionado.")

with col_motivo:
    st.subheader("⚠️ Top 10 Razões Declaradas (Banco)", help="Lista os códigos/motivos técnicos enviados pelos bancos emissores via Adyen que justificaram o cancelamento.")
    st.markdown("<span style='color: #64748B; font-weight: 500; font-size: 0.85rem;'>Motivos reportados pelo emissor na Adyen</span>", unsafe_allow_html=True)
    
    df_motivo = df.groupby(['adyen_dispute_reason', 'Tipo de Chargeback'])['Valor_Float'].sum().reset_index()
    if not df_motivo.empty:
        motivos_top = df_motivo.groupby('adyen_dispute_reason')['Valor_Float'].sum().nlargest(10).index
        df_motivo = df_motivo[df_motivo['adyen_dispute_reason'].isin(motivos_top)]
        
        fig_motivo = px.bar(df_motivo, x='Valor_Float', y='adyen_dispute_reason', color='Tipo de Chargeback',
                          color_discrete_map=COLORS, orientation='h')
        
        fig_motivo.update_traces(hovertemplate='Valor: R$ %{x:,.2f}<extra></extra>')
        fig_motivo = apply_premium_layout(fig_motivo)
        fig_motivo.update_layout(xaxis_title="", yaxis_title="", yaxis={'categoryorder':'total ascending'})
        fig_motivo.update_yaxes(showgrid=False)
        fig_motivo.update_xaxes(showgrid=True, gridcolor='#1E293B')
        st.plotly_chart(fig_motivo, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados de motivos Adyen.")

st.markdown("<br>", unsafe_allow_html=True)

# ─── Linha 4: Análise Operacional e de Localização ────────────────────────────
col_uf, col_entrega = st.columns(2)

with col_uf:
    st.subheader("📍 Concentração de Perdas (UF)", help="Mapeia os estados do Brasil que concentram o maior volume financeiro de chargebacks, baseado na UF de entrega/faturamento da VTEX.")
    df_uf = df.groupby(['vtex_uf', 'Tipo de Chargeback'])['Valor_Float'].sum().reset_index()
    df_uf = df_uf[df_uf['vtex_uf'] != 'NC']
    if not df_uf.empty:
        uf_totals = df_uf.groupby('vtex_uf')['Valor_Float'].sum().nlargest(10).index
        df_uf = df_uf[df_uf['vtex_uf'].isin(uf_totals)]
        
        fig_uf = px.bar(df_uf, x='vtex_uf', y='Valor_Float', color='Tipo de Chargeback',
                        color_discrete_map=COLORS)
        
        fig_uf.update_traces(hovertemplate='Estado: %{x}<br>Valor: R$ %{y:,.2f}<extra></extra>')
        fig_uf = apply_premium_layout(fig_uf)
        fig_uf.update_layout(xaxis_title="", yaxis_title="", xaxis={'categoryorder':'total descending'})
        st.plotly_chart(fig_uf, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados de estado.")

with col_entrega:
    st.subheader("🚚 Risco por Modalidade de Entrega", help="Cruza as informações logísticas para mostrar o percentual de ocorrências em entregas domiciliares versus retiradas em loja física.")
    df_entrega = df.groupby(['vtex_forma_entrega', 'Tipo de Chargeback'])['Valor_Float'].sum().reset_index()
    df_entrega = df_entrega[df_entrega['vtex_forma_entrega'] != 'Sem Registro']
    
    if not df_entrega.empty:
        fig_entrega = px.bar(df_entrega, y='vtex_forma_entrega', x='Valor_Float', color='Tipo de Chargeback',
                             color_discrete_map=COLORS, orientation='h')
        
        fig_entrega.update_traces(hovertemplate='Valor: R$ %{x:,.2f}<extra></extra>')
        fig_entrega = apply_premium_layout(fig_entrega)
        fig_entrega.update_layout(xaxis_title="", yaxis_title="", yaxis={'categoryorder':'total ascending'})
        fig_entrega.update_yaxes(showgrid=False)
        fig_entrega.update_xaxes(showgrid=True, gridcolor='#1E293B')
        st.plotly_chart(fig_entrega, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados de entrega.")
