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
        df = pd.read_excel('relatorio_chargeback_consolidado.xlsx', sheet_name='Consolidado Completo', engine='calamine')
        
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
    st.warning("O arquivo `relatorio_chargeback_consolidado.xlsx` não foi encontrado ou está vazio.")
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
        for key in ['filtro_data', 'filtro_adyen', 'filtro_cat_adyen', 'filtro_sap', 'filtro_entrega']:
            if key in st.session_state:
                del st.session_state[key]
        st.rerun()

col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)

with col_f1:
    min_date = df_raw['Data_Real'].min() if not df_raw['Data_Real'].isna().all() else datetime(2025, 1, 1)
    max_date = df_raw['Data_Real'].max() if not df_raw['Data_Real'].isna().all() else datetime.today()
    
    date_range = st.date_input(
        "📅 Período (Lançamento):",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
        format="DD/MM/YYYY",
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
        "📑 SAP:", 
        options=categorias_sap,
        default=default_sap,
        placeholder="Todas (Clique para filtrar)",
        key='filtro_sap'
    )

with col_f5:
    formas_entrega = sorted([x for x in df_raw['vtex_forma_entrega'].unique() if x != 'Sem Registro' and pd.notna(x)])
    entrega_selecionada = st.multiselect(
        "🚚 Forma de Entrega:", 
        options=formas_entrega,
        placeholder="Todas",
        key='filtro_entrega'
    )

# Validação do date_range
if len(date_range) == 2:
    start_date, end_date = date_range
else:
    start_date, end_date = min_date.date(), max_date.date()

# Aplicação dos Filtros
df = df_raw[
    (df_raw['Data_Real'].dt.date >= start_date) & (df_raw['Data_Real'].dt.date <= end_date)
]
if tipo_selecionado:
    df = df[df['adyen_record_type'].isin(tipo_selecionado)]
if tipo_cb_selecionado:
    df = df[df['Tipo de Chargeback'].isin(tipo_cb_selecionado)]
if cat_sap_selecionada:
    df = df[df['Categoria SAP'].isin(cat_sap_selecionada)]
if entrega_selecionada:
    df = df[df['vtex_forma_entrega'].isin(entrega_selecionada)]

# Helpers de Formatação
def fmt_currency(val):
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_currency_int(val):
    return f"R$ {val:,.0f}".replace(",", ".")

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
        <div class="kpi-value text-indigo">{fmt_currency_int(total_valor)}</div>
        <div class="kpi-subtitle">Total de {fmt_number(total_registros)} transações</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">
            Prejuízo por Fraude
            <div class="kpi-tooltip">❔<span class="kpi-tooltiptext">Soma financeira perdida confirmada por transações não reconhecidas (cartão clonado, fraude, etc).</span></div>
        </div>
        <div class="kpi-value text-pink">{fmt_currency_int(fraude_valor)}</div>
        <div class="kpi-subtitle">Perda por cartões roubados</div>
    </div>
    <div class="kpi-card">
        <div class="kpi-title">
            Desacordo Comercial
            <div class="kpi-tooltip">❔<span class="kpi-tooltiptext">Valores em disputa devido a devoluções, problemas operacionais ou mercadorias danificadas.</span></div>
        </div>
        <div class="kpi-value text-emerald">{fmt_currency_int(comercial_valor)}</div>
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
    df_evol = df.groupby(['Data_MesAno'])['Valor_Float'].sum().reset_index()
    if not df_evol.empty:
        fig_evol = px.bar(df_evol, x='Data_MesAno', y='Valor_Float',
                          color_discrete_sequence=['#818CF8'])
        fig_evol.update_traces(hovertemplate='Lançamento: %{x}<br>Valor: R$ %{y:,.2f}<extra></extra>')
        fig_evol = apply_premium_layout(fig_evol)
        fig_evol.update_xaxes(type='category')
        fig_evol.update_layout(xaxis_title="", yaxis_title="")
        st.plotly_chart(fig_evol, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados para o período selecionado.")

with col_evol2:
    st.subheader("🛒 Origem: Mês de Referência da Venda", help="Mostra os mesmos chargebacks do período filtrado, mas reposicionados no mês em que a Venda Original (Pedido) ocorreu.")
    df_venda = df.groupby(['Data_Venda_MesAno'])['Valor_Float'].sum().reset_index()
    if not df_venda.empty:
        df_venda = df_venda.sort_values('Data_Venda_MesAno')
        fig_venda = px.bar(df_venda, x='Data_Venda_MesAno', y='Valor_Float',
                           color_discrete_sequence=['#818CF8'])
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
        fig_sap = px.pie(df_sap_comp, values='Valor_Abs', names='Categoria SAP', hole=0.7)
        fig_sap.update_traces(
            textposition='inside', 
            textinfo='percent',
            hovertemplate='%{label}<br>R$ %{value:,.2f}<extra></extra>',
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
    
    df_temp = df[df['vtex_tipo_produto'].notna() & (df['vtex_tipo_produto'] != 'Sem Registro')].copy()
    if not df_temp.empty:
        # Truncar o nome para 35 caracteres e agrupar para evitar barras empilhadas no Plotly
        df_temp['Produto_Curto'] = df_temp['vtex_tipo_produto'].apply(lambda x: (x[:35] + '...') if isinstance(x, str) and len(x) > 35 else str(x))
        df_prod = df_temp.groupby('Produto_Curto')['Valor_Float'].sum().reset_index()
        
        df_prod = df_prod.sort_values('Valor_Float', ascending=False).head(10).sort_values('Valor_Float', ascending=True)
        
        fig_prod = px.bar(df_prod, x='Valor_Float', y='Produto_Curto', orientation='h', color_discrete_sequence=['#818cf8'])
        fig_prod.update_traces(texttemplate='R$ %{x:,.0f}', textposition='outside', hovertemplate='Valor: R$ %{x:,.2f}<extra></extra>')
        fig_prod = apply_premium_layout(fig_prod)
        fig_prod.update_layout(xaxis_title="", yaxis_title="", margin=dict(r=50))
        st.plotly_chart(fig_prod, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Nenhum produto rastreado no período selecionado.")

with col_motivo:
    st.subheader("⚠️ Top 10 Razões Declaradas (Banco)", help="Lista os códigos/motivos técnicos enviados pelos bancos emissores via Adyen que justificaram o cancelamento.")
    st.markdown("<span style='color: #64748B; font-weight: 500; font-size: 0.85rem;'>Motivos reportados pelo emissor na Adyen</span>", unsafe_allow_html=True)
    
    df_motivo = df.groupby(['adyen_dispute_reason'])['Valor_Float'].sum().reset_index()
    if not df_motivo.empty:
        df_motivo = df_motivo.sort_values('Valor_Float', ascending=False).head(10).sort_values('Valor_Float', ascending=True)
        fig_motivo = px.bar(df_motivo, x='Valor_Float', y='adyen_dispute_reason', orientation='h', color_discrete_sequence=['#EC4899'])
        fig_motivo.update_traces(texttemplate='R$ %{x:,.0f}', textposition='outside', hovertemplate='Valor: R$ %{x:,.2f}<extra></extra>')
        fig_motivo = apply_premium_layout(fig_motivo)
        fig_motivo.update_layout(xaxis_title="", yaxis_title="", margin=dict(r=50))
        st.plotly_chart(fig_motivo, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados de motivos Adyen.")

st.markdown("<br>", unsafe_allow_html=True)

# ─── Linha 4: Análise Operacional e de Localização ────────────────────────────
col_uf, col_entrega = st.columns(2)

with col_uf:
    st.subheader("📍 Concentração de Perdas (UF)", help="Mapeia os estados do Brasil que concentram o maior volume financeiro de chargebacks, baseado na UF de entrega/faturamento da VTEX.")
    df_uf = df.groupby(['vtex_uf'])['Valor_Float'].sum().reset_index()
    df_uf = df_uf[df_uf['vtex_uf'] != 'NC']
    if not df_uf.empty:
        df_uf = df_uf.sort_values('Valor_Float', ascending=False).head(10)
        fig_uf = px.bar(df_uf, x='vtex_uf', y='Valor_Float', color_discrete_sequence=['#10B981'])
        fig_uf.update_traces(texttemplate='R$ %{y:,.0f}', textposition='outside', hovertemplate='Valor: R$ %{y:,.2f}<extra></extra>')
        fig_uf = apply_premium_layout(fig_uf)
        fig_uf.update_layout(xaxis_title="", yaxis_title="", margin=dict(t=30))
        st.plotly_chart(fig_uf, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados de estado.")

with col_entrega:
    st.subheader("🚚 Risco por Modalidade de Entrega", help="Cruza as informações logísticas para mostrar o percentual de ocorrências em entregas domiciliares versus retiradas em loja física.")
    df_entrega = df.groupby(['vtex_forma_entrega'])['Valor_Float'].sum().reset_index()
    df_entrega = df_entrega[df_entrega['vtex_forma_entrega'] != 'Sem Registro']
    
    if not df_entrega.empty:
        df_entrega = df_entrega.sort_values('Valor_Float', ascending=True)
        fig_entrega = px.bar(df_entrega, x='Valor_Float', y='vtex_forma_entrega', orientation='h', color_discrete_sequence=['#F59E0B'])
        fig_entrega.update_traces(texttemplate='R$ %{x:,.0f}', textposition='outside', hovertemplate='Valor: R$ %{x:,.2f}<extra></extra>')
        fig_entrega = apply_premium_layout(fig_entrega)
        fig_entrega.update_layout(xaxis_title="", yaxis_title="", margin=dict(r=50))
        st.plotly_chart(fig_entrega, use_container_width=True, config={'displayModeBar': False})
    else:
        st.info("Sem dados de entrega.")

st.markdown("<br><hr style='border-color: #334155;'>", unsafe_allow_html=True)

# ─── Linha 5: Tabela de Dados Brutos ──────────────────────────────────────────
st.subheader("📋 Detalhamento dos Pedidos Contestados", help="Tabela com o detalhamento das transações que compõem o filtro atual. Útil para auditoria pontual de pedidos e valores.")

# Format columns for display
cols_disponiveis = df.columns.tolist()
cols_exibir = []

# Tenta exibir essas colunas principais, se existirem
preferencias = [
    ('Data_Real', 'Lançamento SAP'),
    ('Data_Venda', 'Data da Venda'),
    ('adyen_vtex_order_id', 'Pedido VTEX'),
    ('Categoria SAP', 'Contábil SAP'),
    ('Tipo de Chargeback', 'Classificação Analítica'),
    ('Valor_Float', 'Valor em R$'),
    ('vtex_tipo_produto', 'Tipo de Produto'),
    ('adyen_dispute_reason', 'Motivo Adyen')
]

col_map = {}
for col, rename in preferencias:
    if col in cols_disponiveis:
        cols_exibir.append(col)
        col_map[col] = rename

if cols_exibir:
    df_table = df[cols_exibir].copy()
    
    # Formata as datas para não ficarem com horas se não precisar
    if 'Data_Real' in df_table.columns:
        df_table['Data_Real'] = df_table['Data_Real'].dt.strftime('%Y-%m-%d')
    if 'Data_Venda' in df_table.columns:
        df_table['Data_Venda'] = df_table['Data_Venda'].dt.strftime('%Y-%m-%d')
        
    df_table = df_table.rename(columns=col_map)
    st.dataframe(df_table, use_container_width=True, hide_index=True)
else:
    st.info("Tabela de dados não disponível com as colunas atuais.")
