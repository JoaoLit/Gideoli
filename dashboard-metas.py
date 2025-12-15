import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from datetime import datetime

st.set_page_config(
    page_title="Dashboard de Metas",
    layout="wide",
    page_icon="📊",
    initial_sidebar_state="expanded"
)

COLORS = {
    'primary': '#2563eb',
    'success': '#10b981',
    'warning': '#f59e0b',
    'danger': '#ef4444',
    'info': '#3b82f6',
    'neutral': '#64748b',
    'bg_light': '#f8fafc',
    'card_bg': '#ffffff',
    'text_dark': '#0f172a',
    'text_muted': '#64748b',
    'border': '#e2e8f0'
}

COL_EMISSAO = 'EMISSÃO'
COL_VALOR = 'VALOR'
COL_CONTAGEM = 'CONTAGEM'
COL_VENDEDOR = 'VENDEDOR'
COL_META_INICIAL = 'Meta Inicial'
COL_META_MENSAL = 'Meta Mensal'
COL_META_ACUMULADO = 'Acumulado'
COL_MES = 'MÊS'
COL_MES_NUM = 'Mes_Num'
COL_NOME_MES = 'Nome_Mes'

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    * {{
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }}
    
    .stApp {{
        background: {COLORS['bg_light']};
    }}
    
    .main-title {{
        font-size: 2.25rem;
        font-weight: 700;
        color: {COLORS['text_dark']};
        text-align: center;
        margin: 1rem 0 2rem 0;
        letter-spacing: -0.025em;
    }}
    
    .section-title {{
        font-size: 1.25rem;
        font-weight: 600;
        color: {COLORS['text_dark']};
        margin: 2.5rem 0 1rem 0;
        padding: 0.75rem 1rem;
        background: white;
        border-radius: 8px;
        border-left: 4px solid {COLORS['primary']};
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }}
    
    div[data-testid="metric-container"] {{
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        border: 1px solid {COLORS['border']};
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: all 0.2s ease;
    }}
    
    div[data-testid="metric-container"]:hover {{
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        transform: translateY(-2px);
    }}
    
    div[data-testid="metric-container"] label {{
        font-size: 0.875rem;
        font-weight: 500;
        color: {COLORS['text_muted']};
        text-transform: uppercase;
        letter-spacing: 0.025em;
    }}
    
    div[data-testid="metric-container"] [data-testid="stMetricValue"] {{
        font-size: 1.875rem;
        font-weight: 700;
        color: {COLORS['text_dark']};
    }}
    
    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        background: white;
        padding: 0.5rem;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }}
    
    .stTabs [data-baseweb="tab"] {{
        height: 50px;
        padding: 0 1.5rem;
        border-radius: 8px;
        font-weight: 500;
        color: {COLORS['text_muted']};
    }}
    
    .stTabs [aria-selected="true"] {{
        background: {COLORS['primary']};
        color: white;
    }}
    
    div[data-testid="stExpander"] {{
        background: white;
        border: 1px solid {COLORS['border']};
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }}
    
    .insight-box {{
        background: linear-gradient(135deg, {COLORS['primary']}15 0%, {COLORS['info']}10 100%);
        padding: 1.25rem;
        border-radius: 12px;
        border: 1px solid {COLORS['primary']}40;
        margin: 1rem 0;
    }}
    
    .insight-title {{
        font-size: 1rem;
        font-weight: 600;
        color: {COLORS['primary']};
        margin-bottom: 0.5rem;
    }}
    
    .insight-text {{
        font-size: 0.9375rem;
        color: {COLORS['text_dark']};
        line-height: 1.6;
    }}
    
    /* Sidebar styling */
    section[data-testid="stSidebar"] {{
        background: white;
        border-right: 1px solid {COLORS['border']};
    }}
    
    section[data-testid="stSidebar"] h2 {{
        color: {COLORS['text_dark']};
        font-weight: 600;
    }}
</style>
""", unsafe_allow_html=True)

def formatar_moeda(valor):
    """Formata valor em moeda brasileira."""
    return f"R$ {valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

def carregar_e_processar_dados(arq_vendas, arq_metas):
    """Carrega e processa planilhas de vendas e metas."""
    df_vendas = pd.read_excel(arq_vendas)
    df_metas = pd.read_excel(arq_metas, sheet_name='metas')
    df_metas_vendedores = pd.read_excel(arq_metas, sheet_name='Planilha1')  # Carregar dados individuais dos vendedores
    
    # Processar vendas
    df_vendas[COL_EMISSAO] = pd.to_datetime(df_vendas[COL_EMISSAO], errors='coerce')
    df_vendas = df_vendas.dropna(subset=[COL_EMISSAO])
    df_vendas[COL_MES_NUM] = df_vendas[COL_EMISSAO].dt.month
    
    mapa_meses = {
        'Janeiro': 1, 'Fevereiro': 2, 'Março': 3, 'Abril': 4, 
        'Maio': 5, 'Junho': 6, 'Julho': 7, 'Agosto': 8,
        'Setembro': 9, 'Outubro': 10, 'Novembro': 11, 'Dezembro': 12
    }
    df_metas[COL_MES_NUM] = df_metas['Mês'].map(mapa_meses)
    
    df_metas = df_metas.rename(columns={
        'Mensal': COL_META_INICIAL,
        'Acumulado': COL_META_MENSAL
    })
    # Adicionar coluna Acumulado igual a Meta Mensal para compatibilidade
    df_metas[COL_META_ACUMULADO] = df_metas[COL_META_MENSAL]
    
    # Processar dados dos vendedores
    df_metas_vendedores[COL_MES_NUM] = df_metas_vendedores['Mês'].map(mapa_meses)
    
    df_metas_vendedores = df_metas_vendedores.rename(columns={
        'Meta Mensal Acumulada': COL_META_ACUMULADO
    })
    
    return df_vendas, df_metas, df_metas_vendedores

def criar_pizza_atingimento(valor_real, valor_meta, titulo, mostrar_rotulos=True):
    """Cria gráfico de pizza mostrando atingimento da meta."""
    if valor_meta == 0 or pd.isna(valor_meta):
        return go.Figure().add_annotation(
            text="Meta não definida",
            showarrow=False,
            font=dict(size=16, color=COLORS['text_muted'])
        )
    
    percentual = (valor_real / valor_meta) * 100
    
    if valor_real >= valor_meta:
        labels = ['Meta Atingida', 'Superação']
        values = [valor_meta, valor_real - valor_meta]
        colors = [COLORS['success'], COLORS['primary']]
        center_color = COLORS['success']
    else:
        labels = ['Realizado', 'Falta Atingir']
        values = [valor_real, valor_meta - valor_real]
        colors = [COLORS['info'], COLORS['border']]
        center_color = COLORS['warning'] if percentual >= 70 else COLORS['danger']
    
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=0.65,
        marker=dict(colors=colors, line=dict(color='white', width=3)),
        textinfo='label+percent' if mostrar_rotulos else 'percent',
        textposition='outside' if mostrar_rotulos else 'auto',
        textfont=dict(size=15, family='Inter', color=COLORS['text_dark']),
        hovertemplate='<b>%{label}</b><br>%{value:,.2f}<br>%{percent}<extra></extra>'
    )])
    
    fig.add_annotation(
        text=f"<b>{percentual:.1f}%</b>",
        x=0.5, y=0.55,
        font=dict(size=32, color=center_color, family='Inter'),
        showarrow=False
    )
    
    fig.add_annotation(
        text="da Meta",
        x=0.5, y=0.45,
        font=dict(size=14, color=COLORS['text_muted'], family='Inter'),
        showarrow=False
    )
    
    fig.update_layout(
        title=dict(
            text=titulo,
            x=0.5,
            xanchor='center',
            font=dict(size=16, color=COLORS['text_dark'], family='Inter', weight=600)
        ),
        height=380,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.15,
            xanchor="center",
            x=0.5,
            font=dict(size=12, family='Inter')
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(t=60, b=80, l=20, r=20)
    )
    
    return fig

def criar_pizza_distribuicao(df, mostrar_rotulos=True):
    """Cria gráfico de pizza com distribuição por mês."""
    color_scale = [COLORS['primary'], COLORS['success'], COLORS['info'], 
                   COLORS['warning'], '#8b5cf6', '#ec4899', '#06b6d4', '#f59e0b']
    
    fig = go.Figure(data=[go.Pie(
        labels=df[COL_NOME_MES],
        values=df[COL_VALOR],
        hole=0.5,
        marker=dict(colors=color_scale[:len(df)], line=dict(color='white', width=2)),
        textinfo='label+percent' if mostrar_rotulos else 'percent',
        textposition='outside' if mostrar_rotulos else 'auto',
        textfont=dict(size=14, family='Inter'),
        hovertemplate='<b>%{label}</b><br>R$ %{value:,.2f}<br>%{percent}<extra></extra>'
    )])
    
    total = df[COL_VALOR].sum()
    fig.add_annotation(
        text=f"<b>Total</b><br>{formatar_moeda(total)}",
        x=0.5, y=0.5,
        font=dict(size=14, color=COLORS['text_dark'], family='Inter'),
        showarrow=False
    )
    
    fig.update_layout(
        title=dict(
            text='Distribuição de Faturamento por Mês',
            x=0.5,
            xanchor='center',
            font=dict(size=16, color=COLORS['text_dark'], family='Inter', weight=600)
        ),
        height=380,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.15,
            xanchor="center",
            x=0.5,
            font=dict(size=11, family='Inter')
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(t=60, b=80, l=20, r=20)
    )
    
    return fig

def criar_grafico_barras(df: pd.DataFrame, em_percentual: bool = False, mostrar_rotulos: bool = True) -> go.Figure:
    """Cria gráfico de barras comparando realizado vs meta."""
    fig = go.Figure()
    
    # Ordenar por mês para garantir ordem correta
    df = df.sort_values(COL_MES_NUM).copy()
    
    # Determinar cores baseadas no atingimento
    cores_barras = []
    for _, row in df.iterrows():
        if pd.isna(row[COL_META_INICIAL]) or row[COL_META_INICIAL] == 0:
            cores_barras.append(COLORS['neutral'])
        else:
            perc = (row[COL_VALOR] / row[COL_META_INICIAL]) * 100
            if perc >= 100:
                cores_barras.append(COLORS['success'])
            elif perc >= 70:
                cores_barras.append(COLORS['warning'])
            else:
                cores_barras.append(COLORS['danger'])
    
    if em_percentual:
        y_valores = []
        for _, row in df.iterrows():
            if pd.isna(row[COL_META_INICIAL]) or row[COL_META_INICIAL] == 0:
                y_valores.append(0)
            else:
                y_valores.append((row[COL_VALOR] / row[COL_META_INICIAL]) * 100)
        
        fig.add_trace(go.Bar(
            x=df[COL_NOME_MES],
            y=y_valores,
            name='Atingimento',
            marker=dict(color=cores_barras, line=dict(color='white', width=1)),
            text=[f'{v:.1f}%' for v in y_valores] if mostrar_rotulos else None,
            textposition='outside' if mostrar_rotulos else None,
            textfont=dict(size=14, family='Inter', color=COLORS['text_dark']) if mostrar_rotulos else None,
            hovertemplate='<b>%{x}</b><br>Atingimento: %{y:.1f}%<extra></extra>'
        ))
        
        # Linha de referência em 100%
        fig.add_hline(
            y=100,
            line=dict(color=COLORS['success'], width=2, dash='dash'),
            annotation=dict(text="Meta (100%)", font=dict(size=11, color=COLORS['success']))
        )
        
        y_title = 'Percentual de Atingimento (%)'
    else:
        fig.add_trace(go.Bar(
            x=df[COL_NOME_MES],
            y=df[COL_VALOR],
            name='Realizado',
            marker=dict(color=cores_barras, line=dict(color='white', width=1)),
            text=df[COL_VALOR].apply(lambda x: formatar_moeda(x)) if mostrar_rotulos else None,
            textposition='outside' if mostrar_rotulos else None,
            textfont=dict(size=14, family='Inter', color=COLORS['text_dark']) if mostrar_rotulos else None,
            hovertemplate='<b>%{x}</b><br>Realizado: R$ %{y:,.2f}<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=df[COL_NOME_MES],
            y=df[COL_META_INICIAL],
            name='Meta',
            mode='lines+markers+text' if mostrar_rotulos else 'lines+markers',
            line=dict(color=COLORS['primary'], width=3),
            marker=dict(size=10, color=COLORS['primary'], line=dict(color='white', width=2)),
            text=df[COL_META_INICIAL].apply(lambda x: formatar_moeda(x)) if mostrar_rotulos else None,
            textposition='top center' if mostrar_rotulos else None,
            textfont=dict(size=14, family='Inter', color=COLORS['primary']) if mostrar_rotulos else None,
            hovertemplate='<b>%{x}</b><br>Meta: R$ %{y:,.2f}<extra></extra>'
        ))
        
        y_title = 'Valor (R$)'
    
    fig.update_layout(
        title=dict(
            text='Desempenho Mensal vs Meta',
            x=0.5,
            xanchor='center',
            font=dict(size=18, color=COLORS['text_dark'], family='Inter', weight=600)
        ),
        xaxis=dict(
            title='Mês',
            title_font=dict(size=13, color=COLORS['text_dark']),
            tickfont=dict(size=12, color=COLORS['text_muted']),
            showgrid=False
        ),
        yaxis=dict(
            title=y_title,
            title_font=dict(size=13, color=COLORS['text_dark']),
            tickfont=dict(size=11, color=COLORS['text_muted']),
            showgrid=True,
            gridcolor=COLORS['border']
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=420,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=12, family='Inter')
        ),
        margin=dict(t=80, b=60, l=60, r=40),
        hovermode='x unified'
    )
    
    return fig

def criar_grafico_cumulativo(df: pd.DataFrame, mostrar_rotulos: bool = True) -> go.Figure:
    """Cria gráfico de área com evolução cumulativa."""
    df = df.sort_values(COL_MES_NUM).copy()
    
    df['Realizado_Acum'] = df[COL_VALOR].cumsum()
    df['Meta_Acum'] = df[COL_META_ACUMULADO].cumsum()
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(
        x=df[COL_NOME_MES],
        y=df['Realizado_Acum'],
        name='Realizado Acumulado',
        mode='lines+markers+text' if mostrar_rotulos else 'lines',
        line=dict(color=COLORS['success'], width=3),
        marker=dict(size=8, color=COLORS['success']) if mostrar_rotulos else None,
        fill='tozeroy',
        fillcolor=f"rgba(16, 185, 129, 0.15)",
        text=df['Realizado_Acum'].apply(lambda x: formatar_moeda(x)) if mostrar_rotulos else None,
        textposition='top center' if mostrar_rotulos else None,
        textfont=dict(size=14, family='Inter', color=COLORS['success']) if mostrar_rotulos else None,
        hovertemplate='<b>%{x}</b><br>Acumulado: R$ %{y:,.2f}<extra></extra>'
    ))
    
    fig.add_trace(go.Scatter(
        x=df[COL_NOME_MES],
        y=df['Meta_Acum'],
        name='Meta Acumulada',
        mode='lines+markers+text' if mostrar_rotulos else 'lines+markers',
        line=dict(color=COLORS['primary'], width=3, dash='dot'),
        marker=dict(size=8, color=COLORS['primary']),
        text=df['Meta_Acum'].apply(lambda x: formatar_moeda(x)) if mostrar_rotulos else None,
        textposition='bottom center' if mostrar_rotulos else None,
        textfont=dict(size=14, family='Inter', color=COLORS['primary']) if mostrar_rotulos else None,
        hovertemplate='<b>%{x}</b><br>Meta: R$ %{y:,.2f}<extra></extra>'
    ))

    fig.update_layout(
        title=dict(
            text='Evolução Acumulada no Ano',
            x=0.5,
            xanchor='center',
            font=dict(size=18, color=COLORS['text_dark'], family='Inter', weight=600)
        ),
        xaxis=dict(
            title='Mês',
            title_font=dict(size=13, color=COLORS['text_dark']),
            tickfont=dict(size=12, family='Inter'),
            showgrid=False
        ),
        yaxis=dict(
            title='Valor Acumulado (R$)',
            title_font=dict(size=13, color=COLORS['text_dark']),
            tickformat=',.0f',
            tickfont=dict(size=11, family='Inter'),
            showgrid=True,
            gridcolor='rgba(0,0,0,0.05)'
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=420,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=12, family='Inter')
        ),
        margin=dict(t=80, b=60, l=60, r=40),
        hovermode='x unified'
    )
    
    return fig

def criar_grafico_barras_acumulado(df: pd.DataFrame, em_percentual: bool = False, mostrar_rotulos: bool = True) -> go.Figure:
    """Cria gráfico de barras comparando realizado vs meta acumulada."""
    fig = go.Figure()
    
    # Ordenar por mês para garantir ordem correta
    df = df.sort_values(COL_MES_NUM).copy()
    
    # Determinar cores baseadas no atingimento
    cores_barras = []
    for _, row in df.iterrows():
        if pd.isna(row[COL_META_ACUMULADO]) or row[COL_META_ACUMULADO] == 0:
            cores_barras.append(COLORS['neutral'])
        else:
            perc = (row[COL_VALOR] / row[COL_META_ACUMULADO]) * 100
            if perc >= 100:
                cores_barras.append(COLORS['success'])
            elif perc >= 70:
                cores_barras.append(COLORS['warning'])
            else:
                cores_barras.append(COLORS['danger'])
    
    if em_percentual:
        y_valores = []
        for _, row in df.iterrows():
            if pd.isna(row[COL_META_ACUMULADO]) or row[COL_META_ACUMULADO] == 0:
                y_valores.append(0)
            else:
                y_valores.append((row[COL_VALOR] / row[COL_META_ACUMULADO]) * 100)
        
        fig.add_trace(go.Bar(
            x=df[COL_NOME_MES],
            y=y_valores,
            name='Atingimento',
            marker=dict(color=cores_barras, line=dict(color='white', width=1)),
            text=[f'{v:.1f}%' for v in y_valores] if mostrar_rotulos else None,
            textposition='outside' if mostrar_rotulos else None,
            textfont=dict(size=14, family='Inter', color=COLORS['text_dark']) if mostrar_rotulos else None,
            hovertemplate='<b>%{x}</b><br>Atingimento: %{y:.1f}%<extra></extra>'
        ))
        
        # Linha de referência em 100%
        fig.add_hline(
            y=100,
            line=dict(color=COLORS['success'], width=2, dash='dash'),
            annotation=dict(text="Meta (100%)", font=dict(size=11, color=COLORS['success']))
        )
        
        y_title = 'Percentual de Atingimento (%)'
    else:
        fig.add_trace(go.Bar(
            x=df[COL_NOME_MES],
            y=df[COL_VALOR],
            name='Realizado',
            marker=dict(color=cores_barras, line=dict(color='white', width=1)),
            text=df[COL_VALOR].apply(lambda x: formatar_moeda(x)) if mostrar_rotulos else None,
            textposition='outside' if mostrar_rotulos else None,
            textfont=dict(size=14, family='Inter', color=COLORS['text_dark']) if mostrar_rotulos else None,
            hovertemplate='<b>%{x}</b><br>Realizado: R$ %{y:,.2f}<extra></extra>'
        ))
        
        fig.add_trace(go.Scatter(
            x=df[COL_NOME_MES],
            y=df[COL_META_ACUMULADO],
            name='Meta Acumulada (Ajustada)',
            mode='lines+markers+text' if mostrar_rotulos else 'lines+markers',
            line=dict(color=COLORS['warning'], width=3),
            marker=dict(size=10, color=COLORS['warning'], line=dict(color='white', width=2)),
            text=df[COL_META_ACUMULADO].apply(lambda x: formatar_moeda(x)) if mostrar_rotulos else None,
            textposition='top center' if mostrar_rotulos else None,
            textfont=dict(size=14, family='Inter', color=COLORS['warning']) if mostrar_rotulos else None,
            hovertemplate='<b>%{x}</b><br>Meta Acumulada: R$ %{y:,.2f}<extra></extra>'
        ))
        
        y_title = 'Valor (R$)'
    
    fig.update_layout(
        title=dict(
            text='Desempenho Mensal vs Meta Acumulada (Ajustada)',
            x=0.5,
            xanchor='center',
            font=dict(size=18, color=COLORS['text_dark'], family='Inter', weight=600)
        ),
        xaxis=dict(
            title='Mês',
            title_font=dict(size=13, color=COLORS['text_dark']),
            tickfont=dict(size=12, color=COLORS['text_muted']),
            showgrid=False
        ),
        yaxis=dict(
            title=y_title,
            title_font=dict(size=13, color=COLORS['text_dark']),
            tickfont=dict(size=11, color=COLORS['text_muted']),
            showgrid=True,
            gridcolor=COLORS['border']
        ),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        height=420,
        showlegend=True,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.2,
            xanchor="center",
            x=0.5,
            font=dict(size=12, family='Inter')
        ),
        margin=dict(t=80, b=60, l=60, r=40),
        hovermode='x unified'
    )
    
    return fig

def gerar_insights(df, total_vendas, total_meta):
    """Gera insights baseados nos dados."""
    insights = []
    
    # Análise de atingimento geral
    if total_meta > 0:
        perc_total = (total_vendas / total_meta) * 100
        if perc_total >= 100:
            insights.append({
                'tipo': 'success',
                'titulo': '🎯 Meta Atingida!',
                'texto': f'Parabéns! Você atingiu {perc_total:.1f}% da meta total, superando em {formatar_moeda(total_vendas - total_meta)}.'
            })
        elif perc_total >= 70:
            falta = total_meta - total_vendas
            insights.append({
                'tipo': 'warning',
                'titulo': '⚡ Quase lá!',
                'texto': f'Você está em {perc_total:.1f}% da meta. Faltam {formatar_moeda(falta)} para atingir o objetivo.'
            })
        else:
            insights.append({
                'tipo': 'danger',
                'titulo': '📊 Atenção Necessária',
                'texto': f'O atingimento atual é de {perc_total:.1f}%. Revise a estratégia para melhorar os resultados.'
            })
    
    # Análise de tendência
    if len(df) >= 2:
        ultimos_3 = df.tail(3)[COL_VALOR].values if len(df) >= 3 else df[COL_VALOR].values
        if len(ultimos_3) >= 2:
            tendencia = ultimos_3[-1] > ultimos_3[-2]
            if tendencia:
                crescimento = ((ultimos_3[-1] / ultimos_3[-2]) - 1) * 100
                insights.append({
                    'tipo': 'success',
                    'titulo': '📈 Tendência Positiva',
                    'texto': f'Crescimento de {crescimento:.1f}% no último mês analisado em relação ao anterior.'
                })
            else:
                queda = ((ultimos_3[-2] / ultimos_3[-1]) - 1) * 100
                insights.append({
                    'tipo': 'warning',
                    'titulo': '📉 Atenção à Queda',
                    'texto': f'Redução de {queda:.1f}% no último mês. Considere ações corretivas.'
                })
    
    # Melhor mês
    if not df.empty:
        melhor_mes = df.loc[df[COL_VALOR].idxmax()]
        insights.append({
            'tipo': 'info',
            'titulo': '🏆 Melhor Performance',
            'texto': f'{melhor_mes[COL_NOME_MES]} foi o melhor mês com {formatar_moeda(melhor_mes[COL_VALOR])} em vendas.'
        })
    
    return insights

def criar_heatmap_faturamento(df):
    """Cria mapa de calor mostrando intensidade de faturamento por mês."""
    # Preparar dados
    df_sorted = df.sort_values(COL_MES_NUM).copy()
    
    # Criar matriz para heatmap (1 linha com todos os meses)
    valores = df_sorted[COL_VALOR].values.reshape(1, -1)
    meses = df_sorted[COL_NOME_MES].values
    
    # Criar texto para hover
    texto_hover = [[formatar_moeda(val)] for val in df_sorted[COL_VALOR].values]
    texto_hover = [texto_hover]  # Transformar em matriz 2D
    
    fig = go.Figure(data=go.Heatmap(
        z=valores,
        x=meses,
        y=['Faturamento'],
        colorscale=[
            [0, '#e0e7ff'],      # Azul muito claro
            [0.25, '#c7d2fe'],   # Azul claro
            [0.5, '#818cf8'],    # Azul médio
            [0.75, '#6366f1'],   # Azul
            [1, '#4f46e5']       # Azul escuro
        ],
        text=texto_hover,
        texttemplate='%{text}',
        textfont=dict(size=13, color='white', family='Inter', weight=600),
        hovertemplate='<b>%{x}</b><br>Faturamento: %{text}<extra></extra>',
        showscale=True,
        colorbar=dict(
            title="Valor (R$)",
            tickformat=",.0f",
            len=0.7,
            thickness=15
        )
    ))
    
    fig.update_layout(
        title=dict(
            text='Mapa de Calor - Intensidade de Faturamento',
            x=0.5,
            xanchor='center',
            font=dict(size=16, color=COLORS['text_dark'], family='Inter', weight=600)
        ),
        height=200,
        xaxis=dict(
            title='',
            tickfont=dict(size=12, family='Inter'),
            showgrid=False
        ),
        yaxis=dict(
            title='',
            showticklabels=False,
            showgrid=False
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=100, t=50, b=50),
        font=dict(family='Inter', size=12, color=COLORS['text_dark'])
    )
    
    return fig

def criar_histograma_faturamento(df):
    """Cria histograma mostrando distribuição de valores de faturamento."""
    df_sorted = df.sort_values(COL_MES_NUM).copy()
    
    fig = go.Figure(data=[go.Bar(
        x=df_sorted[COL_NOME_MES],
        y=df_sorted[COL_VALOR],
        marker=dict(
            color=df_sorted[COL_VALOR],
            colorscale='Blues',
            showscale=False,
            line=dict(color='white', width=2)
        ),
        text=[formatar_moeda(val) for val in df_sorted[COL_VALOR]],
        textposition='outside',
        textfont=dict(size=13, color=COLORS['text_dark'], family='Inter', weight=600),
        hovertemplate='<b>%{x}</b><br>Faturamento: R$ %{y:,.2f}<extra></extra>'
    )])
    
    fig.update_layout(
        title=dict(
            text='Histograma - Valores de Faturamento por Mês',
            x=0.5,
            xanchor='center',
            font=dict(size=16, color=COLORS['text_dark'], family='Inter', weight=600)
        ),
        height=380,
        xaxis=dict(
            title='Mês',
            tickfont=dict(size=12, family='Inter'),
            showgrid=False
        ),
        yaxis=dict(
            title='Faturamento (R$)',
            tickformat=',.0f',
            tickfont=dict(size=12, family='Inter'),
            showgrid=True,
            gridcolor='rgba(0,0,0,0.05)'
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(t=60, b=60, l=80, r=40)
    )
    
    return fig

st.markdown('<h1 class="main-title">📊 Dashboard de Análise de Metas</h1>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ Configurações")
    st.markdown("---")
    
    f_vendas = st.file_uploader("📁 Planilha de Vendas", type=['xlsx'], help="Envie a planilha com os dados de vendas", key="uploader_vendas")
    f_metas = st.file_uploader("🎯 Planilha de Metas", type=['xlsx'], help="Envie a planilha com as metas definidas", key="uploader_metas")
    
    if f_vendas and f_metas:
        st.markdown("---")
        st.subheader("🔍 Filtros")
        
        mostrar_rotulos = st.toggle("Mostrar rótulos nos gráficos", value=True, help="Exibir valores diretamente nos gráficos")
        
        df_vendas, df_metas, df_metas_vendedores = carregar_e_processar_dados(f_vendas, f_metas)
        
        meses_disponiveis = sorted(df_vendas[COL_MES_NUM].unique())
        meses_abrev = {
            1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
            7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
        }
        
        meses_nomes = {
            1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
            5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
            9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
        }
        
        meses_sel = st.multiselect(
            "Selecione os meses:",
            options=meses_disponiveis,
            format_func=lambda x: meses_nomes[x],
            default=meses_disponiveis,
            help="Escolha os meses que deseja analisar"
        )
        
        if not meses_sel:
            st.warning("⚠️ Selecione pelo menos um mês")
            st.stop()
        
        mostrar_percentual = st.checkbox(
            "Exibir em percentual",
            value=False,
            help="Mostra o gráfico em percentual de atingimento ao invés de valores absolutos"
        )

if f_vendas and f_metas and meses_sel:
    df_filtrado = df_vendas[df_vendas[COL_MES_NUM].isin(meses_sel)].copy()
    
    df_agrupado = df_filtrado.groupby(COL_MES_NUM).agg({
        COL_VALOR: 'sum',
        COL_CONTAGEM: 'sum'
    }).reset_index()
    
    df_consolidado = pd.merge(df_agrupado, df_metas, on=COL_MES_NUM, how='left').sort_values(COL_MES_NUM)
    df_consolidado[COL_NOME_MES] = df_consolidado[COL_MES_NUM].map(meses_abrev)
    
    total_vendas = df_consolidado[COL_VALOR].sum()
    total_meta_mensal = df_consolidado[COL_META_MENSAL].sum()
    total_meta_acumulado = df_consolidado[COL_META_ACUMULADO].sum()
    total_meta_inicial = df_consolidado[COL_META_INICIAL].sum()
    total_pedidos = df_consolidado[COL_CONTAGEM].sum()
    ticket_medio = total_vendas / total_pedidos if total_pedidos > 0 else 0
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            label="Faturamento Total",
            value=formatar_moeda(total_vendas)
        )
    
    with col2:
        if total_meta_mensal > 0:
            perc_meta = (total_vendas / total_meta_mensal) * 100
            delta_meta = perc_meta - 100
            st.metric(
                label="vs Meta Mensal",
                value=f"{perc_meta:.1f}%",
                delta=f"{delta_meta:+.1f}%"
            )
        else:
            st.metric(label="vs Meta Mensal", value="N/A")
    
    with col3:
        if total_meta_inicial > 0:
            perc_meta_inicial = (total_vendas / total_meta_inicial) * 100
            delta_inicial = perc_meta_inicial - 100
            st.metric(
                label="vs Meta Inicial",
                value=f"{perc_meta_inicial:.1f}%",
                delta=f"{delta_inicial:+.1f}%"
            )
        else:
            st.metric(label="vs Meta Inicial", value="N/A")
    
    with col4:
        if total_meta_acumulado > 0:
            falta_meta = total_meta_acumulado - total_vendas
            st.metric(
                label="Falta para Meta",
                value=formatar_moeda(-falta_meta),
                delta=f"{(falta_meta/total_meta_acumulado)*100:.1f}% da meta" if falta_meta > 0 else "Meta atingida! 🎉"
            )
        else:
            st.metric(label="Falta para Meta", value="N/A")
    
    with col5:
        st.metric(
            label="Ticket Médio",
            value=formatar_moeda(ticket_medio),
            delta=f"{total_pedidos} pedidos"
        )
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    tab_geral, tab_vendedor = st.tabs(["📊 Visão Geral", "👤 Por Vendedor"])
    
    with tab_geral:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            fig_pizza1 = criar_pizza_atingimento(total_vendas, total_meta_inicial, "Vs Meta Inicial", mostrar_rotulos)
            st.plotly_chart(fig_pizza1, use_container_width=True)
        
        with col2:
            fig_pizza_vs_meta = criar_pizza_atingimento(total_vendas, total_meta_mensal, "Vs Meta Mensal", mostrar_rotulos)
            st.plotly_chart(fig_pizza_vs_meta, use_container_width=True)
        
        with col3:
            fig_pizza2 = criar_pizza_distribuicao(df_consolidado, mostrar_rotulos)
            st.plotly_chart(fig_pizza2, use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown('<div class="section-title">📊 Visualização de Distribuição de Faturamento</div>', unsafe_allow_html=True)
        
        fig_heatmap = criar_heatmap_faturamento(df_consolidado)
        st.plotly_chart(fig_heatmap, use_container_width=True)
        
        col_hist1, col_hist2 = st.columns([2, 1])
        
        with col_hist1:
            fig_histograma = criar_histograma_faturamento(df_consolidado)
            st.plotly_chart(fig_histograma, use_container_width=True)
        
        with col_hist2:
            st.markdown("""
            <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                        padding: 20px; border-radius: 12px; color: white; height: 100%;">
                <h3 style="margin: 0 0 15px 0; font-size: 18px;">📈 Estatísticas</h3>
            """, unsafe_allow_html=True)
            
            media_faturamento = df_consolidado[COL_VALOR].mean()
            max_faturamento = df_consolidado[COL_VALOR].max()
            min_faturamento = df_consolidado[COL_VALOR].min()
            
            st.markdown(f"""
                <div style="margin: 10px 0;">
                    <div style="font-size: 12px; opacity: 0.9;">Média Mensal</div>
                    <div style="font-size: 20px; font-weight: bold;">{formatar_moeda(media_faturamento)}</div>
                </div>
                <div style="margin: 10px 0;">
                    <div style="font-size: 12px; opacity: 0.9;">Maior Faturamento</div>
                    <div style="font-size: 20px; font-weight: bold;">{formatar_moeda(max_faturamento)}</div>
                </div>
                <div style="margin: 10px 0;">
                    <div style="font-size: 12px; opacity: 0.9;">Menor Faturamento</div>
                    <div style="font-size: 20px; font-weight: bold;">{formatar_moeda(min_faturamento)}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        fig_barras = criar_grafico_barras(df_consolidado, mostrar_percentual, mostrar_rotulos)
        st.plotly_chart(fig_barras, use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        fig_barras_acumulado = criar_grafico_barras_acumulado(df_consolidado, mostrar_percentual, mostrar_rotulos)
        st.plotly_chart(fig_barras_acumulado, use_container_width=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        fig_cumulativo = criar_grafico_cumulativo(df_consolidado, mostrar_rotulos)
        st.plotly_chart(fig_cumulativo, use_container_width=True)
    
    with tab_vendedor:
        st.markdown('<div class="section-title">👤 Análise Individual por Vendedor</div>', unsafe_allow_html=True)
        
        vendedores = sorted(df_vendas[COL_VENDEDOR].unique())
        vendedor_selecionado = st.selectbox(
            "Selecione um vendedor para análise detalhada:",
            ["Selecione..."] + vendedores,
            help="Escolha um vendedor para ver suas métricas individuais"
        )
        
        if vendedor_selecionado != "Selecione...":
            df_vendas_vendedor = df_vendas[df_vendas[COL_VENDEDOR] == vendedor_selecionado]
            
            df_metas_vendedor = df_metas_vendedores[df_metas_vendedores[COL_VENDEDOR] == vendedor_selecionado].copy()
            
            if df_metas_vendedor.empty:
                st.warning(f"⚠️ Nenhuma meta encontrada para '{vendedor_selecionado}' na Planilha1")
                st.stop()
            
            vendas_vendedor = df_vendas_vendedor.groupby(COL_MES_NUM).agg({
                COL_VALOR: 'sum',
                COL_CONTAGEM: 'sum'
            }).reset_index()
            
            metas_vendedor = df_metas_vendedor.groupby(COL_MES_NUM).agg({
                COL_META_INICIAL: 'sum',
                COL_META_MENSAL: 'sum',
                COL_META_ACUMULADO: 'sum' # Adicionando agregação para Meta Acumulada
            }).reset_index()
            
            df_vendedor = pd.merge(vendas_vendedor, metas_vendedor, on=COL_MES_NUM, how='left').sort_values(COL_MES_NUM)
            df_vendedor[COL_NOME_MES] = df_vendedor[COL_MES_NUM].map(meses_abrev)
            
            total_vendas_v = df_vendedor[COL_VALOR].sum()
            total_meta_mensal_v = df_vendedor[COL_META_MENSAL].sum()
            total_meta_acumulado_v = df_metas_vendedor[COL_META_ACUMULADO].sum() if COL_META_ACUMULADO in df_metas_vendedor.columns else total_meta_mensal_v
            total_meta_inicial_v = df_vendedor[COL_META_INICIAL].sum()
            total_pedidos_v = df_vendedor[COL_CONTAGEM].sum()
            ticket_medio_v = total_vendas_v / total_pedidos_v if total_pedidos_v > 0 else 0
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric(
                    label="Faturamento",
                    value=formatar_moeda(total_vendas_v)
                )
            
            with col2:
                if total_meta_mensal_v > 0:
                    perc_meta_v = (total_vendas_v / total_meta_mensal_v) * 100
                    delta_v = perc_meta_v - 100
                    st.metric(
                        label="vs Meta Mensal",
                        value=f"{perc_meta_v:.1f}%",
                        delta=f"{delta_v:+.1f}%"
                    )
                else:
                    st.metric(label="vs Meta Mensal", value="N/A")
            
            with col3:
                if total_meta_inicial_v > 0:
                    perc_meta_inicial_v = (total_vendas_v / total_meta_inicial_v) * 100
                    delta_inicial_v = perc_meta_inicial_v - 100
                    st.metric(
                        label="vs Meta Inicial",
                        value=f"{perc_meta_inicial_v:.1f}%",
                        delta=f"{delta_inicial_v:+.1f}%"
                    )
                else:
                    st.metric(label="vs Meta Inicial", value="N/A")
            
            with col4:
                st.metric(
                    label="Ticket Médio",
                    value=formatar_moeda(ticket_medio_v),
                    delta=f"{total_pedidos_v} pedidos"
                )
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            col_v1, col_v2 = st.columns(2)
            
            with col_v1:
                fig_pizza_v = criar_pizza_atingimento(
                    total_vendas_v,
                    total_meta_mensal_v,
                    f"Atingimento - {vendedor_selecionado}",
                    mostrar_rotulos
                )
                st.plotly_chart(fig_pizza_v, use_container_width=True)
            
            with col_v2:
                fig_dist_v = criar_pizza_distribuicao(df_vendedor, mostrar_rotulos)
                st.plotly_chart(fig_dist_v, use_container_width=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            
            with st.expander("📋 Ver Detalhamento das Vendas"):
                df_detalhe = df_vendas_vendedor[[COL_EMISSAO, COL_VALOR, COL_CONTAGEM]].copy()
                df_detalhe = df_detalhe.sort_values(COL_EMISSAO, ascending=False)
                df_detalhe[COL_VALOR] = df_detalhe[COL_VALOR].apply(formatar_moeda)
                
                st.dataframe(
                    df_detalhe,
                    use_container_width=True,
                    hide_index=True
                )
else:
    st.info("👋 Bem-vindo! Por favor, envie as planilhas de **Vendas** e **Metas** na barra lateral para iniciar a análise.")
    
    st.markdown("""
    ### Como usar este dashboard:
    
    1. **Envie as planilhas** - Use os campos na barra lateral para carregar seus arquivos Excel
    2. **Filtre os dados** - Selecione os meses que deseja analisar
    3. **Explore as análises** - Navegue entre as abas para ver diferentes visões dos dados
    4. **Analise os insights** - Veja recomendações automáticas baseadas nos dados
    
    ---
    
    💡 **Dica:** Use a opção "Exibir em percentual" para facilitar comparações entre períodos diferentes.
    """)
