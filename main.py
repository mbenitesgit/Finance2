import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import os

def processar_extratos_bi(arquivo_excel):
    """
    Processa o arquivo Excel específico do BI com múltiplas abas mensais
    """
    # Ler todas as abas do Excel (exceto a última que está vazia)
    xl = pd.ExcelFile(arquivo_excel)
    abas = [sheet for sheet in xl.sheet_names if sheet != 'Planilha2']
    
    dados_combinados = []
    for aba in abas:
        try:
            df = pd.read_excel(arquivo_excel, sheet_name=aba)
            
            # Padronizar nomes de colunas
            df.columns = [col.lower().strip() for col in df.columns]
            
            # Verificar estrutura das colunas e renomear se necessário
            if 'valor (r$)' in df.columns:
                df = df.rename(columns={'valor (r$)': 'valor'})
            elif 'valor' not in df.columns:
                # Tentar identificar a coluna de valor
                for col in df.columns:
                    if 'valor' in col.lower():
                        df = df.rename(columns={col: 'valor'})
                        break
            
            # Adicionar colunas de identificação
            df['origem'] = aba
            df['mes_ano'] = aba
            
            # Converter data
            df['data'] = pd.to_datetime(df['data'])
            
            dados_combinados.append(df)
        except Exception as e:
            print(f"Erro ao processar aba {aba}: {e}")
            continue
    
    if not dados_combinados:
        raise ValueError("Nenhuma aba válida foi encontrada")
    
    df_completo = pd.concat(dados_combinados, ignore_index=True)
    
    # Processar tipos de transação
    df_completo['tipo'] = df_completo['tipo'].str.strip()
    
    # Identificar gastos e receitas baseado no tipo e valor
    def classificar_movimentacao(tipo, valor):
        if 'enviado' in str(tipo).lower() or (valor < 0 and 'recebido' not in str(tipo).lower()):
            return 'Gasto'
        elif 'recebido' in str(tipo).lower() or valor > 0:
            return 'Receita'
        else:
            return 'Outro'
    
    df_completo['tipo_movimentacao'] = df_completo.apply(
        lambda x: classificar_movimentacao(x['tipo'], x['valor']), axis=1
    )
    
    # Garantir que valores negativos sejam tratados como gastos
    df_completo['valor_absoluto'] = df_completo['valor'].abs()
    
    # Extrair período
    df_completo['mes'] = df_completo['data'].dt.month
    df_completo['ano'] = df_completo['data'].dt.year
    df_completo['mes_ano_period'] = df_completo['data'].dt.to_period('M')
    
    return df_completo

def criar_categorias_automaticas(df):
    """
    Cria categorias automáticas baseadas nos nomes dos destinatários
    """
    # Palavras-chave para categorização
    categorias = {
        'Educação': ['COLEGIO', 'ESCOLA', 'FACULDADE', 'HEBE MARSIGLIA'],
        'Alimentação': ['ZAFFARI', 'ATACADÃO', 'SUPERMERCADO', 'BK BRASIL', 'IFOOD', 'COMERCIAL'],
        'Serviços Públicos': ['CIA RIOGRANDENSE', 'CIA ESTADUAL', 'SANEMEN', 'ENERGIA', 'AGUA'],
        'Transporte': ['UBER', 'TRANSPESSOAL', 'BUS2', 'ESTACIONAMENTO', 'REK PARKING'],
        'Saúde': ['FARMÁCIA', 'MEDICAMENTOS', 'BRAIR', 'PLANO DE SAÚDE'],
        'Compras Online': ['SHOPEE', 'AMERICANAS', 'NETSHOES', 'MERCADO LIVRE'],
        'Serviços Financeiros': ['SERASA', 'OPP SERVICOS', 'FINANCEIRO', 'BANCO', 'ITAU'],
        'Família': ['MAURICIO BENITES', 'DEBORA APARECIDA', 'SELMA FURTADO', 'GISELE BORGES', 'JOÃO VITOR'],
        'Lazer': ['CROSS EXPERIENCE', 'ACADEMIA', 'CINEMA', 'RESTAURANTE'],
        'Impostos/Taxas': ['IPVA', 'SEFAZ', 'DETRAN', 'GAD/E', 'IMPOSTO'],
        'Telecomunicações': ['CLARO', 'TELEFONE', 'INTERNET'],
        'Outros': []
    }
    
    def classificar_categoria(destinatario):
        destinatario_upper = str(destinatario).upper()
        for categoria, palavras in categorias.items():
            for palavra in palavras:
                if palavra in destinatario_upper:
                    return categoria
        return 'Outros'
    
    df['categoria'] = df['destinatário/pagador'].apply(classificar_categoria)
    return df

def criar_dashboard_html_bi(df):
    """
    Cria dashboard interativo em HTML específico para os dados do BI
    """
    # Filtrar apenas gastos para análise
    gastos = df[df['tipo_movimentacao'] == 'Gasto']
    receitas = df[df['tipo_movimentacao'] == 'Receita']
    
    # 1. Métricas Principais
    total_gastos = gastos['valor_absoluto'].sum()
    total_receitas = receitas['valor_absoluto'].sum()
    saldo = total_receitas - total_gastos
    
    # 2. Gastos por Categoria
    gastos_por_categoria = gastos.groupby('categoria')['valor_absoluto'].sum().sort_values(ascending=False)
    
    # 3. Principais Destinatários (Top 15)
    principais_destinos = gastos.groupby('destinatário/pagador')['valor_absoluto'].sum().nlargest(15)
    
    # 4. Principais Fontes Pagadoras (Top 15)
    principais_fontes = receitas.groupby('destinatário/pagador')['valor_absoluto'].sum().nlargest(15)
    
    # 5. Evolução Mensal
    evolucao_mensal = df.groupby(['mes_ano_period', 'tipo_movimentacao'])['valor_absoluto'].sum().unstack(fill_value=0)
    evolucao_mensal.index = evolucao_mensal.index.astype(str)
    
    # 6. Gastos Mensais por Categoria (Top 5 categorias)
    top_categorias = gastos_por_categoria.head(5).index
    gastos_mensais_categoria = gastos[gastos['categoria'].isin(top_categorias)].groupby(
        ['mes_ano_period', 'categoria'])['valor_absoluto'].sum().unstack(fill_value=0)
    gastos_mensais_categoria.index = gastos_mensais_categoria.index.astype(str)
    
    # Criar dashboard com múltiplos gráficos
    fig = make_subplots(
        rows=3, cols=2,
        subplot_titles=(
            'Evolução Mensal - Gastos vs Receitas',
            'Distribuição de Gastos por Categoria',
            'Principais Destinatários de Gastos',
            'Principais Fontes de Receitas', 
            'Evolução Mensal das Principais Categorias de Gastos',
            'Distribuição de Gastos por Mês'
        ),
        specs=[
            [{"type": "scatter"}, {"type": "pie"}],
            [{"type": "bar"}, {"type": "bar"}],
            [{"type": "scatter", "colspan": 2}, None]
        ],
        vertical_spacing=0.08,
        horizontal_spacing=0.08
    )
    
    # Gráfico 1: Evolução Mensal
    if 'Gasto' in evolucao_mensal.columns:
        fig.add_trace(
            go.Scatter(x=evolucao_mensal.index, y=evolucao_mensal['Gasto'], 
                      name='Gastos', line=dict(color='red'), mode='lines+markers'),
            row=1, col=1
        )
    if 'Receita' in evolucao_mensal.columns:
        fig.add_trace(
            go.Scatter(x=evolucao_mensal.index, y=evolucao_mensal['Receita'], 
                      name='Receitas', line=dict(color='green'), mode='lines+markers'),
            row=1, col=1
        )
    
    # Gráfico 2: Pizza - Gastos por Categoria
    fig.add_trace(
        go.Pie(labels=gastos_por_categoria.index, values=gastos_por_categoria.values,
               name='Categorias', hole=0.4),
        row=1, col=2
    )
    
    # Gráfico 3: Barras - Principais Destinatários
    fig.add_trace(
        go.Bar(x=principais_destinos.values, y=principais_destinos.index,
               orientation='h', name='Destinatários', marker_color='coral'),
        row=2, col=1
    )
    
    # Gráfico 4: Barras - Principais Fontes
    fig.add_trace(
        go.Bar(x=principais_fontes.values, y=principais_fontes.index,
               orientation='h', name='Fontes Pagadoras', marker_color='lightgreen'),
        row=2, col=2
    )
    
    # Gráfico 5: Evolução das Categorias
    for categoria in top_categorias:
        if categoria in gastos_mensais_categoria.columns:
            fig.add_trace(
                go.Scatter(x=gastos_mensais_categoria.index, 
                          y=gastos_mensais_categoria[categoria],
                          name=categoria, mode='lines+markers'),
                row=3, col=1
            )
    
    fig.update_layout(
        height=1400,
        title_text="Dashboard Financeiro - Análise Completa dos Extratos",
        showlegend=True,
        template="plotly_white"
    )
    
    # Criar HTML completo
    html_content = f"""
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard Financeiro - Extratos Bancários</title>
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #f5f5f5;
            }}
            .container {{
                max-width: 1400px;
                margin: 0 auto;
                background: white;
                padding: 20px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }}
            .header {{
                text-align: center;
                margin-bottom: 30px;
                border-bottom: 2px solid #e0e0e0;
                padding-bottom: 20px;
            }}
            .metrics {{
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 20px;
                margin: 30px 0;
            }}
            .metric-card {{
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 25px;
                border-radius: 10px;
                text-align: center;
                box-shadow: 0 4px 15px rgba(0,0,0,0.1);
            }}
            .metric-value {{
                font-size: 28px;
                font-weight: bold;
                margin: 10px 0;
            }}
            .metric-label {{
                font-size: 14px;
                opacity: 0.9;
            }}
            .positive {{ color: #4CAF50; }}
            .negative {{ color: #f44336; }}
            .charts-container {{
                margin-top: 30px;
            }}
            .summary-section {{
                background: #f8f9fa;
                padding: 20px;
                border-radius: 8px;
                margin: 20px 0;
            }}
            .summary-title {{
                font-size: 18px;
                font-weight: bold;
                margin-bottom: 15px;
                color: #333;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>📊 Dashboard Financeiro - Análise de Extratos</h1>
                <p>Período: {df['data'].min().strftime('%d/%m/%Y')} a {df['data'].max().strftime('%d/%m/%Y')}</p>
            </div>
            
            <div class="metrics">
                <div class="metric-card">
                    <div class="metric-label">Total de Receitas</div>
                    <div class="metric-value positive">R$ {total_receitas:,.2f}</div>
                    <div>{receitas.shape[0]} transações</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Total de Gastos</div>
                    <div class="metric-value negative">R$ {total_gastos:,.2f}</div>
                    <div>{gastos.shape[0]} transações</div>
                </div>
                <div class="metric-card">
                    <div class="metric-label">Saldo Final</div>
                    <div class="metric-value { 'positive' if saldo >= 0 else 'negative' }">R$ {saldo:,.2f}</div>
                    <div>Saldo acumulado</div>
                </div>
            </div>
            
            <div class="summary-section">
                <div class="summary-title">📈 Insights Principais</div>
                <p><strong>Maior Destinatário:</strong> {principais_destinos.index[0]} - R$ {principais_destinos.iloc[0]:,.2f}</p>
                <p><strong>Principal Fonte de Receita:</strong> {principais_fontes.index[0]} - R$ {principais_fontes.iloc[0]:,.2f}</p>
                <p><strong>Categoria com Maior Gasto:</strong> {gastos_por_categoria.index[0]} - R$ {gastos_por_categoria.iloc[0]:,.2f}</p>
                <p><strong>Período Analisado:</strong> {len(df['mes_ano_period'].unique())} meses</p>
            </div>
            
            <div id="graficos"></div>
        </div>
        
        <script>
            var graphs = {fig.to_json()};
            Plotly.newPlot('graficos', graphs.data, graphs.layout);
        </script>
    </body>
    </html>
    """
    
    with open('dashboard_financeiro_bi.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    return html_content

def exportar_resumos_excel_bi(df):
    """
    Exporta resumos detalhados para Excel
    """
    with pd.ExcelWriter('resumo_financeiro_bi.xlsx', engine='openpyxl') as writer:
        
        # 1. Resumo Mensal Detalhado
        resumo_mensal = df.groupby(['ano', 'mes', 'mes_ano_period', 'tipo_movimentacao'])['valor_absoluto'].sum().unstack()
        if 'Gasto' in resumo_mensal.columns and 'Receita' in resumo_mensal.columns:
            resumo_mensal['Saldo'] = resumo_mensal['Receita'] - resumo_mensal['Gasto']
        resumo_mensal.to_excel(writer, sheet_name='Resumo Mensal')
        
        # 2. Gastos por Categoria
        gastos_categoria = df[df['tipo_movimentacao'] == 'Gasto'].groupby('categoria')['valor_absoluto'].sum().sort_values(ascending=False)
        gastos_categoria.to_excel(writer, sheet_name='Gastos por Categoria')
        
        # 3. Principais Destinatários (Top 50)
        destinatarios = df[df['tipo_movimentacao'] == 'Gasto'].groupby('destinatário/pagador')['valor_absoluto'].sum().nlargest(50)
        destinatarios.to_excel(writer, sheet_name='Principais Destinatários')
        
        # 4. Fontes Pagadoras (Top 50)
        fontes = df[df['tipo_movimentacao'] == 'Receita'].groupby('destinatário/pagador')['valor_absoluto'].sum().nlargest(50)
        fontes.to_excel(writer, sheet_name='Fontes Pagadoras')
        
        # 5. Evolução Mensal por Categoria
        evol_categoria = df[df['tipo_movimentacao'] == 'Gasto'].pivot_table(
            values='valor_absoluto', 
            index='mes_ano_period', 
            columns='categoria', 
            aggfunc='sum'
        ).fillna(0)
        evol_categoria.to_excel(writer, sheet_name='Evolução Categorias')
        
        # 6. Top Transações (Maiores Valores)
        top_transacoes = df.nlargest(50, 'valor_absoluto')[['data', 'tipo', 'destinatário/pagador', 'valor', 'categoria', 'mes_ano']]
        top_transacoes.to_excel(writer, sheet_name='Top Transações', index=False)
        
        # 7. Dados Completos
        df.to_excel(writer, sheet_name='Dados Completos', index=False)

# Executar a análise
if __name__ == "__main__":
    try:
        print("Processando arquivo BI.xlsx...")
        
        # Processar dados
        df = processar_extratos_bi("Bi.xlsx")
        
        # Criar categorias automáticas
        df = criar_categorias_automaticas(df)
        
        print(f"Dados processados: {len(df)} transações")
        print(f"Período: {df['data'].min()} a {df['data'].max()}")
        print(f"Total de meses: {len(df['mes_ano_period'].unique())}")
        
        # Criar dashboard HTML
        print("Criando dashboard HTML...")
        criar_dashboard_html_bi(df)
        
        # Exportar resumos para Excel
        print("Exportando resumos para Excel...")
        exportar_resumos_excel_bi(df)
        
        print("\n✅ Arquivos gerados com sucesso!")
        print("📊 dashboard_financeiro_bi.html - Dashboard interativo completo")
        print("📋 resumo_financeiro_bi.xlsx - Relatórios detalhados em Excel")
        print("\n📈 Principais estatísticas:")
        print(f"   - Total de transações: {len(df):,}")
        print(f"   - Total receitas: R$ {df[df['tipo_movimentacao']=='Receita']['valor_absoluto'].sum():,.2f}")
        print(f"   - Total gastos: R$ {df[df['tipo_movimentacao']=='Gasto']['valor_absoluto'].sum():,.2f}")
        print(f"   - Saldo: R$ {(df[df['tipo_movimentacao']=='Receita']['valor_absoluto'].sum() - df[df['tipo_movimentacao']=='Gasto']['valor_absoluto'].sum()):,.2f}")
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        import traceback
        traceback.print_exc()