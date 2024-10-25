# dashboard.py

from dash import Dash, dcc, html, dash_table
import plotly.express as px
from dash.dependencies import Input, Output
import pandas as pd
from threading import Timer
import webbrowser
from auth import autenticar_google_sheets
from data_processing import carregar_dados, ler_aba
from utils import formatar_moeda, simplificar_colunas

# Variável global para o orçamento total, que poderá ser atualizada
orcamento_total = 500000  # Valor inicial padrão

def set_orcamento_total(novo_orcamento):
    """
    Atualiza o valor do orçamento total.
    """
    global orcamento_total
    orcamento_total = novo_orcamento

def iniciar_dashboard():
    """
    Inicializa o dashboard com dados do Google Sheets e configura o layout e os callbacks.
    """
    # Conexão com Google Sheets
    client, sheet = autenticar_google_sheets('credentials.json', 'Relatórios PPG Geografia (Responses)')
    df = carregar_dados(sheet)

    # Processamento dos dados financeiros e resumo
    df_financeiro = df[df['Tipo do requerimento'] == 'Recursos Financeiros']
    df_financeiro['[Recursos financeiros]   Data do lançamento'] = pd.to_datetime(df_financeiro['[Recursos financeiros]   Data do lançamento'])
    df_financeiro = df_financeiro.sort_values(by='[Recursos financeiros]   Data do lançamento')
    gastos_totais = df_financeiro['[Recursos financeiros]   Valor'].sum()
    saldo_final = orcamento_total - gastos_totais  # Calcula o saldo com o orçamento atualizado

    # Resumo para o gráfico de pizza
    df_resumo = df_financeiro.groupby('[Recursos financeiros]   Tipo de lançamento').agg({'[Recursos financeiros]   Valor': 'sum'}).reset_index()

    # Carregar dados das outras abas
    df_defesas = ler_aba(client, 'Relatórios PPG Geografia (Responses)', 'Defesas')
    df_periodico = ler_aba(client, 'Relatórios PPG Geografia (Responses)', 'Periódico')
    df_jornalerevista = ler_aba(client, 'Relatórios PPG Geografia (Responses)', 'Jornal e Revista')

    # Configurações do Dash
    app = Dash(__name__)

    # Layout do Dashboard
    app.layout = html.Div(children=[
        html.H1(children='Relatório PPG Geografia'),

        html.Div([
            # Gráfico de linha e pizza
            html.Div([
                html.Label('Tipo de Lançamento:', style={'font-size': '15px', 'display': 'inline-block', 'marginRight': '15px'}),
                dcc.Dropdown(
                    id='dropdown-tipo-lancamento',
                    options=[{'label': 'Todos os Tipos', 'value': 'todos'}] +
                             [{'label': tipo, 'value': tipo} for tipo in df_financeiro['[Recursos financeiros]   Tipo de lançamento'].unique()],
                    value='todos',
                    clearable=False
                ),
                dcc.Graph(id='graph-valores-tempo')
            ], style={'width': '33%', 'display': 'inline-block'}),

            html.Div([
                dcc.Graph(id='graph-pizza')
            ], style={'width': '33%', 'display': 'inline-block'}),

            html.Div([
                dash_table.DataTable(
                    id='table-resumo',
                    columns=[
                        {"name": "Tipo de Lançamento", "id": '[Recursos financeiros]   Tipo de lançamento'},
                        {"name": "Soma dos Valores", "id": '[Recursos financeiros]   Valor'}
                    ],
                    data=df_resumo.to_dict('records'),
                    style_table={'maxHeight': '400px', 'overflowY': 'auto'}
                ),
                html.P(f"Total de Gastos: {formatar_moeda(gastos_totais)}"),
                html.P(f"Orçamento Total: {formatar_moeda(orcamento_total)}"),
                html.P(f"Saldo Final: {formatar_moeda(saldo_final)}")
            ], style={'width': '33%', 'display': 'inline-block', 'verticalAlign': 'top'})
        ], style={'display': 'flex'}),

        # Seção das tabelas inferiores
        html.Div([
            # Tabela de Status de Compensação
            html.Div([
                html.H2(children='Status de Compensação'),
                dcc.Dropdown(
                    id='dropdown-compensacao',
                    options=[
                        {'label': 'Todos os Nomes', 'value': 'todos'},
                        {'label': 'Pagantes', 'value': 'Sim'},
                        {'label': 'Pendente de Pagamento', 'value': 'não'}
                    ],
                    value='todos',
                    clearable=False
                ),
                dash_table.DataTable(
                    id='table-compensacao',
                    columns=[
                        {"name": "Nome", "id": '[Recursos financeiros]   Nome do beneficiado'},
                        {"name": "Tipo de Lançamento", "id": '[Recursos financeiros]   Tipo de lançamento'},
                        {"name": "Valor", "id": '[Recursos financeiros]   Valor'},
                        {"name": "Pago", "id": '[Recursos financeiros]   Compensação'}
                    ],
                    data=df_financeiro.to_dict('records'),
                    style_table={'maxHeight': '280px', 'overflowY': 'auto'}
                )
            ], style={'width': '45%', 'display': 'inline-block', 'marginRight': '5%'}),

            # Tabela de Visualização de Outras Abas
            html.Div([
                html.H2(children='Visualização de Outras Abas'),
                dcc.Dropdown(
                    id='dropdown-visualizacao',
                    options=[
                        {'label': 'Defesas', 'value': 'defesas'},
                        {'label': 'Periódico', 'value': 'periodico'},
                        {'label': 'Jornal e Revista', 'value': 'jornalerevista'}
                    ],
                    value='defesas',
                    clearable=False
                ),
                html.Div(id='tabela-aba')
            ], style={'width': '45%', 'display': 'inline-block'})
        ], style={'display': 'flex', 'flex-direction': 'row', 'marginTop': '30px'}),
    ])

    # Callback para atualizar o gráfico de linha
    @app.callback(
        Output('graph-valores-tempo', 'figure'),
        [Input('dropdown-tipo-lancamento', 'value')]
    )
    def atualizar_grafico(tipo_lancamento):
        df_filtrado = df_financeiro if tipo_lancamento == 'todos' else df_financeiro[df_financeiro['[Recursos financeiros]   Tipo de lançamento'] == tipo_lancamento]
        df_filtrado['Valor Acumulado'] = df_filtrado['[Recursos financeiros]   Valor'].astype(float).cumsum()
        fig = px.line(df_filtrado, x='[Recursos financeiros]   Data do lançamento', y='Valor Acumulado')
        return fig

    # Callback para atualizar o gráfico de pizza
    @app.callback(
        Output('graph-pizza', 'figure'),
        [Input('dropdown-tipo-lancamento', 'value')]
    )
    def atualizar_grafico_pizza(tipo_lancamento):
        fig = px.pie(df_resumo, values='[Recursos financeiros]   Valor', names='[Recursos financeiros]   Tipo de lançamento')
        return fig

    # Callback para atualizar a tabela de compensação com base no status de compensação
    @app.callback(
        Output('table-compensacao', 'data'),
        [Input('dropdown-compensacao', 'value')]
    )
    def atualizar_tabela_compensacao(status_compensacao):
        df_filtrado = df_financeiro if status_compensacao == 'todos' else df_financeiro[df_financeiro['[Recursos financeiros]   Compensação'] == status_compensacao]
        return df_filtrado.to_dict('records')

    # Callback para atualizar a tabela com base na seleção do dropdown de visualização
    @app.callback(
        Output('tabela-aba', 'children'),
        [Input('dropdown-visualizacao', 'value')]
    )
    def atualizar_tabela(aba_selecionada):
        if aba_selecionada == 'defesas':
            df_selecionado = simplificar_colunas(df_defesas)
        elif aba_selecionada == 'periodico':
            df_selecionado = simplificar_colunas(df_periodico)
        else:
            df_selecionado = simplificar_colunas(df_jornalerevista)

        return dash_table.DataTable(
            columns=[{"name": i, "id": i} for i in df_selecionado.columns],
            data=df_selecionado.to_dict('records'),
            style_table={'maxHeight': '400px', 'overflowY': 'auto'}
        )

    # Abrir o navegador
    def open_browser():
        webbrowser.open_new('http://127.0.0.1:8050/')
    Timer(1, open_browser).start()
    app.run_server(debug=True, use_reloader=False)
