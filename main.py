from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
from dash import Dash, dcc, html
import dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import webbrowser
from threading import Timer

# Definir o escopo da API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# Autenticação com as credenciais
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Acessar a planilha pelo nome ou URL
sheet = client.open('Relatórios PPG Geografia (Responses)').sheet1  # ou use .worksheet('NOME_DA_ABA')

# Pegar todos os dados da planilha
data = sheet.get_all_records()

# Converter para DataFrame
df = pd.DataFrame(data)


#------------------------------------------------#
#Tratamento de dados do sheets 
def atualizar_abas_com_colunas_personalizadas(df, coluna_criterio, client, nome_planilha):
    """
    Atualiza abas no Google Sheets com dados da aba principal, separando por valor de uma coluna
    e copiando colunas específicas de acordo com o valor da coluna critério, evitando duplicatas.
    """

    # Defina os conjuntos de colunas de interesse com base no valor da coluna critério
    colunas_interesse_dict = {
        'Recursos financeiros': [
            '[Recursos financeiros]   Nome do beneficiado',
            '[Recursos financeiros]   Documento do beneficiado',
            '[Recursos financeiros]   Tipo de lançamento\n',
            '[Recursos financeiros]   Data do lançamento ',
            '[Recursos financeiros]   Valor',
            '[Recursos financeiros]   Compensação',
            '[Recursos financeiros]   Descrição',
            '[Recursos financeiros]   Enviar email automático de recebimento do valor?'
        ],
        'Defesas': [
            '[Defesas]   Data da defesa',
            '[Defesas]   Autor',
            '[Defesas]   Tipo de Defesa',
            '[Defesas]   Título '
        ],
        'Periódico': [
            '[Periódico]   Autor',
            '[Periódico]   Ano',
            '[Periódico]   Título',
            '[Periódico]   Referencia'
        ],
        'Jornal e Revista': [
            '[Jornal e Revista]   Autor',
            '[Jornal e Revista]   Ano',
            '[Jornal e Revista]   Titulo',
        ]
    }

    # Obter os valores únicos da coluna que serve como critério
    valores_unicos = df[coluna_criterio].unique()

    # Acessar a planilha no Google Sheets
    sheet = client.open(nome_planilha)
    
    # Para cada valor único na coluna de critério, processar os dados
    for valor in valores_unicos:
        # Verifica se há colunas de interesse definidas para o valor
        colunas_interesse = colunas_interesse_dict.get(valor)
        
        if colunas_interesse is None:
            print(f"Não há colunas de interesse definidas para o valor '{valor}' na coluna critério.")
            continue

        # Filtrar o DataFrame pelos dados do valor da coluna_criterio
        df_filtrado = df[df[coluna_criterio] == valor][colunas_interesse]
        
        # Substituir NaN por string vazia para evitar erro de JSON
        df_filtrado = df_filtrado.fillna('')

        # Nome da aba correspondente ao valor da coluna_criterio
        aba_nome = str(valor)

        # Tentar acessar a aba, senão criar uma nova
        try:
            worksheet = sheet.worksheet(aba_nome)
            # Pegar todos os dados existentes na aba e transformá-los em um DataFrame
            dados_existentes = pd.DataFrame(worksheet.get_all_records())

            # Combinar novos dados com os existentes, removendo duplicatas
            df_atualizado = pd.concat([dados_existentes, df_filtrado]).drop_duplicates()

        except:
            # Se a aba não existir, criar uma nova aba e considerar todos os dados como novos
            worksheet = sheet.add_worksheet(title=aba_nome, rows="1000", cols="20")
            df_atualizado = df_filtrado

        # Limpar o conteúdo existente na aba e escrever os dados atualizados
        worksheet.clear()

        # Atualizar a aba com os dados (inclui cabeçalhos)
        worksheet.update([df_atualizado.columns.values.tolist()] + df_atualizado.values.tolist())

    print(f"Dados atualizados nas abas com base na coluna {coluna_criterio}, sem duplicatas.")

# Chamada da função
atualizar_abas_com_colunas_personalizadas(df, 'Tipo do requerimento', client, 'Relatórios PPG Geografia (Responses)')

#------------------------------------------------#


from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
from dash import Dash, dcc, html
import dash_table
from dash.dependencies import Input, Output
import plotly.express as px
import webbrowser
from threading import Timer

# Definir o escopo da API
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

# Autenticação com as credenciais
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Função para ler os dados das abas
def ler_aba(nome_planilha, nome_aba):
    """
    Função para ler os dados de uma aba específica no Google Sheets e retornar um DataFrame.
    """
    sheet = client.open(nome_planilha).worksheet(nome_aba)
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    return df

# Carregar os dados de Recursos Financeiros
df_financeiro = ler_aba('Relatórios PPG Geografia (Responses)', 'Recursos financeiros')
df_defesas = ler_aba('Relatórios PPG Geografia (Responses)', 'Defesas')
df_periodico = ler_aba('Relatórios PPG Geografia (Responses)', 'Periódico')
df_jornalerevista = ler_aba('Relatórios PPG Geografia (Responses)', 'Jornal e Revista')

# Converter a coluna de Data do Lançamento para o formato de data
df_financeiro['[Recursos financeiros]   Data do lançamento'] = pd.to_datetime(df_financeiro['[Recursos financeiros]   Data do lançamento'])

# Ordenar os dados pela Data do Lançamento
df_financeiro = df_financeiro.sort_values(by='[Recursos financeiros]   Data do lançamento')

# Criar uma coluna de status de compensação ajustada
df_financeiro['[Recursos financeiros]   Compensação'] = df_financeiro['[Recursos financeiros]   Compensação'].apply(
    lambda x: 'sim' if x.lower() == 'sim' else 'não'
)

# Tabela de resumo por tipo de lançamento
df_resumo = df_financeiro.groupby('[Recursos financeiros]   Tipo de lançamento\n').agg({'[Recursos financeiros]   Valor': 'sum'}).reset_index()

# Valor arbitrário do orçamento
orcamento_total = 500000
gastos_totais = df_resumo['[Recursos financeiros]   Valor'].sum()
saldo_final = orcamento_total - gastos_totais

# Função para remover prefixos de colunas para a visualização de abas
def simplificar_colunas(df):
    df_simplificado = df.rename(columns=lambda x: x.split('] ')[-1] if ']' in x else x)
    return df_simplificado

# Criar o app Dash
app = Dash(__name__)

# Layout do Dashboard
app.layout = html.Div(children=[
    html.H1(children='Dashboard de Relatórios PPG Geografia'),

    # Layout da parte superior com gráfico de linha, gráfico de pizza e tabela
    html.Div([
        # Gráfico de linha na parte superior esquerda
        html.Div([
            html.Label('Tipo de Lançamento:', style={'font-size': '12px', 'display': 'inline-block', 'marginRight': '15px'}),
            dcc.Dropdown(
                id='dropdown-tipo-lancamento',
                options=[{'label': 'Todos os Tipos', 'value': 'todos'}] +
                         [{'label': tipo, 'value': tipo} for tipo in df_financeiro['[Recursos financeiros]   Tipo de lançamento\n'].unique()],
                value='todos',  # Valor inicial: Todos os Tipos
                clearable=False,
                style={'width': 'auto', 'display': 'inline-block', 'min-width': '400px', 'max-width': '600px'}
            ),
            dcc.Graph(id='graph-valores-tempo', config={'displayModeBar': False}, style={'height': '300px'})
        ], style={'width': '33%', 'display': 'inline-block'}),  # Tamanho ajustado para ficar ao lado do gráfico de pizza e da tabela

        # Gráfico de pizza no meio
        html.Div([
            dcc.Graph(id='graph-pizza', config={'displayModeBar': False}, style={'height': '300px'})
        ], style={'width': '33%', 'display': 'inline-block'}),  # Tamanho ajustado para ficar entre o gráfico de linha e a tabela

        # Tabela de resumo de tipos de lançamento à direita
        html.Div([
            dash_table.DataTable(
                id='table-resumo',
                columns=[
                    {"name": "Tipo de Lançamento", "id": '[Recursos financeiros]   Tipo de lançamento\n'},
                    {"name": "Soma dos Valores", "id": '[Recursos financeiros]   Valor'}
                ],
                data=df_resumo.to_dict('records'),
                style_table={
                    'maxHeight': '400px',
                    'overflowY': 'auto',
                    'border': '1px solid black'
                },
                fixed_rows={'headers': True},
                style_cell={'textAlign': 'left', 'padding': '8px', 'whiteSpace': 'normal', 'height': 'auto'}
            ),
            html.Br(),
            html.P(f"Total de Gastos: R$ {gastos_totais:,.2f}", style={'font-size': '12px'}),
            html.P(f"Orçamento Total: R$ {orcamento_total:,.2f}", style={'font-size': '12px'}),
            html.P(f"Saldo Final: R$ {saldo_final:,.2f}", style={'font-size': '12px'})
        ], style={'width': '33%', 'display': 'inline-block', 'verticalAlign': 'top'})
    ], style={'display': 'flex', 'flex-direction': 'row'}),

    # Seletor e visualização das duas tabelas lado a lado
    html.Div([
        # Tabela de compensação à esquerda
        html.Div([
            html.H2(children='Status de Compensação', style={'display': 'inline-block', 'marginRight': '20px'}),
            dcc.Dropdown(
                id='dropdown-compensacao',
                options=[
                    {'label': 'Todos os Nomes', 'value': 'todos'},
                    {'label': 'Pagantes', 'value': 'sim'},
                    {'label': 'Pendente de Pagamento', 'value': 'não'}
                ],
                value='todos',
                clearable=False,
                style={'width': '200px', 'display': 'inline-block'}
            ),
            dash_table.DataTable(
                id='table-compensacao',
                columns=[
                    {"name": "Nome", "id": '[Recursos financeiros]   Nome do beneficiado'},
                    {"name": "Tipo de Lançamento", "id": '[Recursos financeiros]   Tipo de lançamento\n'},
                    {"name": "Valor", "id": '[Recursos financeiros]   Valor'},
                    {"name": "Compensação", "id": '[Recursos financeiros]   Compensação'}
                ],
                style_table={'maxHeight': '280px', 'overflowY': 'auto', 'border': '1px solid black'},
                fixed_rows={'headers': True},
                style_cell={'textAlign': 'left', 'padding': '10px', 'whiteSpace': 'normal', 'height': 'auto'}
            )
        ], style={'width': '45%', 'display': 'inline-block', 'marginRight': '5%'}),

        # Tabela de visualização das outras abas à direita
        html.Div([
            html.H2(children='Visualização de Outras Abas', style={'display': 'inline-block', 'marginRight': '20px'}),
            dcc.Dropdown(
                id='dropdown-visualizacao',
                options=[
                    {'label': 'Defesas', 'value': 'defesas'},
                    {'label': 'Periódico', 'value': 'periodico'},
                    {'label': 'Jornal e Revista', 'value': 'jornalerevista'}
                ],
                value='defesas',
                clearable=False,
                style={'width': '200px', 'display': 'inline-block'}
            ),
            html.Div(id='tabela-aba')
        ], style={'width': '45%', 'display': 'inline-block'})
    ], style={'display': 'flex', 'flex-direction': 'row', 'marginTop': '30px'})
])

# Callback para atualizar o gráfico de linha com base no tipo de lançamento selecionado
@app.callback(
    Output('graph-valores-tempo', 'figure'),
    [Input('dropdown-tipo-lancamento', 'value')]
)
def atualizar_grafico(tipo_lancamento):
    if tipo_lancamento == 'todos':
        df_filtrado = df_financeiro.copy()  # Todos os tipos
    else:
        df_filtrado = df_financeiro[df_financeiro['[Recursos financeiros]   Tipo de lançamento\n'] == tipo_lancamento]

    df_filtrado['Valor Acumulado'] = df_filtrado['[Recursos financeiros]   Valor'].cumsum()

    fig = px.line(df_filtrado, 
                  x='[Recursos financeiros]   Data do lançamento', 
                  y='Valor Acumulado', 
                  title='')  # Remover título do gráfico
    return fig

# Callback para atualizar o gráfico de pizza com os tipos de lançamento
@app.callback(
    Output('graph-pizza', 'figure'),
    [Input('dropdown-tipo-lancamento', 'value')]
)
def atualizar_grafico_pizza(tipo_lancamento):
    fig = px.pie(df_resumo, values='[Recursos financeiros]   Valor', names='[Recursos financeiros]   Tipo de lançamento\n', 
                 title='Distribuição dos Tipos de Lançamento', color_discrete_sequence=px.colors.qualitative.Set3)
    return fig

# Callback para atualizar a tabela de compensação com base no status de compensação
@app.callback(
    Output('table-compensacao', 'data'),
    [Input('dropdown-compensacao', 'value')]
)
def atualizar_tabela_compensacao(status_compensacao):
    if status_compensacao == 'todos':
        df_filtrado = df_financeiro
    else:
        df_filtrado = df_financeiro[df_financeiro['[Recursos financeiros]   Compensação'] == status_compensacao]

    return df_filtrado.to_dict('records')

# Callback para atualizar a tabela com base na seleção do dropdown e simplificar colunas
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
        style_table={'maxHeight': '400px', 'overflowY': 'auto', 'border': '1px solid black'},
        fixed_rows={'headers': True},
        style_cell={'textAlign': 'left', 'padding': '10px', 'whiteSpace': 'normal', 'height': '60px'}
    )

# Função para abrir o navegador automaticamente
def open_browser():
    webbrowser.open_new('http://127.0.0.1:8050/')

# Rodar o servidor e abrir o navegador automaticamente
if __name__ == '__main__':
    Timer(1, open_browser).start()  # Abre o navegador automaticamente após 1 segundo
    app.run_server(debug=True)
