from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
import plotly.express as px

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

def atualizar_abas_com_colunas_personalizadas(df, coluna_criterio, client, nome_planilha):
    """
    Atualiza abas no Google Sheets com dados da aba principal, separando por valor de uma coluna
    e copiando colunas específicas de acordo com o valor da coluna critério, evitando duplicatas.
    
    :param df: DataFrame da aba principal.
    :param coluna_criterio: Nome da coluna que serve de critério para separar os dados.
    :param client: Cliente autorizado do gspread.
    :param nome_planilha: Nome da planilha do Google Sheets.
    """

    # Defina os conjuntos de colunas de interesse com base no valor da coluna critério
    colunas_interesse_dict = {
        'Recursos financeiros': [
            '[Recursos financeiros]   Nome do beneficiado',
            '[Recursos financeiros]   Documento do beneficiado',
            '[Recursos financeiros]   Tipo de lançamento\n',
            '[Recursos financeiros]   Data do lançamento ',
            '[Recursos financeiros]  Valor',
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
            '[Periódico]   Título'
        ],
        'Jornal e Revista': [
            '[Jornal e Revista]   Titulo',
            '[Jornal e Revista]   Autor',
            '[Jornal e Revista]   Ano'
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

atualizar_abas_com_colunas_personalizadas(df, 'Tipo do requerimento', client, 'Relatórios PPG Geografia (Responses)')
