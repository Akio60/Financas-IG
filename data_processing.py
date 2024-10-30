# data_processing.py

import pandas as pd

def carregar_dados(sheet):
    """
    Carrega todos os dados da aba principal e retorna um DataFrame.
    """
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def ler_aba(client, nome_planilha, nome_aba):
    """
    Função para ler os dados de uma aba específica no Google Sheets e retornar um DataFrame.
    """
    sheet = client.open(nome_planilha).worksheet(nome_aba)
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def atualizar_abas_com_colunas_personalizadas(df, coluna_criterio, client, nome_planilha):
    """
    Atualiza abas no Google Sheets com dados da aba principal, separando por valor de uma coluna
    e copiando colunas específicas de acordo com o valor da coluna critério, evitando duplicatas.
    """

    # Verificar se a coluna existe
    if coluna_criterio not in df.columns:
        print(f"Erro: A coluna '{coluna_criterio}' não foi encontrada no DataFrame. Colunas disponíveis: {df.columns}")
        return  # Saia da função se a coluna não existir
    
    # Defina os conjuntos de colunas de interesse com base no valor da coluna critério
    colunas_interesse_dict = {
        'Recursos Financeiros': [
            'Tipo do requerimento',
            '[Recursos financeiros]   Nome do beneficiado',
            '[Recursos financeiros]   Documento do beneficiado',
            '[Recursos financeiros]   Tipo de lançamento',
            '[Recursos financeiros]   Data do lançamento',
            '[Recursos financeiros]   Valor',
            '[Recursos financeiros]   Compensação',
            '[Recursos financeiros]   Descrição',
            '[Recursos financeiros]   Email do beneficiado'
        ],
        'Defesas': [
            'Tipo do requerimento',
            '[Defesas]   Data da defesa',
            '[Defesas]   Autor',
            '[Defesas]   Tipo de Defesa',
            '[Defesas]   Título '
        ],
        'Periódico': [
            'Tipo do requerimento',
            '[Periódico]   Autor',
            '[Periódico]   Ano',
            '[Periódico]   Título',
            '[Periódico]   Referencia'
        ],
        'Jornal e Revista': [
            'Tipo do requerimento',
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
            df_atualizado = df_atualizado.fillna('')  # Substituir NaN por string vazia no DataFrame final

        except:
            # Se a aba não existir, criar uma nova aba e considerar todos os dados como novos
            worksheet = sheet.add_worksheet(title=aba_nome, rows="1000", cols="20")
            df_atualizado = df_filtrado

        # Limpar o conteúdo existente na aba e escrever os dados atualizados
        worksheet.clear()

        # Atualizar a aba com os dados (inclui cabeçalhos)
        worksheet.update([df_atualizado.columns.values.tolist()] + df_atualizado.values.tolist())

    print(f"Dados atualizados nas abas com base na coluna {coluna_criterio}, sem duplicatas.")
