# utils.py

import pandas as pd

def simplificar_colunas(df):
    """
    Simplifica os nomes das colunas removendo prefixos.
    """
    return df.rename(columns=lambda x: x.split('] ')[-1] if ']' in x else x)

def formatar_moeda(valor):
    """
    Formata um valor como moeda brasileira.
    """
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def preparar_df_financeiro(df_financeiro):
    """
    Processa o DataFrame financeiro: conversão de datas, ordenação, e ajuste de compensação.
    """
    df_financeiro['[Recursos financeiros]   Data do lançamento'] = pd.to_datetime(df_financeiro['[Recursos financeiros]   Data do lançamento'])
    df_financeiro = df_financeiro.sort_values(by='[Recursos financeiros]   Data do lançamento')
    df_financeiro['[Recursos financeiros]   Compensação'] = df_financeiro['[Recursos financeiros]   Compensação'].apply(
        lambda x: 'Sim' if x.lower() == 'sim' else 'não'
    )
    return df_financeiro
