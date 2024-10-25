# auth.py

import gspread
from oauth2client.service_account import ServiceAccountCredentials

def autenticar_google_sheets(credentials_file, sheet_name):
    """
    Autentica o Google Sheets usando credenciais e retorna o cliente e a planilha.
    """
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
    client = gspread.authorize(creds)
    
    try:
        # Tenta abrir a planilha
        sheet = client.open(sheet_name).sheet1
    except gspread.SpreadsheetNotFound:
        # Cria a planilha se ela não existir
        sheet = client.create(sheet_name).sheet1
        sheet.update("A1", [["Link Personalizado", "Orçamento Total"]])  # Cabeçalhos
        sheet.update("A2", [["https://www.exemplo.com", 500000]])  # Valores padrão
    return client, sheet
