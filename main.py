# main.py

import tkinter as tk
from app_gui import iniciar_interface  # Importar a função de interface do Tkinter

# Inicializar a aplicação Tkinter com a interface gráfica
#if __name__ == '__main__':
#    iniciar_interface()

import tkinter as tk
from tkinter import ttk
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials

# Conectar com o Google Sheets usando gspread
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Acesse a planilha
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1sNwhkq0nCuTMRhs2HmahV88uIn9KiXY1ex0vlOwC0O8/edit?usp=sharing").sheet1

# Função para carregar dados
def load_data():
    records = sheet.get_all_records()
    return pd.DataFrame(records)

# Função para atualizar o status de uma célula específica
def update_status(row, new_status):
    cell = sheet.find(row['Status'])
    sheet.update_cell(cell.row, cell.col, new_status)

# Interface gráfica
class App:
    def __init__(self, root):
        self.root = root
        root.title("Aplicativo de Visualização e Atualização de Dados")
        
        # Botão para visualizar dados
        self.view_button = tk.Button(root, text="Visualizar Dados", command=self.open_data_view)
        self.view_button.pack(pady=20)
        
    def open_data_view(self):
        # Janela para visualização de dados
        view_window = tk.Toplevel(self.root)
        view_window.title("Dados da Planilha")
        
        # Carregar dados
        data = load_data()
        
        # Treeview para exibir dados
        tree = ttk.Treeview(view_window, columns=data.columns.tolist(), show="headings")
        for col in data.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")
        for _, row in data.iterrows():
            tree.insert("", "end", values=row.tolist())
        tree.pack(fill="both", expand=True)
        
        # Botão de atualização de status
        update_button = tk.Button(view_window, text="Atualizar Status", command=lambda: self.update_selected_status(tree))
        update_button.pack(pady=10)
        
    def update_selected_status(self, tree):
        # Obter item selecionado
        selected_item = tree.selection()[0]
        selected_data = tree.item(selected_item, "values")
        
        # Atualizar coluna Status na planilha
        new_status = "Novo Status"
        update_status(selected_data, new_status)
        
        # Atualização visual
        tree.item(selected_item, values=[*selected_data[:-1], new_status])
        tk.messagebox.showinfo("Sucesso", "Status atualizado com sucesso.")

# Inicializar aplicação
root = tk.Tk()
app = App(root)
root.mainloop()
