# main.py

import tkinter as tk
from tkinter import ttk, messagebox, StringVar, OptionMenu
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

# Carregar dados para obter a lista de colunas
data = load_data()
all_columns = data.columns.tolist()

# Lista de colunas a serem exibidas (edite esta lista para mostrar ou ocultar colunas)
columns_to_display = [
    'Nome completo (sem abreviações):',
    'Curso:',
    'Orientador',
    'Possui bolsa?',
    'Motivo da solicitação',
    'Status'
]

# Lista de colunas a serem exibidas na visualização dos detalhes (edite esta lista para personalizar os detalhes)
detail_columns_to_display = [
    'Nome completo (sem abreviações):',
    'Curso:',
    'Orientador',
    'Possui bolsa?',
    'Status'
]

# Se a lista estiver vazia, exibir todas as colunas
if not columns_to_display:
    columns_to_display = all_columns

# Interface gráfica
class App:
    def __init__(self, root):
        self.root = root
        root.title("Aplicativo de Visualização de Dados")
        root.geometry("1000x600")

        # Frame principal para dividir a tela
        main_frame = tk.Frame(root)
        main_frame.pack(fill="both", expand=True)

        # Frame para os botões à esquerda
        button_frame = tk.Frame(main_frame, width=200, bg="lightgrey")
        button_frame.pack(side="left", fill="y")

        # Botões de filtro
        self.view_paid_button = tk.Button(button_frame, text="Visualizar Status: Pago", command=lambda: self.update_table(status_filter="Pago"))
        self.view_paid_button.pack(pady=10, padx=10, fill="x")
        
        self.view_waiting_docs_button = tk.Button(button_frame, text="Visualizar Status: Aguardando documentação", command=lambda: self.update_table(status_filter="Aguardando documentação"))
        self.view_waiting_docs_button.pack(pady=10, padx=10, fill="x")
        
        self.view_authorized_button = tk.Button(button_frame, text="Visualizar Status: Autorizado", command=lambda: self.update_table(status_filter="Autorizado"))
        self.view_authorized_button.pack(pady=10, padx=10, fill="x")
        
        self.view_processing_button = tk.Button(button_frame, text="Visualizar Status: Em processamento", command=lambda: self.update_table(status_filter="Em processamento"))
        self.view_processing_button.pack(pady=10, padx=10, fill="x")
        
        self.view_canceled_button = tk.Button(button_frame, text="Visualizar Status: Cancelado", command=lambda: self.update_table(status_filter="Cancelado"))
        self.view_canceled_button.pack(pady=10, padx=10, fill="x")
        
        self.view_ready_payment_button = tk.Button(button_frame, text="Visualizar Status: Pronto para pagamento", command=lambda: self.update_table(status_filter="Pronto para pagamento"))
        self.view_ready_payment_button.pack(pady=10, padx=10, fill="x")
        
        self.view_all_button = tk.Button(button_frame, text="Visualizar Todos os Dados", command=lambda: self.update_table())
        self.view_all_button.pack(pady=10, padx=10, fill="x")

        # Frame para a tabela à direita
        table_frame = tk.Frame(main_frame)
        table_frame.pack(side="right", fill="both", expand=True)

        # Treeview para exibir dados
        self.tree = ttk.Treeview(table_frame, columns=columns_to_display, show="headings")
        for col in columns_to_display:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")
        self.tree.pack(fill="both", expand=True)

        # Scrollbar para o Treeview
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Adicionar evento de clique no Treeview para exibir detalhes
        self.tree.bind("<Double-1>", lambda event: self.on_treeview_click(event))

        # Atualizar tabela inicialmente com todos os dados
        self.update_table()

    def update_table(self, status_filter=None):
        # Limpar a tabela atual
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Filtrar apenas as colunas desejadas
        try:
            data_filtered = data[columns_to_display]
        except KeyError as e:
            messagebox.showerror("Erro", f"Coluna não encontrada: {e}")
            return

        # Aplicar filtro de status, se especificado
        if status_filter:
            data_filtered = data_filtered[data_filtered['Status'] == status_filter]

        # Inserir os dados filtrados na tabela
        for idx, row in data_filtered.iterrows():
            self.tree.insert("", "end", iid=idx, values=row.tolist())

    def on_treeview_click(self, event):
        # Obter item selecionado
        selected_item = self.tree.selection()[0]
        row_index = int(selected_item)
        row_data = data.loc[row_index]

        # Abrir nova janela com os detalhes selecionados
        self.show_details_window(row_data, row_index)

    def show_details_window(self, row_data, row_index):
        # Janela para exibir os detalhes do item selecionado
        details_window = tk.Toplevel(self.root)
        details_window.title("Detalhes do Item Selecionado")

        # Exibir apenas as colunas definidas em detail_columns_to_display
        for col in detail_columns_to_display:
            if col in row_data:
                value = row_data[col]
                label = tk.Label(details_window, text=f"{col}: {value}")
                label.pack(anchor="w", padx=10, pady=2)

        # Adicionar dropdown para editar o status
        status_options = [
            "Pago",
            "Aguardando documentação",
            "Autorizado",
            "Em processamento",
            "Cancelado",
            "Pronto para pagamento"
        ]
        status_var = StringVar(details_window)
        status_var.set(row_data['Status'])

        status_label = tk.Label(details_window, text="Editar Status:")
        status_label.pack(anchor="w", padx=10, pady=5)
        status_dropdown = OptionMenu(details_window, status_var, *status_options)
        status_dropdown.pack(anchor="w", padx=10, pady=5)

        # Botão para confirmar a edição do status
        def confirm_edit_status():
            new_status = status_var.get()
            confirm = messagebox.askyesno("Confirmação", f"Deseja realmente alterar o status para '{new_status}'?")
            if confirm:
                # Atualizar o status na planilha do Google Sheets
                cell = sheet.find(row_data['Nome completo (sem abreviações):'])
                sheet.update_cell(cell.row, all_columns.index('Status') + 1, new_status)
                messagebox.showinfo("Sucesso", "Status atualizado com sucesso!")
                # Recarregar os dados e atualizar a tabela na interface
                global data
                data = load_data()
                self.update_table()

        edit_button = tk.Button(details_window, text="Confirmar Edição", command=confirm_edit_status)
        edit_button.pack(pady=10)

        # Botão de envio de e-mail (ainda sem função)
        email_button = tk.Button(details_window, text="Enviar E-mail")
        email_button.pack(pady=10)

# Inicializar aplicação
root = tk.Tk()
app = App(root)
root.mainloop()
