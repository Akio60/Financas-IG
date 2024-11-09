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

# Acessar a planilha
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1sNwhkq0nCuTMRhs2HmahV88uIn9KiXY1ex0vlOwC0O8/edit?usp=sharing").sheet1

# Obter cabeçalhos diretamente da planilha
sheet_headers = sheet.row_values(1)

# Mapear nome da coluna para índice
column_indices = {name: idx + 1 for idx, name in enumerate(sheet_headers)}

# Função para carregar dados
def load_data():
    records = sheet.get_all_records()
    return pd.DataFrame(records)

# Carregar dados para obter a lista de colunas
data = load_data()
all_columns = data.columns.tolist()

# Lista de colunas a serem exibidas (edite esta lista para mostrar ou ocultar colunas)
columns_to_display = [
    'Carimbo de data/hora',  # Certifique-se de incluir o carimbo de data e hora
    'Status',
    'Nome completo (sem abreviações):',
    'Curso:',
    'Orientador',
    'Possui bolsa?',
    'Motivo da solicitação'
]

# Lista de colunas a serem exibidas na visualização dos detalhes
detail_columns_to_display = [
    'Carimbo de data/hora',
    'Status',
    'Nome completo (sem abreviações):',
    'Curso:',
    'Orientador',
    'Possui bolsa?'
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
        self.status_options = [
            "Pago",
            "Aguardando documentação",
            "Autorizado",
            "Em processamento",
            "Cancelado",
            "Pronto para pagamento"
        ]

        for status in self.status_options:
            btn = tk.Button(button_frame, text=f"Visualizar Status: {status}", command=lambda s=status: self.update_table(status_filter=s))
            btn.pack(pady=5, padx=10, fill="x")

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
        self.tree.bind("<Double-1>", self.on_treeview_click)

        # Atualizar tabela inicialmente com todos os dados
        self.update_table()

    def update_table(self, status_filter=None):
        # Limpar a tabela atual
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Recarregar dados
        global data
        data = load_data()

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
        selected_item = self.tree.selection()
        if selected_item:
            row_index = int(selected_item[0])
            row_data = data.loc[row_index]
            # Abrir nova janela com os detalhes selecionados
            self.show_details_window(row_data)

    def show_details_window(self, row_data):
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
        status_var = StringVar(details_window)
        status_var.set(row_data['Status'])

        status_label = tk.Label(details_window, text="Editar Status:")
        status_label.pack(anchor="w", padx=10, pady=5)
        status_dropdown = OptionMenu(details_window, status_var, *self.status_options)
        status_dropdown.pack(anchor="w", padx=10, pady=5)

        # Botão para confirmar a edição do status
        def confirm_edit_status():
            new_status = status_var.get()
            confirm = messagebox.askyesno("Confirmação", f"Deseja realmente alterar o status para '{new_status}'?")
            if confirm:
                try:
                    # Recarregar dados
                    global data
                    data = load_data()
                    # Usar 'Carimbo de data/hora' para encontrar a linha correta
                    timestamp_value = row_data['Carimbo de data/hora']
                    # Encontrar a linha que corresponde ao carimbo de data/hora
                    cell_list = sheet.col_values(column_indices['Carimbo de data/hora'])
                    # Como o get_all_records pode ter convertido o timestamp, precisamos garantir que o formato corresponde
                    for idx, cell_value in enumerate(cell_list[1:], start=2):  # Começa em 2 para pular o cabeçalho
                        if cell_value == timestamp_value:
                            row_number = idx
                            break
                    else:
                        messagebox.showerror("Erro", "Não foi possível encontrar o registro na planilha.")
                        return

                    # Atualizar o status na planilha do Google Sheets
                    sheet.update_cell(row_number, column_indices['Status'], new_status)
                    messagebox.showinfo("Sucesso", "Status atualizado com sucesso!")
                    # Atualizar a tabela na interface
                    self.update_table()
                    details_window.destroy()
                except Exception as e:
                    messagebox.showerror("Erro", f"Ocorreu um erro ao atualizar o status: {e}")

        edit_button = tk.Button(details_window, text="Confirmar Edição", command=confirm_edit_status)
        edit_button.pack(pady=10)

        # Botão de envio de e-mail
        def send_email():
            # Implementar funcionalidade de envio de e-mail aqui
            messagebox.showinfo("E-mail", "Função de envio de e-mail ainda não implementada.")

        email_button = tk.Button(details_window, text="Enviar E-mail", command=send_email)
        email_button.pack(pady=10)

# Inicializar aplicação
root = tk.Tk()
app = App(root)
root.mainloop()
