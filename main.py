import os
import tkinter as tk
from tkinter import ttk, messagebox, Text, Entry
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Carregar variável de ambiente para a senha do email
os.environ['EMAIL_PASSWORD'] = 'senha'

# Lista de colunas da planilha
ALL_COLUMNS_detail = [
    'Valor', 'Carimbo de data/hora', 'Endereço de e-mail', 'Nome completo (sem abreviações):',
    'Ano de ingresso o PPG:', 'Curso:', 'Orientador', 'Possui bolsa?', 'Qual a agência de fomento?',
    'Título do projeto do qual participa:', 'Motivo da solicitação',
    'Nome do evento ou, se atividade de campo, motivos da realização\n* caso não se trate de evento ou viagem de campo, preencher N/A',
    'Local de realização do evento', 'Período de realização da atividade. Indique as datas (dd/mm/aaaa)',
    'Descrever detalhadamente os itens a serem financiados. Por ex: inscrição em evento, diárias (para transporte, hospedagem e alimentação), passagem aérea, pagamento de análises e traduções, etc.\n',
    'Telefone de contato:', 'E-mail DAC:', 'Endereço completo (logradouro, número, bairro, cidade e estado)',
    'CPF:', 'RG/RNE:', 'Dados bancários (banco, agência e conta) '
]

ALL_COLUMNS = [
    'EmailRecebimento', 'EmailProcessando', 'EmailCancelamento', 'EmailAutorizado', 'EmailDocumentacao',
    'EmailPago', 'Valor', 'Status', 'Ultima Atualizacao', 'Carimbo de data/hora', 'Endereço de e-mail',
    'Declaro que li e estou ciente das regras e obrigações dispostas neste formulário', 'Nome completo (sem abreviações):',
    'Ano de ingresso o PPG:', 'Curso:', 'Orientador', 'Possui bolsa?', 'Qual a agência de fomento?',
    'Título do projeto do qual participa:', 'Motivo da solicitação',
    'Nome do evento ou, se atividade de campo, motivos da realização\n* caso não se trate de evento ou viagem de campo, preencher N/A',
    'Local de realização do evento', 'Período de realização da atividade. Indique as datas (dd/mm/aaaa)',
    'Descrever detalhadamente os itens a serem financiados. Por ex: inscrição em evento, diárias (para transporte, hospedagem e alimentação), passagem aérea, pagamento de análises e traduções, etc.\n',
    'E-mail DAC:', 'Endereço completo (logradouro, número, bairro, cidade e estado)', 'Telefone de contato:', 'CPF:', 'RG/RNE:',
    'Dados bancários (banco, agência e conta) '
]

# Google Sheets Handler
class GoogleSheetsHandler:
    def __init__(self, credentials_file, sheet_url):
        scope = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        self.client = gspread.authorize(creds)
        self.sheet = self.client.open_by_url(sheet_url).sheet1
        self.column_indices = {name: idx + 1 for idx, name in enumerate(self.sheet.row_values(1))}

    def load_data(self):
        records = self.sheet.get_all_records()
        return pd.DataFrame(records)

    def update_status(self, timestamp_value, new_status):
        cell_list = self.sheet.col_values(self.column_indices['Carimbo de data/hora'])
        for idx, cell_value in enumerate(cell_list[1:], start=2):
            if cell_value == timestamp_value:
                row_number = idx
                self.sheet.update_cell(row_number, self.column_indices['Status'], new_status)
                return True
        return False

    def update_value(self, timestamp_value, new_value):
        cell_list = self.sheet.col_values(self.column_indices['Carimbo de data/hora'])
        for idx, cell_value in enumerate(cell_list[1:], start=2):
            if cell_value == timestamp_value:
                row_number = idx
                self.sheet.update_cell(row_number, self.column_indices['Valor'], new_value)
                return True
        return False

    def update_cell(self, timestamp_value, column_name, new_value):
        cell_list = self.sheet.col_values(self.column_indices['Carimbo de data/hora'])
        for idx, cell_value in enumerate(cell_list[1:], start=2):
            if cell_value == timestamp_value:
                row_number = idx
                if column_name in self.column_indices:
                    col_number = self.column_indices[column_name]
                    self.sheet.update_cell(row_number, col_number, new_value)
                    return True
        return False

# Email Sender
class EmailSender:
    def __init__(self, smtp_server, smtp_port, sender_email):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.sender_email = sender_email
        self.sender_password = os.getenv('EMAIL_PASSWORD')

    def send_email(self, recipient, subject, body):
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            # Enviar o e-mail
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            server.sendmail(self.sender_email, recipient, msg.as_string())
            server.quit()

            messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao enviar o e-mail: {e}")

# Main Application Class
class App:
    def __init__(self, root, sheets_handler, email_sender):
        self.root = root
        self.sheets_handler = sheets_handler
        self.email_sender = email_sender
        self.data = self.sheets_handler.load_data()
        self.columns_to_display_base = [
            'Status',
            'Nome completo (sem abreviações):',
            'Curso:',
            'Orientador',
            'Possui bolsa?',
            'Motivo da solicitação',
            'Local de realização do evento'
        ]
        self.detail_columns_to_display = ALL_COLUMNS_detail.copy()
        self.columns_to_display = self.columns_to_display_base.copy()
        self.main_frame = None
        self.details_frame = None
        self.detail_widgets = {}
        self.current_row_data = None
        self.selected_button = None  # Para rastrear o botão selecionado
        self.setup_ui()
        self.update_table()

    def setup_ui(self):
        self.root.title("Aplicativo de Visualização de Dados")
        self.root.geometry("1000x700")

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)

        button_frame = tk.Frame(self.main_frame, width=200, bg="lightgrey")
        button_frame.pack(side="left", fill="y")

        title_label = tk.Label(button_frame, text="Lista de Status", font=("Helvetica", 14, "bold"), bg="lightgrey")
        title_label.pack(pady=20, padx=10)

        # Botões para visualizar pendências, pagos, dados vazios e todos os dados
        self.empty_status_button = tk.Button(button_frame, text="Status Vazio", command=lambda: self.select_view("Vazio"))
        self.empty_status_button.pack(pady=10, padx=10, fill="x")

        self.pending_button = tk.Button(button_frame, text="Pendências", command=lambda: self.select_view("Pendências"))
        self.pending_button.pack(pady=10, padx=10, fill="x")

        self.paid_button = tk.Button(button_frame, text="Pago", command=lambda: self.select_view("Pago"))
        self.paid_button.pack(pady=10, padx=10, fill="x")

        self.view_all_button = tk.Button(button_frame, text="Visualização Completa", command=lambda: self.select_view("Todos"))
        self.view_all_button.pack(pady=10, padx=10, fill="x")

        bottom_frame = tk.Frame(self.root, bg="lightgrey")
        bottom_frame.pack(side="bottom", fill="x")

        credits_label = tk.Label(bottom_frame, text="Desenvolvido por: Vitor Akio & Leonardo Macedo",
                                 font=("Helvetica", 10), bg="lightgrey")
        credits_label.pack(side="right", padx=10, pady=10)

        table_frame = tk.Frame(self.main_frame)
        table_frame.pack(side="right", fill="both", expand=True)

        self.table_title_label = tk.Label(table_frame, text="Controle de Orçamento IG - PPG UNICAMP",
                                          font=("Helvetica", 16, "bold"))
        self.table_title_label.pack(pady=10)

        # Inicializar a Treeview sem definir as colunas aqui
        self.tree = ttk.Treeview(table_frame, show="headings")
        self.tree.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self.on_treeview_click)

        # Botão "Voltar"
        # Aumentar o tamanho da fonte e ajustar largura e altura
        self.back_button = tk.Button(
            self.root,
            text="Voltar",
            command=self.back_to_main_view,
            bg="darkgrey",
            fg="white",
            font=("Helvetica", 16, "bold"),  # Aumentar o tamanho da fonte
            width=20,  # Aumentar a largura do botão
            height=2   # Aumentar a altura do botão
        )

    def select_view(self, view_name):
        # Atualiza a tabela com o filtro selecionado
        if view_name == "Vazio":
            status_filter = "Vazio"
        elif view_name == "Pendências":
            status_filter = "Pendências"
        elif view_name == "Pago":
            status_filter = "Pago"
        else:
            status_filter = None  # Visualização Completa

        self.update_table(status_filter=status_filter)
        self.update_selected_button(view_name)

    def update_selected_button(self, view_name):
        # Resetar a cor do botão previamente selecionado
        if self.selected_button:
            self.selected_button.config(bg="SystemButtonFace")

        # Atualizar o botão selecionado e mudar sua cor
        if view_name == "Vazio":
            self.selected_button = self.empty_status_button
        elif view_name == "Pendências":
            self.selected_button = self.pending_button
        elif view_name == "Pago":
            self.selected_button = self.paid_button
        else:
            self.selected_button = self.view_all_button

        self.selected_button.config(bg="lightblue")

    def back_to_main_view(self):
        # Oculta a visualização detalhada e exibe a principal
        self.details_frame.pack_forget()
        self.main_frame.pack(fill="both", expand=True)
        self.back_button.pack_forget()
        self.update_table()

    def update_table(self, status_filter=None):
        # Certificar-se de que a main_frame está visível
        self.main_frame.pack(fill="both", expand=True)
        if self.details_frame:
            self.details_frame.pack_forget()
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Atualizar a lista de colunas com base no filtro
        self.columns_to_display = self.columns_to_display_base.copy()
        if status_filter == "Vazio":
            if 'Status' in self.columns_to_display:
                self.columns_to_display.remove('Status')
        elif 'Status' not in self.columns_to_display:
            self.columns_to_display.insert(0, 'Status')  # Reinsere 'Status' na primeira posição

        # Atualizar as colunas da Treeview
        self.tree["columns"] = self.columns_to_display
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")

        # Carregar os dados completos
        self.data = self.sheets_handler.load_data()

        # Aplicar o filtro de status no DataFrame completo
        if status_filter == "Pendências":
            data_filtered = self.data[self.data['Status'] != 'Pago']
        elif status_filter == "Vazio":
            data_filtered = self.data[self.data['Status'] == '']
        elif status_filter == "Pago":
            data_filtered = self.data[self.data['Status'] == 'Pago']
        else:
            data_filtered = self.data.copy()  # Sem filtro específico

        # Selecionar as colunas a serem exibidas
        try:
            data_filtered = data_filtered[self.columns_to_display]
        except KeyError as e:
            messagebox.showerror("Erro", f"Coluna não encontrada: {e}")
            return

        # Inserir os dados na Treeview
        for idx, row in data_filtered.iterrows():
            self.tree.insert("", "end", iid=idx, values=row.tolist())

        self.back_button.pack_forget()  # Esconde o botão de voltar quando na visualização principal
        self.table_title_label.config(text="Controle de Orçamento IG - PPG UNICAMP")

    def on_treeview_click(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            row_index = int(selected_item[0])
            self.current_row_data = self.sheets_handler.load_data().loc[row_index]
            self.show_details_in_place(self.current_row_data)

    def show_details_in_place(self, row_data):
        # Ocultar a visualização principal
        self.main_frame.pack_forget()

        # Mostrar os detalhes do item selecionado em abas
        self.details_frame = tk.Frame(self.root)
        self.details_frame.pack(fill="both", expand=True)

        self.detail_widgets = {}

        # Criar um Notebook (interface de abas)
        notebook = ttk.Notebook(self.details_frame)
        notebook.pack(fill='both', expand=True)

        # Definir estilos
        label_style = ttk.Style()
        label_style.configure("Bold.TLabel", font=("Helvetica", 12, "bold"))

        value_style = ttk.Style()
        value_style.configure("Regular.TLabel", font=("Helvetica", 12))

        # Agrupar campos em seções
        sections = {
            "Informações Pessoais": [
                'Nome completo (sem abreviações):',
                'Endereço de e-mail',
                'Telefone de contato:',
                'CPF:',
                'RG/RNE:',
                'Endereço completo (logradouro, número, bairro, cidade e estado)'
            ],
            "Informações Acadêmicas": [
                'Ano de ingresso o PPG:',
                'Curso:',
                'Orientador',
                'Possui bolsa?',
                'Qual a agência de fomento?',
                'Título do projeto do qual participa:',
            ],
            "Detalhes da Solicitação": [
                'Motivo da solicitação',
                'Nome do evento ou, se atividade de campo, motivos da realização\n* caso não se trate de evento ou viagem de campo, preencher N/A',
                'Local de realização do evento',
                'Período de realização da atividade. Indique as datas (dd/mm/aaaa)',
                'Descrever detalhadamente os itens a serem financiados. Por ex: inscrição em evento, diárias (para transporte, hospedagem e alimentação), passagem aérea, pagamento de análises e traduções, etc.\n',
            ],
            "Informações Financeiras": [
                'Valor',
                'Dados bancários (banco, agência e conta) ',
            ],
        }

        for section_name, fields in sections.items():
            # Criar um novo frame para a aba
            tab_frame = ttk.Frame(notebook)
            notebook.add(tab_frame, text=section_name)

            row_idx = 0
            for col in fields:
                if col in row_data:
                    label = ttk.Label(tab_frame, text=f"{col}:", style="Bold.TLabel")
                    label.grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
                    value = ttk.Label(tab_frame, text=row_data[col], style="Regular.TLabel", wraplength=600, justify="left")
                    value.grid(row=row_idx, column=1, sticky='w', padx=10, pady=5)
                    self.detail_widgets[col] = {'label': label, 'value': value, 'tab_frame': tab_frame}
                    row_idx += 1

        # Botões no final
        button_frame = ttk.Frame(self.details_frame)
        button_frame.pack(side='bottom', pady=10)

        # Adicionar botões específicos para a visualização 'Status Vazio'
        if row_data['Status'] == '':
            # Entry para "Valor" na aba "Informações Financeiras"
            financial_tab = None
            for child in notebook.winfo_children():
                if notebook.tab(child, "text") == "Informações Financeiras":
                    financial_tab = child
                    break

            if financial_tab:
                row_idx = len(sections['Informações Financeiras'])
                value_label = ttk.Label(financial_tab, text="Valor (R$):", style="Bold.TLabel")
                value_label.grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
                value_entry = ttk.Entry(financial_tab, width=50)
                value_entry.grid(row=row_idx, column=1, sticky='w', padx=10, pady=5)

                # Salvar o widget value_entry para uso posterior
                self.value_entry = value_entry

                def autorizar_auxilio():
                    new_status = 'Autorizado'
                    new_value = self.value_entry.get()
                    self.sheets_handler.update_status(row_data['Carimbo de data/hora'], new_status)
                    self.sheets_handler.update_value(row_data['Carimbo de data/hora'], new_value)
                    self.ask_send_email(row_data, new_status, new_value)
                    self.update_table()
                    self.back_to_main_view()

                def negar_auxilio():
                    new_status = 'Negado'
                    self.sheets_handler.update_status(row_data['Carimbo de data/hora'], new_status)
                    self.ask_send_email(row_data, new_status)
                    self.update_table()
                    self.back_to_main_view()

                # Botões de ação
                autorizar_button = ttk.Button(button_frame, text="Autorizar Auxílio", command=autorizar_auxilio)
                autorizar_button.pack(side="left", padx=10)

                negar_button = ttk.Button(button_frame, text="Negar Auxílio", command=negar_auxilio)
                negar_button.pack(side="left", padx=10)

        # Exibir o botão "Voltar" centralizado e com tamanho aumentado
        self.back_button.pack(side="bottom", pady=20)
        self.back_button.place(relx=0.5, rely=1.0, anchor='s')

    def ask_send_email(self, row_data, new_status, new_value=None):
        send_email = messagebox.askyesno("Enviar E-mail", "Deseja enviar um e-mail notificando a alteração de status?")
        if send_email:
            email_window = tk.Toplevel(self.root)
            email_window.title("Enviar E-mail")

            recipient_label = tk.Label(email_window, text="Destinatário:")
            recipient_label.pack(anchor="w", padx=10, pady=5)
            recipient_email = row_data['Endereço de e-mail']
            recipient_entry = tk.Entry(email_window, width=50)
            recipient_entry.insert(0, recipient_email)
            recipient_entry.pack(anchor="w", padx=10, pady=5)

            email_body_label = tk.Label(email_window, text="Corpo do E-mail:")
            email_body_label.pack(anchor="w", padx=10, pady=5)
            email_body_text = Text(email_window, width=60, height=15)
            email_body = f"Olá {row_data['Nome completo (sem abreviações):']},\n\nSeu status foi alterado para: {new_status}."
            if new_value:
                email_body += f"\nValor do auxílio: R$ {new_value}."
            email_body += f"\nCurso: {row_data['Curso:']}.\nOrientador: {row_data['Orientador']}.\n\nAtt,\nEquipe de Suporte"
            email_body_text.insert(tk.END, email_body)
            email_body_text.pack(anchor="w", padx=10, pady=5)

            def send_email():
                recipient = recipient_entry.get()
                subject = "Atualização de Status"
                body = email_body_text.get("1.0", tk.END)
                self.email_sender.send_email(recipient, subject, body)
                email_window.destroy()

            send_button = tk.Button(email_window, text="Enviar E-mail", command=send_email)
            send_button.pack(pady=10)

# Inicializar aplicação
if __name__ == "__main__":
    credentials_file = "credentials.json"
    sheet_url = "https://docs.google.com/spreadsheets/d/1sNwhkq0nCuTMRhs2HmahV88uIn9KiXY1ex0vlOwC0O8/edit?usp=sharing"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = "financas.ig.nubia@gmail.com"

    sheets_handler = GoogleSheetsHandler(credentials_file, sheet_url)
    email_sender = EmailSender(smtp_server, smtp_port, sender_email)

    root = tk.Tk()
    app = App(root, sheets_handler, email_sender)
    root.mainloop()
