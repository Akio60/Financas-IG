import os
import tkinter as tk
from tkinter import ttk, messagebox, StringVar, OptionMenu, Text, Entry
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PIL import Image, ImageTk

# Carregar variável de ambiente para a senha do email
os.environ['EMAIL_PASSWORD'] = 'senha'

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
        self.detail_columns_to_display = [
            'Nome completo (sem abreviações):',
            'Curso:',
            'Orientador',
            'Possui bolsa?',
            'Motivo da solicitação',
            'Endereço de e-mail',
            'Telefone',
            'Data de nascimento',
            'RG',
            'CPF',
            'Endereço completo',
            'Cidade',
            'Estado',
            'CEP',
            'Status',
            'Local de realização do evento',
            'Valor'
        ]
        self.columns_to_display = self.columns_to_display_base.copy()
        self.main_frame = None
        self.details_frame = None
        self.detail_widgets = {}
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
        empty_status_button = tk.Button(button_frame, text="Status Vazio", command=lambda: self.update_table(status_filter="Vazio"))
        empty_status_button.pack(pady=10, padx=10, fill="x")

        pending_button = tk.Button(button_frame, text="Pendências", command=lambda: self.update_table(status_filter="Pendências"))
        pending_button.pack(pady=10, padx=10, fill="x")

        paid_button = tk.Button(button_frame, text="Pago", command=lambda: self.update_table(status_filter="Pago"))
        paid_button.pack(pady=10, padx=10, fill="x")

        view_all_button = tk.Button(button_frame, text="Visualização Completa", command=lambda: self.update_table())
        view_all_button.pack(pady=10, padx=10, fill="x")

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
        self.back_button = tk.Button(self.root, text="Voltar", command=self.back_to_main_view, bg="darkgrey", fg="white")

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
        elif status_filter:
            data_filtered = self.data[self.data['Status'] == status_filter]
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
            row_data = self.data.loc[row_index]
            self.show_details_in_place(row_data)

    def show_details_in_place(self, row_data):
        # Ocultar a visualização principal
        self.main_frame.pack_forget()

        # Mostrar os detalhes do item selecionado em caixas de texto
        self.details_frame = tk.Frame(self.root)
        self.details_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.detail_widgets = {}  # Dicionário para armazenar os widgets de detalhe

        row_idx = 0
        for col in self.detail_columns_to_display:
            if col in row_data:
                label = tk.Label(self.details_frame, text=f"{col}:", font=("Helvetica", 12, "bold"))
                label.grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
                value_label = tk.Label(self.details_frame, text=row_data[col], font=("Helvetica", 12))
                value_label.grid(row=row_idx, column=1, padx=10, pady=5)
                # Armazenar os widgets para futura edição
                self.detail_widgets[col] = {'label': label, 'value': value_label}
                row_idx += 1

        # Botão "Editar"
        edit_button = tk.Button(self.details_frame, text="Editar", command=self.enable_editing)
        edit_button.grid(row=row_idx, column=0, padx=10, pady=10)
        row_idx += 1

        # Adicionar botões específicos para a visualização 'Status Vazio'
        if row_data['Status'] == '':
            # Caixa de texto para inserir valor
            value_label = tk.Label(self.details_frame, text="Valor (R$):", font=("Helvetica", 12, "bold"))
            value_label.grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
            value_entry = tk.Entry(self.details_frame, width=50)
            value_entry.grid(row=row_idx, column=1, padx=10, pady=5)
            row_idx += 1

            def autorizar_auxilio():
                new_status = 'Autorizado'
                new_value = value_entry.get()
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

            autorizar_button = tk.Button(self.details_frame, text="Autorizar Auxílio", command=autorizar_auxilio, bg="green", fg="white")
            autorizar_button.grid(row=row_idx, column=0, padx=10, pady=10)

            negar_button = tk.Button(self.details_frame, text="Negar Auxílio", command=negar_auxilio, bg="red", fg="white")
            negar_button.grid(row=row_idx, column=1, padx=10, pady=10)
            row_idx += 1

        # Exibir o botão "Voltar"
        self.back_button.pack(side="bottom", anchor="w", padx=10, pady=10)

    def enable_editing(self):
        # Transformar os Labels em Entries para permitir edição
        for col, widgets in self.detail_widgets.items():
            value = widgets['value'].cget('text')
            widgets['value'].destroy()  # Remove o Label
            entry = tk.Entry(self.details_frame, width=50)
            entry.insert(0, value)
            entry.grid(row=widgets['label'].grid_info()['row'], column=1, padx=10, pady=5)
            self.detail_widgets[col]['value'] = entry  # Atualiza para o Entry

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
