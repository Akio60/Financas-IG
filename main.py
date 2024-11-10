import tkinter as tk
from tkinter import ttk, messagebox, StringVar, OptionMenu, Text
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PIL import Image, ImageTk

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
#    'Carimbo de data/hora',  # Certifique-se de incluir o carimbo de data e hora
    'Status',
    'Nome completo (sem abreviações):',
    'Curso:',
    'Orientador',
    'Possui bolsa?',
    'Motivo da solicitação',
    'Local de realização do evento'
]

# Lista de colunas a serem exibidas na visualização dos detalhes
detail_columns_to_display = [
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
    'Local de realização do evento'
]

# Se a lista estiver vazia, exibir todas as colunas
if not columns_to_display:
    columns_to_display = all_columns

# Interface gráfica
class App:
    def __init__(self, root):
        self.root = root
        root.title("Aplicativo de Visualização de Dados")
        root.geometry("1000x700")

        # Frame principal para dividir a tela
        main_frame = tk.Frame(root)
        main_frame.pack(fill="both", expand=True)

        # Frame para os botões à esquerda
        button_frame = tk.Frame(main_frame, width=200, bg="lightgrey")
        button_frame.pack(side="left", fill="y")

        # Título do aplicativo acima dos botões
        title_label = tk.Label(button_frame, text="Lista de Status", font=("Helvetica", 14, "bold"), bg="lightgrey")
        title_label.pack(pady=20, padx=10)

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

        # Frame para imagem e créditos na parte inferior
        bottom_frame = tk.Frame(root, bg="lightgrey")
        bottom_frame.pack(side="bottom", fill="x")

        # Logo na parte inferior
        try:
            logo_image_bottom = Image.open("logo_unicamp.png")
            logo_image_bottom = logo_image_bottom.resize((50, 50), Image.LANCZOS)
            logo_photo_bottom = ImageTk.PhotoImage(logo_image_bottom)
            logo_label_bottom = tk.Label(bottom_frame, image=logo_photo_bottom, bg="lightgrey")
            logo_label_bottom.image = logo_photo_bottom
            logo_label_bottom.pack(side="left", padx=10, pady=10)
        except Exception as e:
            logo_label_bottom = tk.Label(bottom_frame, text="Logo não encontrada", font=("Helvetica", 10, "italic"), bg="lightgrey")
            logo_label_bottom.pack(side="left", padx=10, pady=10)

        # Texto de créditos na parte inferior
        credits_label = tk.Label(bottom_frame, text="Desenvolvido por: Vitor Akio & Leonardo Macedo", font=("Helvetica", 10), bg="lightgrey")
        credits_label.pack(side="right", padx=10, pady=10)

        # Frame para a tabela à direita
        table_frame = tk.Frame(main_frame)
        table_frame.pack(side="right", fill="both", expand=True)

        # Logo acima da tabela
        logo_frame = tk.Frame(table_frame)
        logo_frame.pack(side="top", pady=10)
        try:
            logo_image = Image.open("logo_ig.png")
            logo_image = logo_image.resize((100, 100), Image.LANCZOS)
            logo_photo = ImageTk.PhotoImage(logo_image)
            logo_label = tk.Label(logo_frame, image=logo_photo)
            logo_label.image = logo_photo
            logo_label.pack()
        except Exception as e:
            logo_label = tk.Label(logo_frame, text="Logo não encontrada", font=("Helvetica", 10, "italic"))
            logo_label.pack()

        # Título da tabela
        table_title_label = tk.Label(table_frame, text="Controle de Orçamento IG - PPG UNICAMP", font=("Helvetica", 16, "bold"))
        table_title_label.pack(pady=10)

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
                    for idx, cell_value in enumerate(cell_list[1:], start=2):
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
        def send_email_window():
            email_window = tk.Toplevel(details_window)
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
            email_body = f"Olá {row_data['Nome completo (sem abreviações):']},\n\nSeu status atual é: {row_data['Status']}.\nCurso: {row_data['Curso:']}.\nOrientador: {row_data['Orientador']}.\n\nAtt,\nEquipe de Suporte"
            email_body_text.insert(tk.END, email_body)
            email_body_text.pack(anchor="w", padx=10, pady=5)

            def send_email():
                try:
                    # Configurações do servidor SMTP (exemplo com Gmail)
                    smtp_server = "smtp.gmail.com"
                    smtp_port = 587
                    sender_email = "financas.ig.nubia@gmail.com"
                    import os
                    sender_password = 'upxu mpbq mbce mdti'

                    recipient = recipient_entry.get()
                    subject = "Atualização de Status"
                    body = email_body_text.get("1.0", tk.END)

                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = recipient
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'plain'))

                    # Enviar o e-mail
                    server = smtplib.SMTP(smtp_server, smtp_port)
                    server.starttls()
                    server.login(sender_email, sender_password)
                    server.sendmail(sender_email, recipient, msg.as_string())
                    server.quit()

                    messagebox.showinfo("Sucesso", "E-mail enviado com sucesso!")
                    email_window.destroy()
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao enviar o e-mail: {e}")

            send_button = tk.Button(email_window, text="Enviar E-mail", command=send_email)
            send_button.pack(pady=10)

        email_button = tk.Button(details_window, text="Enviar E-mail", command=send_email_window)
        email_button.pack(pady=10)

# Inicializar aplicação
# Certifique-se de definir a variável de ambiente 'EMAIL_PASSWORD' antes de executar o aplicativo
root = tk.Tk()
app = App(root)
root.mainloop()
