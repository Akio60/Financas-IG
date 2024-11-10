import os
import tkinter as tk
from tkinter import ttk, messagebox, Text
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PIL import Image, ImageTk  # Importar Pillow para manipulação de imagens

# Carregar variável de ambiente para a senha do email
os.environ['EMAIL_PASSWORD'] = 'upxu mpbq mbce mdti'  # Substitua 'senha' pela sua senha real ou configure a variável de ambiente

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

# Classe para manipular o Google Sheets
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

# Classe para enviar e-mails
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

# Classe principal da aplicação
class App:
    def __init__(self, root, sheets_handler, email_sender):
        self.root = root
        self.sheets_handler = sheets_handler
        self.email_sender = email_sender
        self.data = self.sheets_handler.load_data()
        self.columns_to_display_base = [
            'Carimbo de data/hora',
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
        self.detail_widgets = {}
        self.current_row_data = None
        self.selected_button = None  # Para rastrear o botão selecionado
        self.details_frame = None    # Inicializar self.details_frame como None
        self.current_view = None     # Para rastrear a visualização atual
        self.setup_ui()
        # Remover a chamada para self.update_table() para não exibir a tabela ao iniciar
        # self.update_table()

    def setup_ui(self):
        self.root.title("Aplicativo de Visualização de Dados")
        self.root.geometry("1000x700")

        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True)

        # Frame esquerdo (painel de botões)
        self.left_frame = tk.Frame(self.main_frame, width=200, bg="lightgrey")
        self.left_frame.pack(side="left", fill="y")

        title_label = tk.Label(self.left_frame, text="Lista de Status", font=("Helvetica", 14, "bold"), bg="lightgrey")
        title_label.pack(pady=20, padx=10)

        # Botões para visualizar pendências, pagos, dados vazios e todos os dados
        self.empty_status_button = tk.Button(self.left_frame, text="Aguardando aprovação", command=lambda: self.select_view("Aguardando aprovação"))
        self.empty_status_button.pack(pady=10, padx=10, fill="x")

        self.pending_button = tk.Button(self.left_frame, text="Pendências", command=lambda: self.select_view("Pendências"))
        self.pending_button.pack(pady=10, padx=10, fill="x")

        self.paid_button = tk.Button(self.left_frame, text="Pago", command=lambda: self.select_view("Pago"))
        self.paid_button.pack(pady=10, padx=10, fill="x")

        self.ready_for_payment_button = tk.Button(self.left_frame, text="Pronto para pagamento", command=lambda: self.select_view("Pronto para pagamento"))
        self.ready_for_payment_button.pack(pady=10, padx=10, fill="x")

        self.view_all_button = tk.Button(self.left_frame, text="Visualização Completa", command=lambda: self.select_view("Todos"))
        self.view_all_button.pack(pady=10, padx=10, fill="x")

        # Botão Página Inicial
        self.home_button = tk.Button(self.left_frame, text="🏠 Página Inicial", command=self.go_to_home)
        self.home_button.pack(pady=10, padx=10, fill="x")

        bottom_frame = tk.Frame(self.root, bg="lightgrey")
        bottom_frame.pack(side="bottom", fill="x")

        credits_label = tk.Label(bottom_frame, text="Desenvolvido por: Vitor Akio & Leonardo Macedo",
                                 font=("Helvetica", 10), bg="lightgrey")
        credits_label.pack(side="right", padx=10, pady=10)

        # Frame direito (conteúdo principal)
        self.content_frame = tk.Frame(self.main_frame)
        self.content_frame.pack(side="left", fill="both", expand=True)

        # Criar o welcome_frame para a tela inicial
        self.welcome_frame = tk.Frame(self.content_frame)
        self.welcome_frame.pack(fill="both", expand=True)

        # Configurar a tela de boas-vindas
        self.setup_welcome_screen()

        # Frame da tabela dentro do content_frame (não empacotar agora)
        self.table_frame = tk.Frame(self.content_frame)
        # self.table_frame.pack(fill="both", expand=True)  # Não empacotar agora

        self.table_title_label = tk.Label(self.table_frame, text="Controle de Orçamento IG - PPG UNICAMP",
                                          font=("Helvetica", 16, "bold"))
        self.table_title_label.pack(pady=10)

        # Inicializar a Treeview
        self.tree = ttk.Treeview(self.table_frame, show="headings")
        self.tree.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self.on_treeview_click)

        # Botão "Voltar"
        self.back_button = tk.Button(
            self.content_frame,
            text="Voltar",
            command=self.back_to_main_view,
            bg="darkgrey",
            fg="white",
            font=("Helvetica", 16, "bold"),  # Aumentar o tamanho da fonte
            width=20,  # Aumentar a largura do botão
            height=2   # Aumentar a altura do botão
        )

    def setup_welcome_screen(self):
        # Carregar as imagens e redimensioná-las
        try:
            img_ig = Image.open('logo_ig.png')
            img_unicamp = Image.open('logo_unicamp.png')

            # Redimensionar as imagens para um tamanho adequado
            img_ig = img_ig.resize((100, 100), Image.LANCZOS)
            img_unicamp = img_unicamp.resize((100, 100), Image.LANCZOS)

            # Converter as imagens para PhotoImage
            logo_ig = ImageTk.PhotoImage(img_ig)
            logo_unicamp = ImageTk.PhotoImage(img_unicamp)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar as imagens: {e}")
            return

        # Criar frames para as logos e texto
        logos_frame = tk.Frame(self.welcome_frame)
        logos_frame.pack(pady=20)

        # Exibir as logos lado a lado
        ig_label = tk.Label(logos_frame, image=logo_ig)
        ig_label.image = logo_ig  # Manter referência
        ig_label.pack(side='left', padx=10)

        unicamp_label = tk.Label(logos_frame, image=logo_unicamp)
        unicamp_label.image = logo_unicamp  # Manter referência
        unicamp_label.pack(side='left', padx=10)

        # Texto resumindo o programa
        summary_text = (
            "Este aplicativo permite a visualização e gerenciamento das solicitações de auxílio financeiro "
            "do Programa de Pós-Graduação do IG - UNICAMP. Utilize os botões ao lado para filtrar e visualizar "
            "as solicitações."
        )

        summary_label = tk.Label(self.welcome_frame, text=summary_text, font=("Helvetica", 12), wraplength=600, justify='center')
        summary_label.pack(pady=20)

    def select_view(self, view_name):
        self.current_view = view_name  # Adicionado para rastrear a visualização atual
        # Ocultar o welcome_frame se estiver visível
        if self.welcome_frame:
            self.welcome_frame.pack_forget()

        # Certificar-se de que o table_frame está visível
        if not self.table_frame.winfo_ismapped():
            self.table_frame.pack(fill="both", expand=True)

        # Atualiza a tabela com o filtro selecionado
        if view_name == "Aguardando aprovação":
            status_filter = "Vazio"
        elif view_name == "Pendências":
            status_filter = "Pendências"
        elif view_name == "Pago":
            status_filter = "Pago"
        elif view_name == "Pronto para pagamento":
            status_filter = "Pronto para pagamento"
        else:
            status_filter = None  # Visualização Completa

        self.update_table(status_filter=status_filter)
        self.update_selected_button(view_name)

    def update_selected_button(self, view_name):
        # Resetar a cor do botão previamente selecionado
        if self.selected_button:
            self.selected_button.config(bg="SystemButtonFace")

        # Atualizar o botão selecionado e mudar sua cor
        if view_name == "Aguardando aprovação":
            self.selected_button = self.empty_status_button
        elif view_name == "Pendências":
            self.selected_button = self.pending_button
        elif view_name == "Pago":
            self.selected_button = self.paid_button
        elif view_name == "Pronto para pagamento":
            self.selected_button = self.ready_for_payment_button
        elif view_name == "Todos":
            self.selected_button = self.view_all_button
        else:
            self.selected_button = None

        if self.selected_button:
            self.selected_button.config(bg="lightblue")

    def go_to_home(self):
        # Ocultar outros frames
        if self.table_frame.winfo_ismapped():
            self.table_frame.pack_forget()
        if self.details_frame and self.details_frame.winfo_ismapped():
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None
        # Mostrar o welcome_frame
        self.welcome_frame.pack(fill="both", expand=True)
        # Resetar o botão selecionado
        self.update_selected_button(None)

    def back_to_main_view(self):
        # Oculta a visualização detalhada e exibe a principal
        if self.details_frame:
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None

        # Esconder o botão "Voltar"
        self.back_button.pack_forget()

        # Mostrar o frame da tabela
        self.table_frame.pack(fill="both", expand=True)

    def update_table(self, status_filter=None):
        # Certificar-se de que o frame da tabela está visível
        if self.details_frame:
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None
        self.table_frame.pack(fill="both", expand=True)

        for item in self.tree.get_children():
            self.tree.delete(item)

        # Atualizar a lista de colunas com base no filtro
        self.columns_to_display = self.columns_to_display_base.copy()
        if status_filter == "Vazio":
            if 'Status' in self.columns_to_display:
                self.columns_to_display.remove('Status')
        elif 'Status' not in self.columns_to_display:
            self.columns_to_display.insert(1, 'Status')  # Reinsere 'Status' na segunda posição

        # Atualizar as colunas da Treeview
        self.tree["columns"] = self.columns_to_display
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")

        # Carregar os dados completos
        self.data = self.sheets_handler.load_data()

        # Converter 'Carimbo de data/hora' para datetime e formatar
        self.data['Carimbo de data/hora'] = pd.to_datetime(self.data['Carimbo de data/hora'], errors='coerce')
        self.data['Carimbo de data/hora'] = self.data['Carimbo de data/hora'].dt.strftime('%d/%m/%Y')

        # Aplicar o filtro de status no DataFrame completo
        if status_filter == "Pendências":
            data_filtered = self.data[(self.data['Status'] != 'Pago') & (self.data['Status'] != '')]
        elif status_filter == "Vazio":
            data_filtered = self.data[self.data['Status'] == '']
        elif status_filter == "Pago":
            data_filtered = self.data[self.data['Status'] == 'Pago']
        elif status_filter == "Pronto para pagamento":
            data_filtered = self.data[self.data['Status'] == 'Pronto para pagamento']
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

        # Esconde o botão de voltar quando na visualização principal
        self.back_button.pack_forget()
        self.table_title_label.config(text="Controle de Orçamento IG - PPG UNICAMP")

    def on_treeview_click(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            row_index = int(selected_item[0])
            self.current_row_data = self.sheets_handler.load_data().loc[row_index]
            self.show_details_in_place(self.current_row_data)

    def show_details_in_place(self, row_data):
        # Ocultar o frame da tabela
        self.table_frame.pack_forget()

        # Mostrar os detalhes do item selecionado em abas
        self.details_frame = tk.Frame(self.content_frame)
        self.details_frame.pack(fill="both", expand=True)

        # Adicionar o título na sessão de detalhes
        self.details_title_label = tk.Label(self.details_frame, text="Controle de Orçamento IG - PPG UNICAMP",
                                            font=("Helvetica", 16, "bold"))
        self.details_title_label.pack(pady=10)

        self.detail_widgets = {}

        # Definir estilos
        label_style = ttk.Style()
        label_style.configure("Bold.TLabel", font=("Helvetica", 12, "bold"))

        value_style = ttk.Style()
        value_style.configure("Regular.TLabel", font=("Helvetica", 12))

        # Estilo para as abas do Notebook com tamanho aumentado
        notebook_style = ttk.Style()
        notebook_style.theme_use('default')

        # Copiar o layout existente do TNotebook para o novo estilo
        notebook_style.layout("CustomNotebook.TNotebook", notebook_style.layout("TNotebook"))

        # Configurar o novo estilo das abas
        notebook_style.configure("CustomNotebook.TNotebook.Tab", font=("Helvetica", 13))

        # Criar um Notebook (interface de abas) com estilo personalizado
        notebook = ttk.Notebook(self.details_frame, style="CustomNotebook.TNotebook")
        notebook.pack(fill='both', expand=True)

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

            # Adicionar campos e botões específicos na aba "Informações Financeiras" se o status estiver vazio
            if section_name == "Informações Financeiras" and row_data['Status'] == '':
                # Entry para "Valor (R$)"
                value_label = ttk.Label(tab_frame, text="Valor (R$):", style="Bold.TLabel")
                value_label.grid(row=row_idx, column=0, sticky='w', padx=10, pady=5)
                value_entry = ttk.Entry(tab_frame, width=50)
                value_entry.grid(row=row_idx, column=1, sticky='w', padx=10, pady=5)
                self.value_entry = value_entry

                row_idx += 1  # Próxima linha para os botões

                # Funções para os botões
                def autorizar_auxilio():
                    new_value = self.value_entry.get()
                    if not new_value.strip():
                        messagebox.showwarning("Aviso", "Por favor, insira um valor antes de autorizar o auxílio.")
                        return
                    new_status = 'Aguardando documentação'
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

                # Criar botões usando tk.Button
                autorizar_button = tk.Button(tab_frame, text="Autorizar Auxílio", command=autorizar_auxilio,
                                             bg="green", fg="white", font=("Helvetica", 13))
                negar_button = tk.Button(tab_frame, text="Negar Auxílio", command=negar_auxilio,
                                         bg="red", fg="white", font=("Helvetica", 13))

                # Posicionar os botões abaixo do campo "Valor (R$)"
                autorizar_button.grid(row=row_idx, column=0, padx=10, pady=10, sticky='w')
                negar_button.grid(row=row_idx, column=1, padx=10, pady=10, sticky='w')

        # Adicionar a aba "Ações" se estiver na visualização de Pendências
        if self.current_view == "Pendências":
            # Criar a aba "Ações"
            actions_tab = ttk.Frame(notebook)
            notebook.add(actions_tab, text="Ações")

            # Botão "Requerir Documentos"
            def request_documents():
                subject = "Requisição de Documentos"
                body = f"Olá {row_data['Nome completo (sem abreviações):']},\n\n" \
                       f"Precisamos que você envie os documentos X, Y e Z para prosseguirmos com sua solicitação.\n\n" \
                       f"Atenciosamente,\nEquipe Financeira"
                self.send_custom_email(row_data['Endereço de e-mail'], subject, body)

            request_button = tk.Button(actions_tab, text="Requerir Documentos", command=request_documents, bg="orange", fg="white", font=("Helvetica", 13))
            request_button.pack(pady=10)

            # Botão "Autorizar Pagamento"
            def authorize_payment():
                new_status = 'Pronto para pagamento'
                self.sheets_handler.update_status(row_data['Carimbo de data/hora'], new_status)
                subject = "Pagamento Autorizado"
                body = f"Olá {row_data['Nome completo (sem abreviações):']},\n\n" \
                       f"Seu pagamento foi autorizado e está pronto para ser processado.\n\n" \
                       f"Atenciosamente,\nEquipe Financeira"
                self.send_custom_email(row_data['Endereço de e-mail'], subject, body)
                # Atualizar tabela e voltar à visualização principal
                self.update_table()
                self.back_to_main_view()

            authorize_button = tk.Button(actions_tab, text="Autorizar Pagamento", command=authorize_payment, bg="green", fg="white", font=("Helvetica", 13))
            authorize_button.pack(pady=10)

        # Exibir o botão "Voltar" no final da sessão de detalhes
        self.back_button.pack(side='bottom', pady=20)

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

    def send_custom_email(self, recipient, subject, body):
        email_window = tk.Toplevel(self.root)
        email_window.title("Enviar E-mail")

        recipient_label = tk.Label(email_window, text="Destinatário:")
        recipient_label.pack(anchor="w", padx=10, pady=5)
        recipient_entry = tk.Entry(email_window, width=50)
        recipient_entry.insert(0, recipient)
        recipient_entry.pack(anchor="w", padx=10, pady=5)

        email_body_label = tk.Label(email_window, text="Corpo do E-mail:")
        email_body_label.pack(anchor="w", padx=10, pady=5)
        email_body_text = Text(email_window, width=60, height=15)
        email_body_text.insert(tk.END, body)
        email_body_text.pack(anchor="w", padx=10, pady=5)

        def send_email():
            recipient_addr = recipient_entry.get()
            email_body = email_body_text.get("1.0", tk.END)
            self.email_sender.send_email(recipient_addr, subject, email_body)
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
