import os
import tkinter as tk
from tkinter import ttk, messagebox, Text
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from PIL import Image, ImageTk
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import json

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
                # Atualizar a coluna 'Ultima Atualizacao' com o timestamp atual
                current_timestamp = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                self.sheet.update_cell(row_number, self.column_indices['Ultima Atualizacao'], current_timestamp)
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
        self.detail_columns_to_display = ALL_COLUMNS_detail.copy()
        self.columns_to_display = []  # Será definido em cada visualização
        self.detail_widgets = {}
        self.current_row_data = None
        self.selected_button = None  # Para rastrear o botão selecionado
        self.details_frame = None    # Inicializar self.details_frame como None
        self.current_view = None     # Para rastrear a visualização atual
        self.treeview_data = None    # Para armazenar os dados exibidos na Treeview
        self.email_templates = {}
        self.load_email_templates()

        # Definir cores base
        self.bg_color = '#e0f0ff'        # Fundo azul claro
        self.button_bg_color = '#add8e6'  # Botões azul
        self.frame_bg_color = '#d0e0ff'   # Frames azul claro

        # Definir cores para os status com melhor contraste
        self.status_colors = {
            'Aguardando documentação': '#B8860B',  # DarkGoldenrod
            'Autorizado': '#006400',               # DarkGreen
            'Negado': '#8B0000',                   # DarkRed
            'Pronto para pagamento': '#00008B',    # DarkBlue
            'Pago': '#4B0082',                     # Indigo
            'Cancelado': '#696969',                # DimGray
            # Adicione outros status se necessário
        }

        # Mapeamento de nomes de colunas para exibição
        self.column_display_names = {
            'Carimbo de data/hora_str': 'Data do requerimento',
            # Adicione outros mapeamentos se necessário
        }

        self.setup_ui()

    def load_email_templates(self):
        # Carregar os templates de e-mail de um arquivo JSON ou usar padrões
        try:
            with open('email_templates.json', 'r', encoding='utf-8') as f:
                self.email_templates = json.load(f)
        except FileNotFoundError:
            # Templates padrão
            self.email_templates = {
                'Trabalho de Campo': 'Prezado(a) {Nome},\n\nPor favor, envie os documentos necessários para o Trabalho de Campo.\n\nAtt,\nEquipe Financeira',
                'Participação em eventos': 'Prezado(a) {Nome},\n\nPor favor, envie os documentos necessários para a Participação em Eventos.\n\nAtt,\nEquipe Financeira',
                'Visita técnica': 'Prezado(a) {Nome},\n\nPor favor, envie os documentos necessários para a Visita Técnica.\n\nAtt,\nEquipe Financeira',
                'Outros': 'Prezado(a) {Nome},\n\nPor favor, envie os documentos necessários.\n\nAtt,\nEquipe Financeira',
                'Aprovação': 'Prezado(a) {Nome},\n\nSua solicitação foi aprovada.\n\nAtt,\nEquipe Financeira',
                'Pagamento': 'Prezado(a) {Nome},\n\nSeu pagamento foi efetuado com sucesso.\n\nAtt,\nEquipe Financeira',
            }

    def save_email_templates(self):
        # Salvar os templates de e-mail em um arquivo JSON
        with open('email_templates.json', 'w', encoding='utf-8') as f:
            json.dump(self.email_templates, f, ensure_ascii=False, indent=4)

    def setup_ui(self):
        self.root.title("Controle de Orçamento IG - PPG UNICAMP")
        self.root.geometry("1000x700")
        self.root.configure(bg=self.bg_color)

        self.main_frame = tk.Frame(self.root, bg=self.bg_color)
        self.main_frame.pack(fill="both", expand=True)

        # Frame esquerdo (painel de botões)
        self.left_frame = tk.Frame(self.main_frame, width=200, bg=self.frame_bg_color)
        self.left_frame.pack(side="left", fill="y")

        title_label = tk.Label(self.left_frame, text="Lista de Status", font=("Helvetica", 14, "bold"), bg=self.frame_bg_color)
        title_label.pack(pady=20, padx=10)

        # Botão Página Inicial (apenas símbolo de casa)
        self.home_button = tk.Button(self.left_frame, text="🏠", command=self.go_to_home, bg=self.button_bg_color)
        self.home_button.pack(pady=10, padx=10, fill="x")

        # Botões de visualização
        self.empty_status_button = tk.Button(self.left_frame, text="Solicitações recebidas", command=lambda: self.select_view("Aguardando aprovação"), bg=self.button_bg_color)
        self.empty_status_button.pack(pady=10, padx=10, fill="x")

        self.pending_button = tk.Button(self.left_frame, text="Solicitações em andamento", command=lambda: self.select_view("Pendências"), bg=self.button_bg_color)
        self.pending_button.pack(pady=10, padx=10, fill="x")

        self.ready_for_payment_button = tk.Button(self.left_frame, text="Solicitações pré efetuadas", command=lambda: self.select_view("Pronto para pagamento"), bg=self.button_bg_color)
        self.ready_for_payment_button.pack(pady=10, padx=10, fill="x")

        # Frame para os botões inferiores
        bottom_buttons_frame = tk.Frame(self.left_frame, bg=self.frame_bg_color)
        bottom_buttons_frame.pack(side='bottom', fill='x', pady=10)

        # Botão Configurações (apenas símbolo de engrenagem)
        self.settings_button = tk.Button(self.left_frame, text='⚙', command=self.open_settings, bg=self.button_bg_color)
        self.settings_button.pack(side='bottom', pady=10, padx=10, fill='x')

        # Barra de busca e botão 'Pesquisar' (movidos para baixo)
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(self.left_frame, textvariable=self.search_var, width=25)
        self.search_entry.pack(pady=5, padx=10)

        self.search_button = tk.Button(self.left_frame, text="Pesquisar", command=self.perform_search, bg=self.button_bg_color)
        self.search_button.pack(pady=5, padx=10, fill="x")
        search_label = tk.Label(bottom_buttons_frame, text="Pesquisar:", bg=self.frame_bg_color)

        # Botão para 'Configurações'
        self.settings_button = tk.Button(self.left_frame, text="Configurações", command=self.open_settings, bg=self.button_bg_color)
        self.settings_button.pack(pady=10, padx=10, fill="x")
        # Botão para 'Estatísticas' (movido para baixo)

        bottom_frame = tk.Frame(self.root, bg=self.bg_color)
        bottom_frame.pack(side="bottom", fill="x")

        credits_label = tk.Label(bottom_frame, text="Desenvolvido por: Vitor Akio & Leonardo Macedo",
                                 font=("Helvetica", 10), bg=self.bg_color)
        credits_label.pack(side="right", padx=10, pady=10)

        # Frame direito (conteúdo principal)
        self.content_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        self.content_frame.pack(side="left", fill="both", expand=True)

        # Criar o welcome_frame para a tela inicial
        self.welcome_frame = tk.Frame(self.content_frame, bg=self.bg_color)
        self.welcome_frame.pack(fill="both", expand=True)

        # Configurar a tela de boas-vindas
        self.setup_welcome_screen()

        # Frame da tabela dentro do content_frame (não empacotar agora)
        self.table_frame = tk.Frame(self.content_frame, bg=self.bg_color)

        self.table_title_label = tk.Label(self.table_frame, text="Controle de Orçamento IG - PPG UNICAMP",
                                          font=("Helvetica", 16, "bold"), bg=self.bg_color)
        self.table_title_label.pack(pady=10)

        # Inicializar a Treeview
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background=self.bg_color, foreground="black", rowheight=25, fieldbackground=self.bg_color)
        style.map('Treeview', background=[('selected', '#347083')])

        self.tree = ttk.Treeview(self.table_frame, show="headings", style="Treeview")
        self.tree.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.tree.bind("<Double-1>", self.on_treeview_click)

        # Definir tags para linhas ímpares e pares
        self.tree.tag_configure('oddrow', background='#f0f8ff')   # AliceBlue
        self.tree.tag_configure('evenrow', background='#ffffff')  # White

        # Botão "Voltar"
        self.back_button = tk.Button(
            self.content_frame,
            text="Voltar",
            command=self.back_to_main_view,
            bg=self.button_bg_color,
            fg="white",
            font=("Helvetica", 16, "bold"),
            width=20,
            height=2
        )

    def setup_welcome_screen(self):
        # Carregar as imagens e redimensioná-las
        try:
            img_ig = Image.open('logo_unicamp.png')
            img_unicamp = Image.open('logo_ig.png')

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
        logos_frame = tk.Frame(self.welcome_frame, bg=self.bg_color)
        logos_frame.pack(pady=20)

        # Exibir as logos lado a lado
        ig_label = tk.Label(logos_frame, image=logo_ig, bg=self.bg_color)
        ig_label.image = logo_ig  # Manter referência
        ig_label.pack(side='left', padx=10)

        unicamp_label = tk.Label(logos_frame, image=logo_unicamp, bg=self.bg_color)
        unicamp_label.image = logo_unicamp  # Manter referência
        unicamp_label.pack(side='left', padx=10)

        # Texto resumindo o programa
        summary_text = (
            "Este aplicativo permite a visualização e gerenciamento das solicitações de auxílio financeiro "
            "do Programa de Pós-Graduação do IG - UNICAMP. Utilize os botões ao lado para filtrar e visualizar "
            "as solicitações."
        )

        summary_label = tk.Label(self.welcome_frame, text=summary_text, font=("Helvetica", 12), wraplength=600, justify='center', bg=self.bg_color)
        summary_label.pack(pady=20)

    def select_view(self, view_name):
        self.current_view = view_name  # Adicionado para rastrear a visualização atual
        # Ocultar o welcome_frame se estiver visível
        if self.welcome_frame:
            self.welcome_frame.pack_forget()

        # Ocultar o frame de estatísticas se estiver visível
        if hasattr(self, 'statistics_frame') and self.statistics_frame.winfo_ismapped():
            self.statistics_frame.pack_forget()

        # Ocultar detalhes se estiverem visíveis
        if self.details_frame and self.details_frame.winfo_ismapped():
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None

        # Certificar-se de que o table_frame está visível
        if not self.table_frame.winfo_ismapped():
            self.table_frame.pack(fill="both", expand=True)

        self.update_table()
        self.update_selected_button(view_name)

    def perform_search(self):
        self.current_view = "Search"
        # Ocultar outros frames
        if self.welcome_frame.winfo_ismapped():
            self.welcome_frame.pack_forget()
        if hasattr(self, 'statistics_frame') and self.statistics_frame.winfo_ismapped():
            self.statistics_frame.pack_forget()
        if self.details_frame and self.details_frame.winfo_ismapped():
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None

        self.update_table()

    def update_selected_button(self, view_name):
        # Resetar a cor do botão previamente selecionado
        if self.selected_button:
            self.selected_button.config(bg=self.button_bg_color)

        # Atualizar o botão selecionado e mudar sua cor
        if view_name == "Aguardando aprovação":
            self.selected_button = self.empty_status_button
        elif view_name == "Pendências":
            self.selected_button = self.pending_button
        elif view_name == "Pronto para pagamento":
            self.selected_button = self.ready_for_payment_button
        elif view_name == "Todos":
            self.selected_button = self.view_all_button
        elif view_name == "Estatísticas":
            self.selected_button = self.statistics_button
        else:
            self.selected_button = None

        if self.selected_button:
            self.selected_button.config(bg="#87CEFA")  # Cor diferente para indicar seleção

    def go_to_home(self):
        # Ocultar outros frames
        if self.table_frame.winfo_ismapped():
            self.table_frame.pack_forget()
        if self.details_frame and self.details_frame.winfo_ismapped():
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None
        if hasattr(self, 'statistics_frame') and self.statistics_frame.winfo_ismapped():
            self.statistics_frame.pack_forget()
        # Mostrar o welcome_frame
        self.welcome_frame.pack(fill="both", expand=True)
        # Resetar o botão selecionado
        self.update_selected_button(None)
        # Limpar a barra de busca
        self.search_var.set('')

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

    def update_table(self):
        # Certificar-se de que o frame da tabela está visível
        if self.details_frame:
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None
        if hasattr(self, 'statistics_frame') and self.statistics_frame.winfo_ismapped():
            self.statistics_frame.pack_forget()
        self.table_frame.pack(fill="both", expand=True)

        for item in self.tree.get_children():
            self.tree.delete(item)

        # Definir as colunas a serem exibidas com base na visualização atual
        if self.current_view == "Aguardando aprovação":
            self.columns_to_display = [
                'Endereço de e-mail', 'Nome completo (sem abreviações):', 'Curso:', 'Orientador', 'Qual a agência de fomento?',
                'Título do projeto do qual participa:', 'Motivo da solicitação',
                'Local de realização do evento', 'Período de realização da atividade. Indique as datas (dd/mm/aaaa)',
                'Telefone de contato:'
            ]
        elif self.current_view == "Pendências":
            self.columns_to_display = [
                'Carimbo de data/hora_str', 'Status', 'Nome completo (sem abreviações):', 'Ultima Atualizacao', 'Valor', 'Curso:', 'Orientador', 'E-mail DAC:'
            ]
        elif self.current_view == "Pronto para pagamento":
            self.columns_to_display = [
                'Carimbo de data/hora_str', 'Nome completo (sem abreviações):', 'Ultima Atualizacao', 'Valor', 'Telefone de contato:', 'E-mail DAC:',
                'Endereço completo (logradouro, número, bairro, cidade e estado)', 'CPF:', 'RG/RNE:', 'Dados bancários (banco, agência e conta) '
            ]
        else:  # Visualização completa e pesquisa
            self.columns_to_display = [
                'Carimbo de data/hora_str', 'Nome completo (sem abreviações):', 'Ultima Atualizacao', 'Valor', 'Status'
            ]

        # Atualizar as colunas da Treeview
        self.tree["columns"] = self.columns_to_display
        for col in self.tree["columns"]:
            display_name = self.column_display_names.get(col, col)
            self.tree.heading(col, text=display_name, command=lambda _col=col: self.treeview_sort_column(self.tree, _col, False))
            self.tree.column(col, anchor="center", width=150)

        # Carregar os dados completos
        self.data = self.sheets_handler.load_data()

        # Converter 'Carimbo de data/hora' para datetime e formatar
        self.data['Carimbo de data/hora'] = pd.to_datetime(
            self.data['Carimbo de data/hora'],
            format='%d/%m/%Y %H:%M:%S',
            errors='coerce'
        )
        self.data['Carimbo de data/hora_str'] = self.data['Carimbo de data/hora'].dt.strftime('%d/%m/%Y')

        # Aplicar o filtro de status no DataFrame completo
        if self.current_view == "Search":
            data_filtered = self.data.copy()
        else:
            if self.current_view == "Pendências":
                data_filtered = self.data[self.data['Status'].isin(['Autorizado', 'Aguardando documentação'])]
            elif self.current_view == "Aguardando aprovação":
                data_filtered = self.data[self.data['Status'] == '']
            elif self.current_view == "Pronto para pagamento":
                data_filtered = self.data[self.data['Status'] == 'Pronto para pagamento']
            else:
                data_filtered = self.data.copy()  # Visualização completa ou outra

        # Aplicar filtro de busca se necessário
        search_term = self.search_var.get().lower()
        if search_term:
            data_filtered = data_filtered[
                data_filtered[self.columns_to_display].apply(
                    lambda row: row.astype(str).str.lower().str.contains(search_term).any(), axis=1
                )
            ]

        # Selecionar as colunas a serem exibidas
        try:
            data_filtered = data_filtered[self.columns_to_display + ['Carimbo de data/hora']]
        except KeyError as e:
            messagebox.showerror("Erro", f"Coluna não encontrada: {e}")
            return

        # Armazenar os dados exibidos na Treeview
        self.treeview_data = data_filtered.copy()

        # Inserir os dados na Treeview
        for idx, row in data_filtered.iterrows():
            values = row[self.columns_to_display].tolist()
            # Determinar a tag para cores alternadas
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            # Obter a cor do status
            status_value = row.get('Status', '')
            status_color = self.status_colors.get(status_value, '#000000')  # Cor padrão preta

            # Configurar uma tag específica para a cor do texto
            self.tree.tag_configure(f'status_tag_{idx}', foreground=status_color)
            # Inserir o item com a tag
            self.tree.insert("", "end", iid=str(idx), values=values, tags=(tag, f'status_tag_{idx}'))

        # Esconde o botão de voltar quando na visualização principal
        self.back_button.pack_forget()
        self.table_title_label.config(text="Controle de Orçamento IG - PPG UNICAMP")

    def treeview_sort_column(self, tv, col, reverse):
        # Obter os dados da coluna
        data_list = [(tv.set(k, col), k) for k in tv.get_children('')]
        # Tentar converter para datetime ou numérico
        try:
            data_list.sort(key=lambda t: datetime.strptime(t[0], '%d/%m/%Y'), reverse=reverse)
        except ValueError:
            try:
                data_list.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
            except ValueError:
                data_list.sort(reverse=reverse)
        # Reordenar os itens na Treeview
        for index, (val, k) in enumerate(data_list):
            tv.move(k, '', index)
        # Alternar o estado de reverse para a próxima ordenação
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

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
        self.details_frame = tk.Frame(self.content_frame, bg=self.bg_color)
        self.details_frame.pack(fill="both", expand=True)

        # Adicionar o título na sessão de detalhes
        self.details_title_label = tk.Label(self.details_frame, text="Controle de Orçamento IG - PPG UNICAMP",
                                            font=("Helvetica", 16, "bold"), bg=self.bg_color)
        self.details_title_label.pack(pady=10)

        self.detail_widgets = {}

        # Estilo para as abas do Notebook com tamanho aumentado
        notebook_style = ttk.Style()
        notebook_style.theme_use('default')

        notebook_style.configure("CustomNotebook.TNotebook", background=self.bg_color)
        notebook_style.configure("CustomNotebook.TNotebook.Tab", background=self.bg_color, font=("Helvetica", 13))

        # Definir estilos para frames e labels
        frame_style_name = 'Custom.TFrame'
        notebook_style.configure(frame_style_name, background=self.bg_color)

        notebook_style.configure("CustomBold.TLabel", font=("Helvetica", 12, "bold"), background=self.bg_color)
        notebook_style.configure("CustomRegular.TLabel", font=("Helvetica", 12), background=self.bg_color)

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
            tab_frame = ttk.Frame(notebook, style=frame_style_name)
            notebook.add(tab_frame, text=section_name)

            # Configurar as colunas
            tab_frame.columnconfigure(0, weight=1, minsize=200)
            tab_frame.columnconfigure(1, weight=3)

            row_idx = 0
            for col in fields:
                if col in row_data:
                    label = ttk.Label(tab_frame, text=f"{col}:", style="CustomBold.TLabel", wraplength=200, justify='left')
                    label.grid(row=row_idx, column=0, sticky='nw', padx=10, pady=5)
                    value = ttk.Label(tab_frame, text=row_data[col], style="CustomRegular.TLabel", wraplength=600, justify="left")
                    value.grid(row=row_idx, column=1, sticky='nw', padx=10, pady=5)
                    self.detail_widgets[col] = {'label': label, 'value': value, 'tab_frame': tab_frame}
                    row_idx += 1

            # Adicionar campos e botões específicos na aba "Informações Financeiras" se o status estiver vazio
            if section_name == "Informações Financeiras" and row_data['Status'] == '':
                # Entry para "Valor (R$)"
                value_label = ttk.Label(tab_frame, text="Valor (R$):", style="CustomBold.TLabel")
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
                    new_status = 'Autorizado'  # Alterado para 'Autorizado'
                    timestamp_str = row_data['Carimbo de data/hora'].strftime('%d/%m/%Y %H:%M:%S')
                    self.sheets_handler.update_status(timestamp_str, new_status)
                    self.sheets_handler.update_value(timestamp_str, new_value)
                    self.ask_send_email(row_data, new_status, new_value)
                    self.update_table()
                    self.back_to_main_view()

                def negar_auxilio():
                    confirm = messagebox.askyesno("Confirmação", "Tem certeza que deseja recusar/cancelar o auxílio?")
                    if confirm:
                        new_status = 'Cancelado'
                        timestamp_str = row_data['Carimbo de data/hora'].strftime('%d/%m/%Y %H:%M:%S')
                        self.sheets_handler.update_status(timestamp_str, new_status)
                        self.ask_send_email(row_data, new_status)
                        self.update_table()
                        self.back_to_main_view()

                # Criar botões usando tk.Button com largura fixa
                button_width = 25
                autorizar_button = tk.Button(tab_frame, text="Autorizar Auxílio", command=autorizar_auxilio,
                                             bg="green", fg="white", font=("Helvetica", 13), width=button_width)
                negar_button = tk.Button(tab_frame, text="Recusar/Cancelar Auxílio", command=negar_auxilio,
                                         bg="red", fg="white", font=("Helvetica", 13), width=button_width)

                # Posicionar os botões abaixo do campo "Valor (R$)"
                autorizar_button.grid(row=row_idx, column=0, padx=10, pady=10, sticky='w')
                negar_button.grid(row=row_idx, column=1, padx=10, pady=10, sticky='w')

        # Adicionar a aba "Ações" se estiver na visualização de Pendências ou Pronto para pagamento
        if self.current_view in ["Pendências", "Pronto para pagamento"]:
            # Criar a aba "Ações"
            actions_tab = ttk.Frame(notebook, style=frame_style_name)
            notebook.add(actions_tab, text="Ações")

            if self.current_view == "Pendências":
                # Botão "Requerir Documentos"
                def request_documents():
                    new_status = 'Aguardando documentação'  # Alterado para 'Aguardando documentação'
                    timestamp_str = row_data['Carimbo de data/hora'].strftime('%d/%m/%Y %H:%M:%S')
                    self.sheets_handler.update_status(timestamp_str, new_status)

                    motivo = row_data.get('Motivo da solicitação', 'Outros')
                    motivo = motivo.strip()

                    # Obter o template de e-mail correspondente
                    email_template = self.email_templates.get(motivo, self.email_templates['Outros'])
                    subject = "Requisição de Documentos"
                    body = email_template.format(Nome=row_data['Nome completo (sem abreviações):'])

                    self.send_custom_email(row_data['Endereço de e-mail'], subject, body)
                    self.update_table()
                    self.back_to_main_view()

                # Botão "Autorizar Pagamento"
                def authorize_payment():
                    new_status = 'Pronto para pagamento'
                    timestamp_str = row_data['Carimbo de data/hora'].strftime('%d/%m/%Y %H:%M:%S')
                    self.sheets_handler.update_status(timestamp_str, new_status)

                    # Obter o template de e-mail de aprovação
                    email_template = self.email_templates.get('Aprovação', 'Sua solicitação foi aprovada.')
                    subject = "Pagamento Autorizado"
                    body = email_template.format(Nome=row_data['Nome completo (sem abreviações):'])

                    self.send_custom_email(row_data['Endereço de e-mail'], subject, body)
                    # Atualizar tabela e voltar à visualização principal
                    self.update_table()
                    self.back_to_main_view()

                # Botão "Recusar/Cancelar Auxílio"
                def cancel_auxilio():
                    confirm = messagebox.askyesno("Confirmação", "Tem certeza que deseja recusar/cancelar o auxílio?")
                    if confirm:
                        new_status = 'Cancelado'
                        timestamp_str = row_data['Carimbo de data/hora'].strftime('%d/%m/%Y %H:%M:%S')
                        self.sheets_handler.update_status(timestamp_str, new_status)
                        subject = "Auxílio Cancelado"
                        body = f"Olá {row_data['Nome completo (sem abreviações):']},\n\n" \
                               f"Seu auxílio foi cancelado.\n\n" \
                               f"Atenciosamente,\nEquipe Financeira"
                        self.send_custom_email(row_data['Endereço de e-mail'], subject, body)
                        self.update_table()
                        self.back_to_main_view()

                # Criar botões com largura fixa
                button_width = 25
                request_button = tk.Button(actions_tab, text="Requerir Documentos", command=request_documents,
                                           bg="orange", fg="white", font=("Helvetica", 13), width=button_width)
                request_button.pack(pady=10)

                authorize_button = tk.Button(actions_tab, text="Autorizar Pagamento", command=authorize_payment,
                                             bg="green", fg="white", font=("Helvetica", 13), width=button_width)
                authorize_button.pack(pady=10)

                cancel_button = tk.Button(actions_tab, text="Recusar/Cancelar Auxílio", command=cancel_auxilio,
                                          bg="red", fg="white", font=("Helvetica", 13), width=button_width)
                cancel_button.pack(pady=10)

            elif self.current_view == "Pronto para pagamento":
                # Botão "Pagamento Efetuado"
                def payment_made():
                    new_status = 'Pago'
                    timestamp_str = row_data['Carimbo de data/hora'].strftime('%d/%m/%Y %H:%M:%S')
                    self.sheets_handler.update_status(timestamp_str, new_status)

                    # Obter o template de e-mail de pagamento
                    email_template = self.email_templates.get('Pagamento', 'Seu pagamento foi efetuado com sucesso.')
                    subject = "Pagamento Efetuado"
                    body = email_template.format(Nome=row_data['Nome completo (sem abreviações):'])

                    self.send_custom_email(row_data['Endereço de e-mail'], subject, body)
                    # Atualizar tabela e voltar à visualização principal
                    self.update_table()
                    self.back_to_main_view()

                # Botão "Recusar/Cancelar Auxílio"
                def cancel_auxilio():
                    confirm = messagebox.askyesno("Confirmação", "Tem certeza que deseja recusar/cancelar o auxílio?")
                    if confirm:
                        new_status = 'Cancelado'
                        timestamp_str = row_data['Carimbo de data/hora'].strftime('%d/%m/%Y %H:%M:%S')
                        self.sheets_handler.update_status(timestamp_str, new_status)
                        subject = "Auxílio Cancelado"
                        body = f"Olá {row_data['Nome completo (sem abreviações):']},\n\n" \
                               f"Seu auxílio foi cancelado.\n\n" \
                               f"Atenciosamente,\nEquipe Financeira"
                        self.send_custom_email(row_data['Endereço de e-mail'], subject, body)
                        self.update_table()
                        self.back_to_main_view()

                # Criar botões com largura fixa
                button_width = 25
                payment_button = tk.Button(actions_tab, text="Pagamento Efetuado", command=payment_made,
                                           bg="green", fg="white", font=("Helvetica", 13), width=button_width)
                payment_button.pack(pady=10)

                cancel_button = tk.Button(actions_tab, text="Recusar/Cancelar Auxílio", command=cancel_auxilio,
                                          bg="red", fg="white", font=("Helvetica", 13), width=button_width)
                cancel_button.pack(pady=10)

        # Adicionando a aba "Histórico de solicitações"
        history_tab = ttk.Frame(notebook, style=frame_style_name)
        notebook.add(history_tab, text="Histórico de solicitações")

        # Recuperar CPF do dado atual
        cpf = str(row_data.get('CPF:', '')).strip()

        # Carregar todos os dados
        all_data = self.sheets_handler.load_data()
        all_data['CPF:'] = all_data['CPF:'].astype(str).str.strip()

        # Filtrar dados pelo CPF
        history_data = all_data[all_data['CPF:'] == cpf]

        # Selecionar colunas a serem exibidas
        history_columns = ['Carimbo de data/hora', 'Ultima Atualizacao', 'Valor', 'Status']

        history_tree = ttk.Treeview(history_tab, columns=history_columns, show='headings', height=10)
        history_tree.pack(fill='both', expand=True)

        for col in history_columns:
            history_tree.heading(col, text=col, command=lambda _col=col: self.treeview_sort_column(history_tree, _col, False))
            history_tree.column(col, anchor='center', width=100)

        # Armazenar os dados para acesso posterior
        self.history_tree_data = history_data.copy()

        # Inserir dados na treeview com índices reais
        for idx, hist_row in history_data.iterrows():
            values = [hist_row[col] for col in history_columns]
            history_tree.insert("", "end", iid=str(idx), values=values)

        # Vincular evento de duplo clique
        history_tree.bind("<Double-1>", self.on_history_treeview_click)

        # Se não houver dados, a tabela estará vazia (com cabeçalhos)
        if history_data.empty:
            # A tabela já estará vazia, não é necessário fazer nada
            pass

        # Exibir o botão "Voltar" no final da sessão de detalhes
        self.back_button.pack(side='bottom', pady=20)

    def on_history_treeview_click(self, event):
        selected_item = event.widget.selection()
        if selected_item:
            item_iid = selected_item[0]
            selected_row = self.history_tree_data.loc[int(item_iid)]
            self.show_details_in_new_window(selected_row)

    def show_details_in_new_window(self, row_data):
        # Criar uma nova janela
        detail_window = tk.Toplevel(self.root)
        detail_window.title("Detalhes da Solicitação")
        detail_window.geometry("800x600")
        detail_window.configure(bg=self.bg_color)

        detail_frame = tk.Frame(detail_window, bg=self.bg_color)
        detail_frame.pack(fill="both", expand=True)

        # Adicionar o título na sessão de detalhes
        details_title_label = tk.Label(detail_frame, text="Detalhes da Solicitação",
                                       font=("Helvetica", 16, "bold"), bg=self.bg_color)
        details_title_label.pack(pady=10)

        # Replicar a criação do notebook e das abas para exibir os detalhes
        # Estilo para as abas do Notebook com tamanho aumentado
        notebook_style = ttk.Style()
        notebook_style.theme_use('default')

        notebook_style.configure("CustomNotebook.TNotebook", background=self.bg_color)
        notebook_style.configure("CustomNotebook.TNotebook.Tab", background=self.bg_color, font=("Helvetica", 13))

        # Definir estilos para frames e labels
        frame_style_name = 'Custom.TFrame'
        notebook_style.configure(frame_style_name, background=self.bg_color)

        notebook_style.configure("CustomBold.TLabel", font=("Helvetica", 12, "bold"), background=self.bg_color)
        notebook_style.configure("CustomRegular.TLabel", font=("Helvetica", 12), background=self.bg_color)

        # Criar um Notebook (interface de abas) com estilo personalizado
        notebook = ttk.Notebook(detail_frame, style="CustomNotebook.TNotebook")
        notebook.pack(fill='both', expand=True)

        # Reutilizar o dicionário 'sections' para criar as abas e exibir os detalhes
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
            tab_frame = ttk.Frame(notebook, style=frame_style_name)
            notebook.add(tab_frame, text=section_name)

            # Configurar as colunas
            tab_frame.columnconfigure(0, weight=1, minsize=200)
            tab_frame.columnconfigure(1, weight=3)

            row_idx = 0
            for col in fields:
                if col in row_data:
                    label = ttk.Label(tab_frame, text=f"{col}:", style="CustomBold.TLabel", wraplength=200, justify='left')
                    label.grid(row=row_idx, column=0, sticky='nw', padx=10, pady=5)
                    value = ttk.Label(tab_frame, text=row_data[col], style="CustomRegular.TLabel", wraplength=600, justify="left")
                    value.grid(row=row_idx, column=1, sticky='nw', padx=10, pady=5)
                    row_idx += 1

        # Botão para fechar a janela
        close_button = tk.Button(detail_frame, text="Fechar", command=detail_window.destroy, bg=self.button_bg_color)
        close_button.pack(pady=10)

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

    def show_statistics(self):
        # Ocultar outros frames
        if self.welcome_frame.winfo_ismapped():
            self.welcome_frame.pack_forget()
        if self.table_frame.winfo_ismapped():
            self.table_frame.pack_forget()
        if self.details_frame and self.details_frame.winfo_ismapped():
            self.details_frame.pack_forget()
            self.details_frame.destroy()
            self.details_frame = None

        # Esconder o botão "Voltar" se estiver visível
        if self.back_button.winfo_ismapped():
            self.back_button.pack_forget()

        # Atualizar o botão selecionado
        self.update_selected_button("Estatísticas")

        # Criar o frame de estatísticas
        self.statistics_frame = tk.Frame(self.content_frame, bg=self.bg_color)
        self.statistics_frame.pack(fill='both', expand=True)
        self.display_statistics()

    def display_statistics(self):
        # Limpar o frame
        for widget in self.statistics_frame.winfo_children():
            widget.destroy()

        # Recarregar os dados
        self.data = self.sheets_handler.load_data()

        # Converter 'Valor' para numérico
        self.data['Valor'] = self.data['Valor'].astype(str).str.replace(',', '.').str.extract(r'(\d+\.?\d*)')[0]
        self.data['Valor'] = pd.to_numeric(self.data['Valor'], errors='coerce').fillna(0)

        total_requests = len(self.data)
        pending_requests = len(self.data[
            self.data['Status'].isin(['Autorizado', 'Aguardando documentação'])
        ])
        awaiting_payment_requests = len(self.data[self.data['Status'] == 'Pronto para pagamento'])
        paid_requests = len(self.data[self.data['Status'] == 'Pago'])
        total_paid_values = self.data[self.data['Status'] == 'Pago']['Valor'].sum()
        total_released_values = self.data[
            self.data['Status'].isin(['Pago', 'Pronto para pagamento'])
        ]['Valor'].sum()

        # Exibir as estatísticas
        stats_text = f"""
    Número total de solicitações: {total_requests}
    Número de solicitações pendentes: {pending_requests}
    Número de solicitações aguardando pagamento: {awaiting_payment_requests}
    Número de solicitações pagas: {paid_requests}
    Soma dos valores já pagos: R$ {total_paid_values:.2f}
    Soma dos valores já liberados: R$ {total_released_values:.2f}
    """

        stats_label = tk.Label(self.statistics_frame, text=stats_text, font=("Helvetica", 12), bg=self.bg_color, justify='left')
        stats_label.pack(pady=10, padx=10, anchor='w')

        # Criar gráficos adicionais
        fig, axes = plt.subplots(2, 2, figsize=(10, 8))
        plt.subplots_adjust(hspace=0.5, wspace=0.3)

        # Gráfico de barras dos valores pagos ao longo do tempo
        paid_data = self.data[self.data['Status'] == 'Pago'].copy()
        if not paid_data.empty:
            paid_data['Ultima Atualizacao'] = pd.to_datetime(paid_data['Ultima Atualizacao'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
            paid_data = paid_data.dropna(subset=['Ultima Atualizacao'])
            paid_data_grouped = paid_data.groupby(paid_data['Ultima Atualizacao'].dt.date)['Valor'].sum()
            paid_data_grouped.plot(kind='bar', ax=axes[0, 0])
            axes[0, 0].set_title('Valores Pagos ao Longo do Tempo')
            axes[0, 0].set_xlabel('Data')
            axes[0, 0].set_ylabel('Valor Pago (R$)')
        else:
            axes[0, 0].text(0.5, 0.5, 'Nenhum pagamento realizado ainda.', horizontalalignment='center', verticalalignment='center')
            axes[0, 0].set_title('Valores Pagos ao Longo do Tempo')
            axes[0, 0].axis('off')

        # Gráfico de pizza por Motivo da Solicitação
        motivo_counts = self.data['Motivo da solicitação'].value_counts()
        motivo_counts.plot(kind='pie', ax=axes[0, 1], autopct='%1.1f%%')
        axes[0, 1].set_ylabel('')
        axes[0, 1].set_title('Distribuição por Motivo da Solicitação')

        # Gráfico de pizza por Agência de Fomento
        agencia_counts = self.data['Qual a agência de fomento?'].value_counts()
        agencia_counts.plot(kind='pie', ax=axes[1, 0], autopct='%1.1f%%')
        axes[1, 0].set_ylabel('')
        axes[1, 0].set_title('Distribuição por Agência de Fomento')

        # Gráfico de barras do número de solicitações por Status
        status_counts = self.data['Status'].value_counts()
        status_counts.plot(kind='bar', ax=axes[1, 1])
        axes[1, 1].set_title('Solicitações por Status')
        axes[1, 1].set_xlabel('Status')
        axes[1, 1].set_ylabel('Número de Solicitações')

        fig.tight_layout()

        # Exibir o gráfico no Tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.statistics_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10)

    def open_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Configurações")
        settings_window.geometry("600x600")
        settings_window.configure(bg=self.bg_color)

        instructions = ("Nesta seção, você pode editar os modelos de e-mail utilizados.\n"
                        "Use chaves para inserir dados do formulário. Por exemplo:\n"
                        "{Nome} será substituído pelo nome do solicitante.")

        instructions_label = tk.Label(settings_window, text=instructions, bg=self.bg_color, font=("Helvetica", 12), justify='left')
        instructions_label.pack(pady=10, padx=10)

        # Criar botões para editar cada modelo de e-mail
        for motivo in self.email_templates.keys():
            button = tk.Button(settings_window, text=f"Editar E-mail para {motivo}", command=lambda m=motivo: self.edit_email_template(m), bg=self.button_bg_color)
            button.pack(pady=5, padx=10, fill='x')

    def edit_email_template(self, motivo):
        template_window = tk.Toplevel(self.root)
        template_window.title(f"Editar Modelo de E-mail para {motivo}")
        template_window.geometry("600x400")
        template_window.configure(bg=self.bg_color)

        label = tk.Label(template_window, text=f"Editando o modelo de e-mail para {motivo}:", bg=self.bg_color, font=("Helvetica", 12))
        label.pack(pady=10)

        text_widget = Text(template_window, width=70, height=15)
        text_widget.pack(pady=10, padx=10)
        text_widget.insert(tk.END, self.email_templates[motivo])

        def save_template():
            new_template = text_widget.get("1.0", tk.END).strip()
            self.email_templates[motivo] = new_template
            self.save_email_templates()
            template_window.destroy()

        save_button = tk.Button(template_window, text="Salvar", command=save_template, bg="green", fg="white")
        save_button.pack(pady=10)

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
