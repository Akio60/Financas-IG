# app_gui.py

import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk
from dashboard import iniciar_dashboard, set_orcamento_total
from data_processing import ler_aba, atualizar_abas_com_colunas_personalizadas  # Adicionar importação necessária
import webbrowser
import pandas as pd
from auth import autenticar_google_sheets
import threading
from datetime import datetime
import gspread

# Variáveis globais para o link, orçamento e dimensões das janelas
link_personalizado = None
orcamento_total = None
LARGURA_JANELA = 800
ALTURA_JANELA = 600

def carregar_configuracoes():
    global link_personalizado, orcamento_total
    client, _ = autenticar_google_sheets('credentials.json', 'Configurações')
    
    try:
        sheet_config = client.open("Configurações").sheet1
    except gspread.SpreadsheetNotFound:
        sheet_config = client.create("Configurações").sheet1
        sheet_config.update("A1", [["Link Personalizado", "Orçamento Total"]])
        sheet_config.update("A2", [["https://www.exemplo.com", 500000]])
        link_personalizado = "https://www.exemplo.com"
        orcamento_total = 500000
        set_orcamento_total(orcamento_total)
        return
    
    config_data = sheet_config.get_all_records()
    link_personalizado = config_data[0]["Link Personalizado"]
    orcamento_total = float(config_data[0]["Orçamento Total"])
    set_orcamento_total(orcamento_total)

def salvar_configuracoes():
    global link_personalizado, orcamento_total
    client, _ = autenticar_google_sheets('credentials.json', 'Configurações')
    sheet_config = client.open("Configurações").sheet1
    sheet_config.update("A1", [["Link Personalizado", "Orçamento Total"]])
    sheet_config.update("A2", [[link_personalizado, orcamento_total]])

def abrir_link():
    global link_personalizado
    webbrowser.open(link_personalizado)

def registrar_alteracao(client, acao, dados_anteriores):
    """
    Registra a ação de edição ou remoção na planilha Registros.
    """
    try:
        sheet_registros = client.open("Registros").sheet1
    except gspread.SpreadsheetNotFound:
        # Criar a planilha "Registros" se ela não existir
        sheet_registros = client.create("Registros").sheet1
        sheet_registros.update("A1", [["Data", "Ação", "Dados Anteriores"]])
        
    data_atual = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    registro = [[data_atual, acao, str(dados_anteriores)]]
    sheet_registros.append_rows(registro)

def carregar_dados_e_atualizar_abas():
    """
    Carrega os dados do formulário e atualiza as abas no Google Sheets.
    """
    client, _ = autenticar_google_sheets('credentials.json', 'Relatórios PPG Geografia (Responses)')
    # Obter dados da aba principal
    sheet = client.open('Relatórios PPG Geografia (Responses)').sheet1
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    
    # Atualizar as abas no Google Sheets usando a função do módulo data_processing
    atualizar_abas_com_colunas_personalizadas(df, 'Tipo do requerimento', client, 'Relatórios PPG Geografia (Responses)')

def abrir_configuracoes():
    config_janela = tk.Toplevel()
    config_janela.title("Configurações")
    config_janela.geometry("400x300")
    config_janela.configure(bg="#2C3E50")

    tk.Label(config_janela, text="Configurações do Aplicativo", font=("Helvetica", 14), bg="#2C3E50", fg="#ECF0F1").pack(pady=10)
    tk.Label(config_janela, text="Link Personalizado:", bg="#2C3E50", fg="#ECF0F1").pack()
    entry_link = tk.Entry(config_janela, width=50)
    entry_link.insert(0, link_personalizado)
    entry_link.pack()

    tk.Label(config_janela, text="Orçamento Total (R$):", bg="#2C3E50", fg="#ECF0F1").pack()
    entry_orcamento = tk.Entry(config_janela, width=20)
    entry_orcamento.insert(0, str(orcamento_total))
    entry_orcamento.pack()

    def salvar_alteracoes():
        global link_personalizado, orcamento_total
        link_personalizado = entry_link.get()
        orcamento_total = float(entry_orcamento.get())
        set_orcamento_total(orcamento_total)
        salvar_configuracoes()
        messagebox.showinfo("Configurações Salvas", "Configurações atualizadas com sucesso!")
        config_janela.destroy()

    tk.Button(config_janela, text="Salvar Configurações", command=salvar_alteracoes, bg="#27AE60", fg="#ECF0F1").pack(pady=10)

def iniciar_dashboard_thread():
    threading.Thread(target=iniciar_dashboard, daemon=True).start()

def mostrar_dados_aba(nome_aba, client):
    """
    Cria uma janela fixa e maximizada para mostrar os dados da aba.
    Inclui uma caixa de busca, permite edição e remoção dos dados selecionados, com confirmação.
    """
    # Obter dados da aba
    df = ler_aba(client, 'Relatórios PPG Geografia (Responses)', nome_aba)
    aba_principal = client.open("Relatórios PPG Geografia (Responses)").sheet1

    # Configurar a janela
    dados_janela = tk.Toplevel()
    dados_janela.title(f"Dados - {nome_aba}")
    dados_janela.geometry(f"{LARGURA_JANELA}x{ALTURA_JANELA}")
    dados_janela.configure(bg="#34495E")

    # Caixa de busca e botão buscar
    frame_search = tk.Frame(dados_janela, bg="#34495E")
    frame_search.pack(fill="x", padx=10, pady=5)
    
    entry_busca = tk.Entry(frame_search, width=30)
    entry_busca.grid(row=0, column=0, padx=(0, 5))
    
    def buscar():
        termo = entry_busca.get().lower()
        dados_filtrados = df[df.apply(lambda row: row.astype(str).str.lower().str.contains(termo).any(), axis=1)]
        atualizar_tabela(dados_filtrados)

    tk.Button(frame_search, text="Buscar", command=buscar, bg="#2980B9", fg="#ECF0F1", width=10).grid(row=0, column=1)

    # Tabela de dados com opção de ordenação
    frame_table = tk.Frame(dados_janela, bg="#34495E")
    frame_table.pack(fill="both", expand=True)
    
    colunas = df.columns.tolist()
    tabela = ttk.Treeview(frame_table, columns=colunas, show="headings", height=15)
    tabela.pack(expand=True, fill="both")

    for col in colunas:
        tabela.heading(col, text=col, command=lambda _col=col: ordenar_tabela(_col))

    def ordenar_tabela(col):
        df.sort_values(by=col, ascending=True, inplace=True)
        atualizar_tabela(df)

    def atualizar_tabela(dados):
        tabela.delete(*tabela.get_children())
        for _, row in dados.iterrows():
            tabela.insert("", "end", values=row.tolist())

    atualizar_tabela(df)

    def editar_item():
        selecionado = tabela.selection()
        if not selecionado:
            messagebox.showwarning("Seleção vazia", "Por favor, selecione um item para editar.")
            return

        valores = tabela.item(selecionado, "values")
        editar_janela = tk.Toplevel()
        editar_janela.title("Editar Item")
        editar_janela.geometry(f"{LARGURA_JANELA}x{ALTURA_JANELA}")
        editar_janela.configure(bg="#2C3E50")

        entradas = {}
        for i, valor in enumerate(valores):
            tk.Label(editar_janela, text=colunas[i], bg="#2C3E50", fg="#ECF0F1").grid(row=i, column=0, padx=10, pady=5, sticky="w")
            entrada = tk.Entry(editar_janela, width=30)
            entrada.grid(row=i, column=1, padx=10, pady=5)
            entrada.insert(0, valor)
            entradas[colunas[i]] = entrada

        def salvar_edicao():
            novo_valor = [entradas[col].get() for col in colunas]
            try:
                # Remover o registro antigo e adicionar o novo na aba correspondente
                index = df.index[df[colunas[0]] == valores[0]].tolist()[0]
                aba_principal.delete_rows(index + 2)  # Remove a linha original

                # Adiciona o registro editado
                aba_principal.append_row(novo_valor)
                
                # Atualiza abas específicas
                atualizar_abas_com_colunas_personalizadas(df, 'Tipo do requerimento', client, 'Relatórios PPG Geografia (Responses)')
                registrar_alteracao(client, "Edição", valores)
                
                messagebox.showinfo("Salvo", "Dados atualizados com sucesso!")
                editar_janela.destroy()
            except IndexError:
                messagebox.showerror("Erro", "Erro ao encontrar o índice do item selecionado.")


        tk.Button(editar_janela, text="Salvar Item", command=salvar_edicao, bg="#27AE60", fg="#ECF0F1").grid(row=len(valores), column=0, columnspan=2, pady=10)

    # Função para remover o item selecionado com confirmação
    def remover_item():
        selecionado = tabela.selection()
        if not selecionado:
            messagebox.showwarning("Seleção vazia", "Por favor, selecione um item para remover.")
            return

        resposta = messagebox.askyesno("Confirmar Remoção", "Deseja realmente remover o item selecionado?")
        if resposta:
            valores = tabela.item(selecionado, "values")
            tabela.delete(selecionado)
            aba_principal.delete_rows(df.index[df[colunas[0]] == valores[0]].tolist()[0] + 2)
            atualizar_abas_com_colunas_personalizadas(df, 'Tipo do requerimento', client, 'Relatórios PPG Geografia (Responses)')
            registrar_alteracao(client, "Remoção", valores)
            messagebox.showinfo("Removido", "Item removido com sucesso!")

    frame_actions = tk.Frame(dados_janela, bg="#34495E")
    frame_actions.pack(fill="x", padx=10, pady=10)

    tk.Button(frame_actions, text="Editar Item Selecionado", command=editar_item, bg="#27AE60", fg="#ECF0F1", width=20).pack(side="left", padx=5)
    tk.Button(frame_actions, text="Remover Item Selecionado", command=remover_item, bg="#E74C3C", fg="#ECF0F1", width=20).pack(side="right", padx=5)

def iniciar_interface():
    carregar_dados_e_atualizar_abas()
    carregar_configuracoes()

    root = tk.Tk()
    root.title("Controle de Finanças - Nubia")
    root.geometry(f"{LARGURA_JANELA}x{ALTURA_JANELA}")
    root.resizable(False, False)
    root.configure(bg="#2C3E50")

    frame = tk.Frame(root, bd=2, relief="ridge", bg="#34495E")
    frame.pack(pady=20, padx=20, fill="both", expand=True)

    imagem_esquerda = Image.open("logo_unicamp.png").resize((80, 80), Image.LANCZOS)
    imagem_direita = Image.open("logo_ig.png").resize((80, 80), Image.LANCZOS)
    imagem_esquerda_tk = ImageTk.PhotoImage(imagem_esquerda)
    imagem_direita_tk = ImageTk.PhotoImage(imagem_direita)

    tk.Label(root, image=imagem_esquerda_tk, bg="#2C3E50").place(x=25, y=25)
    tk.Label(root, image=imagem_direita_tk, bg="#2C3E50").place(x=690, y=25)

    label_title = tk.Label(frame, text="Aplicação de Gerenciamento", font=("Helvetica", 16, "bold"), bg="#34495E", fg="#ECF0F1")
    label_title.pack(pady=10)

    client, _ = autenticar_google_sheets('credentials.json', 'Relatórios PPG Geografia (Responses)')

    frame_left = tk.Frame(frame, bg="#34495E")
    frame_right = tk.Frame(frame, bg="#34495E")
    frame_left.pack(side="left", fill="both", expand=True, padx=10)
    frame_right.pack(side="right", fill="both", expand=True, padx=10)

    botao_acoes_esquerda = {
        "Dashboard": iniciar_dashboard_thread,
        "Adicionar registro": abrir_link,
        "Envio de Emails": lambda: print("Funcionalidade de envio de emails"),
        "Configurações": abrir_configuracoes
    }
    for label, command in botao_acoes_esquerda.items():
        tk.Button(frame_left, text=label, font=("Helvetica", 12), bg="#2980B9", fg="#ECF0F1",
                  command=command, relief="flat", width=20, height=2).pack(pady=5)

    botoes_abas = {
        "Defesas": lambda: mostrar_dados_aba("Defesas", client),
        "Periódico": lambda: mostrar_dados_aba("Periódico", client),
        "Jornal e Revista": lambda: mostrar_dados_aba("Jornal e Revista", client),
        "Recursos Financeiros": lambda: mostrar_dados_aba("Recursos Financeiros", client)
    }
    for label, command in botoes_abas.items():
        tk.Button(frame_right, text=label, font=("Helvetica", 12), bg="#2980B9", fg="#ECF0F1",
                  command=command, relief="flat", width=20, height=2).pack(pady=5)

    root.mainloop()
