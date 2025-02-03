import tkinter as tk
from tkinter import messagebox
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


# Definir usuários e senhas predefinidos
USUARIOS = {
    'teste': 'teste',
    'rafael': 'rafael',
    'isabela': 'isabela',
    'fernanda': 'fernanda'
}

# Configurações do Google Sheets
CREDENTIALS_FILE = "credenciais.json"  # Substitua com o caminho correto para seu arquivo de credenciais
SHEET_NAME = "ponto"  # Nome da planilha


# Funções de autenticação e interação com o Google Sheets
def conectar_googlesheets():
    """Função para autenticar e conectar ao Google Sheets"""
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
    client = gspread.authorize(creds)
    planilha = client.open(SHEET_NAME)
    return planilha


def criar_aba_usuario(usuario):
    """Função para criar aba do usuário se ela não existir"""
    planilha = conectar_googlesheets()

    try:
        aba = planilha.worksheet(usuario)  # Tenta acessar a aba do usuário
    except gspread.exceptions.WorksheetNotFound:
        aba = planilha.add_worksheet(title=usuario, rows="100", cols="7")
        aba.append_row(['Usuário', 'Data', 'Hora Entrada', 'Hora Saída', 'Hora Almoço Início', 'Hora Almoço Fim', 'Tipo'])  # Cabeçalhos
    return aba


def ponto_ja_marcado(usuario, data_atual, tipo_ponto):
    """Função para verificar se o ponto já foi marcado"""
    planilha = conectar_googlesheets()
    aba = planilha.worksheet(usuario)

    registros = aba.get_all_records()
    for registro in registros:
        if registro['Usuário'] == usuario and registro['Data'] == data_atual and registro['Tipo'] == tipo_ponto:
            return True
    return False


# Funções de marcação de ponto
def realizar_login():
    """Função de login"""
    usuario = entry_usuario.get()
    senha = entry_senha.get()

    if USUARIOS.get(usuario) == senha:
        messagebox.showinfo("Sucesso", "Login realizado com sucesso!")
        login_frame.pack_forget()
        marcar_ponto_frame.pack(pady=20)
        user_atual.set(usuario)

        # Criar a aba do usuário se não existir
        criar_aba_usuario(usuario)
        atualizar_botoes()
    else:
        messagebox.showerror("Erro", "Usuário ou senha inválidos.")


def atualizar_botoes():
    """Função para atualizar o estado dos botões"""
    usuario = user_atual.get()
    data_atual = datetime.datetime.now().strftime("%d/%m/%Y")

    # Verificar quais pontos já foram marcados hoje
    pontos = ["Entrada", "Saída", "Início Almoço", "Fim Almoço"]
    botoes = [entrada_button, saida_button, almoco_inicio_button, almoco_fim_button]

    for ponto, botao in zip(pontos, botoes):
        if ponto_ja_marcado(usuario, data_atual, ponto):
            botao.config(state=tk.DISABLED)
        else:
            botao.config(state=tk.NORMAL)


def marcar_ponto(tipo_ponto):
    """Função para marcar o ponto"""
    usuario = user_atual.get()
    data_atual = datetime.datetime.now().strftime("%d/%m/%Y")
    hora_atual = datetime.datetime.now().strftime("%H:%M:%S")

    if ponto_ja_marcado(usuario, data_atual, tipo_ponto):
        messagebox.showwarning("Aviso", f"Você já marcou o ponto de {tipo_ponto} hoje!")
        return

    planilha = conectar_googlesheets()
    aba = planilha.worksheet(usuario)

    pontos = {
        "entrada": [usuario, data_atual, hora_atual, "", "", "", "Entrada"],
        "saida": [usuario, data_atual, "", hora_atual, "", "", "Saída"],
        "almoco_inicio": [usuario, data_atual, "", "", hora_atual, "", "Início Almoço"],
        "almoco_fim": [usuario, data_atual, "", "", "", hora_atual, "Fim Almoço"]
    }

    aba.append_row(pontos[tipo_ponto])

    messagebox.showinfo("Sucesso", f"Ponto de {tipo_ponto} marcado com sucesso!")
    atualizar_botoes()


# Função para exportar o extrato mensal em PDF
def exportar_pdf():
    """Função para exportar o extrato mensal em PDF"""
    usuario = user_atual.get()  # Obtém o usuário atual
    data_atual = datetime.datetime.now().strftime("%m/%Y")  # Obtém o mês e ano atual
    
    # Conectar à planilha e acessar a aba do usuário
    planilha = conectar_googlesheets()
    aba = planilha.worksheet(usuario)
    
    # Obter todos os registros da aba do usuário
    registros = aba.get_all_records()
    
    # Filtrar registros do mês atual
    registros_mes = [registro for registro in registros if datetime.datetime.strptime(registro['Data'], "%d/%m/%Y").strftime("%m/%Y") == data_atual]
    
    # Caso não haja registros, exibir um aviso
    if not registros_mes:
        messagebox.showwarning("Aviso", "Nenhum ponto registrado para este mês.")
        return
    
    # Definir o nome do PDF
    nome_pdf = f"{usuario}_extrato_{data_atual.replace('/', '-')}.pdf"
    
    # Criar o objeto canvas para gerar o PDF
    c = canvas.Canvas(nome_pdf, pagesize=letter)
    c.drawString(100, 750, f"Extrato de Ponto - {usuario} - {data_atual}")
    
    # Adicionar cada registro ao PDF
    y = 730
    for registro in registros_mes:
        linha = f"Data: {registro['Data']} | Entrada: {registro['Hora Entrada']} | Almoço: {registro['Hora Almoço Início']}{registro['Hora Almoço Fim']} | Saída: {registro['Hora Saída']}"
        c.drawString(100, y, linha)
        y -= 20  # Descer a posição para a próxima linha
        
        # Se o conteúdo ultrapassar o limite da página, cria uma nova página
        if y < 100:
            c.showPage()  # Adiciona uma nova página
            c.drawString(100, 750, f"Extrato de Ponto - {usuario} - {data_atual}")
            y = 730  # Resetar a posição 'y' para o topo da nova página
    
    # Salvar o PDF
    c.save()
    
    # Mostrar uma mensagem de sucesso
    messagebox.showinfo("Sucesso", f"Extrato exportado para {nome_pdf}")


# Configuração da janela principal
root = tk.Tk()

# Ajuste o tamanho da janela e centralize na tela
largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()
largura_janela = 600  # Defina o tamanho desejado
altura_janela = 400  # Defina o tamanho desejado
pos_x = (largura_tela - largura_janela) // 2
pos_y = (altura_tela - altura_janela) // 2
root.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

root.title("Marcador de Ponto")

# Definindo a fonte maior
font_padrao = ("Helvetica", 14)

# Frame de login
login_frame = tk.Frame(root)
login_frame.pack(padx=10, pady=20, fill=tk.BOTH, expand=True)

tk.Label(login_frame, text="Usuário:", font=font_padrao).grid(row=0, column=0, padx=5, pady=10)
entry_usuario = tk.Entry(login_frame, font=font_padrao)
entry_usuario.grid(row=0, column=1, padx=5, pady=10)

tk.Label(login_frame, text="Senha:", font=font_padrao).grid(row=1, column=0, padx=5, pady=10)
entry_senha = tk.Entry(login_frame, show="*", font=font_padrao)
entry_senha.grid(row=1, column=1, padx=5, pady=10)

login_button = tk.Button(login_frame, text="Entrar", command=realizar_login, font=font_padrao)
login_button.grid(row=2, columnspan=2, pady=20)

# Frame para marcação de ponto
marcar_ponto_frame = tk.Frame(root)

user_atual = tk.StringVar()

tk.Label(marcar_ponto_frame, text="Usuário:", font=font_padrao).grid(row=0, column=0, padx=5, pady=10)
tk.Label(marcar_ponto_frame, textvariable=user_atual, font=font_padrao).grid(row=0, column=1, padx=5, pady=10)

entrada_button = tk.Button(marcar_ponto_frame, text="Entrada", command=lambda: marcar_ponto("entrada"), font=font_padrao)
entrada_button.grid(row=1, column=0, padx=5, pady=5)

saida_button = tk.Button(marcar_ponto_frame, text="Saída", command=lambda: marcar_ponto("saida"), font=font_padrao)
saida_button.grid(row=1, column=1, padx=5, pady=5)

almoco_inicio_button = tk.Button(marcar_ponto_frame, text="Início Almoço", command=lambda: marcar_ponto("almoco_inicio"), font=font_padrao)
almoco_inicio_button.grid(row=2, column=0, padx=5, pady=5)

almoco_fim_button = tk.Button(marcar_ponto_frame, text="Fim Almoço", command=lambda: marcar_ponto("almoco_fim"), font=font_padrao)
almoco_fim_button.grid(row=2, column=1, padx=5, pady=5)

exportar_button = tk.Button(marcar_ponto_frame, text="Exportar PDF", command=exportar_pdf, font=font_padrao)
exportar_button.grid(row=3, columnspan=2, pady=20)

# Iniciar a interface
root.mainloop()
