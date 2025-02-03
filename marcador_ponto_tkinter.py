import tkinter as tk
from tkinter import messagebox
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import os.path
import datetime

# Configurações do Google Sheets
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
CREDENTIALS_FILE = "c:/Users/israd/OneDrive/Área de Trabalho/marcadordeponto/credenciais.json" # Substitua pelo caminho do seu arquivo JSON
SPREADSHEET_NAME = "PONTO-DANTEC"  # Nome da planilha no Google Sheets

# Função para autenticação do OAuth2
def autenticar_google_sheets():
    creds = None
    # Verifica se o token já existe
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # Se não houver credenciais válidas, solicita uma nova autorização
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Salva o token para a próxima execução
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    client = gspread.authorize(creds)
    return client.open(SPREADSHEET_NAME)

# Função para acessar a aba do usuário com login e senha
def acessar_aba_usuario():
    login = entry_login.get()
    senha = entry_senha.get()

    if login == senha and login:  # Verifica se login e senha são iguais e não estão vazios
        try:
            planilha = autenticar_google_sheets()
            # Tenta acessar a aba
            aba = planilha.worksheet(login)
            mostrar_botoes_ponto(aba, login)
        except gspread.exceptions.WorksheetNotFound:
            messagebox.showerror("Erro", f"A aba para o usuário {login} não existe.")
    else:
        messagebox.showwarning("Aviso", "Login ou senha inválidos!")

# Função para verificar se o ponto já foi registrado no dia
def ponto_registrado(aba, tipo_ponto):
    data_atual = datetime.datetime.now().strftime("%d/%m/%Y")
    registros = aba.get_all_records()
    for registro in registros:
        if registro['Data'] == data_atual and registro['Tipo'] == tipo_ponto:
            return True
    return False

# Função para registrar o ponto na planilha
def registrar_ponto(tipo, aba, btn_entrada, btn_saida, btn_almoco, btn_volta_almoco, btn_saida_final):
    if ponto_registrado(aba, tipo):
        messagebox.showwarning("Aviso", f"Você já registrou o ponto de {tipo} hoje.")
    else:
        hora_atual = datetime.datetime.now().strftime("%H:%M:%S")
        data_atual = datetime.datetime.now().strftime("%d/%m/%Y")
        aba.append_row([data_atual, hora_atual, tipo])
        messagebox.showinfo("Sucesso", f"Ponto {tipo} registrado com sucesso!")
        # Desativar o botão após o registro
        if tipo == "Entrada":
            btn_entrada.config(state=tk.DISABLED)
        elif tipo == "Saída":
            btn_saida.config(state=tk.DISABLED)
        elif tipo == "Saída Almoço":
            btn_almoco.config(state=tk.DISABLED)
        elif tipo == "Volta Almoço":
            btn_volta_almoco.config(state=tk.DISABLED)
        elif tipo == "Saída Final":
            btn_saida_final.config(state=tk.DISABLED)

# Função para exibir os botões de registro de ponto
def mostrar_botoes_ponto(aba, nome_usuario):
    # Esconde a tela de login
    frame_login.pack_forget()

    # Tela de opções de ponto
    frame_ponto = tk.Frame(root)
    frame_ponto.pack(padx=20, pady=20)

    # Exibe o nome do usuário
    tk.Label(frame_ponto, text=f"Bem-vindo, {nome_usuario}!", font=("Arial", 16)).pack(pady=10)

    # Botões de registro de ponto
    btn_entrada = tk.Button(frame_ponto, text="Entrada", command=lambda: registrar_ponto("Entrada", aba, btn_entrada, btn_saida, btn_almoco, btn_volta_almoco, btn_saida_final))
    btn_entrada.pack(padx=20, pady=5)

    btn_saida = tk.Button(frame_ponto, text="Saída", command=lambda: registrar_ponto("Saída", aba, btn_entrada, btn_saida, btn_almoco, btn_volta_almoco, btn_saida_final))
    btn_saida.pack(padx=20, pady=5)

    btn_almoco = tk.Button(frame_ponto, text="Saída para Almoço", command=lambda: registrar_ponto("Saída Almoço", aba, btn_entrada, btn_saida, btn_almoco, btn_volta_almoco, btn_saida_final))
    btn_almoco.pack(padx=20, pady=5)

    btn_volta_almoco = tk.Button(frame_ponto, text="Volta do Almoço", command=lambda: registrar_ponto("Volta Almoço", aba, btn_entrada, btn_saida, btn_almoco, btn_volta_almoco, btn_saida_final))
    btn_volta_almoco.pack(padx=20, pady=5)

    btn_saida_final = tk.Button(frame_ponto, text="Saída Final", command=lambda: registrar_ponto("Saída Final", aba, btn_entrada, btn_saida, btn_almoco, btn_volta_almoco, btn_saida_final))
    btn_saida_final.pack(padx=20, pady=5)

    # Botão para voltar
    btn_voltar = tk.Button(frame_ponto, text="Voltar", command=lambda: voltar_para_acesso())
    btn_voltar.pack(pady=20)

# Função para voltar à tela de login
def voltar_para_acesso():
    frame_ponto.pack_forget()
    frame_login.pack(padx=20, pady=20)

# Interface gráfica
root = tk.Tk()
root.title("Sistema de Registro de Ponto")

# Tela de login
frame_login = tk.Frame(root)
frame_login.pack(padx=20, pady=20)

tk.Label(frame_login, text="Login:").grid(row=0, column=0, padx=5, pady=5)
entry_login = tk.Entry(frame_login)
entry_login.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_login, text="Senha:").grid(row=1, column=0, padx=5, pady=5)
entry_senha = tk.Entry(frame_login, show="*")
entry_senha.grid(row=1, column=1, padx=5, pady=5)

btn_entrar = tk.Button(frame_login, text="Entrar", command=acessar_aba_usuario)
btn_entrar.grid(row=2, column=0, columnspan=2, padx=5, pady=5)

# Iniciar a interface
root.mainloop()
