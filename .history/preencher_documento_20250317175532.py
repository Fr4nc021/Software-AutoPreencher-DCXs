import customtkinter as ctk
from tkinter import messagebox
from docx import Document
import os
import sys

# Configuração do tema
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Obtendo o caminho correto, dependendo se está rodando no .exe ou no .py
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(os.path.dirname(__file__))

# Definindo a pasta onde os modelos estão
pasta_modelos = os.path.join(base_path, "modelos")

# Garante que a pasta de modelos existe
if not os.path.exists(pasta_modelos):
    os.makedirs(pasta_modelos)

# Mapeando os modelos disponíveis
modelos_disponiveis = {
    "1": os.path.join(pasta_modelos, "1.ACERVO.docx"),
    "2": os.path.join(pasta_modelos, "2.DSA.docx"),
    "3": os.path.join(pasta_modelos, "3.IDONEIDADE.docx"),
    "4": os.path.join(pasta_modelos, "4.COMPETIÇÃO.docx"),
    "5": os.path.join(pasta_modelos, "5.Procuração.docx"),
}

# Criando a interface
tela = ctk.CTk()
tela.title("Gerador de Documentos")
tela.geometry("500x600")

ctk.CTkLabel(tela, text="📄 Gerador de Documentos", font=("Arial", 18, "bold")).pack(pady=(15, 10))

# Adicionando um frame com scrollbar
frame_scroll = ctk.CTkFrame(tela)
frame_scroll.pack(fill="both", expand=True, padx=10, pady=10)

canvas = ctk.CTkCanvas(frame_scroll, height=450)
scrollbar = ctk.CTkScrollbar(frame_scroll, command=canvas.yview)
frame_conteudo = ctk.CTkFrame(canvas)

frame_conteudo.bind(
    "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=frame_conteudo, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# Checkbox de seleção de modelos
modelo_checkboxes = {}
frame_check = ctk.CTkFrame(frame_conteudo)
frame_check.pack(pady=10)
for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = ctk.BooleanVar()
    check = ctk.CTkCheckBox(frame_check, text=os.path.basename(modelo_nome), variable=var)
    check.pack(anchor="w")
    modelo_checkboxes[modelo_id] = var

# Campos de entrada com limite de texto
campo_labels = [
    ("Nome Completo", "NOME COMPLETO", 50),
    ("RG", "RG", 15),
    ("CPF", "CPF", 14),
    ("Endereço completo", "ENDEREÇO COMPLETO", 100),
    ("Cidade de Nascimento", "CIDADE DE NASCIMENTO", 30),
    ("Estado de nascimento", "ESTADO DE NASCIMENTO", 2),
    ("Data de Nascimento", "DATA DE NASCIMENTO", 10),
    ("Nome do Pai", "NOME PAI", 50),
    ("Nome da Mãe", "NOME MÃE", 50),
    ("Telefone", "TELEFONE", 15),
    ("Data de expedição do RG", "DATA DE EXPEDIÇÃO RG", 10),
    ("Órgão expedidor do RG", "ÓRGÃO EXPEDIDOR RG", 10),
]

entradas = {}

def limitar_texto(entrada, limite):
    def callback(texto):
        if len(texto) > limite:
            return False
        return True
    return tela.register(callback)

frame_campos = ctk.CTkFrame(frame_conteudo)
frame_campos.pack(pady=(10, 20))

for label_texto, chave, limite in campo_labels:
    frame = ctk.CTkFrame(frame_campos)
    frame.pack(fill="x", padx=10, pady=5)
    ctk.CTkLabel(frame, text=label_texto + ":").pack(side="left")
    entrada = ctk.CTkEntry(frame, width=250, validate="key", validatecommand=(limitar_texto(entrada, limite), "%P"))
    entrada.pack(side="right", fill="x", expand=True)
    entradas[chave] = entrada

# Funções de ação
def coletar_dados():
    dados = {chave: entrada.get() for chave, entrada in entradas.items()}
    modelos_selecionados = [modelo for modelo, selecionado in modelo_checkboxes.items() if selecionado.get()]
    if modelos_selecionados:
        gerar_documento(dados, modelos_selecionados)
    else:
        messagebox.showerror("Erro", "Selecione pelo menos um modelo!")

def selecionar_todos():
    for modelo in modelo_checkboxes.values():
        modelo.set(True)

def selecionar_e_gerar():
    selecionar_todos()
    coletar_dados()

# Botões estilizados
frame_botoes = ctk.CTkFrame(frame_conteudo)
frame_botoes.pack(pady=20)
ctk.CTkButton(frame_botoes, text="Gerar Selecionados", command=coletar_dados, width=200).pack(side="left", padx=5)
ctk.CTkButton(frame_botoes, text="Gerar Todos", command=selecionar_e_gerar, width=200).pack(side="right", padx=5)

# Executando o app
tela.mainloop()
