import customtkinter as ctk
from tkinter import messagebox
from docx import Document
import os

ctk.set_appearance_mode("dark")  # Configura o modo de aparência para 'dark'
ctk.set_default_color_theme("blue")

import sys

if getattr(sys, 'frozen', False):  # Se o programa estiver rodando como um executável
    pasta_modelos = os.path.join(sys._MEIPASS, "modelos")  
else:
    pasta_modelos = os.path.dirname(__file__)  

modelos_disponiveis = {
    "1": os.path.join(pasta_modelos, "1.ACERVO.docx"),
    "2": os.path.join(pasta_modelos, "2.DSA.docx"),
    "3": os.path.join(pasta_modelos, "3.IDONEIDADE.docx"),
    "4": os.path.join(pasta_modelos, "4.COMPETIÇÃO.docx"),
    "5": os.path.join(pasta_modelos, "5.Procuração.docx"),
}

def gerar_documento(dados, modelos_selecionados):
    try:
        pasta_destino = "Documentos Gerados"
        os.makedirs(pasta_destino, exist_ok=True)
        
        for modelo in modelos_selecionados:
            modelo_path = modelos_disponiveis.get(modelo)
            if not modelo_path or not os.path.exists(modelo_path):
                messagebox.showerror("Erro", f"Modelo {modelo} não encontrado!")
                return
            
            doc = Document(modelo_path)
            for paragrafo in doc.paragraphs:
                for run in paragrafo.runs:  # Itera sobre as partes do texto dentro do parágrafo
                    for chave, valor in dados.items():
                        if f"{{{{{chave}}}}}" in run.text:
                            run.text = run.text.replace(f"{{{{{chave}}}}}", valor)  # Substitui sem perder a formatação
            
            doc.save(os.path.join(pasta_destino, os.path.basename(modelo_path)))
        
        messagebox.showinfo("Sucesso", "Documentos gerados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar documento: {str(e)}")

# Interface
tela = ctk.CTk()
tela.title("Gerador de Documentos")
tela.geometry("480x600")  # Janela um pouco maior para acomodar as duas colunas
tela.config(bg="black")  # Define o fundo da tela para preto

# Frame com barra de rolagem
frame_rolagem = ctk.CTkFrame(tela, bg_color="black")
frame_rolagem.pack(fill="both", expand=True, padx=20, pady=20)  # Centralizando com margens

# Adicionando a barra de rolagem
canvas = ctk.CTkCanvas(frame_rolagem, bg="black")
canvas.pack(side="left", fill="both", expand=True)

scrollbar = ctk.CTkScrollbar(frame_rolagem, command=canvas.yview)  # Removido 'orient'
scrollbar.pack(side="right", fill="y")

canvas.configure(yscrollcommand=scrollbar.set)

# Criando um frame dentro do Canvas
frame_conteudo = ctk.CTkFrame(canvas, bg_color="black")
canvas.create_window((0, 0), window=frame_conteudo, anchor="center")

# Ajustando para a rolagem
frame_conteudo.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

ctk.CTkLabel(frame_conteudo, text="📄 Gerador de Documentos", font=("Arial", 16, "bold"), text_color="white").pack(pady=10)

modelo_checkboxes = {}
frame_check = ctk.CTkFrame(frame_conteudo, bg_color="black")
frame_check.pack(pady=10, padx=15, fill="x")  # Margem lateral
for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = ctk.BooleanVar()
    ctk.CTkCheckBox(frame_check, text=os.path.basename(modelo_nome), variable=var, text_color="white").pack(anchor="center", padx=5)

# Dividindo os campos de entrada em 2 colunas
frame_campos = ctk.CTkFrame(frame_conteudo, bg_color="black")
frame_campos.pack(pady=10, padx=15, fill="x")  # Margem lateral

campo_labels = [
    ("Nome Completo", "NOME COMPLETO"), ("RG", "RG"), ("CPF", "CPF"),
    ("Endereço", "ENDEREÇO COMPLETO"), ("Cidade", "CIDADE DE NASCIMENTO"),
    ("Estado", "ESTADO DE NASCIMENTO"), ("Data Nasc.", "DATA DE NASCIMENTO"),
    ("Pai", "NOME PAI"), ("Mãe", "NOME MÃE"), ("Telefone", "TELEFONE"),
    ("Expedição RG", "DATA DE EXPEDIÇÃO DOCUMENTO"), ("Órgão Expedidor", "ÓRGÃO EXPEDIDOR"), 
    ("Estado Civil", "ESTADO CIVIL"),("Profissão", "PROFISSÃO")
]

entradas = {}

# Ajustando para duas colunas
for i, (label_texto, chave) in enumerate(campo_labels):
    frame = ctk.CTkFrame(frame_campos, bg_color="black")
    frame.grid(row=i//2, column=i%2, padx=10, pady=5, sticky="ew")  # Grid para 2 colunas
    ctk.CTkLabel(frame, text=label_texto+":", text_color="white").pack(side="left", padx=5)
    entrada = ctk.CTkEntry(frame, width=180)  # Ajustado para ser mais estreito
    entrada.pack(side="right", fill="x", expand=True)
    entradas[chave] = entrada

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

frame_botoes = ctk.CTkFrame(frame_conteudo, bg_color="black")
frame_botoes.pack(pady=10, padx=15, fill="x")  # Margem lateral
ctk.CTkButton(frame_botoes, text="Gerar Selecionados", command=coletar_dados, width=140).pack(side="left", padx=5)
ctk.CTkButton(frame_botoes, text="Gerar Todos", command=selecionar_e_gerar, width=140).pack(side="right", padx=5)

tela.mainloop()
