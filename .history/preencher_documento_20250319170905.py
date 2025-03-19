import customtkinter as ctk
from tkinter import messagebox
from docx import Document
import os

ctk.set_appearance_mode("System")
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
tela.geometry("380x600")  # Janela mais estreita

# Adicionando um Canvas para a rolagem
canvas = ctk.CTkCanvas(tela)
canvas.pack(side="left", fill="both", expand=True)

# Adicionando a barra de rolagem
scrollbar = ctk.CTkScrollbar(tela, orient="vertical", command=canvas.yview)
scrollbar.pack(side="right", fill="y")

# Configurando a rolagem do Canvas
canvas.configure(yscrollcommand=scrollbar.set)

# Criando um frame dentro do Canvas
frame = ctk.CTkFrame(canvas)
canvas.create_window((0, 0), window=frame, anchor="nw")

# Adicionando um evento para redimensionar a rolagem
frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

ctk.CTkLabel(frame, text="📄 Gerador de Documentos", font=("Arial", 16, "bold")).pack(pady=10)

modelo_checkboxes = {}
frame_check = ctk.CTkFrame(frame)
frame_check.pack(pady=10, padx=15, fill="x")  # Margem lateral
for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = ctk.BooleanVar()
    ctk.CTkCheckBox(frame_check, text=os.path.basename(modelo_nome), variable=var).pack(anchor="w", padx=5)

frame_campos = ctk.CTkFrame(frame)
frame_campos.pack(pady=10, padx=15, fill="x")  # Margem lateral

campo_labels = [
    ("Nome Completo", "NOME COMPLETO"), ("RG", "RG"), ("CPF", "CPF"),
    ("Endereço", "ENDEREÇO COMPLETO"), ("Cidade", "CIDADE DE NASCIMENTO"),
    ("Estado", "ESTADO DE NASCIMENTO"), ("Data Nasc.", "DATA DE NASCIMENTO"),
    ("Pai", "NOME PAI"), ("Mãe", "NOME MÃE"), ("Telefone", "TELEFONE"),
    ("Expedição RG", "DATA DE EXPEDIÇÃO DOCUMENTO"), ("Órgão Expedidor", "ÓRGÃO EXPEDIDOR"), ("Estado Civil", "ESTADO CIVIL"),("Profissão", "PROFISSÃO")
]
entradas = {}

for label_texto, chave in campo_labels:
    frame = ctk.CTkFrame(frame_campos)
    frame.pack(fill="x", padx=10, pady=5)
    ctk.CTkLabel(frame, text=label_texto+":").pack(side="left", padx=5)
    entrada = ctk.CTkEntry(frame, width=180)  # Ajustado para ser mais estreito
    entrada.pack(side="right", fill="x" width=250, expand=True)
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

frame_botoes = ctk.CTkFrame(frame)
frame_botoes.pack(pady=10, padx=15, fill="x")  # Margem lateral
ctk.CTkButton(frame_botoes, text="Gerar Selecionados", command=coletar_dados, width=140).pack(side="left", padx=5)
ctk.CTkButton(frame_botoes, text="Gerar Todos", command=selecionar_e_gerar, width=140).pack(side="right", padx=5)

tela.mainloop()
