import customtkinter as ctk
from tkinter import messagebox
from docx import Document
import os
import sys

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

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

# Tamanho da tela (ajustado para evitar que ultrapasse a tela)
screen_width = tela.winfo_screenwidth()
screen_height = tela.winfo_screenheight()

# Ajuste o tamanho máximo da janela para 80% da tela
max_width = int(screen_width * 0.8)
max_height = int(screen_height * 0.8)

tela.geometry(f"500x600")  # Tamanho inicial
tela.minsize(400, 500)  # Tamanho mínimo
tela.maxsize(max_width, max_height)  # Tamanho máximo

ctk.CTkLabel(tela, text="📄 Gerador de Documentos", font=("Arial", 18, "bold")).pack(pady=10)

modelo_checkboxes = {}
frame_check = ctk.CTkFrame(tela)
frame_check.pack(pady=10)
for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = ctk.BooleanVar()
    ctk.CTkCheckBox(frame_check, text=os.path.basename(modelo_nome), variable=var).pack(anchor="w")
    modelo_checkboxes[modelo_id] = var

frame_campos = ctk.CTkFrame(tela)
frame_campos.pack(pady=10)

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
    ctk.CTkLabel(frame, text=label_texto+":").pack(side="left")
    entrada = ctk.CTkEntry(frame, width=250)
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

frame_botoes = ctk.CTkFrame(tela)
frame_botoes.pack(pady=10)
ctk.CTkButton(frame_botoes, text="Gerar Selecionados", command=coletar_dados, width=200).pack(side="left", padx=5)
ctk.CTkButton(frame_botoes, text="Gerar Todos", command=selecionar_e_gerar, width=200).pack(side="right", padx=5)

tela.mainloop()
