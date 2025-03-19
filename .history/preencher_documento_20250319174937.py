import customtkinter as ctk
from tkinter import messagebox
from docx import Document
import os
import sys

# Configuração geral
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

if getattr(sys, 'frozen', False):
    pasta_modelos = os.path.join(sys._MEIPASS, "modelos")
else:
    pasta_modelos = os.path.dirname(_file_)

modelos_disponiveis = {
    "1": os.path.join(pasta_modelos, "1.ACERVO.docx"),
    "2": os.path.join(pasta_modelos, "2.DSA.docx"),
    "3": os.path.join(pasta_modelos, "3.IDONEIDADE.docx"),
    "4": os.path.join(pasta_modelos, "4.COMPETIÇÃO.docx"),
    "5": os.path.join(pasta_modelos, "5.Procuração.docx"),
}

# Função para gerar documentos
def gerar_documento(dados, modelos_selecionados):
    try:
        os.makedirs("Documentos Gerados", exist_ok=True)
        
        for modelo in modelos_selecionados:
            modelo_path = modelos_disponiveis.get(modelo)
            if not modelo_path or not os.path.exists(modelo_path):
                messagebox.showerror("Erro", f"Modelo {modelo} não encontrado!")
                return
            
            doc = Document(modelo_path)
            for paragrafo in doc.paragraphs:
                for run in paragrafo.runs:
                    for chave, valor in dados.items():
                        if f"{{{{{chave}}}}}" in run.text:
                            run.text = run.text.replace(f"{{{{{chave}}}}}", valor)
            
            doc.save(os.path.join("Documentos Gerados", os.path.basename(modelo_path)))
        
        messagebox.showinfo("Sucesso", "Documentos gerados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar documento: {str(e)}")

# Interface Principal
tela = ctk.CTk()
tela.title("Gerador de Documentos")
tela.geometry("800x600")  # Tamanho maior para acomodar melhor
tela.configure(bg="black")

# Frame principal centralizado
frame_geral = ctk.CTkFrame(tela, fg_color="black")
frame_geral.pack(expand=True, pady=10, padx=20)

ctk.CTkLabel(frame_geral, text="📄 Gerador de Documentos", font=("Arial", 18, "bold"), text_color="white").pack(pady=10)

# Frame para checkboxes
frame_check = ctk.CTkFrame(frame_geral, fg_color="black")
frame_check.pack(pady=10, padx=15, fill="x")

modelo_checkboxes = {}
for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = ctk.BooleanVar()
    modelo_checkboxes[modelo_id] = var
    ctk.CTkCheckBox(frame_check, text=os.path.basename(modelo_nome), variable=var, text_color="white", fg_color="black").pack(anchor="w", padx=5)

# Frame para os campos de entrada em duas colunas
frame_campos = ctk.CTkFrame(frame_geral, fg_color="black")
frame_campos.pack(pady=10, padx=15, fill="x")

campo_labels = [
    ("Nome Completo", "NOME COMPLETO"), ("RG", "RG"), ("CPF", "CPF"),
    ("Endereço", "ENDEREÇO COMPLETO"), ("Cidade", "CIDADE DE NASCIMENTO"),
    ("Estado", "ESTADO DE NASCIMENTO"), ("Data Nasc.", "DATA DE NASCIMENTO"),
    ("Pai", "NOME PAI"), ("Mãe", "NOME MÃE"), ("Telefone", "TELEFONE"),
    ("Expedição RG", "DATA DE EXPEDIÇÃO DOCUMENTO"), ("Órgão Expedidor", "ÓRGÃO EXPEDIDOR"), 
    ("Estado Civil", "ESTADO CIVIL"), ("Profissão", "PROFISSÃO")
]

entradas = {}

# Distribuir em duas colunas
for i, (label_texto, chave) in enumerate(campo_labels):
    row, col = divmod(i, 2)  # Divide os campos em 2 colunas
    label = ctk.CTkLabel(frame_campos, text=label_texto+":", text_color="white")
    label.grid(row=row, column=col*2, padx=5, pady=3, sticky="w")
    entrada = ctk.CTkEntry(frame_campos, width=180)
    entrada.grid(row=row, column=col*2+1, padx=5, pady=3, sticky="ew")
    entradas[chave] = entrada

frame_campos.columnconfigure(1, weight=1)
frame_campos.columnconfigure(3, weight=1)

# Função para coletar dados
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

# Botões
frame_botoes = ctk.CTkFrame(frame_geral, fg_color="black")
frame_botoes.pack(pady=10, padx=15, fill="x")

ctk.CTkButton(frame_botoes, text="Gerar Selecionados", command=coletar_dados, width=140).pack(side="left", padx=5)
ctk.CTkButton(frame_botoes, text="Gerar Todos", command=selecionar_e_gerar, width=140).pack(side="right", padx=5)

tela.mainloop()