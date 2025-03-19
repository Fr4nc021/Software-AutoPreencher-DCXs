import customtkinter as ctk
from tkinter import messagebox
from docx import Document
import os
import sys

# Configuração do tema
ctk.set_appearance_mode("System")  # Alterna entre modo claro/escuro automaticamente
ctk.set_default_color_theme("blue")  # Define o tema azul

# Obtendo o caminho correto, dependendo se está rodando no .exe ou no .py
if getattr(sys, 'frozen', False):  # Se estiver rodando como .exe
    base_path = sys._MEIPASS
else:
    base_path = os.path.abspath(os.path.dirname(__file__))  # Se estiver rodando como .py

# Definindo a pasta onde os modelos estão
pasta_modelos = os.path.join(base_path, "modelos")

# Verifica se a pasta de modelos existe
if not os.path.exists(pasta_modelos):
    os.makedirs(pasta_modelos)  # Cria a pasta caso não exista

# Mapeando os modelos disponíveis
modelos_disponiveis = {
    "1": os.path.join(pasta_modelos, "1.ACERVO.docx"),
    "2": os.path.join(pasta_modelos, "2.DSA.docx"),
    "3": os.path.join(pasta_modelos, "3.IDONEIDADE.docx"),
    "4": os.path.join(pasta_modelos, "4.COMPETIÇÃO.docx"),
    "5": os.path.join(pasta_modelos, "5.Procuração.docx"),
}

# Teste para verificar se os arquivos existem
arquivos_faltando = [modelo for modelo in modelos_disponiveis.values() if not os.path.exists(modelo)]
if arquivos_faltando:
    print(f"Erro: Alguns arquivos de modelos não foram encontrados: {arquivos_faltando}")

# Função para preencher documentos Word
def gerar_documento(dados, modelos_selecionados):
    try:
        pasta_destino = os.path.join(base_path, "Documentos Gerados")
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        for modelo_selecionado in modelos_selecionados:
            modelo_path = modelos_disponiveis.get(modelo_selecionado)
            if not modelo_path or not os.path.exists(modelo_path):
                messagebox.showerror("Erro", f"Modelo {modelo_selecionado} não encontrado!")
                return

            doc = Document(modelo_path)
            for paragrafo in doc.paragraphs:
                for chave, valor in dados.items():
                    placeholder = f"{{{{{chave}}}}}"
                    if placeholder in paragrafo.text:
                        paragrafo.text = paragrafo.text.replace(placeholder, valor)

            caminho_arquivo = os.path.join(pasta_destino, f"{os.path.basename(modelo_path)}")
            doc.save(caminho_arquivo)

        messagebox.showinfo("Sucesso", "Documentos gerados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar documento: {str(e)}")

# Criando a interface
tela = ctk.CTk()
tela.title("Gerador de Documentos")
tela.geometry("500x600")

ctk.CTkLabel(tela, text="📄 Gerador de Documentos", font=("Arial", 18, "bold")).pack(pady=(15, 20))

# Checkbox de seleção de modelos
modelo_checkboxes = {}
frame_check = ctk.CTkFrame(tela)
frame_check.pack(pady=(10, 20))
for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = ctk.BooleanVar()
    check = ctk.CTkCheckBox(frame_check, text=os.path.basename(modelo_nome), variable=var)
    check.pack(anchor="w")
    modelo_checkboxes[modelo_id] = var

# Campos de entrada
frame_campos = ctk.CTkFrame(tela)
frame_campos.pack(pady=(10, 20))  # 10px acima, 20px abaixo
campo_labels = [
    ("Nome Completo", "NOME COMPLETO"), 
    ("RG", "RG"), 
    ("CPF", "CPF"),
    ("Endereço completo", "ENDEREÇO COMPLETO"), 
    ("Cidade de Nascimento", "CIDADE DE NASCIMENTO"),
    ("Estado de nascimento", "ESTADO DE NASCIMENTO"), 
    ("Data de Nascimento", "DATA DE NASCIMENTO"),
    ("Nome do Pai", "NOME PAI"), 
    ("Nome da Mãe", "NOME MÃE"), 
    ("Telefone", "TELEFONE"),
    ("Data de expedição do RG", "DATA DE EXPEDIÇÃO RG"),
    ("Órgão expedidor do RG", "ÓRGÃO EXPEDIDOR RG"),
]
entradas = {}
for label_texto, chave in campo_labels:
    frame = ctk.CTkFrame(frame_campos)
    frame.pack(fill="x", padx=15, pady=20)
    ctk.CTkLabel(frame, text=label_texto + ":").pack(side="left")
    entrada = ctk.CTkEntry(frame, width=250)
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
frame_botoes = ctk.CTkFrame(tela)
frame_botoes.pack(pady=20)
ctk.CTkButton(frame_botoes, text="Gerar Selecionados", command=coletar_dados, width=200).pack(side="left", padx=5)
ctk.CTkButton(frame_botoes, text="Gerar Todos", command=selecionar_e_gerar, width=200).pack(side="right", padx=5)

# Executando o app
tela.mainloop()
