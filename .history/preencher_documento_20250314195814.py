import tkinter as tk
from tkinter import ttk, messagebox
from docx import Document
import os

# Obtém o diretório onde o script está sendo executado
pasta_modelos = os.path.dirname(__file__)

# Lista de modelos disponíveis com caminhos corretos
modelos_disponiveis = {
    "1": os.path.join(pasta_modelos, "1.ACERVO.docx"),
    "2": os.path.join(pasta_modelos, "2.DSA.docx"),
    "3": os.path.join(pasta_modelos, "3.IDONEIDADE.docx"),
    "4": os.path.join(pasta_modelos, "4.COMPETIÇÃO.docx"),
    "5": os.path.join(pasta_modelos, "5.Procuração.docx"),
}

# Função para preencher o documento Word com os dados
def gerar_documento(dados, modelos_selecionados):
    try:
        pasta_destino = "Documentos Gerados"
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

         for modelo_id, modelo_path in modelos_disponiveis.items():
    var = tk.BooleanVar()
    nome_modelo = os.path.basename(modelo_path)  # Exibir apenas o nome do arquivo
    check = tk.Checkbutton(checkbox_frame, text=nome_modelo, variable=var, bg="#f4f4f4")
    check.pack(anchor="w")
    modelo_checkboxes[modelo_id] = var
            caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
            doc.save(caminho_arquivo)

        messagebox.showinfo("Sucesso", "Documentos gerados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar documento: {str(e)}")

# Função para coletar os dados
def coletar_dados():
    dados = {chave: entrada.get() for chave, entrada in entradas.items()}
    modelos_selecionados = [modelo for modelo, selecionado in modelo_checkboxes.items() if selecionado.get()]
    if modelos_selecionados:
        gerar_documento(dados, modelos_selecionados)
    else:
        messagebox.showerror("Erro", "Selecione pelo menos um modelo!")

# Função para selecionar todos os modelos
def selecionar_todos():
    for modelo in modelo_checkboxes.values():
        modelo.set(True)

def selecionar_e_gerar():
    selecionar_todos()
    coletar_dados()

# Criando interface gráfica
root = tk.Tk()
root.title("Gerador de Documentos")
root.geometry("500x550")
root.configure(bg="#f4f4f4")

style = ttk.Style()
style.configure("TButton", font=("Arial", 12), padding=5)
style.configure("TLabel", font=("Arial", 11), background="#f4f4f4")
style.configure("TCombobox", font=("Arial", 11))

label_titulo = tk.Label(root, text="📄 Gerador de Documentos", font=("Arial", 16, "bold"), bg="#f4f4f4")
label_titulo.pack(pady=10)

checkbox_frame = tk.Frame(root, bg="#f4f4f4")
checkbox_frame.pack(pady=5)
modelo_checkboxes = {}
for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = tk.BooleanVar()
    check = tk.Checkbutton(checkbox_frame, text=modelo_nome, variable=var, bg="#f4f4f4")
    check.pack(anchor="w")
    modelo_checkboxes[modelo_id] = var

frame_campos = tk.Frame(root, bg="#f4f4f4")
frame_campos.pack(pady=5)
campos = [
    ("Nome Completo", "NOME COMPLETO"), 
    ("RG", "RG"), 
    ("CPF", "CPF"),
    ("Endereço Completo", "ENDEREÇO COMPLETO"), 
    ("Cidade de Nascimento", "CIDADE DE NASCIMENTO"),
    ("Estado de Nascimento", "ESTADO DE NASCIMENTO"), 
    ("Data de Nascimento", "DATA DE NASCIMENTO"),
    ("Nome do Pai", "NOME PAI"), 
    ("Nome da Mãe", "NOME MÃE"), 
    ("Data de Expedição", "DATA DE EXPEDIÇÃO DOCUMENTO"),
    ("Órgão Expedidor", "ÓRGÃO EXPEDIDOR"), 
    ("Telefone", "TELEFONE"),
    ("Estado Civil", "ESTADO CIVIL"),
    ("Profissão", "PROFISSÃO"),
]
entradas = {}
for label_texto, chave in campos:
    frame = tk.Frame(frame_campos, bg="#f4f4f4")
    frame.pack(fill="x", padx=10, pady=2)
    tk.Label(frame, text=label_texto + ":", bg="#f4f4f4").pack(side="left")
    entrada = tk.Entry(frame)
    entrada.pack(side="right", fill="x", expand=True)
    entradas[chave] = entrada

btn_frame = tk.Frame(root, bg="#f4f4f4")
btn_frame.pack(pady=10)

ttk.Button(btn_frame, text="Gerar Documentos Selecionados", command=coletar_dados).pack(side="left", padx=5)
ttk.Button(btn_frame, text="Gerar Todos os Documentos", command=selecionar_e_gerar).pack(side="right", padx=5)

root.mainloop()
