import tkinter as tk
from tkinter import messagebox
from docx import Document
import os

# Lista de modelos disponíveis com caminhos de arquivo relativos
modelos_disponiveis = {
    "1": "1.ACERVO.docx",
    "2": "2.DSA.docx",
    "3": "3.IDONEIDADE.docx",
    "4": "4.COMPETIÇÃO.docx",   
}

# Função para preencher o documento Word com os dados
def gerar_documento(dados, modelos_selecionados):
    try:
        pasta_destino = "Documentos Gerados"
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        for modelo_selecionado in modelos_selecionados:
            modelo_path = modelos_disponiveis.get(modelo_selecionado)  # Obtém o caminho correto do modelo

            if not modelo_path or not os.path.exists(modelo_path):
                messagebox.showerror("Erro", f"Modelo {modelo_selecionado} não encontrado!")
                return

            doc = Document(modelo_path)

            # Substituir os placeholders pelos dados inseridos
            for paragrafo in doc.paragraphs:
                for chave, valor in dados.items():
                    placeholder = f"{{{{{chave}}}}}"  # Formato {{CHAVE}}
                    if placeholder in paragrafo.text:
                        paragrafo.text = paragrafo.text.replace(placeholder, valor)

            # Salvar o documento com o nome do próprio modelo
            nome_arquivo = f"{os.path.basename(modelo_path)}"  # Usa o nome do arquivo do modelo
            caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
            doc.save(caminho_arquivo)

        messagebox.showinfo("Sucesso", "Documentos gerados com sucesso!")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao gerar documento: {str(e)}")

# Função para coletar os dados e chamar a função de geração
def coletar_dados():
    # Coleta de dados do formulário
    dados = {
        "NOME COMPLETO": entry_nome.get(),
        "RG": entry_rg.get(),
        "CPF": entry_cpf.get(),
        "ENDEREÇO COMPLETO": entry_endereco.get(),
        "CIDADE DE NASCIMENTO": entry_cidade.get(),
        "ESTADO DE NASCIMENTO": entry_estado.get(),
        "DATA DE NASCIMENTO": entry_data.get(),
        "NOME PAI": entry_pai.get(),
        "NOME MÃE": entry_mae.get(),
        "DATA DE EXPEDIÇÃO DOCUMENTO": entry_dataexp.get(),
        "ÓRGÃO EXPEDIDOR": entry_orgao.get(),
        "TELEFONE": entry_telefone.get(),
    }

    # Verifica se pelo menos um modelo foi selecionado
    modelos_selecionados = [modelo for modelo, selecionado in modelo_checkboxes.items() if selecionado.get()]
    
    # Se nenhum modelo for selecionado, mostra um erro
    if modelos_selecionados:
        gerar_documento(dados, modelos_selecionados)
    else:
        messagebox.showerror("Erro", "Selecione pelo menos um modelo!")

# Função para selecionar todos os checkboxes (gerar todos os documentos)
def selecionar_todos():
    for modelo in modelo_checkboxes.values():
        modelo.set(True)  # Marca todos os checkboxes

# Criar interface gráfica com Tkinter
root = tk.Tk()
root.title("Gerador de Documentos")
root.geometry("500x500") #tamanho das janelas
root.configure(bg="#f4f4f4") #cor de fundo

#Estilos para deixar a interface mais moderna
style = tk.ttk.Style()
style.configure("TButton", font=("Arial", 12), padding=5, width=20)
style.configure("TLabel", font=("Arial", 11),background="#f4f4f4")
style.configure("TCombobox", font=("Arial", 11))

# Ajustar o layout para acomodar todos os elementos de forma organizada
root.grid_rowconfigure(0, weight=1)
root.grid_rowconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_columnconfigure(0, weight=1, minsize=150)

# Criando checkboxes para selecionar quais modelos gerar
modelo_checkboxes = {}
checkbox_frame = tk.Frame(root)
checkbox_frame.grid(row=0, column=0, sticky="w", padx=10, pady=5)

for modelo_id, modelo_nome in modelos_disponiveis.items():
    var = tk.BooleanVar()
    check = tk.Checkbutton(checkbox_frame, text=modelo_nome, variable=var)
    check.grid(row=int(modelo_id)-1, column=0, sticky="w", padx=10, pady=5)
    modelo_checkboxes[modelo_id] = var

# Criando os campos de entrada
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
    ("Data de Expedição do Documento", "DATA DE EXPEDIÇÃO DOCUMENTO"),
    ("Órgão Expedidor", "ÓRGÃO EXPEDIDOR"),
    ("Telefone", "TELEFONE"),
]

# Dicionário para armazenar os campos de entrada
entradas = {}

# Criar os labels e campos de entrada dinamicamente
entrada_frame = tk.Frame(root)
entrada_frame.grid(row=1, column=0, sticky="w", padx=10, pady=5)

for i, (label_texto, chave) in enumerate(campos):
    label = tk.Label(entrada_frame, text=label_texto + ":")
    label.grid(row=i, column=0, sticky="w", padx=10, pady=5)
    entry = tk.Entry(entrada_frame)
    entry.grid(row=i, column=1, padx=10, pady=5)
    entradas[chave] = entry

# Mapear entradas para variáveis globais
entry_nome = entradas["NOME COMPLETO"]
entry_rg = entradas["RG"]
entry_cpf = entradas["CPF"]
entry_endereco = entradas["ENDEREÇO COMPLETO"]
entry_cidade = entradas["CIDADE DE NASCIMENTO"]
entry_estado = entradas["ESTADO DE NASCIMENTO"]
entry_data = entradas["DATA DE NASCIMENTO"]
entry_pai = entradas["NOME PAI"]
entry_mae = entradas["NOME MÃE"]
entry_dataexp = entradas["DATA DE EXPEDIÇÃO DOCUMENTO"]
entry_orgao = entradas["ÓRGÃO EXPEDIDOR"]
entry_telefone = entradas["TELEFONE"]

# Criar botões
btn_frame = tk.Frame(root)
btn_frame.grid(row=2, column=0, sticky="w", padx=10, pady=5)

btn_gerar = tk.Button(btn_frame, text="Gerar Documentos Selecionados", command=coletar_dados)
btn_gerar.grid(row=0, column=0, pady=10)

btn_gerar_todos = tk.Button(btn_frame, text="Gerar Todos os Documentos", command=selecionar_todos)
btn_gerar_todos.grid(row=0, column=1, padx=10, pady=10)

# Iniciar interface
root.mainloop()
