import tkinter as tk
from tkinter import messagebox
from docx import Document
import os

# Lista de modelos disponíveis com caminhos de arquivo relativos
modelos_disponiveis = {
    "1": "1.ACERVO.docx",
    "2": "2.DSA.docx",
    "3": "3. IDONEIDADE.docx",
    "4": "4. DECLARAÇÃO DE COMPETIÇÃO.docx",   
}

# Função para preencher o documento Word com os dados
def gerar_documento(dados, modelos_selecionados):
    try:
        pasta_destino = "documentos"
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

            # Salvar o documento preenchido em uma pasta "documentos"
            nome_arquivo = f"{modelo_selecionado}_{dados['NOME COMPLETO'].replace(' ', '_')}.docx"
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
    }

    # Verifica se todos os campos obrigatórios estão preenchidos
    for chave, valor in dados.items():
        if not valor.strip():
            messagebox.showwarning("Campos incompletos", f"O campo {chave} deve ser preenchido!")
            return

    # Pega os modelos selecionados nos checkboxes
    modelos_selecionados = [modelo for modelo, var in checkboxes.items() if var.get()]

    if not modelos_selecionados:
        messagebox.showwarning("Seleção de modelo", "Selecione pelo menos um modelo para gerar!")
        return

    gerar_documento(dados, modelos_selecionados)

# Criar interface gráfica com Tkinter
root = tk.Tk()
root.title("Gerador de Documentos")

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
]

# Dicionário para armazenar os campos de entrada
entradas = {}

# Criar os labels e campos de entrada dinamicamente
for i, (label_texto, chave) in enumerate(campos):
    label = tk.Label(root, text=label_texto + ":")
    label.grid(row=i, column=0, sticky="w", padx=10, pady=5)
    entry = tk.Entry(root)
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

# Criar menu suspenso para selecionar o modelo
label_modelo = tk.Label(root, text="Escolha os modelos:")
label_modelo.grid(row=len(campos), column=0, padx=10, pady=5)

# Criar checkboxes para seleção dos documentos a serem gerados
checkboxes = {}
for i, (modelo, nome_arquivo) in enumerate(modelos_disponiveis.items()):
    var = tk.BooleanVar()
    checkboxes[modelo] = var
    checkbox = tk.Checkbutton(root, text=f"{modelo} - {nome_arquivo}", variable=var)
    checkbox.grid(row=len(campos) + i + 1, column=0, columnspan=2, padx=10, pady=5)

# Criar botões
btn_gerar = tk.Button(root, text="Gerar Documento", command=coletar_dados)
btn_gerar.grid(row=len(campos) + len(modelos_disponiveis) + 1, column=0, columnspan=2, pady=10)

# Iniciar interface
root.mainloop()
