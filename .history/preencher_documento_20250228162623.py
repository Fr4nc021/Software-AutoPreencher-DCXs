import os
from docx import Document

# Caminho do arquivo de modelo
modelo_path = os.path.join(r"C:\Users\franc\OneDrive\Área de Trabalho\Software Docs", "modelo.docx")

# Função para preencher o documento
def preencher_documento(modelo_path, dados, output_path):
    if not os.path.exists(modelo_path):
        print(f"Erro: O arquivo '{modelo_path}' não foi encontrado.")
        return

    # Carrega o modelo do documento
    doc = Document(modelo_path)

    # Substitui as chaves pelos valores reais
    for paragrafo in doc.paragraphs:
        for chave, valor in dados.items():
            paragrafo.text = paragrafo.text.replace(f'{{{{{chave}}}}}', valor)

    # Salva o documento preenchido
    doc.save(output_path)
    print(f"Documento gerado com sucesso: {output_path}")

# Dados do cliente (exemplo)
dados_cliente = {
    "NOME COMPLETO": "João da Silva",
    "RG": "123456789",
    "CPF": "987.654.321-00",
    "ENDEREÇO COMPLETO": "Rua das Flores, 123 - Bairro Centro, São Paulo - SP"
}

# Caminho de saída
output_path = os.path.join(r"C:\Users\franc\OneDrive\Área de Trabalho\Software Docs", "Documento_Preenchido.docx")

# Executa a função
preencher_documento(modelo_path, dados_cliente, output_path)