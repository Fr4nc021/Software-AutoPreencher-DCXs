from docx import Document

# Função para preencher o documento
def preencher_documento(modelo_path, dados, output_path):
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

# Chamar a função passando o modelo e o nome do arquivo de saída
preencher_documento("MODELO.DOCX", dados_cliente, "Documento_Preenchido.docx")