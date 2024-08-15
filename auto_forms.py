import pandas as pd
from docxtpl import DocxTemplate
from pathlib import Path
from docx2pdf import convert

# Função para formatar o CPF
def formatar_cpf(cpf):
    cpf_str = str(cpf).zfill(11)  # Converter para string e garantir que tenha 11 dígitos com zeros à esquerda
    return f"{cpf_str[:3]}.{cpf_str[3:6]}.{cpf_str[6:9]}-{cpf_str[9:]}"



# Ler a planilha
planilha = pd.read_excel("Teste.xlsx")

# Dicionário de templates
templates = {
    'RH NOSSA': Path("C:/Users/mathe/Downloads/Auto/FORM_NOSSA.docx"),
    'CWBEM': Path("C:/Users/mathe/Downloads/Auto/FORM_CWBEM.docx"),
    'FOUR HANDS': Path("C:/Users/mathe/Downloads/Auto/FORM_FOUR_HANDS.docx")
}


# Caminho da pasta onde os formulários serão salvos
pasta_caminho_docx = Path("C:/Users/mathe/Downloads/FORMULARIOS/TESTES")
pasta_caminho_pdf = Path("C:/Users/mathe/Downloads/FORMULARIOS/PDFs")

# Iterar sobre as linhas da planilha
for index, row in planilha.iterrows():
    empresa = row.get('EMPRESA')
    nome = row.get('NOME')
    cpf = row.get('CPF')
    identificador = row.get('NOMEDOC')

    # Verificar se todas as colunas necessárias têm valores
    if pd.notna(empresa) and pd.notna(nome) and pd.notna(cpf) and pd.notna(identificador):
        cpf_formatado = formatar_cpf(cpf)

        # Verificar se o template existe para a empresa
        template_path = templates.get(empresa)
        if template_path and Path(template_path).exists():
            try:
                template = DocxTemplate(template_path)

                context = {
                    'nome': nome,
                    'cpf': cpf_formatado
                }

                template.render(context)
                caminho_arq = pasta_caminho_docx / f'{identificador}.docx'

                # Salvar o documento
                template.save(caminho_arq)
                print(f"Documento salvo: {caminho_arq}")
                convert(pasta_caminho_docx, pasta_caminho_pdf)

            except Exception as e:
                print(f"Erro ao processar o template para a empresa {empresa}: {e}")
        else:
            print(f"Template não encontrado para a empresa: {empresa}")
    else:
        print(f"Dados ausentes na linha {index}")

print("Geração de formulários concluída.")
