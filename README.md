# Gerador de Formulários Automatizado

Este script automatiza a geração de formulários personalizados(nesse caso, CPF e Nome) a partir de um template do Word (.docx) e de dados contidos em uma planilha Excel.

## Funcionalidades

- **Leitura de Dados**: Lê uma planilha Excel com dados de empresas, nomes, CPFs, e identificadores.
- **Formatação de CPF**: Formata os CPFs no padrão `000.000.000-00`.
- **Geração de Formulários**: Para cada linha da planilha, gera um documento Word (.docx) com base em um template específico para a empresa.
- **Armazenamento**: Salva os formulários gerados em uma pasta designada.

```bash
python auto_forms.py
