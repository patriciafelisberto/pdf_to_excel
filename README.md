# Conversor de PDF para Excel e Processador de Excel
Este script Python fornece ferramentas para converter tabelas de arquivos PDF em formato Excel. Além disso, processa um tipo específico de arquivo Excel para rearranjar seus dados de acordo com regras definidas.

## Funcionalidades:
1. **Converter PDFs em Lote para Excel**: O script pode escanear um diretório em busca de todos os arquivos PDF e converter tabelas encontradas nesses PDFs em arquivos Excel individuais.
2. **Processar Arquivo Excel**: Para um formato de arquivo Excel dado, o script extrai informações específicas e as reorganiza de uma forma estruturada.

## Dependências:
- **pandas**: Usado para manipulação de dados.
- **pdfplumber**: Usado para extrair tabelas de arquivos PDF.

## Instalação e Configuração:
Antes de começar, é recomendado criar um ambiente virtual (venv) para o projeto. Isso irá isolar as dependências do projeto e evitar conflitos com outros projetos ou pacotes instalados globalmente em sua máquina.

```shell
python3 -m venv nome-do-venv
```

Em seguida, ative o venv com o seguinte comando:

```shell
source nome-do-venv/bin/activate
```

Instalar dependências.

```shell
pip install -r requirements.txt
```

## Como Usar:

### Converter PDF para Excel:
1. Coloque seus arquivos PDF em um diretório chamado `0_pdf` (ou altere a variável `source_directory` de acordo).
2. Execute o script. Os arquivos Excel serão salvos no diretório `1_excel`.

### Processar um Arquivo Excel Específico:
1. Certifique-se de ter um arquivo Excel no diretório `1_excel` com o formato apropriado.
2. Ajuste a variável `input_filename` para o nome do seu arquivo Excel.
3. Execute o script. Os arquivos Excel processados serão salvos no diretório `2_excel_processed`.

## Estrutura do Código:
1. **pdf_to_excel(pdf_path, excel_path)**: Uma função que converte um único arquivo PDF em um arquivo Excel.
2. **convert_all_pdfs_in_directory(source_directory, destination_directory)**: Uma função que encontra todos os arquivos PDF em um diretório e os converte para o formato Excel.
3. **process_and_save_excel_file(excel_path, output_directory)**: Uma função que processa um formato específico de arquivo Excel e salva a versão processada em um diretório designado.

## Observação:
Foi feito a tratativa de formatação apenas para a planinha do arquivo 'ABCB4 - BCO ABC BRASIL S.A. - Valores Mobiliários Negociados e Detidos 2023-06-01 - prot 1121359', para que fosse possível ter uma ideia de como ficaria um arquivo mais bem estruturado após a conversão para excel.

