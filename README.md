# Automa-o-PDF
Automaçã# 📄 Extração de Dados de PDFs para Excel

# DESCRIÇÃO:
### 🧠 Do que se trata?
Este script automatiza a leitura de arquivos PDF em uma pasta específica, extrai informações como número da fatura e data, e salva esses dados em uma planilha Excel. É útil para organizar documentos financeiros ou administrativos de forma estruturada.

### 🏗️ Como foi construído?
Utiliza a biblioteca pdfplumber para leitura de PDFs.
Usa openpyxl para criação e manipulação de arquivos Excel.
Emprega expressões regulares (re) para extrair padrões específicos de texto.
Organiza os dados em colunas: Invoice #, Date, File Name, Status.

### 🧰 Com o que foi construído?
os: manipulação de diretórios e arquivos.
pdfplumber: leitura de conteúdo textual de PDFs.
re: busca de padrões com expressões regulares.
openpyxl: criação e edição de planilhas Excel.
datetime: geração de timestamp para nomear arquivos.

### ❓ Por que foi construído?
Para automatizar a coleta de dados de documentos PDF e facilitar a análise e organização em planilhas, economizando tempo e reduzindo erros manuais.

### 🔁 O que ele recebe e o que ele retorna?
Entrada:
Arquivos PDF localizados na pasta Pdf_pasta.
Saída:
Um arquivo Excel com os dados extraídos, nomeado com a data e hora da execução.

:

### 📊 Estrutura de Dados Utilizada:

1. Lista de Arquivos
Tipo: list
Origem: os.listdir(diretory)
Finalidade: Armazena os nomes dos arquivos PDF encontrados na pasta Pdf_pasta.
2. Planilha Excel (Workbook e Worksheet)
Tipo: openpyxl.Workbook e Worksheet
Finalidade: Armazena os dados extraídos dos PDFs em formato tabular.
Estrutura das colunas:
A: Invoice Number (INVOICE #)
B: Date (DATE)
C: Nome do arquivo PDF
D: Status da leitura (ex: "Finalizado" ou "Exception")
3. Strings com Expressões Regulares
Tipo: str + re.Match
Finalidade: Captura padrões específicos no texto dos PDFs:
Número da fatura: INVOICE #(\d+)
Data da fatura: DATE: (\d{2}/\d{2}/\d{4})
4. Contador de Linha
Variável: last_empty_line
Tipo: int
Finalidade: Controla a próxima linha disponível na planilha para inserção de dados.



.md = markdown

# Instrução de instalação:

### ⚙️ Pré-requisitos

Instale as bibliotecas necessárias com:


pip install pdfplumber openpyxl

### 🚀 Instruções de Uso
Crie uma pasta chamada Pdf_pasta e coloque os PDFs nela.
Execute o script com Python:

python seu_script.py





# Instrução de Uso:
DETALHES TECNICOS = DOCUMENTAÇÃO
DETALHES COMO EXECUTAR
DO QUE SE TRATA 
COMO CONSTRIBUIR o PDF
