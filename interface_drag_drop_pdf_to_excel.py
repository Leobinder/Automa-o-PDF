from flask import Flask, request, render_template_string, redirect, url_for
import os
import pdfplumber
import re
from openpyxl import Workbook
from datetime import datetime

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
PDF_FOLDER = 'Pdf_pasta'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

extracted_data = []

HTML_TEMPLATE = '''
<!doctype html>
<html lang="pt">
<head>
    <meta charset="utf-8">
    <title>Armazenar Arquivos</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #eef2f3;
        }
        header {
            background-color: #2c3e50;
            color: white;
            padding: 20px;
            font-size: 24px;
            display: flex;
            align-items: center;
        }
        header a {
            color: white;
            text-decoration: none;
            margin-right: 20px;
            font-size: 20px;
        }
        .container {
            max-width: 900px;
            margin: 40px auto;
            background-color: white;
            padding: 30px;
            border-radius: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        h2 {
            color: #34495e;
            margin-bottom: 20px;
        }
        input[type=file], input[type=submit], button {
            padding: 10px;
            font-size: 16px;
            margin-top: 10px;
            border-radius: 6px;
            border: 1px solid #ccc;
        }
        input[type=submit], button {
            background-color: #3498db;
            color: white;
            border: none;
            cursor: pointer;
        }
        input[type=submit]:hover, button:hover {
            background-color: #2980b9;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 30px;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }
        th {
            background-color: #f9f9f9;
            color: #2c3e50;
        }
        .buttons {
            margin-top: 20px;
        }
        .buttons form {
            display: inline-block;
            margin-right: 10px;
        }
        .success {
            color: green;
            font-weight: bold;
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <header>
        <a href="/">←</a> ARMAZENAR ARQUIVOS
    </header>
    <div class="container">
        <h2>Envio de Arquivos</h2>
        <form method="post" enctype="multipart/form-data">
            <input type="file" name="pdfs" multiple required>
            <input type="submit" value="Upload">
        </form>

        <form method="post">
            <button type="submit" name="processar_pasta" value="true">Processar PDFs da Pasta</button>
        </form>

        {% if data %}
        <table>
            <tr><th>Invoice #</th><th>Date</th><th>File Name</th><th>Status</th></tr>
            {% for row in data %}
            <tr>
                <td>{{ row[0] }}</td>
                <td>{{ row[1] }}</td>
                <td>{{ row[2] }}</td>
                <td>{{ row[3] }}</td>
            </tr>
            {% endfor %}
        </table>
        <div class="buttons">
            <form action="/confirm" method="post">
                <button type="submit" name="confirm" value="yes">Deseja mesmo enviar essas informações?</button>
            </form>
            <form action="/" method="get">
                <button type="submit">Não</button>
            </form>
        </div>
        {% endif %}

        {% if success %}
        <p class="success">✅ Dados enviados com sucesso!</p>
        {% endif %}
    </div>
</body>
</html>
'''


def extrair_dados_dos_pdfs(pasta=PDF_FOLDER):
    dados_extraidos = []
    wb = Workbook()
    ws = wb.active
    ws.title = 'DADOS-Pdf'
    ws.append(['Invoice #', 'Date', 'File Name', 'Status'])

    arquivos = os.listdir(pasta)
    if not arquivos:
        raise Exception("Não existem arquivos na pasta.")

    for file in arquivos:
        caminho = os.path.join(pasta, file)
        try:
            with pdfplumber.open(caminho) as pdf:
                texto = pdf.pages[0].extract_text()

            inv_number_re_pattern = r'INVOICE #(\d+)'
            inv_date_re_pattern = r'DATE: (\d{2}/\d{2}/\d{4})'

            numero = re.search(inv_number_re_pattern, texto)
            data = re.search(inv_date_re_pattern, texto)

            invoice_number = numero.group(1) if numero else 'Não encontrado'
            invoice_date = data.group(1) if data else 'Não encontrado'

            dados_extraidos.append([invoice_number, invoice_date, file, 'Finalizado'])
            ws.append([invoice_number, invoice_date, file, 'Finalizado'])

        except Exception as e:
            dados_extraidos.append(['Erro', 'Erro', file, f'Erro: {str(e)}'])
            ws.append(['Erro', 'Erro', file, f'Erro: {str(e)}'])

    now = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    excel_path = os.path.join(UPLOAD_FOLDER, f'Invoice_{now}.xlsx')
    wb.save(excel_path)

    return dados_extraidos

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    global extracted_data
    success = False

    if request.method == 'POST':
        if 'processar_pasta' in request.form:
            try:
                extracted_data = extrair_dados_dos_pdfs()
            except Exception as e:
                extracted_data = [['Erro', 'Erro', 'Nenhum arquivo', f'Erro: {str(e)}']]
        else:
            uploaded_files = request.files.getlist('pdfs')
            extracted_data = []

            wb = Workbook()
            ws = wb.active
            ws.title = 'DADOS-Pdf'
            ws.append(['Invoice #', 'Date', 'File Name', 'Status'])

            for file in uploaded_files:
                filename = file.filename
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                file.save(file_path)

                try:
                    with pdfplumber.open(file_path) as pdf:
                        page = pdf.pages[0]
                        text = page.extract_text()

                    inv_number_re_pattern = r'INVOICE #(\d+)'
                    inv_date_re_pattern = r'DATE: (\d{2}/\d{2}/\d{4})'

                    match_number = re.search(inv_number_re_pattern, text)
                    match_date = re.search(inv_date_re_pattern, text)

                    invoice_number = match_number.group(1) if match_number else 'Não encontrado'
                    invoice_date = match_date.group(1) if match_date else 'Não encontrado'

                    extracted_data.append([invoice_number, invoice_date, filename, 'Finalizado'])
                    ws.append([invoice_number, invoice_date, filename, 'Finalizado'])

                except Exception as e:
                    extracted_data.append(['Erro', 'Erro', filename, f'Erro: {str(e)}'])
                    ws.append(['Erro', 'Erro', filename, f'Erro: {str(e)}'])

            now = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            excel_filename = f'Invoice_{now}.xlsx'
            excel_path = os.path.join(UPLOAD_FOLDER, excel_filename)
            wb.save(excel_path)

    return render_template_string(HTML_TEMPLATE, data=extracted_data, success=success)

@app.route('/confirm', methods=['POST'])
def confirm():
    global extracted_data
    if request.form.get('confirm') == 'yes':
        extracted_data = []
        return render_template_string(HTML_TEMPLATE, data=[], success=True)
    return redirect(url_for('upload_files'))

if __name__ == '__main__':
    app.run(debug=True)

