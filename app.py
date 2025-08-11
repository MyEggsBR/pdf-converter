#!/usr/bin/env python3
import subprocess
import sys
import os

# Instalar dependências automaticamente
def install_packages():
    packages = [
        'Flask==2.3.3',
        'pdfplumber==0.10.3', 
        'pandas==2.0.3',
        'openpyxl==3.1.2'
    ]
    for package in packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package], 
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except:
            pass

# Tentar importar, se falhar, instalar
try:
    import flask
    import pdfplumber
    import pandas as pd
except ImportError:
    print("Instalando dependências...")
    install_packages()
    import flask
    import pdfplumber
    import pandas as pd

import re
import uuid
from flask import Flask, request, send_file

app = Flask(__name__)

# Configuração
app.config["UPLOAD_FOLDER"] = "/tmp"
app.config["OUTPUT_FOLDER"] = "/tmp"

# Regex para detectar cliente (ajustada para capturar telefone)
regex_cliente = re.compile(r"^(\d+)\s+(.+?)\s+\((\d+)\)(\d{4,5}-\d{4})\s+(.+)$")

# Regex para detectar título
regex_titulo = re.compile(
    r"^(\d+\.\d)\s+(\d{2}/\d{2}/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+(\d+)\s+(\w+)\s+([\w/-]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)$"
)

def processar_pdf(pdf_path, output_path):
    dados = []
    cliente_atual = None
    cidade_atual = None
    telefone_atual = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            linhas = page.extract_text().split("\n")
            for linha in linhas:
                match_cliente = regex_cliente.match(linha)
                if match_cliente:
                    codigo_cliente = match_cliente.group(1).strip()
                    cliente_atual = match_cliente.group(2).strip()
                    ddd = match_cliente.group(3).strip()
                    numero = match_cliente.group(4).strip()
                    telefone_atual = f"({ddd}){numero}"
                    cidade_atual = match_cliente.group(5).strip()
                    continue

                match_titulo = regex_titulo.match(linha)
                if match_titulo and cliente_atual:
                    documento = match_titulo.group(1)
                    emissao = match_titulo.group(2)
                    vencimento = match_titulo.group(3)
                    ats = match_titulo.group(4)
                    tipo = match_titulo.group(5)
                    boleto = match_titulo.group(6)
                    valor_doc = match_titulo.group(7)
                    juros = match_titulo.group(8)
                    multa = match_titulo.group(9)
                    tarifa = match_titulo.group(10)
                    valor_total = match_titulo.group(11)

                    dados.append([
                        cliente_atual, telefone_atual, cidade_atual, documento, emissao, vencimento, ats,
                        tipo, boleto, valor_doc, juros, multa, tarifa, valor_total
                    ])

    colunas = [
        "Cliente", "Telefone", "Cidade", "Documento", "Emissão", "Vencimento", "ATS",
        "Tipo", "Boleto", "Valor Documento", "Juros", "Multa", "Tarifa", "Valor Total"
    ]
    df = pd.DataFrame(dados, columns=colunas)

    # Ajusta valores numéricos
    for col in ["Valor Documento", "Juros", "Multa", "Tarifa", "Valor Total"]:
        df[col] = df[col].str.replace(".", "", regex=False).str.replace(",", ".", regex=False).astype(float)

    # Salva Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Pendências")

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return "Nenhum arquivo enviado", 400

        file = request.files["file"]
        if file.filename == "":
            return "Nenhum arquivo selecionado", 400

        if file:
            filename = f"{uuid.uuid4().hex}.pdf"
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)

            output_filename = filename.replace(".pdf", ".xlsx")
            output_filepath = os.path.join(app.config["OUTPUT_FOLDER"], output_filename)

            try:
                processar_pdf(filepath, output_filepath)
                
                # Ler arquivo para envio
                with open(output_filepath, 'rb') as f:
                    file_data = f.read()
                
                # Limpar arquivos
                os.remove(filepath)
                os.remove(output_filepath)
                
                # Criar resposta
                from flask import Response
                response = Response(
                    file_data,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename=converted_{output_filename}'}
                )
                return response
                
            except Exception as e:
                return f"Erro ao processar PDF: {str(e)}", 500

    # HTML inline
    html_content = '''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8"/>
    <meta content="width=device-width, initial-scale=1.0" name="viewport"/>
    <title>Conversor PDF → Excel</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet"/>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: 'Inter', sans-serif; line-height: 1.6; color: #ffffff; background: #081f8e; position: relative; }
        body::before { content: ''; position: fixed; inset: 0; background-image: url('https://images.pexels.com/photos/6801874/pexels-photo-6801874.jpeg?auto=compress&cs=tinysrgb&w=1260&h=750&dpr=2'); background-size: cover; background-position: center; background-repeat: no-repeat; z-index: -2; opacity: 0.2; }
        body::after { content: ''; position: fixed; inset: 0; background: linear-gradient(135deg, #081f8e 0%, #013ae2 50%, #262626 100%); z-index: -1; }
        .app-container { display: grid; grid-template-columns: 1fr 1fr; min-height: 100vh; position: relative; z-index: 1; }
        .hero-section { position: relative; display: flex; align-items: center; justify-content: center; padding: 3rem 2rem; }
        .hero-content { position: relative; z-index: 10; text-align: center; color: #ffffff; max-width: 28rem; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; }
        .logo-container { margin-bottom: 2rem; }
        .logo-container img { max-width: 200px; height: auto; border-radius: 0.75rem; box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3); background: rgba(255, 255, 255, 0.1); padding: 1rem; }
        .hero-title { font-size: 3.5rem; font-weight: 700; line-height: 1.1; margin-bottom: 1.5rem; color: #ffffff; text-shadow: 0 4px 12px rgba(0, 0, 0, 0.3); }
        .hero-subtitle { font-size: 1.5rem; color: #ffffff; margin-bottom: 2rem; line-height: 1.6; font-weight: 500; opacity: 0.9; }
        .form-section { display: flex; align-items: center; justify-content: center; padding: 3rem 2rem; position: relative; z-index: 10; }
        .form-container { width: 100%; max-width: 28rem; }
        .upload-form { background: rgba(255, 255, 255, 0.95); backdrop-filter: blur(20px); border-radius: 1.5rem; box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.3); padding: 2rem; border: 1px solid rgba(255, 255, 255, 0.2); }
        .form-header { text-align: center; margin-bottom: 2rem; }
        .form-title { font-size: 1.875rem; font-weight: 700; color: #081f8e; margin-bottom: 0.5rem; }
        .form-subtitle { color: #262626; font-size: 1rem; }
        .input-group { margin-bottom: 1.5rem; }
        .input-label { display: block; font-size: 0.875rem; font-weight: 600; color: #262626; margin-bottom: 0.5rem; }
        .upload-area { position: relative; border: 2px dashed #d1d5db; border-radius: 0.75rem; padding: 2rem 1.5rem; text-align: center; cursor: pointer; transition: all 0.2s ease; background: rgba(250, 250, 250, 0.8); }
        .upload-area:hover { border-color: #013ae2; background: rgba(240, 248, 255, 0.9); }
        .upload-icon { color: #013ae2; margin: 0 auto 1rem; }
        .upload-text { color: #262626; font-weight: 500; margin-bottom: 0.25rem; }
        .upload-subtext { color: #262626; font-size: 0.875rem; opacity: 0.7; }
        .file-input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
        .submit-button { width: 100%; background: linear-gradient(135deg, #013ae2 0%, #081f8e 100%); color: #ffffff; border: none; padding: 1rem 1.5rem; border-radius: 0.75rem; font-size: 1.125rem; font-weight: 600; cursor: pointer; transition: all 0.2s ease; }
        .submit-button:hover { background: linear-gradient(135deg, #012bb8 0%, #06186e 100%); transform: translateY(-1px); box-shadow: 0 10px 25px -5px rgba(1, 58, 226, 0.4); }
        .form-disclaimer { font-size: 0.75rem; color: #262626; text-align: center; margin-top: 1rem; line-height: 1.4; opacity: 0.7; }
        @media (max-width: 1024px) { .app-container { grid-template-columns: 1fr; } .hero-section { min-height: 40vh; padding: 2rem 1rem; } .hero-title { font-size: 3rem; } .form-section { padding: 2rem 1rem; } }
        @media (max-width: 768px) { .hero-title { font-size: 2.5rem; } .logo-container img { max-width: 150px; } .upload-form { padding: 1.5rem; } .form-title { font-size: 1.5rem; } }
    </style>
</head>
<body>
    <div class="app-container">
        <div class="hero-section">
            <div class="hero-content">
                <div class="logo-container">
                    <img src="https://www.suportedata.com.br/logo.png" alt="Logo">
                </div>
                <h1 class="hero-title">Conversor PDF → Excel</h1>
                <p class="hero-subtitle">Transforme dados de PDF em planilhas organizadas</p>
            </div>
        </div>
        <div class="form-section">
            <div class="form-container">
                <form class="upload-form" action="/" method="POST" enctype="multipart/form-data">
                    <div class="form-header">
                        <h2 class="form-title">Envie seu PDF</h2>
                        <p class="form-subtitle">Faça upload do arquivo PDF para conversão</p>
                    </div>
                    <div class="input-group">
                        <label class="input-label">Arquivo PDF</label>
                        <div class="upload-area">
                            <svg class="upload-icon" fill="none" height="48" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" width="48">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                                <polyline points="7,10 12,15 17,10"></polyline>
                                <line x1="12" x2="12" y1="15" y2="3"></line>
                            </svg>
                            <div class="upload-text">Clique ou arraste seu arquivo PDF aqui</div>
                            <div class="upload-subtext">Apenas arquivos PDF são aceitos</div>
                            <input accept=".pdf" class="file-input" name="file" type="file" required/>
                        </div>
                    </div>
                    <button class="submit-button" type="submit">Converter para Excel</button>
                    <p class="form-disclaimer">Seu PDF será processado e convertido automaticamente</p>
                </form>
            </div>
        </div>
    </div>
</body>
</html>'''
    
    return html_content

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 55555))
    app.run(host="0.0.0.0", port=port, debug=False)
