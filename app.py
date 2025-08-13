#!/usr/bin/env python3
import subprocess
import sys
import os
import tempfile
import logging

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
            logger.info(f"Instalando {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package], 
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            logger.error(f"Erro ao instalar {package}: {e}")

# Tentar importar, se falhar, instalar
try:
    from flask import Flask, request, Response, jsonify
    import pdfplumber
    import pandas as pd
except ImportError:
    print("Instalando dependências...")
    install_packages()
    from flask import Flask, request, Response, jsonify
    import pdfplumber
    import pandas as pd

import re
import uuid
from io import BytesIO

app = Flask(__name__)

# Configuração melhorada
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB max
app.config["UPLOAD_FOLDER"] = tempfile.gettempdir()
app.config["OUTPUT_FOLDER"] = tempfile.gettempdir()

# Regex otimizadas para o formato específico do PDF
try:
    # Regex para detectar cliente - formato: "CODIGO NOME (DDD)TELEFONE CIDADE"
    regex_cliente = re.compile(r"^(\d+)\s+([A-Z\s\.,&-]+?)\s+\((\d{2})\)(\d{4,5}[-\s]?\d{4})\s+([A-Z\s]+)$")
    
    # Regex para títulos com boleto bancário (BANC)
    regex_titulo_banc = re.compile(
        r"^(\d+\.\d+)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d+)\s+(BANC)\s+(\d+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)$"
    )
    
    # Regex para títulos cartão (CART) - formato padrão
    regex_titulo_cart = re.compile(
        r"^(\d+\.\d+)\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d{1,2}/\d{1,2}/\d{4})\s+(\d+)\s+(CART)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)$"
    )
    
    # Regex para detectar linhas de separação ou totais que devem ser ignoradas
    regex_ignorar = re.compile(r"^(TOTAL|Limite|[-=]+|Docto|Vencidas|Total|Financeiro|LB Palmas|Pendencia|Rota|Somente).*")
    
except re.error as e:
    logger.error(f"Erro nas expressões regulares: {e}")

def limpar_valor_numerico(valor_str):
    """Limpa e converte string para float"""
    try:
        if not valor_str or valor_str.strip() == '':
            return 0.0
        # Remove pontos como separadores de milhares e substitui vírgula por ponto decimal
        valor_limpo = valor_str.replace(".", "").replace(",", ".")
        return float(valor_limpo)
    except (ValueError, AttributeError):
        return 0.0

def processar_pdf(pdf_path, output_path):
    """Processa PDF e gera Excel com tratamento de erros melhorado"""
    dados = []
    cliente_atual = None
    cidade_atual = None
    telefone_atual = None
    codigo_cliente_atual = None
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            logger.info(f"Processando PDF com {len(pdf.pages)} páginas")
            
            for page_num, page in enumerate(pdf.pages, 1):
                try:
                    texto = page.extract_text()
                    if not texto:
                        logger.warning(f"Página {page_num} não contém texto extraível")
                        continue
                        
                    linhas = texto.split("\n")
                    logger.info(f"Processando {len(linhas)} linhas da página {page_num}")
                    
                    for linha_num, linha in enumerate(linhas, 1):
                        linha = linha.strip()
                        
                        # Ignorar linhas vazias ou separadores
                        if not linha or len(linha) < 5:
                            continue
                            
                        # Ignorar linhas que sabemos que não contêm dados úteis
                        if regex_ignorar.match(linha):
                            continue
                            
                        # Debug: mostrar as primeiras linhas de cada página
                        if linha_num <= 10:
                            logger.debug(f"Página {page_num}, Linha {linha_num}: {linha}")
                            
                        # Tentar match com regex de cliente
                        try:
                            match_cliente = regex_cliente.match(linha)
                            if match_cliente:
                                codigo_cliente_atual = match_cliente.group(1).strip()
                                cliente_atual = match_cliente.group(2).strip()
                                ddd = match_cliente.group(3).strip()
                                numero = match_cliente.group(4).strip()
                                telefone_atual = f"({ddd}){numero}"
                                cidade_atual = match_cliente.group(5).strip()
                                logger.info(f"Cliente encontrado: {codigo_cliente_atual} - {cliente_atual}")
                                continue
                        except Exception as e:
                            logger.debug(f"Erro ao processar cliente na linha {linha_num}: {e}")

                        # Tentar match com títulos BANC (bancário)
                        try:
                            match_banc = regex_titulo_banc.match(linha)
                            if match_banc and cliente_atual:
                                documento = match_banc.group(1)
                                emissao = match_banc.group(2)
                                vencimento = match_banc.group(3)
                                ats = match_banc.group(4)
                                tipo = match_banc.group(5)
                                boleto = match_banc.group(6)
                                valor_doc = match_banc.group(7)
                                juros = match_banc.group(8)
                                multa = match_banc.group(9)
                                tarifa = match_banc.group(10)
                                valor_total = match_banc.group(11)

                                dados.append([
                                    codigo_cliente_atual, cliente_atual, telefone_atual, cidade_atual, 
                                    documento, emissao, vencimento, ats, tipo, boleto, 
                                    valor_doc, juros, multa, tarifa, valor_total
                                ])
                                logger.info(f"Título BANC encontrado: {documento} - {cliente_atual}")
                                continue
                        except Exception as e:
                            logger.debug(f"Erro ao processar título BANC na linha {linha_num}: {e}")

                        # Tentar match com títulos CART (cartão)
                        try:
                            match_cart = regex_titulo_cart.match(linha)
                            if match_cart and cliente_atual:
                                documento = match_cart.group(1)
                                emissao = match_cart.group(2)
                                vencimento = match_cart.group(3)
                                ats = match_cart.group(4)
                                tipo = match_cart.group(5)
                                boleto = ""  # CART não tem boleto
                                valor_doc = match_cart.group(6)
                                juros = match_cart.group(7)
                                multa = match_cart.group(8)
                                tarifa = match_cart.group(9)
                                valor_total = match_cart.group(10)

                                dados.append([
                                    codigo_cliente_atual, cliente_atual, telefone_atual, cidade_atual, 
                                    documento, emissao, vencimento, ats, tipo, boleto, 
                                    valor_doc, juros, multa, tarifa, valor_total
                                ])
                                logger.info(f"Título CART encontrado: {documento} - {cliente_atual}")
                                continue
                        except Exception as e:
                            logger.debug(f"Erro ao processar título CART na linha {linha_num}: {e}")
                            
                except Exception as e:
                    logger.error(f"Erro ao processar página {page_num}: {e}")
                    continue

    except Exception as e:
        logger.error(f"Erro ao abrir PDF: {e}")
        raise Exception(f"Erro ao processar PDF: {str(e)}")

    if not dados:
        raise Exception("Nenhum dado foi extraído do PDF. Verifique se o formato está correto.")

    # Criar DataFrame com colunas ajustadas
    colunas = [
        "Código Cliente", "Cliente", "Telefone", "Cidade", "Documento", "Emissão", 
        "Vencimento", "ATS", "Tipo", "Boleto", "Valor Documento", "Juros", "Multa", "Tarifa", "Valor Total"
    ]
    
    try:
        df = pd.DataFrame(dados, columns=colunas)
        logger.info(f"DataFrame criado com {len(df)} registros")

        # Ajustar valores numéricos com tratamento de erro
        colunas_numericas = ["Valor Documento", "Juros", "Multa", "Tarifa", "Valor Total"]
        for col in colunas_numericas:
            if col in df.columns:
                df[col] = df[col].apply(limpar_valor_numerico)

        # Salvar Excel com formatação melhorada
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Pendências")
            
            # Ajustar larguras das colunas
            worksheet = writer.sheets["Pendências"]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
            
        logger.info(f"Excel salvo em: {output_path}")
        return len(dados)
        
    except Exception as e:
        logger.error(f"Erro ao criar DataFrame ou salvar Excel: {e}")
        raise Exception(f"Erro ao processar dados: {str(e)}")

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            # Verificar se arquivo foi enviado
            if "file" not in request.files:
                return jsonify({"error": "Nenhum arquivo enviado"}), 400

            file = request.files["file"]
            if file.filename == "":
                return jsonify({"error": "Nenhum arquivo selecionado"}), 400

            # Verificar tipo de arquivo
            if not file.filename.lower().endswith('.pdf'):
                return jsonify({"error": "Apenas arquivos PDF são aceitos"}), 400

            # Salvar arquivo temporariamente
            filename = f"{uuid.uuid4().hex}.pdf"
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            
            try:
                file.save(filepath)
                logger.info(f"Arquivo salvo temporariamente: {filepath}")

                # Verificar se arquivo foi salvo corretamente
                if not os.path.exists(filepath) or os.path.getsize(filepath) == 0:
                    raise Exception("Erro ao salvar arquivo temporário")

                # Processar PDF
                output_filename = filename.replace(".pdf", ".xlsx")
                output_filepath = os.path.join(app.config["OUTPUT_FOLDER"], output_filename)

                registros_processados = processar_pdf(filepath, output_filepath)
                
                # Verificar se Excel foi criado
                if not os.path.exists(output_filepath):
                    raise Exception("Erro ao gerar arquivo Excel")

                # Ler arquivo para envio
                with open(output_filepath, 'rb') as f:
                    file_data = f.read()
                
                # Limpar arquivos temporários
                try:
                    os.remove(filepath)
                    os.remove(output_filepath)
                except Exception as e:
                    logger.warning(f"Erro ao limpar arquivos temporários: {e}")
                
                # Criar resposta com nome mais descritivo
                original_name = file.filename.replace('.pdf', '')
                download_name = f"pendencias_{original_name}_{uuid.uuid4().hex[:8]}.xlsx"
                
                response = Response(
                    file_data,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={
                        'Content-Disposition': f'attachment; filename="{download_name}"',
                        'Content-Length': str(len(file_data))
                    }
                )
                
                logger.info(f"Conversão concluída: {registros_processados} registros processados")
                return response
                
            except Exception as e:
                # Limpar arquivo temporário em caso de erro
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                except:
                    pass
                    
                logger.error(f"Erro no processamento: {str(e)}")
                return jsonify({"error": f"Erro ao processar PDF: {str(e)}"}), 500

        except Exception as e:
            logger.error(f"Erro geral: {str(e)}")
            return jsonify({"error": f"Erro interno do servidor: {str(e)}"}), 500

    # HTML com melhorias visuais
    html_content = '''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="utf-8"/>
    <meta content="width=device-width, initial-scale=1.0" name="viewport"/>
    <title>Conversor PDF → Excel - Pendências</title>
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
        .features { margin-top: 2rem; text-align: left; }
        .feature { display: flex; align-items: center; margin-bottom: 1rem; }
        .feature-icon { margin-right: 0.5rem; font-size: 1.2rem; }
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
        .upload-area.dragover { border-color: #013ae2; background: rgba(240, 248, 255, 0.9); transform: scale(1.02); }
        .upload-icon { color: #013ae2; margin: 0 auto 1rem; }
        .upload-text { color: #262626; font-weight: 500; margin-bottom: 0.25rem; }
        .upload-subtext { color: #262626; font-size: 0.875rem; opacity: 0.7; }
        .file-input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; }
        .file-info { margin-top: 1rem; padding: 0.75rem; background: rgba(1, 58, 226, 0.1); border-radius: 0.5rem; color: #013ae2; font-size: 0.875rem; display: none; }
        .submit-button { width: 100%; background: linear-gradient(135deg, #013ae2 0%, #081f8e 100%); color: #ffffff; border: none; padding: 1rem 1.5rem; border-radius: 0.75rem; font-size: 1.125rem; font-weight: 600; cursor: pointer; transition: all 0.2s ease; }
        .submit-button:hover { background: linear-gradient(135deg, #012bb8 0%, #06186e 100%); transform: translateY(-1px); box-shadow: 0 10px 25px -5px rgba(1, 58, 226, 0.4); }
        .submit-button:disabled { background: #ccc; cursor: not-allowed; transform: none; }
        .loading { display: none; text-align: center; margin-top: 1rem; color: #013ae2; }
        .spinner { border: 2px solid #f3f3f3; border-top: 2px solid #013ae2; border-radius: 50%; width: 20px; height: 20px; animation: spin 1s linear infinite; margin: 0 auto 0.5rem; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .error-message { display: none; margin-top: 1rem; padding: 0.75rem; background: rgba(220, 38, 38, 0.1); border: 1px solid rgba(220, 38, 38, 0.3); border-radius: 0.5rem; color: #dc2626; font-size: 0.875rem; }
        .success-message { display: none; margin-top: 1rem; padding: 0.75rem; background: rgba(34, 197, 94, 0.1); border: 1px solid rgba(34, 197, 94, 0.3); border-radius: 0.5rem; color: #22c55e; font-size: 0.875rem; }
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
                    <img src="https://www.suportedata.com.br/logo.png" alt="Logo" onerror="this.style.display='none'">
                </div>
                <h1 class="hero-title">📊 PDF → Excel</h1>
                <p class="hero-subtitle">Conversor de Pendências Financeiras</p>
                <div class="features">
                    <div class="feature">
                        <span class="feature-icon">✅</span>
                        <span>Extrai clientes e títulos automaticamente</span>
                    </div>
                    <div class="feature">
                        <span class="feature-icon">💰</span>
                        <span>Organiza valores e vencimentos</span>
                    </div>
                    <div class="feature">
                        <span class="feature-icon">📋</span>
                        <span>Gera planilha Excel formatada</span>
                    </div>
                    <div class="feature">
                        <span class="feature-icon">🚀</span>
                        <span>Processamento rápido e seguro</span>
                    </div>
                </div>
            </div>
        </div>
        <div class="form-section">
            <div class="form-container">
                <form class="upload-form" id="uploadForm" action="/" method="POST" enctype="multipart/form-data">
                    <div class="form-header">
                        <h2 class="form-title">Envie seu PDF</h2>
                        <p class="form-subtitle">Relatório de pendências financeiras</p>
                    </div>
                    <div class="input-group">
                        <label class="input-label">Arquivo PDF</label>
                        <div class="upload-area" id="uploadArea">
                            <svg class="upload-icon" fill="none" height="48" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" width="48">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                                <polyline points="7,10 12,15 17,10"></polyline>
                                <line x1="12" x2="12" y1="15" y2="3"></line>
                            </svg>
                            <div class="upload-text">Clique ou arraste seu PDF aqui</div>
                            <div class="upload-subtext">Máximo 50MB • Formato específico de pendências</div>
                            <input accept=".pdf" class="file-input" name="file" type="file" id="fileInput" required/>
                        </div>
                        <div class="file-info" id="fileInfo"></div>
                    </div>
                    <button class="submit-button" type="submit" id="submitBtn">🔄 Converter para Excel</button>
                    <div class="loading" id="loading">
                        <div class="spinner"></div>
                        <div>Processando PDF...</div>
                    </div>
                    <div class="error-message" id="errorMessage"></div>
                    <div class="success-message" id="successMessage"></div>
                    <p class="form-disclaimer">
                        <strong>Formato suportado:</strong> PDFs de pendências com clientes, documentos, valores e vencimentos.
                        Todos os dados são processados localmente e não são armazenados.
                    </p>
                </form>
            </div>
        </div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const submitBtn = document.getElementById('submitBtn');
        const loading = document.getElementById('loading');
        const errorMessage = document.getElementById('errorMessage');
        const successMessage = document.getElementById('successMessage');
        const uploadForm = document.getElementById('uploadForm');

        // Drag and drop
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                showFileInfo(files[0]);
            }
        });

        uploadArea.addEventListener('click', () => {
            fileInput.click();
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                showFileInfo(e.target.files[0]);
            }
        });

        function showFileInfo(file) {
            const maxSize = 50 * 1024 * 1024; // 50MB
            
            if (!file.type.includes('pdf') && !file.name.toLowerCase().endsWith('.pdf')) {
                showError('Apenas arquivos PDF são aceitos.');
                return;
            }
            
            if (file.size > maxSize) {
                showError('Arquivo muito grande. Tamanho máximo: 50MB.');
                return;
            }
            
            fileInfo.innerHTML = `📄 ${file.name} (${formatFileSize(file.size)})`;
            fileInfo.style.display = 'block';
            hideMessages();
        }

        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }

        function showError(message) {
            errorMessage.textContent = message;
            errorMessage.style.display = 'block';
            successMessage.style.display = 'none';
        }

        function showSuccess(message) {
            successMessage.textContent = message;
            successMessage.style.display = 'block';
            errorMessage.style.display = 'none';
        }

        function hideMessages() {
            errorMessage.style.display = 'none';
            successMessage.style.display = 'none';
        }

        uploadForm.addEventListener('submit', (e) => {
            if (!fileInput.files.length) {
                e.preventDefault();
                showError('Por favor, selecione um arquivo PDF.');
                return;
            }
            
            submitBtn.disabled = true;
            submitBtn.textContent = 'Processando...';
            loading.style.display = 'block';
            hideMessages();
            
            // Mostrar mensagem de sucesso após um tempo
            setTimeout(() => {
                showSuccess('PDF processado com sucesso! Download iniciará automaticamente.');
            }, 2000);
        });

        // Reset form after response
        window.addEventListener('load', () => {
            submitBtn.disabled = false;
            submitBtn.textContent = '🔄 Converter para Excel';
            loading.style.display = 'none';
        });
    </script>
</body>
</html>'''
    
    return html_content

@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": "Arquivo muito grande. Tamanho máximo: 50MB."}), 413

@app.errorhandler(500)
def internal_error(e):
    return jsonify({"error": "Erro interno do servidor."}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    logger.info(f"Iniciando servidor na porta {port}")
    logger.info("Conversor PDF → Excel para Pendências Financeiras")
    logger.info("Regex otimizadas para formato específico")
    app.run(host="0.0.0.0", port=port, debug=False)
