import sys
import os
import traceback # Para log detalhado de exceções
import time
import contextlib
from pathlib import Path
import tempfile
# Adiciona o diretório pai de 'src' ao sys.path para permitir importações como 'from src.conversor_olist import ...'
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import io # Para enviar o arquivo em memória
from werkzeug.utils import secure_filename # Para nomes de arquivo seguros

# Importa a função de conversão do outro arquivo .py
from src.conversor_olist import converter_orcamento_para_olist

app = Flask(__name__, static_folder='static', template_folder='static')

# Define o caminho base para os arquivos de dados que estão dentro de 'src'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads') # Para uploads temporários de orçamentos
MAPEAMENTO_PRODUTOS_FILENAME = "PLanilha mapeamento Orçamento Olist.xlsx"
CLIENTES_FILENAME = "clientes.xlsx"
MODELO_SAIDA_OLIST_FILENAME = "formato Olist(SAIDA).xlsx"

MAPEAMENTO_PRODUTOS_PATH = os.path.join(BASE_DIR, MAPEAMENTO_PRODUTOS_FILENAME)
CLIENTES_PATH = os.path.join(BASE_DIR, CLIENTES_FILENAME)
MODELO_SAIDA_OLIST_PATH = os.path.join(BASE_DIR, MODELO_SAIDA_OLIST_FILENAME)

# Criar diretório de uploads se não existir
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/clientes', methods=['GET'])
def get_clientes():
    try:
        if not os.path.exists(CLIENTES_PATH):
            app.logger.error(f"Arquivo de clientes não encontrado em: {CLIENTES_PATH}")
            return jsonify({
                'error': 'Arquivo de clientes não encontrado. Faça o upload na seção de mapeamento.',
                'details': {'path': CLIENTES_PATH}
            }), 404
        
        df_clientes = pd.read_excel(CLIENTES_PATH, sheet_name='CLIENTES')
        if 'ID' in df_clientes.columns and 'Nome' in df_clientes.columns:
            df_clientes = df_clientes.dropna(subset=['Nome'])
            df_clientes['ID'] = df_clientes['ID'].astype(str)
            clientes_list = df_clientes[['ID', 'Nome']].to_dict(orient='records')
            return jsonify({'clientes': clientes_list})
        else:
            return jsonify({'error': 'Estrutura inválida na planilha de clientes.'}), 500
    except Exception as e:
        app.logger.error(f"Erro ao carregar clientes: {str(e)}\n{traceback.format_exc()}")
        return jsonify({'error': str(e), 'details': traceback.format_exc()}), 500

def remove_file_with_retry(file_path, max_retries=3, delay=1):
    """Remove um arquivo com tentativas múltiplas caso esteja em uso."""
    for attempt in range(max_retries):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
            return True
        except PermissionError:
            if attempt < max_retries - 1:
                time.sleep(delay)
                continue
            raise
        except Exception:
            raise
    return False

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    if 'arquivo_excel' not in request.files:
        return jsonify({'error': 'Nenhum arquivo Excel de orçamento enviado.'}), 400
    
    file = request.files['arquivo_excel']
    cliente_id_str = request.form.get('cliente_id')

    if not cliente_id_str:
        return jsonify({'error': 'ID do cliente não fornecido.'}), 400

    if file.filename == '':
        return jsonify({'error': 'Nome de arquivo de orçamento vazio.'}), 400

    if not file or not allowed_file(file.filename):
        return jsonify({'error': 'Tipo de arquivo de orçamento inválido. Use .xlsx'}), 400

    try:
        # Criar arquivo temporário em memória
        input_excel = io.BytesIO(file.read())
        
        # Verificar arquivos de mapeamento
        required_files = {
            'clientes': CLIENTES_PATH,
            'mapeamento': MAPEAMENTO_PRODUTOS_PATH,
            'modelo': MODELO_SAIDA_OLIST_PATH
        }
        
        for file_type, path in required_files.items():
            if not os.path.exists(path):
                return jsonify({
                    'error': f'Arquivo de {file_type} não encontrado.',
                    'details': {'path': path}
                }), 500

        # Verificar estrutura do arquivo de clientes
        df_clientes_check = pd.read_excel(CLIENTES_PATH, sheet_name='CLIENTES')
        if df_clientes_check.empty:
            return jsonify({'error': 'Arquivo de clientes está vazio.'}), 500
        
        if 'ID' not in df_clientes_check.columns:
            return jsonify({'error': 'Coluna ID não encontrada no arquivo de clientes.'}), 500

        # Converter ID do cliente
        try:
            cliente_id_convertido = int(cliente_id_str) if pd.api.types.is_numeric_dtype(df_clientes_check['ID']) else str(cliente_id_str)
        except ValueError:
            return jsonify({'error': 'ID do cliente inválido.'}), 400

        app.logger.info(f"Iniciando conversão para cliente ID: {cliente_id_convertido}")
        
        # Processar arquivo
        df_convertido = converter_orcamento_para_olist(
            input_excel,
            MAPEAMENTO_PRODUTOS_PATH,
            CLIENTES_PATH,
            cliente_id_convertido,
            MODELO_SAIDA_OLIST_PATH
        )

        if df_convertido.empty:
            return jsonify({'error': 'Nenhum dado foi processado.'}), 500

        # Criar arquivo de saída em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_convertido.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='orcamento_convertido_olist.xlsx'
        )

    except Exception as e:
        app.logger.error(f"Erro no processamento: {str(e)}\n{traceback.format_exc()}")
        return jsonify({
            'error': 'Erro no processamento do arquivo.',
            'details': {
                'message': str(e),
                'traceback': traceback.format_exc()
            }
        }), 500

@app.route('/upload_mapeamento', methods=['POST'])
def upload_mapeamento():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado.'}), 400
        
    file_type = request.form.get('file_type')
    if not file_type:
        return jsonify({'error': 'Tipo de arquivo de mapeamento não especificado.'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'Nome de arquivo vazio.'}), 400

    if file and allowed_file(file.filename):
        # Determinar o nome do arquivo de destino com base no file_type
        if file_type == 'clientes':
            target_filename = CLIENTES_FILENAME
            save_path = CLIENTES_PATH
        elif file_type == 'produtos':
            target_filename = MAPEAMENTO_PRODUTOS_FILENAME
            save_path = MAPEAMENTO_PRODUTOS_PATH
        else:
            return jsonify({'error': 'Tipo de arquivo de mapeamento inválido.'}), 400
        
        try:
            # Salva o arquivo diretamente no local correto em 'src'
            file.save(save_path)
            app.logger.info(f"Arquivo de mapeamento '{target_filename}' atualizado com sucesso em '{save_path}'.")
            return jsonify({'message': f'Arquivo {target_filename} atualizado com sucesso!'})
        except Exception as e:
            app.logger.error(f"Erro ao salvar o arquivo de mapeamento '{target_filename}': {str(e)}\n{traceback.format_exc()}")
            return jsonify({'error': f'Erro ao salvar o arquivo {target_filename}: {str(e)}'}), 500
    else:
        return jsonify({'error': 'Tipo de arquivo inválido. Use .xlsx'}), 400

@app.errorhandler(500)
def internal_error(error):
    app.logger.error(f"Erro interno do servidor: {str(error)}\n{traceback.format_exc()}")
    return jsonify({
        'error': 'Erro interno do servidor',
        'details': {
            'message': str(error),
            'traceback': traceback.format_exc()
        }
    }), 500

@app.errorhandler(404)
def not_found_error(error):
    return jsonify({
        'error': 'Recurso não encontrado',
        'details': {'message': str(error)}
    }), 404

# Para desenvolvimento local
if __name__ == '__main__':
    import logging
    import tempfile
    log_file = os.path.join(tempfile.gettempdir(), 'flask_app.log')
    logging.basicConfig(
        filename=log_file,
        level=logging.DEBUG,
        format='%(asctime)s %(levelname)s: %(message)s'
    )
    app.logger.info('Iniciando aplicação...')
    app.run(host='0.0.0.0', port=5000, debug=True)

# Para Vercel - necessário para serverless
app = app

