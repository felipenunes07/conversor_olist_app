import sys
import os
import traceback # Para log detalhado de exceções
import time
import contextlib
from pathlib import Path
import tempfile
# Adiciona o diretório pai de 'src' ao sys.path para permitir importações como 'from src.conversor_olist import ...'
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

# ===== INÍCIO DIAGNÓSTICO =====
print("Diretório atual:", os.getcwd())
# ===== FIM DIAGNÓSTICO =====

from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import io # Para enviar o arquivo em memória
from werkzeug.utils import secure_filename # Para nomes de arquivo seguros

# Importa a função de conversão do outro arquivo .py
from conversor_olist import converter_orcamento_para_olist

app = Flask(__name__, static_folder='static', template_folder='static')

# Define o caminho base para os arquivos de dados que estão dentro de 'src'
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads') # Para uploads temporários de orçamentos
MAPEAMENTO_PRODUTOS_FILENAME = "PLanilha mapeamento Orçamento Olist.xlsx"
CLIENTES_FILENAME = "clientes.xlsx"
MODELO_SAIDA_OLIST_FILENAME = "formato Olist(SAIDA).xlsx"

MAPEAMENTO_PRODUTOS_PATH = os.path.join(DATA_DIR, MAPEAMENTO_PRODUTOS_FILENAME)
CLIENTES_PATH = os.path.join(DATA_DIR, CLIENTES_FILENAME)
MODELO_SAIDA_OLIST_PATH = os.path.join(DATA_DIR, MODELO_SAIDA_OLIST_FILENAME)

# ===== INÍCIO DIAGNÓSTICO =====
print("Caminhos dos arquivos:")
print(f"MAPEAMENTO_PRODUTOS_PATH: {MAPEAMENTO_PRODUTOS_PATH} (Existe: {os.path.exists(MAPEAMENTO_PRODUTOS_PATH)})")
print(f"CLIENTES_PATH: {CLIENTES_PATH} (Existe: {os.path.exists(CLIENTES_PATH)})")
print(f"MODELO_SAIDA_OLIST_PATH: {MODELO_SAIDA_OLIST_PATH} (Existe: {os.path.exists(MODELO_SAIDA_OLIST_PATH)})")
print(f"Arquivos em {DATA_DIR} (Existe: {os.path.exists(DATA_DIR)}):", os.listdir(DATA_DIR) if os.path.exists(DATA_DIR) else "Pasta não existe")
# ===== FIM DIAGNÓSTICO =====

# Criar diretório de uploads se não existir
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def check_required_files():
    """Check if all required files exist and are readable."""
    # Certifique-se de que o diretório DATA_DIR existe
    os.makedirs(DATA_DIR, exist_ok=True)
    print(f"[CHECK] Verificando diretório de dados: {DATA_DIR} (Existe: {os.path.exists(DATA_DIR)})")
    
    required_files = {
        'clientes': CLIENTES_PATH,
        'mapeamento': MAPEAMENTO_PRODUTOS_PATH,
        'modelo': MODELO_SAIDA_OLIST_PATH
    }
    
    missing_files = []
    for file_type, path in required_files.items():
        if not os.path.exists(path):
            missing_files.append(file_type)
            app.logger.error(f"Required file missing: {path}")
            print(f"[CHECK] Arquivo ausente: {path}")
        else:
            print(f"[CHECK] Arquivo encontrado: {path}")
    
    return missing_files

@app.route('/')
def index():
    try:
        # Check for required files on startup
        missing_files = check_required_files()
        if missing_files:
            return render_template('error.html', 
                                error=f"Missing required files: {', '.join(missing_files)}. Please upload them first.")
        return render_template('index.html')
    except Exception as e:
        app.logger.error(f"Error rendering index: {str(e)}\n{traceback.format_exc()}")
        return jsonify({'error': 'Error loading application'}), 500

@app.route('/clientes', methods=['GET'])
def get_clientes():
    try:
        if not os.path.exists(CLIENTES_PATH):
            app.logger.error(f"Client file not found at: {CLIENTES_PATH}")
            return jsonify({
                'error': 'Client file not found',
                'details': {'path': CLIENTES_PATH}
            }), 404
        
        df_clientes = pd.read_excel(CLIENTES_PATH, sheet_name='CLIENTES')
        if 'ID' in df_clientes.columns and 'Nome' in df_clientes.columns:
            df_clientes = df_clientes.dropna(subset=['Nome'])
            df_clientes['ID'] = df_clientes['ID'].astype(str)
            clientes_list = df_clientes[['ID', 'Nome']].to_dict(orient='records')
            return jsonify({'clientes': clientes_list})
        else:
            return jsonify({'error': 'Invalid client file structure'}), 500
    except Exception as e:
        app.logger.error(f"Error loading clients: {str(e)}\n{traceback.format_exc()}")
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
    try:
        # Check required files first
        missing_files = check_required_files()
        if missing_files:
            return jsonify({
                'error': 'Missing required files',
                'details': {'missing': missing_files}
            }), 500

        if 'arquivo_excel' not in request.files:
            return jsonify({'error': 'No Excel file uploaded'}), 400
        
        file = request.files['arquivo_excel']
        cliente_id_str = request.form.get('cliente_id')

        if not cliente_id_str:
            return jsonify({'error': 'No client ID provided'}), 400

        if file.filename == '':
            return jsonify({'error': 'Empty filename'}), 400

        if not file or not allowed_file(file.filename):
            return jsonify({'error': 'Invalid file type. Use .xlsx'}), 400

        # Create in-memory file
        input_excel = io.BytesIO(file.read())
        
        try:
            df_convertido = converter_orcamento_para_olist(
                input_excel,
                MAPEAMENTO_PRODUTOS_PATH,
                CLIENTES_PATH,
                cliente_id_str,
                MODELO_SAIDA_OLIST_PATH
            )

            if df_convertido.empty:
                return jsonify({'error': 'No data processed'}), 500

            # Create output file in memory
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
            app.logger.error(f"Error processing file: {str(e)}\n{traceback.format_exc()}")
            return jsonify({
                'error': 'Error processing file',
                'details': {
                    'message': str(e),
                    'traceback': traceback.format_exc()
                }
            }), 500

    except Exception as e:
        app.logger.error(f"Unexpected error: {str(e)}\n{traceback.format_exc()}")
        return jsonify({
            'error': 'Unexpected error',
            'details': {
                'message': str(e),
                'traceback': traceback.format_exc()
            }
        }), 500

@app.route('/upload_mapeamento', methods=['POST'])
def upload_mapeamento():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
            
        file_type = request.form.get('file_type')
        if not file_type:
            return jsonify({'error': 'No mapping file type specified'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'Empty filename'}), 400

        if file and allowed_file(file.filename):
            if file_type == 'clientes':
                save_path = CLIENTES_PATH
            elif file_type == 'produtos':
                save_path = MAPEAMENTO_PRODUTOS_PATH
            else:
                return jsonify({'error': 'Invalid mapping file type'}), 400
            
            try:
                os.makedirs(os.path.dirname(save_path), exist_ok=True)
                file.save(save_path)
                return jsonify({'message': f'File updated successfully'})
            except Exception as e:
                app.logger.error(f"Error saving mapping file: {str(e)}\n{traceback.format_exc()}")
                return jsonify({'error': f'Error saving file: {str(e)}'}), 500
        else:
            return jsonify({'error': 'Invalid file type. Use .xlsx'}), 400
    except Exception as e:
        app.logger.error(f"Error in upload_mapeamento: {str(e)}\n{traceback.format_exc()}")
        return jsonify({'error': str(e)}), 500

@app.errorhandler(500)
def internal_error(error):
    app.logger.error(f"Internal server error: {str(error)}\n{traceback.format_exc()}")
    return jsonify({
        'error': 'Internal server error',
        'details': {
            'message': str(error),
            'traceback': traceback.format_exc()
        }
    }), 500

@app.errorhandler(404)
def not_found_error(error):
    return jsonify({
        'error': 'Resource not found',
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

