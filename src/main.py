import sys
import os
import traceback # Para log detalhado de exceções
import time
import contextlib
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

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

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
            # Tenta criar um arquivo vazio se não existir para não quebrar a interface inicial
            # Ou retorna um erro mais amigável para o usuário sobre a necessidade de upload
            return jsonify({'error': 'Arquivo de clientes (clientes.xlsx) não encontrado no servidor. Por favor, faça o upload do arquivo de mapeamento de clientes.'}), 404
        
        df_clientes = pd.read_excel(CLIENTES_PATH, sheet_name='CLIENTES')
        if 'ID' in df_clientes.columns and 'Nome' in df_clientes.columns:
            df_clientes = df_clientes.dropna(subset=['Nome'])
            df_clientes['ID'] = df_clientes['ID'].astype(str)
            clientes_list = df_clientes[['ID', 'Nome']].to_dict(orient='records')
            return jsonify({'clientes': clientes_list})
        else:
            return jsonify({'error': 'Colunas ID ou Nome não encontradas na planilha de clientes.'}), 500
    except FileNotFoundError:
        app.logger.error(f"Arquivo de clientes não encontrado em: {CLIENTES_PATH}")
        return jsonify({'error': 'Arquivo de clientes não encontrado. Faça o upload na seção de mapeamento.'}), 404
    except Exception as e:
        app.logger.error(f"Erro ao carregar clientes: {str(e)}\n{traceback.format_exc()}")
        return jsonify({'error': f'Erro interno ao carregar clientes: {str(e)}'}), 500

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
    temp_orcamento_path = None
    try:
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

        filename = secure_filename(file.filename)
        temp_orcamento_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        try:
            file.save(temp_orcamento_path)
            app.logger.info(f"Arquivo temporário salvo em: {temp_orcamento_path}")
            
            # Verificar arquivos de mapeamento
            if not os.path.exists(CLIENTES_PATH):
                remove_file_with_retry(temp_orcamento_path)
                return jsonify({'error': 'Arquivo de mapeamento de clientes (clientes.xlsx) não encontrado. Faça o upload primeiro.'}), 500
            if not os.path.exists(MAPEAMENTO_PRODUTOS_PATH):
                remove_file_with_retry(temp_orcamento_path)
                return jsonify({'error': 'Arquivo de mapeamento de produtos (PLanilha mapeamento Orçamento Olist.xlsx) não encontrado. Faça o upload primeiro.'}), 500
            if not os.path.exists(MODELO_SAIDA_OLIST_PATH):
                remove_file_with_retry(temp_orcamento_path)
                return jsonify({'error': 'Arquivo modelo de saída (formato Olist(SAIDA).xlsx) não encontrado.'}), 500

            df_clientes_check = pd.read_excel(CLIENTES_PATH, sheet_name='CLIENTES')
            if df_clientes_check.empty:
                remove_file_with_retry(temp_orcamento_path)
                return jsonify({'error': 'Arquivo de clientes está vazio.'}), 500
            
            if 'ID' not in df_clientes_check.columns:
                remove_file_with_retry(temp_orcamento_path)
                return jsonify({'error': 'Coluna ID não encontrada no arquivo de clientes.'}), 500

            # Converter ID do cliente para o tipo correto
            if pd.api.types.is_numeric_dtype(df_clientes_check['ID']):
                try:
                    cliente_id_convertido = int(cliente_id_str)
                except ValueError:
                    try:
                        cliente_id_convertido = float(cliente_id_str)
                    except ValueError:
                        remove_file_with_retry(temp_orcamento_path)
                        return jsonify({'error': 'ID do cliente inválido. Deve ser um número.'}), 400
            else:
                cliente_id_convertido = str(cliente_id_str)

            app.logger.info(f"Iniciando conversão para cliente ID: {cliente_id_convertido}")
            
            # Usar context manager para garantir que os arquivos Excel sejam fechados
            with pd.ExcelFile(temp_orcamento_path) as xls:
                df_convertido = converter_orcamento_para_olist(
                    temp_orcamento_path, 
                    MAPEAMENTO_PRODUTOS_PATH, 
                    CLIENTES_PATH, 
                    cliente_id_convertido, 
                    MODELO_SAIDA_OLIST_PATH
                )
            
            # Tentar remover o arquivo temporário com retry
            if not remove_file_with_retry(temp_orcamento_path):
                app.logger.warning(f"Não foi possível remover o arquivo temporário: {temp_orcamento_path}")

            if df_convertido.empty:
                return jsonify({'error': 'Nenhum dado foi processado. Verifique se o arquivo de orçamento está no formato correto.'}), 500

            # Criar o arquivo de saída em memória
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

        except FileNotFoundError as e_fnf:
            app.logger.error(f"Arquivo não encontrado: {str(e_fnf)}\n{traceback.format_exc()}")
            if temp_orcamento_path:
                remove_file_with_retry(temp_orcamento_path)
            return jsonify({'error': f'Arquivo não encontrado: {str(e_fnf)}'}), 500
        except Exception as e_proc:
            app.logger.error(f"Erro durante o processamento: {str(e_proc)}\n{traceback.format_exc()}")
            if temp_orcamento_path:
                remove_file_with_retry(temp_orcamento_path)
            return jsonify({'error': f'Erro durante o processamento: {str(e_proc)}'}), 500
    except Exception as e:
        app.logger.error(f"Erro inesperado: {str(e)}\n{traceback.format_exc()}")
        if temp_orcamento_path and os.path.exists(temp_orcamento_path):
            remove_file_with_retry(temp_orcamento_path)
        return jsonify({'error': f'Erro inesperado: {str(e)}'}), 500

@app.route('/upload_mapeamento', methods=['POST'])
def upload_mapeamento():
    file_type = request.form.get('file_type') # 'clientes' ou 'produtos'
    
    if not file_type:
        return jsonify({'error': 'Tipo de arquivo de mapeamento não especificado.'}), 400

    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado.'}), 400

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
    return jsonify({'error': 'Erro interno do servidor. Verifique os logs para mais detalhes.'}), 500

@app.errorhandler(404)
def not_found_error(error):
    return jsonify({'error': 'Página não encontrada.'}), 404

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

# Adicionar esta linha para a Vercel
app = app

