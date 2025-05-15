import sys
import os
import traceback # Para log detalhado de exceções
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

    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename) # Segurança
        temp_orcamento_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            file.save(temp_orcamento_path)
            
            # Checar se os arquivos de mapeamento existem antes de processar
            if not os.path.exists(CLIENTES_PATH):
                 os.remove(temp_orcamento_path)
                 return jsonify({'error': 'Arquivo de mapeamento de clientes (clientes.xlsx) não encontrado. Faça o upload primeiro.'}), 500
            if not os.path.exists(MAPEAMENTO_PRODUTOS_PATH):
                 os.remove(temp_orcamento_path)
                 return jsonify({'error': 'Arquivo de mapeamento de produtos (PLanilha mapeamento Orçamento Olist.xlsx) não encontrado. Faça o upload primeiro.'}), 500

            df_clientes_check = pd.read_excel(CLIENTES_PATH, sheet_name='CLIENTES')
            if not df_clientes_check.empty and 'ID' in df_clientes_check.columns:
                if pd.api.types.is_numeric_dtype(df_clientes_check['ID']):
                    try:
                        cliente_id_convertido = int(cliente_id_str)
                    except ValueError:
                        cliente_id_convertido = float(cliente_id_str)
                else:
                    cliente_id_convertido = str(cliente_id_str)
            else:
                cliente_id_convertido = cliente_id_str

            df_convertido = converter_orcamento_para_olist(
                temp_orcamento_path, 
                MAPEAMENTO_PRODUTOS_PATH, 
                CLIENTES_PATH, 
                cliente_id_convertido, 
                MODELO_SAIDA_OLIST_PATH
            )
            
            os.remove(temp_orcamento_path)

            if df_convertido.empty:
                return jsonify({'error': 'Falha na conversão ou nenhum dado processado. Verifique os logs e o arquivo de entrada.'}), 500

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
            app.logger.error(f"Arquivo de configuração não encontrado: {e_fnf.filename}\n{traceback.format_exc()}")
            if os.path.exists(temp_orcamento_path): os.remove(temp_orcamento_path)
            return jsonify({'error': f'Arquivo de configuração não encontrado: {e_fnf.filename}. Faça o upload na seção de mapeamento.'}), 500
        except Exception as e_proc:
            app.logger.error(f"Erro durante o processamento: {str(e_proc)}\n{traceback.format_exc()}")
            if os.path.exists(temp_orcamento_path): os.remove(temp_orcamento_path)
            return jsonify({'error': f'Erro durante o processamento: {str(e_proc)}'}), 500
    else:
        return jsonify({'error': 'Tipo de arquivo de orçamento inválido. Use .xlsx'}), 400

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

if __name__ == '__main__':
    import logging
    logging.basicConfig(filename='/home/ubuntu/flask_app.log', level=logging.DEBUG)
    app.run(host='0.0.0.0', port=5000, debug=False) # Debug=False para produção, mas True para desenvolvimento

