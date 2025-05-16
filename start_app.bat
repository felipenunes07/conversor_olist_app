@echo off
echo Iniciando o Conversor Olist...

:: Verificar se o ambiente virtual existe, senão criar
if not exist venv (
    echo Criando ambiente virtual...
    python -m venv venv
)

:: Ativar o ambiente virtual
call venv\Scripts\activate

:: Instalar dependências
echo Instalando dependências...
pip install flask pandas openpyxl flask-sqlalchemy pymysql python-dotenv gunicorn

:: Executar a aplicação
echo Iniciando a aplicação...
python run_local.py

pause 