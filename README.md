# Conversor de Orçamentos para Olist

Uma aplicação web desenvolvida em Flask para converter arquivos de orçamento para o formato Olist.

## Funcionalidades

- Interface web amigável para upload de arquivos
- Conversão de orçamentos para o formato Olist
- Mapeamento de produtos e clientes
- Download do arquivo convertido em formato Excel
- Suporte a múltiplos clientes

## Requisitos

- Python 3.x
- Flask 3.1.0
- Pandas
- Outras dependências listadas em `requirements.txt`

## Instalação

1. Clone o repositório:
```bash
git clone https://github.com/felipenunes07/conversor_olist_app.git
cd conversor_olist_app
```

2. Crie um ambiente virtual e ative-o:
```bash
python -m venv venv
# No Windows:
.\venv\Scripts\activate
# No Linux/Mac:
source venv/bin/activate
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

## Configuração

1. Certifique-se de que os seguintes arquivos estão presentes no diretório `src`:
   - `clientes.xlsx` - Arquivo de mapeamento de clientes
   - `PLanilha mapeamento Orçamento Olist.xlsx` - Arquivo de mapeamento de produtos
   - `formato Olist(SAIDA).xlsx` - Modelo do arquivo de saída

2. Crie o diretório de uploads (se não existir):
```bash
mkdir src/uploads
```

## Uso

1. Inicie a aplicação:
```bash
cd src
python main.py
```

2. Acesse a interface web em: http://localhost:5000

3. Na interface web:
   - Faça upload dos arquivos de mapeamento necessários
   - Selecione um cliente
   - Faça upload do arquivo de orçamento
   - Clique em "Processar" para converter
   - Faça o download do arquivo convertido

## Estrutura do Projeto

```
conversor_olist_app/
├── src/
│   ├── static/
│   │   ├── css/
│   │   ├── js/
│   │   └── index.html
│   ├── uploads/
│   ├── main.py
│   ├── conversor_olist.py
│   └── [arquivos de mapeamento]
├── venv/
├── requirements.txt
└── README.md
```

## Desenvolvimento

- O modo debug está ativado por padrão para desenvolvimento
- Os logs são salvos em um arquivo temporário
- O diretório `uploads` é usado para arquivos temporários durante o processamento

## Suporte

Para reportar problemas ou sugerir melhorias, por favor abra uma issue no GitHub. 