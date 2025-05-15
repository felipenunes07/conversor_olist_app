# Conversor Olist

Aplicação para converter orçamentos para o formato Olist.

## Estrutura do Projeto

```
conversor_olist_app/
├── requirements.txt
├── render.yaml
└── src/
    ├── main.py
    ├── conversor_olist.py
    ├── storage.py
    ├── static/
    │   ├── index.html
    │   └── error.html
    └── data/
        ├── clientes.xlsx
        ├── PLanilha mapeamento Orçamento Olist.xlsx
        └── formato Olist(SAIDA).xlsx
```

## Arquivos Necessários

Os seguintes arquivos Excel são necessários e devem estar na pasta `src/data/`:

1. `clientes.xlsx` - Lista de clientes
2. `PLanilha mapeamento Orçamento Olist.xlsx` - Mapeamento de produtos
3. `formato Olist(SAIDA).xlsx` - Modelo de saída

## Configuração Local

1. Clone o repositório:
```bash
git clone <seu-repositorio>
cd conversor_olist_app
```

2. Crie um ambiente virtual e instale as dependências:
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
pip install -r requirements.txt
```

3. Coloque os arquivos Excel necessários na pasta `src/data/`

4. Execute a aplicação:
```bash
cd src
python main.py
```

## Deploy no Render

1. Faça fork deste repositório no GitHub

2. No Render:
   - Crie uma nova Web Service
   - Conecte ao seu repositório GitHub
   - Selecione o branch principal
   - O arquivo `render.yaml` configurará automaticamente o deploy

3. Após o deploy:
   - Faça upload dos arquivos Excel necessários através da interface da aplicação
   - Verifique se todos os arquivos foram carregados corretamente

## Variáveis de Ambiente

- `PYTHONPATH`: src
- `FLASK_ENV`: production
- `FLASK_DEBUG`: 0

## Suporte

Em caso de problemas:
1. Verifique se todos os arquivos Excel necessários estão presentes
2. Confira os logs da aplicação
3. Certifique-se de que os arquivos Excel estão no formato correto 