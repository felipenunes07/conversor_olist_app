import pandas as pd
import sys
import traceback
import re # Para normalização

def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto_str = str(texto).lower().strip()
    # Remover múltiplos espaços
    texto_str = re.sub(r'\s+', ' ', texto_str)
    return texto_str

def encontrar_linha_cabecalho(df_preview, palavras_chave_cabecalho):
    palavras_chave_normalizadas = [normalizar_texto(pc) for pc in palavras_chave_cabecalho]
    for i, row in df_preview.iterrows():
        valores_linha = [normalizar_texto(x) for x in row.tolist()]
        if all(palavra_chave in valores_linha for palavra_chave in palavras_chave_normalizadas):
            return i
    return None

def converter_orcamento_para_olist(caminho_orcamento_entrada, caminho_mapeamento_produtos, caminho_clientes, id_cliente_selecionado, caminho_modelo_saida_olist_com_dados):
    df_modelo_saida_temp = None
    colunas_modelo_olist = []
    produtos_nao_mapeados_log = [] # Lista para logar produtos não mapeados
    print(f"[CONVERSOR V6] Iniciando conversão. Cliente ID: {id_cliente_selecionado}", file=sys.stderr)
    try:
        print(f"[CONVERSOR V6] Lendo arquivo de mapeamento: {caminho_mapeamento_produtos}", file=sys.stderr)
        df_mapeamento = pd.read_excel(caminho_mapeamento_produtos, sheet_name='CATÁLOGO')
        # Normalizar a coluna de busca no mapeamento
        if 'MODELO' in df_mapeamento.columns:
            df_mapeamento['MODELO_NORMALIZADO_BUSCA'] = df_mapeamento['MODELO'].apply(normalizar_texto)
            print(f"[CONVERSOR V6] Coluna 'MODELO' normalizada para busca em df_mapeamento.", file=sys.stderr)
        else:
            print(f"[CONVERSOR V6] ERRO: Coluna 'MODELO' não encontrada em {caminho_mapeamento_produtos}", file=sys.stderr)
            # Tratar erro ou retornar, pois o mapeamento será impossível
            return pd.DataFrame(columns=colunas_modelo_olist if colunas_modelo_olist else [])

        print(f"[CONVERSOR V6] Lendo arquivo de clientes: {caminho_clientes}", file=sys.stderr)
        df_clientes = pd.read_excel(caminho_clientes, sheet_name='CLIENTES')
        
        print(f"[CONVERSOR V6] Lendo NOVO arquivo modelo de saída com dados: {caminho_modelo_saida_olist_com_dados}", file=sys.stderr)
        xls_modelo_novo = pd.ExcelFile(caminho_modelo_saida_olist_com_dados)
        if xls_modelo_novo.sheet_names:
            df_modelo_saida_temp = pd.read_excel(xls_modelo_novo, sheet_name=0)
            print(f"[CONVERSOR V6] Lida a primeira aba do NOVO modelo de saída: {xls_modelo_novo.sheet_names[0]}", file=sys.stderr)
        else:
            raise ValueError("O NOVO arquivo Excel modelo de saída não contém nenhuma aba.")
        colunas_modelo_olist = df_modelo_saida_temp.columns.tolist()
        print(f"[CONVERSOR V6] Colunas do NOVO modelo Olist: {colunas_modelo_olist}", file=sys.stderr)

        print(f"[CONVERSOR V6] Lendo arquivo de orçamento: {caminho_orcamento_entrada}", file=sys.stderr)
        xls_orc = pd.ExcelFile(caminho_orcamento_entrada)
        sheet_name_orcamento = None
        if 'Orçamento' in xls_orc.sheet_names:
            sheet_name_orcamento = 'Orçamento'
        elif xls_orc.sheet_names:
            sheet_name_orcamento = xls_orc.sheet_names[0]
            print(f"[CONVERSOR V6] Aviso: Aba 'Orçamento' não encontrada. Usando primeira aba: {sheet_name_orcamento}", file=sys.stderr)
        else:
            raise ValueError("O arquivo Excel de orçamento não contém nenhuma aba.")

        df_orc_preview_meta = pd.read_excel(xls_orc, sheet_name=sheet_name_orcamento, nrows=10, header=None) # Ler mais linhas para encontrar cabeçalho
        num_proposta_orc = None
        data_proposta_orc = None
        linha_cabecalho_itens_idx = None

        if len(df_orc_preview_meta) > 0:
            if len(df_orc_preview_meta.columns) > 1:
                # Busca 'Orçamento #' e 'Data' nas primeiras linhas
                for i in range(len(df_orc_preview_meta)):
                    if normalizar_texto(df_orc_preview_meta.iloc[i, 0]) == "orçamento #":
                        num_proposta_orc = df_orc_preview_meta.iloc[i, 1]
                        print(f"[CONVERSOR V6] Número da proposta extraído (linha {i}): {num_proposta_orc}", file=sys.stderr)
                    if normalizar_texto(df_orc_preview_meta.iloc[i, 0]) == "data":
                        data_proposta_orc = df_orc_preview_meta.iloc[i, 1]
                        if isinstance(data_proposta_orc, pd.Timestamp):
                            data_proposta_orc = data_proposta_orc.date()
                        elif isinstance(data_proposta_orc, str):
                            try: data_proposta_orc = pd.to_datetime(data_proposta_orc, dayfirst=True).date()
                            except ValueError: 
                                try: data_proposta_orc = pd.to_datetime(data_proposta_orc).date()
                                except ValueError: pass # Deixar como string se não puder converter
                        print(f"[CONVERSOR V6] Data da proposta extraída (linha {i}): {data_proposta_orc}", file=sys.stderr)
            
            palavras_chave_cabecalho_itens = ["Produto", "Cor", "Qualidade", "Valor Unitário", "Quantidade", "Subtotal"]
            linha_cabecalho_itens_idx = encontrar_linha_cabecalho(df_orc_preview_meta, palavras_chave_cabecalho_itens)
            
            if linha_cabecalho_itens_idx is not None:
                print(f"[CONVERSOR V6] Linha de cabeçalho dos itens identificada no índice: {linha_cabecalho_itens_idx}", file=sys.stderr)
                # header é o índice da linha a ser usada como cabeçalho
                df_orcamento_itens = pd.read_excel(xls_orc, sheet_name=sheet_name_orcamento, header=linha_cabecalho_itens_idx)
                # Remover linhas acima do cabeçalho que foram lidas junto (se header > 0)
                # Se header=0, não há linhas acima. Se header=2, iloc[1:] remove a linha 0 e 1 (relativo ao novo cabeçalho)
                # A linha do cabeçalho se torna a linha 0 do novo DataFrame, então não precisamos mais pular
                # df_orcamento_itens = df_orcamento_itens.iloc[1:].reset_index(drop=True) # Esta linha estava incorreta
            else:
                print("[CONVERSOR V6] Aviso: Cabeçalho dos itens não identificado. Tentando leitura padrão com skiprows=2.", file=sys.stderr)
                df_orcamento_itens = pd.read_excel(xls_orc, sheet_name=sheet_name_orcamento, skiprows=2)
        else:
            print("[CONVERSOR V6] Aviso: Não foi possível ler metadados ou cabeçalho do orçamento.", file=sys.stderr)
            df_orcamento_itens = pd.DataFrame()

        print(f"[CONVERSOR V6] Itens do orçamento lidos. Colunas: {list(df_orcamento_itens.columns)}", file=sys.stderr)
        # Normalizar nomes das colunas do orçamento lido
        df_orcamento_itens.columns = [normalizar_texto(col) for col in df_orcamento_itens.columns]
        print(f"[CONVERSOR V6] Colunas NORMALIZADAS dos itens do orçamento: {list(df_orcamento_itens.columns)}", file=sys.stderr)

    except FileNotFoundError as e:
        print(f"[CONVERSOR V6] Erro Crítico: Arquivo essencial não encontrado: {e.filename}\n{traceback.format_exc()}", file=sys.stderr)
        return pd.DataFrame(columns=colunas_modelo_olist if colunas_modelo_olist else [])
    except Exception as e_read:
        print(f"[CONVERSOR V6] Erro Crítico: Erro inesperado ao ler arquivos: {str(e_read)}\n{traceback.format_exc()}", file=sys.stderr)
        return pd.DataFrame(columns=colunas_modelo_olist if colunas_modelo_olist else [])

    info_cliente_df = pd.DataFrame()
    if not df_clientes.empty and 'ID' in df_clientes.columns:
        try:
            coluna_id_tipo = df_clientes['ID'].dtype
            id_cliente_selecionado_str = str(id_cliente_selecionado)
            if pd.api.types.is_numeric_dtype(coluna_id_tipo):
                try: id_cliente_convertido = int(float(id_cliente_selecionado_str))
                except ValueError: id_cliente_convertido = float(id_cliente_selecionado_str)
            else:
                id_cliente_convertido = id_cliente_selecionado_str
            info_cliente_df = df_clientes[df_clientes['ID'] == id_cliente_convertido]
        except Exception as e_conv_cliente:
            print(f"[CONVERSOR V6] Erro ao buscar cliente: {str(e_conv_cliente)}", file=sys.stderr)
    
    if info_cliente_df.empty:
        print(f"[CONVERSOR V6] Aviso: Cliente com ID '{id_cliente_selecionado}' não encontrado.", file=sys.stderr)
        return pd.DataFrame(columns=colunas_modelo_olist)
    
    info_cliente = info_cliente_df.iloc[0]
    id_contato_cliente = info_cliente['ID']
    nome_contato_cliente = info_cliente['Nome']
    print(f"[CONVERSOR V6] Cliente encontrado: ID {id_contato_cliente}, Nome {nome_contato_cliente}", file=sys.stderr)

    linhas_saida = []
    print(f"[CONVERSOR V6] Processando {len(df_orcamento_itens)} itens do orçamento.", file=sys.stderr)
    for index, linha_item in df_orcamento_itens.iterrows():
        # Usar nomes de coluna NORMALIZADOS
        produto_orcamento_original = linha_item.get('produto') 
        qtde = linha_item.get('quantidade')
        valor_unit = linha_item.get('valor unitário') # ou 'valor unitario' se normalizado

        if pd.isna(produto_orcamento_original) and pd.isna(qtde) and pd.isna(valor_unit):
            # print(f"[CONVERSOR V6] Linha item {index} do orçamento parece vazia (após normalização de colunas), pulando.", file=sys.stderr)
            continue

        produto_orcamento_busca_normalizado = normalizar_texto(produto_orcamento_original)
        print(f"[CONVERSOR V6] Linha Item {index}: ProdutoOriginal='{produto_orcamento_original}', ProdutoBuscaNormalizado='{produto_orcamento_busca_normalizado}', Qtde='{qtde}', ValorUnit='{valor_unit}'", file=sys.stderr)

        id_produto_olist = pd.NA
        descricao_produto_olist = pd.NA
        if produto_orcamento_busca_normalizado and not df_mapeamento.empty and 'MODELO_NORMALIZADO_BUSCA' in df_mapeamento.columns:
            # Busca usando a coluna normalizada
            produto_mapeado_df = df_mapeamento[df_mapeamento['MODELO_NORMALIZADO_BUSCA'] == produto_orcamento_busca_normalizado]
            
            if not produto_mapeado_df.empty:
                produto_mapeado = produto_mapeado_df.iloc[0]
                id_produto_olist = produto_mapeado.get('ID', pd.NA)
                descricao_produto_olist = produto_mapeado.get('MODELO OLIST', pd.NA)
                print(f"[CONVERSOR V6] Produto '{produto_orcamento_busca_normalizado}' MAPEADO: ID Olist='{id_produto_olist}', Descrição Olist='{descricao_produto_olist}'", file=sys.stderr)
            else: 
                print(f"[CONVERSOR V6] Aviso: Produto '{produto_orcamento_busca_normalizado}' (Original: '{produto_orcamento_original}', linha item {index}) NÃO MAPEADO.", file=sys.stderr)
                produtos_nao_mapeados_log.append(f"'{produto_orcamento_busca_normalizado}' (Original: '{produto_orcamento_original}')")
        elif not produto_orcamento_busca_normalizado:
            print(f"[CONVERSOR V6] Aviso: Produto vazio na linha item {index} do orçamento.", file=sys.stderr)
        elif not ('MODELO_NORMALIZADO_BUSCA' in df_mapeamento.columns):
             print(f"[CONVERSOR V6] ERRO INTERNO: Coluna 'MODELO_NORMALIZADO_BUSCA' não existe em df_mapeamento.", file=sys.stderr)

        linha_convertida = {
            'Número da proposta': num_proposta_orc if num_proposta_orc is not None else pd.NA,
            'Data': data_proposta_orc if data_proposta_orc is not None else pd.NA,
            'ID contato': id_contato_cliente,
            'Nome do contato': nome_contato_cliente,
            'ID produto': id_produto_olist,
            'Descrição': descricao_produto_olist, 
            'Quantidade': qtde if pd.notna(qtde) else pd.NA,
            'Valor unitário': valor_unit if pd.notna(valor_unit) else pd.NA
        }
        
        for col in colunas_modelo_olist:
            if col not in linha_convertida:
                linha_convertida[col] = pd.NA
        
        linhas_saida.append(linha_convertida)
    
    if produtos_nao_mapeados_log:
        print(f"[CONVERSOR V6] RESUMO DE PRODUTOS NÃO MAPEADOS ({len(produtos_nao_mapeados_log)} itens): {', '.join(sorted(list(set(produtos_nao_mapeados_log))))}", file=sys.stderr)

    print(f"[CONVERSOR V6] {len(linhas_saida)} linhas de itens processadas para o DataFrame de saída.", file=sys.stderr)
    df_saida_final = pd.DataFrame(columns=colunas_modelo_olist)
    if linhas_saida:
        df_temp = pd.DataFrame(linhas_saida)
        df_saida_final = pd.concat([df_saida_final, df_temp], ignore_index=True)[colunas_modelo_olist]
    
    if not df_saida_final.empty:
        print(f"[CONVERSOR V6] DataFrame de Saída (primeiras linhas APÓS CONCAT):\n{df_saida_final.head().to_string()}", file=sys.stderr)
    else: 
        print("[CONVERSOR V6] DataFrame de Saída está vazio.", file=sys.stderr)
    return df_saida_final

if __name__ == '__main__': 
    pass

