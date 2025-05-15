import pandas as pd
import sys
import traceback
import re # Para normalização
import os
import io
from typing import Union, BinaryIO

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

def converter_orcamento_para_olist(
    arquivo_orcamento: Union[str, BinaryIO],
    caminho_mapeamento_produtos: str,
    caminho_clientes: str,
    id_cliente_selecionado: Union[str, int],
    caminho_modelo_saida_olist_com_dados: str
) -> pd.DataFrame:
    """
    Converte um arquivo de orçamento para o formato Olist.
    
    Args:
        arquivo_orcamento: Caminho do arquivo ou objeto BytesIO contendo o orçamento
        caminho_mapeamento_produtos: Caminho do arquivo de mapeamento de produtos
        caminho_clientes: Caminho do arquivo de clientes
        id_cliente_selecionado: ID do cliente selecionado
        caminho_modelo_saida_olist_com_dados: Caminho do arquivo modelo de saída
        
    Returns:
        DataFrame com o orçamento convertido no formato Olist
    """
    df_modelo_saida_temp = None
    colunas_modelo_olist = []
    produtos_nao_mapeados_log = [] # Lista para logar produtos não mapeados
    
    # Verificar se os arquivos existem antes de tentar abri-los
    for arquivo, descricao in [
        (caminho_orcamento, "orçamento"),
        (caminho_mapeamento_produtos, "mapeamento de produtos"),
        (caminho_clientes, "clientes"),
        (caminho_modelo_saida_olist_com_dados, "modelo de saída")
    ]:
        if not os.path.exists(arquivo):
            erro_msg = f"Arquivo de {descricao} não encontrado: {arquivo}"
            print(f"[CONVERSOR V6] ERRO: {erro_msg}", file=sys.stderr)
            raise FileNotFoundError(erro_msg)

    print(f"[CONVERSOR V6] Iniciando conversão. Cliente ID: {id_cliente_selecionado}", file=sys.stderr)
    try:
        print(f"[CONVERSOR V6] Lendo arquivo de mapeamento: {caminho_mapeamento_produtos}", file=sys.stderr)
        with pd.ExcelFile(caminho_mapeamento_produtos) as xls_map:
            df_mapeamento = pd.read_excel(xls_map, sheet_name='CATÁLOGO')
        
        # Normalizar a coluna de busca no mapeamento
        if 'MODELO' in df_mapeamento.columns:
            df_mapeamento['MODELO_NORMALIZADO_BUSCA'] = df_mapeamento['MODELO'].apply(normalizar_texto)
            print(f"[CONVERSOR V6] Coluna 'MODELO' normalizada para busca em df_mapeamento.", file=sys.stderr)
        else:
            print(f"[CONVERSOR V6] ERRO: Coluna 'MODELO' não encontrada em {caminho_mapeamento_produtos}", file=sys.stderr)
            return pd.DataFrame(columns=colunas_modelo_olist if colunas_modelo_olist else [])

        print(f"[CONVERSOR V6] Lendo arquivo de clientes: {caminho_clientes}", file=sys.stderr)
        with pd.ExcelFile(caminho_clientes) as xls_cli:
            df_clientes = pd.read_excel(xls_cli, sheet_name='CLIENTES')
        
        print(f"[CONVERSOR V6] Lendo NOVO arquivo modelo de saída com dados: {caminho_modelo_saida_olist_com_dados}", file=sys.stderr)
        with pd.ExcelFile(caminho_modelo_saida_olist_com_dados) as xls_modelo_novo:
            if xls_modelo_novo.sheet_names:
                df_modelo_saida_temp = pd.read_excel(xls_modelo_novo, sheet_name=0)
                print(f"[CONVERSOR V6] Lida a primeira aba do NOVO modelo de saída: {xls_modelo_novo.sheet_names[0]}", file=sys.stderr)
            else:
                raise ValueError("O NOVO arquivo Excel modelo de saída não contém nenhuma aba.")
            colunas_modelo_olist = df_modelo_saida_temp.columns.tolist()
            print(f"[CONVERSOR V6] Colunas do NOVO modelo Olist: {colunas_modelo_olist}", file=sys.stderr)

        print(f"[CONVERSOR V6] Lendo arquivo de orçamento", file=sys.stderr)
        if isinstance(arquivo_orcamento, (str, bytes, io.BytesIO)):
            xls_orc = pd.ExcelFile(arquivo_orcamento)
        else:
            raise ValueError("Formato de arquivo de orçamento inválido")
            
        sheet_name_orcamento = 'Orçamento' if 'Orçamento' in xls_orc.sheet_names else xls_orc.sheet_names[0]
        
        # Leitura dos metadados do orçamento
        df_orc_preview_meta = pd.read_excel(xls_orc, sheet_name=sheet_name_orcamento, nrows=10, header=None)
        num_proposta_orc = None
        data_proposta_orc = None
        
        # Extração de metadados
        if len(df_orc_preview_meta) > 0 and len(df_orc_preview_meta.columns) > 1:
            for i in range(len(df_orc_preview_meta)):
                if normalizar_texto(df_orc_preview_meta.iloc[i, 0]) == "orçamento #":
                    num_proposta_orc = df_orc_preview_meta.iloc[i, 1]
                if normalizar_texto(df_orc_preview_meta.iloc[i, 0]) == "data":
                    data_proposta_orc = df_orc_preview_meta.iloc[i, 1]
                    if isinstance(data_proposta_orc, pd.Timestamp):
                        data_proposta_orc = data_proposta_orc.date()
                    elif isinstance(data_proposta_orc, str):
                        try:
                            data_proposta_orc = pd.to_datetime(data_proposta_orc, dayfirst=True).date()
                        except ValueError:
                            try:
                                data_proposta_orc = pd.to_datetime(data_proposta_orc).date()
                            except ValueError:
                                pass
        
        # Identificação do cabeçalho dos itens
        palavras_chave_cabecalho_itens = ["Produto", "Cor", "Qualidade", "Valor Unitário", "Quantidade", "Subtotal"]
        linha_cabecalho_itens_idx = encontrar_linha_cabecalho(df_orc_preview_meta, palavras_chave_cabecalho_itens)
        
        if linha_cabecalho_itens_idx is not None:
            df_orcamento_itens = pd.read_excel(xls_orc, sheet_name=sheet_name_orcamento, header=linha_cabecalho_itens_idx)
        else:
            df_orcamento_itens = pd.read_excel(xls_orc, sheet_name=sheet_name_orcamento, skiprows=2)
        
        # Normalização das colunas
        df_orcamento_itens.columns = [normalizar_texto(col) for col in df_orcamento_itens.columns]
        
        # Busca informações do cliente
        info_cliente_df = pd.DataFrame()
        if not df_clientes.empty and 'ID' in df_clientes.columns:
            try:
                coluna_id_tipo = df_clientes['ID'].dtype
                id_cliente_selecionado_str = str(id_cliente_selecionado)
                
                if pd.api.types.is_numeric_dtype(coluna_id_tipo):
                    try:
                        id_cliente_convertido = int(float(id_cliente_selecionado_str))
                    except ValueError:
                        id_cliente_convertido = float(id_cliente_selecionado_str)
                else:
                    id_cliente_convertido = id_cliente_selecionado_str
                    
                info_cliente_df = df_clientes[df_clientes['ID'] == id_cliente_convertido]
            except Exception as e:
                print(f"[CONVERSOR V6] Erro ao buscar cliente: {str(e)}", file=sys.stderr)
        
        if info_cliente_df.empty:
            raise ValueError(f"Cliente com ID '{id_cliente_selecionado}' não encontrado")
        
        info_cliente = info_cliente_df.iloc[0]
        id_contato_cliente = info_cliente['ID']
        nome_contato_cliente = info_cliente['Nome']
        
        # Processamento dos itens
        linhas_saida = []
        for index, linha_item in df_orcamento_itens.iterrows():
            produto_orcamento_original = linha_item.get('produto')
            qtde = linha_item.get('quantidade')
            valor_unit = linha_item.get('valor unitário')
            
            if pd.isna(produto_orcamento_original) and pd.isna(qtde) and pd.isna(valor_unit):
                continue
                
            produto_orcamento_busca_normalizado = normalizar_texto(produto_orcamento_original)
            
            id_produto_olist = pd.NA
            descricao_produto_olist = pd.NA
            
            if produto_orcamento_busca_normalizado:
                produto_mapeado_df = df_mapeamento[
                    df_mapeamento['MODELO_NORMALIZADO_BUSCA'] == produto_orcamento_busca_normalizado
                ]
                
                if not produto_mapeado_df.empty:
                    produto_mapeado = produto_mapeado_df.iloc[0]
                    id_produto_olist = produto_mapeado.get('ID', pd.NA)
                    descricao_produto_olist = produto_mapeado.get('MODELO OLIST', pd.NA)
                else:
                    produtos_nao_mapeados_log.append(
                        f"'{produto_orcamento_busca_normalizado}' (Original: '{produto_orcamento_original}')"
                    )
            
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
            
            linhas_saida.append({col: linha_convertida.get(col, pd.NA) for col in colunas_modelo_olist})
        
        if produtos_nao_mapeados_log:
            print("[CONVERSOR V6] Produtos não mapeados:", file=sys.stderr)
            for produto in produtos_nao_mapeados_log:
                print(f"  - {produto}", file=sys.stderr)
        
        return pd.DataFrame(linhas_saida)
        
    except Exception as e:
        print(f"[CONVERSOR V6] Erro: {str(e)}\n{traceback.format_exc()}", file=sys.stderr)
        return pd.DataFrame(columns=colunas_modelo_olist if colunas_modelo_olist else [])

if __name__ == '__main__': 
    pass

