from collections import defaultdict
import pandas as pd
import numpy as np
import json

def carregarDados():
    #Ler Planilha
    df_pedidos = pd.read_excel("001-base_completa.xlsx", sheet_name="base_pedidos", usecols="A:P",nrows=400000)
    return df_pedidos

def criarDataframeMcuSolicitante(df_pedidos):    
    # Criar dataframe especifico
    df_mcu_solicitante = df_pedidos.assign(MCU=lambda x: x['NUM_PEDIDO_SISCAP'].astype(str).str[:8],DT_TRANS_STRING=lambda x: x['DT_TRANS'].astype(str).str.split('-').str[::-1].str.join('/'))[['NUM_PEDIDO_SISCAP', 'MCU','DT_TRANS_STRING']].drop_duplicates()
    return df_mcu_solicitante

def criarDataframeMcuPossuiPedidos(df_pedidos):    
    # Criar dataframe especifico
    df_mcu_possui_pedidos = df_pedidos.assign(MCU=lambda x: x['NUM_PEDIDO_SISCAP'].astype(str).str[:8])[['MCU']].drop_duplicates()
    df_mcu_possui_pedidos["FORMATO_JSON"] = "{}"
    return df_mcu_possui_pedidos


def configurarDicionarioMcuSolicitante(df_mcu_solicitante):
    # Trtatar DataFrame
    dict_mcu_solicitante = defaultdict(list)
    for _, linha in df_mcu_solicitante.iterrows():
        id_ = str(linha["MCU"])
        item = {
            "num_pedido_siscap" : linha["NUM_PEDIDO_SISCAP"],
            "dt_trans_string" : linha["DT_TRANS_STRING"]               
        }
        dict_mcu_solicitante[id_].append(item)
    return dict_mcu_solicitante

def configurarDicionarioMcuPossuiPedidos(df_mcu_possui_pedidos):
    dict_mcu_possui_pedidos = defaultdict(list)
    for _, linha in df_mcu_possui_pedidos.iterrows():
        id_ = str(linha["MCU"])
        item = {
            "formato_json" : "{}"
        }
        dict_mcu_possui_pedidos[id_].append(item)
    return dict_mcu_possui_pedidos

# Função para converter objetos Timestamp para string ISO
def json_serializer(obj):
    """Serializa objetos Timestamp para string ISO format."""
    if isinstance(obj, pd.Timestamp):
        return obj.isoformat()
    elif isinstance(obj, (np.integer, np.floating)):
        return float(obj)
    raise TypeError(f"Type {type(obj)} not serializable")

# Função para construir o JSON para um pedido específico conforme estrutura especificada
def construir_json_pedido(mcu,dict_mcu_solicitante):

    #pedido_info = dict_mcu_solicitante.get(mcu, [{}])[0]

    estrutura_json = {
        "MCU" : mcu,
        "PEDIDO" : []
    }

    if mcu in dict_mcu_solicitante:
        for item in dict_mcu_solicitante[mcu]:
            estrutura_json["PEDIDO"].append({
                "NUM_PEDIDO_SISCAP" : str(item.get("num_pedido_siscap","")),
                "DT_TRANS_STRING" : str(item.get("dt_trans_string",""))
            })
    
    # Remover chaves vazias onde não há dados
    for chave in ["PEDIDO"]:
        if not estrutura_json[chave]:
            estrutura_json[chave] = []

    return estrutura_json

# Função para construir o JSON para um pedido específico conforme estrutura especificada
"""
def construir_json_possui_pedido(mcu,dict_mcu_possui_pedido):
    
    estrutura_json = {
        "MCU" : mcu,
        "PEDIDO" : []
    }

    if mcu in dict_mcu_possui_pedido:
        estrutura_json["PEDIDO"].append({
                "NUM_PEDIDO_SISCAP" : "{}"                
        })            
    
    # Remover chaves vazias onde não há dados
    for chave in ["PEDIDO"]:
        if not estrutura_json[chave]:
            estrutura_json[chave] = []

    return estrutura_json
"""

def atualizarDicionarioComDadosJson(dict_mcu_possui_pedido, dict_mcu_solicitante):
    try:
        for mcu in dict_mcu_possui_pedido.keys():
            # construir json para este pedido
            json_data = construir_json_pedido(mcu,dict_mcu_solicitante)

            # Converter para string JSON
            json_str = json.dumps(json_data, default=json_serializer, indent=2, ensure_ascii=False)
        
            # Atualizar o dicionario dict_mcu_possui_pedido:
            if dict_mcu_possui_pedido[mcu]:
                dict_mcu_possui_pedido[mcu][0]["formato_json"] = json_str
        
        return dict_mcu_possui_pedido
    except Exception as e:
        if dict_mcu_possui_pedido[mcu]:
            dict_mcu_possui_pedido[mcu][0]["formato_json"] = "{}"

def criarDataframeParaPlanilhaExcel(dict_mcu_possui_pedido):
    # Criar DataFrame a partir do dicionário atualizado
    df_resultado = pd.DataFrame([
        {
            "MCU": mcu,
            "FORMATO_JSON": dict_mcu_possui_pedido[mcu][0]["formato_json"] if dict_mcu_possui_pedido[mcu] else "{}"
        }
        for mcu in dict_mcu_possui_pedido.keys()
    ])
    return df_resultado

def criarTabelaExcel(worksheet, nome_tabela, nome_exibicao=None):
    """
    Cria uma tabela formatada no Excel.
    """
    from openpyxl.worksheet.table import Table, TableStyleInfo
    
    max_row = worksheet.max_row
    max_col = worksheet.max_column
    
    if max_row > 1:  # Se há dados além do cabeçalho
        ref_range = f"A1:{chr(64 + max_col)}{max_row}"
        
        tab = Table(displayName=nome_tabela, ref=ref_range)
        
        # Estilo moderno
        style = TableStyleInfo(
            name="TableStyleMedium2",  # Estilo azul
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        
        tab.tableStyleInfo = style
        
        # Adicionar nome de exibição se fornecido
        if nome_exibicao:
            tab.name = nome_exibicao
        
        worksheet.add_table(tab)
        
        # Ajustar largura das colunas automaticamente
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

def exportarParaExcel(df_resultado, nome_da_planilha, nome_da_tabela, nome_de_exibicao, nome_do_arquivo):
    # Exportar DataFrame para Excel com formatação de tabela
    with pd.ExcelWriter(nome_do_arquivo, engine='openpyxl') as writer:
        df_resultado.to_excel(writer, sheet_name=nome_da_planilha, index=False)
        worksheet = writer.sheets[nome_da_planilha]
        criarTabelaExcel(worksheet, nome_da_tabela, nome_de_exibicao)

if __name__ == "__main__":
    # Ler planilha excel
    df_pedidos = carregarDados()        
    #cCriar Dataframe
    df_mcu_solicitante = criarDataframeMcuSolicitante(df_pedidos)    
    # Criar Dataframe
    df_mcu_possui_pedidos = criarDataframeMcuPossuiPedidos(df_pedidos)    
    # Criar dicionario
    dict_mcu_solicitante = configurarDicionarioMcuSolicitante(df_mcu_solicitante)    
     # Criar dicionario
    dict_mcu_possui_pedido = configurarDicionarioMcuPossuiPedidos(df_mcu_possui_pedidos)    
    #atualizado a coluna formato_json
    dict_mcu_possui_pedido = atualizarDicionarioComDadosJson(dict_mcu_possui_pedido, dict_mcu_solicitante)
    # Criar DataFrame para exportar para Excel
    df_resultado = criarDataframeParaPlanilhaExcel(dict_mcu_possui_pedido)    
    # Exportar para Excel
    nomeDaPlanilha = "ListaMcu"
    nomeDaTabela = "TabelaMcu"
    nomeDeExibicao = "Tabela de MCU com Pedidos"
    nomeDoArquivo = "pedidos_para_agente_ia_mcu.xlsx"    
    exportarParaExcel(df_resultado, nomeDaPlanilha, nomeDaTabela, nomeDeExibicao, nomeDoArquivo)