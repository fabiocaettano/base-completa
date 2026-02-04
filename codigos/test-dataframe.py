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

def teste(dict_mcu_possui_pedido, dict_mcu_solicitante):
    for mcu in dict_mcu_possui_pedido.keys():
        # construir json para este pedido
        json_data = construir_json_pedido(mcu,dict_mcu_solicitante)

        # Converter para string JSON
        json_str = json.dumps(json_data, default=json_serializer, indent=2, ensure_ascii=False)
        
        # Atualizar o dicionario dict_mcu_possui_pedido:
        if dict_mcu_possui_pedido[mcu]:
            dict_mcu_possui_pedido[mcu][0]["formato_json"] = json_str
        
        return dict_mcu_possui_pedido


if __name__ == "__main__":
    # Ler planilha excel
    df_pedidos = carregarDados()    
    #cCriar Dataframe
    df_mcu_solicitante = criarDataframeMcuSolicitante(df_pedidos)
    df_mcu_possui_pedidos = criarDataframeMcuPossuiPedidos(df_pedidos)
    #criar dicionario
    dict_mcu_solicitante = configurarDicionarioMcuSolicitante(df_mcu_solicitante)
    dict_mcu_possui_pedido = configurarDicionarioMcuPossuiPedidos(df_mcu_possui_pedidos)
    #atualizado a coluna formato_json
    dict_mcu_possui_pedido = teste(dict_mcu_possui_pedido, dict_mcu_solicitante)
    print(dict_mcu_possui_pedido)
    