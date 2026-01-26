import pandas as pd
from collections import defaultdict
import json
from datetime import datetime
import os
import sys
import time
import numpy as np

# Sere como um script de análise de dados para processar informações de pedidos, lotes e objetos postados a partir de uma planilha Excel.
# Serão 05 etapas principais:
# Etapa 01: Carregar os dados das planilhas Excel
# Etapa 02: Ajustar o nome das colunas chaves nos DataFrames 
# Etapa 03: Configurar os DataFrames para análise
# Etapa 04: Criar dicionários com os dados processados
# Etapa 05: Gerar estrutura JSON para integração com API (futuro)
# Etapa 06: Exportar os dados processados para novos arquivos Excel/CSV (futuro)

# Etapa 00


def verificar_e_excluir_arquivo_anterior():
    """
    Verifica se o arquivo pedidos_unicos_com_json.xlsx existe e tenta excluí-lo.
    Se não conseguir excluir, aborta o processamento.
    """
    arquivo_antigo = "pedidos_unicos_com_json.xlsx"
    
    if os.path.exists(arquivo_antigo):
        print(f"⚠️  Arquivo anterior encontrado: {arquivo_antigo}")
        print("Tentando excluir o arquivo anterior...")
        
        try:
            # Tentar fechar qualquer handle que possa estar aberto
            import gc
            gc.collect()
            
            # Tentar excluir o arquivo
            os.remove(arquivo_antigo)
            print(f"✅ Arquivo {arquivo_antigo} excluído com sucesso!")
            return True
            
        except PermissionError:
            print(f"❌ ERRO: O arquivo {arquivo_antigo} está aberto em outro programa.")
            print("Por favor, feche o Excel e tente novamente.")
            return False
            
        except Exception as e:
            print(f"❌ ERRO ao excluir arquivo: {str(e)}")
            return False
    else:
        print(f"✅ Nenhum arquivo anterior encontrado. Pode continuar.")
        return True

# Usar no início do seu script principal
print("="*60)
print("VERIFICAÇÃO DE ARQUIVO ANTERIOR")
print("="*60)

if not verificar_e_excluir_arquivo_anterior():
    print("\n❌ ABORTANDO PROCESSAMENTO!")
    print("Motivo: Não foi possível excluir o arquivo anterior.")
    sys.exit(1)  # Encerra o programa com código de erro

print("="*60)
print("✅ INICIANDO PROCESSAMENTO...")
print("="*60)


# Etapa 01: Carregar os dados das planilhas Excel

## 01.01 Carregar os dados da planilha base_pedidos
df_pedidos = pd.read_excel("001-base_completa.xlsx",
                    sheet_name="base_pedidos",
                    usecols="A:P",nrows=400000)
print(f"✅Total de Registros na Planilha BASE_PEDIDOS: {len(df_pedidos)}")

## 01.02 Carregar os dados da planilha base_lotes
df_lotes = pd.read_excel("001-base_completa.xlsx",
                         sheet_name="base_lotes",
                         usecols="A:Q",
                         nrows=200000)
print(f"✅Total de Registros na Planilha BASE_LOTES: {len(df_lotes)}")

## 01.03 Carregar os dados da planilha base_objeto_postado
df_objetos = pd.read_excel("001-base_completa.xlsx",
                           sheet_name="base_objeto_postado",
                           usecols="A:L",
                           nrows=200000)
print(f"✅Total de Registros na Planilha BASE_OBJSETO_POSTADO: {len(df_objetos)}")


# 01.04 Carregar os dados da planilha base_mcco
df_base_mcco = pd.read_excel("001-base_completa.xlsx",
                           sheet_name="base_mcco",
                           usecols="A:B",
                           nrows=30)
print(f"✅Total de Registro da Planilha BASE_MCCO : {len(df_base_mcco) }")
print(f'='*60)
print(f'')

# Etapa 02: Ajustar o nome das colunas chaves nos DataFrames
df_lotes = df_lotes.rename(columns={
    'lote_pedido': 'LOTE',
    'pedido': 'PEDIDO', 
    'tipo': 'TIPO_PEDIDO',
    'zona': 'ZONA'
})

df_pedidos = df_pedidos.rename(columns={
    'PED': 'PEDIDO',    
    'TP_PED': 'TIPO_PEDIDO'
})

print(f"✅ Ajustando nome das colunas para relacionamento")
print(f'='*60)
print(f'')

# Etapa 03: Configurar os DataFrames para análise

## 03.01 Criar DataFrame com pedidos únicos, filtrar aqueles que começam com '0' e exibir algumas informações
pedidos_unicos = df_pedidos.assign(MCU=lambda x: x['NUM_PEDIDO_SISCAP'].astype(str).str[:8])[['NUM_PEDIDO_SISCAP', 'PEDIDO', 'TIPO_PEDIDO','DT_TRANS', 'MCCO','CLIENTE','NOME_CLI', 'MCU']].drop_duplicates()
df_pedidos_unicos = pedidos_unicos[
    pedidos_unicos['NUM_PEDIDO_SISCAP'].astype(str).str.startswith('0')
]

# Excluir colunas
df_pedidos_unicos = df_pedidos_unicos.drop(["PEDIDO","TIPO_PEDIDO"],axis=1).drop_duplicates()

# merge df_pedidos_unicos com df_base_mcco
df_pedidos_unicos = pd.merge(
    df_pedidos_unicos,
    df_base_mcco[["MCCO","DESCRICAO_MCCO"]],
    on=["MCCO"],
    how="inner"
)

# Incluir coluna Formato JSON
df_pedidos_unicos["FORMATO_JSON"] = "{}"


print(f"✅Pedidos que começam com 0: {len(pedidos_unicos)}")
#print(pedidos_unicos.head())
print(f'='*60)
print(f'')


##03.02  Criar DataFrame somente com pedidos novos (PROX_STATUS = 540)
df_pedidos_novos = df_pedidos[df_pedidos["PROX_STATUS"] == 540]
print(f"✅Total de Pedidos Novos: {len(df_pedidos_novos)}")

df_pedidos_novos_agrupados = df_pedidos_novos.groupby(['NUM_PEDIDO_SISCAP', 'ITEM','DESC_ITEM' ,'UN_MEDIDA'])['QTD'].sum().reset_index().rename(columns={'QTD': 'QTD_TOTAL'})
#print(f"Total de Pedidos Novos Agrupados: {len(df_pedidos_novos_agrupados)}")

#print(df_pedidos_novos_agrupados.head())

#print(f'---'*40)

## 03.03  Criar DataFrame somente com pedidos cancelados
df_pedidos_cancelados = df_pedidos[df_pedidos["ULT_STATUS"].isin([980,982])]
print(f"✅Total de Pedidos Cancelados: {len(df_pedidos_cancelados)}")

df_pedidos_cancelados_agrupados = df_pedidos_cancelados.groupby(['NUM_PEDIDO_SISCAP', 'ITEM','DESC_ITEM' ,'UN_MEDIDA'])['QTD'].sum().reset_index().rename(columns={'QTD': 'QTD_TOTAL'})
#print(f"Total de Pedidos Cancelados Agrupados: {len(df_pedidos_cancelados_agrupados)}")

#print(df_pedidos_cancelados_agrupados.head())

#print(f'---'*40)

## 03.04 Criar DataFrame pedidos em tratamento, ou seja, os que foram vinculados a lotes

# 03.04.01 Exclu ir do DataFrame df_lotes os status 291 e 250 e tipo ZR
df_lotes = df_lotes.drop(df_lotes[df_lotes["status_lote"] == 291].index)
df_lotes = df_lotes.drop(df_lotes[df_lotes["status_lote"] == 250].index)
df_lotes = df_lotes.drop(df_lotes[df_lotes["TIPO_PEDIDO"] == "ZR"].index)

## 03.04.02  Criar coluna do numero da subpauta
df_lotes["SUBPAUTA"] = (df_lotes["LOTE"].astype(str) + df_lotes["PEDIDO"].astype(str) + df_lotes["TIPO_PEDIDO"].astype(str) + df_lotes["ZONA"].astype(str)).astype(str)

## 03.04.03 Merge entre os dados dos lotes e pedidos
df_lotes_relacionados_com_df_pedidos = pd.merge(
    df_lotes,
    df_pedidos[['PEDIDO', 'TIPO_PEDIDO','NUM_PEDIDO_SISCAP','DT_TRANS']],
    on=['PEDIDO', 'TIPO_PEDIDO'],
    how='inner'
)

#print(f"Total de Registros na Planilha BASE_LOTES relacionados com BASE_PEDIDOS antes do tratamento: {len(df_lotes_relacionados_com_df_pedidos)}")
#print(df_lotes_relacionados_com_df_pedidos.head())
#print(f'---'*40)

## 03.04.04 Eilminar duplicidades e ajustar valores NaN na coluna status_impressao
df_lotes_relacionados_com_df_pedidos = df_lotes_relacionados_com_df_pedidos.drop_duplicates()
df_lotes_relacionados_com_df_pedidos = df_lotes_relacionados_com_df_pedidos.replace({'status_impressao': {"": 'III', None: 'III',"NAN": 'III'}})
#print(f"Total de Registros na Planilha BASE_LOTES relacionados com BASE_PEDIDOS: {len(df_lotes_relacionados_com_df_pedidos)}")
#print(df_lotes_relacionados_com_df_pedidos.head())
#print(f'---'*40)

df_lote_em_tratamento = df_lotes_relacionados_com_df_pedidos.groupby(['NUM_PEDIDO_SISCAP','LOTE','status_impressao' ,'item', 'descricao', 'um'])['qtde'].sum().reset_index().rename(columns={'qtde': 'QTD_TOTAL'})
#print(f"Total de Registros de Lotes em Tratamento no WMS: {len(df_lote_em_tratamento)}")

# 03.04.05 Dataframe com lotes sem impressao de pauta (status III)
df_lote_em_tratamento_status_iii = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "III"]
print(f"✅Total de Registros de Lotes em Tratamento no WMS com status III : {len(df_lote_em_tratamento_status_iii)}")
#print(df_lote_em_tratamento_status_iii.head())
#print(f'')

# 03.04.06 Dataframe com lotes com impressao de pauta (status I)
df_lote_em_tratamento_status_i = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "I"]
print(f"✅Total de Registros de Lotes em Tratamento no WMS com status I : {len(df_lote_em_tratamento_status_i)}")
#print(df_lote_em_tratamento_status_i.head())
#print(f'---'*40)
#print(f'')

# 03.04.07 Dataframe com lotes com etiqueta gerada (status II)
df_lote_em_tratamento_status_ii = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "II"]
print(f"✅Total de Registros de Lotes em Tratamento no WMS com status II : {len(df_lote_em_tratamento_status_ii)}")
#print(df_lote_em_tratamento_status_ii.head())
#print(f'---'*40)


## 03.04.08 Concatenar as colunas nota_fiscal e serie em df_objetos

# Preencher NaN com 0 ou outro valor
#df_objetos['NOTA_FISCAL'] = df_objetos['NOTA_FISCAL'].fillna(0)
# Filtrar apenas linhas com NOTA_FISCAL não nulo
df_objetos = df_objetos.dropna(subset=['NOTA_FISCAL']).copy()

# Converter para string (removendo .0 se quiser)
df_objetos['NOTA_FISCAL_STR'] = df_objetos['NOTA_FISCAL'].astype(int).astype(str)

# Criar NOTA_SERIE com zfill
df_objetos['NOTA_SERIE'] = (
    df_objetos['NOTA_FISCAL_STR'] + 
    df_objetos['SERIE'].astype(str).str.zfill(2)
).astype('float64')


## 03.04.09 Merge do dataframe com Lotes em tratamento com o dataframe de objeotos postados

df_lote_possui_objeto_1 = pd.merge(
    df_lotes_relacionados_com_df_pedidos,
    df_objetos[['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA', 'NOTA_SERIE']],
    on=['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA'],
    how='inner'
)

# Eliminar duplicidade para não impactar na quantidade agruapada
df_lote_possui_objeto_1 =  df_lote_possui_objeto_1.drop(["status_lote","status_impressao","posto","lote_consolidado","cep","roteirizador","cod_cliente","descricao_cliente","se","PEDIDO","TIPO_PEDIDO"], axis=1).drop_duplicates()
# Agrupar quantidade atendida
df_lote_possui_objeto_agrupar_itens = df_lote_possui_objeto_1.groupby(['NUM_PEDIDO_SISCAP','SUBPAUTA','ZONA','NOTA_SERIE','item','descricao','um'])['qtde'].sum().reset_index().rename(columns={'qtde': 'QTD_TOTAL'})
#print(df_lote_possui_objeto_agrupar_itens.head())
#print(f'---'*40)



# 03.04.10 Dataframe dos objetos para rastremento

df_lote_possui_objeto_2 = pd.merge(
    df_lotes_relacionados_com_df_pedidos,
    df_objetos[['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA', 'NOTA_SERIE','REGISTRO', 'DATA_EXPEDICAO']],
    on=['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA'],
    how='inner'
)

# 03.04.11 Remover colunas desnecessárias para unificação de rastreamento
df_lote_possui_objeto_unificar_rastreamento = df_lote_possui_objeto_2.drop(["status_lote","status_impressao","posto","lote_consolidado","cep","roteirizador","cod_cliente","descricao_cliente","se","PEDIDO","TIPO_PEDIDO","item","descricao","um","qtde"], axis=1).drop_duplicates()
#print(f'Total de registros entre lotes e objetos unificados por NUM_PEDIDO_SISCAP, ZONA e NOTA_SERIE: {len(df_lote_possui_objeto_unificar_rastreamento)}')
#print(df_lote_possui_objeto_unificar_rastreamento.head())
#print(f'---'*40)

print(f'='*60)
print(f'')

# Etapa 04: Criar dicionários com os dados processados, onde a chave é o NUM_PEDIDO_SISCAP para toodos os dicionários


## 04.01  Criar o dict com o dataframe dos pedidos novos
print(f'✅Criando dicionário de pedidos novos...')
dict_pedidos_novos = defaultdict(list)
for _, linha in df_pedidos_novos_agrupados.iterrows():
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "item": linha["ITEM"],
          "desc_item": linha["DESC_ITEM"],
          "un_medida": linha["UN_MEDIDA"],
          "qtd_total": linha["QTD_TOTAL"]          
    } 
    dict_pedidos_novos[id_].append(item)

#print(dict_pedidos_novos)
#print(f'---'*40)
#print(f'')

## 04.02  Criar o dict com o dataframe dos pedidos cancelados
print(f'✅Criando dicionário de pedidos cancelados...')
dict_pedidos_cancelados =defaultdict(list)  
for _, linha in df_pedidos_cancelados_agrupados.iterrows():
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "item": linha["ITEM"],
          "desc_item": linha["DESC_ITEM"],
          "un_medida": linha["UN_MEDIDA"],
          "qtd_total": linha["QTD_TOTAL"]          
    } 
    dict_pedidos_cancelados[id_].append(item)
#print(dict_pedidos_cancelados)
#print(f'---'*40)
#print(f'')

## 04.03  Criar o dict com o dataframe dos lotes em tratamento status III
print(f'✅Criando dicionário de lotes em tratamento status III...')  
dict_lotes_tratamento_iii =defaultdict(list)  
for _, linha in df_lote_em_tratamento_status_iii.iterrows():
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "lote": linha["LOTE"],
          "item": linha["item"],
          "descricao": linha["descricao"],
          "um": linha["um"],
          "qtd_total": linha["QTD_TOTAL"]          
    } 
    dict_lotes_tratamento_iii[id_].append(item)
#print(dict_lotes_tratamento_iii)
#print(f'---'*40)    
#print(f'')

## 04.04  Criar o dict com o dataframe dos lotes em tratamento status I
print(f'✅Criando dicionário de lotes em tratamento status I...') 
dict_lotes_tratamento_i =defaultdict(list)  
for _, linha in df_lote_em_tratamento_status_i.iterrows():  
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "lote": linha["LOTE"],
          "item": linha["item"],
          "descricao": linha["descricao"],
          "um": linha["um"],
          "qtd_total": linha["QTD_TOTAL"]          
    } 
    dict_lotes_tratamento_i[id_].append(item)
#print(dict_lotes_tratamento_i)
#print(f'---'*40)    
#print(f'') 

## 04.05  Criar o dict com o dataframe do df_lote_possui_objeto_agrupar_iten
print(f'✅Criando dicionário de lotes com objetos postados com itens agrupados...') 
dict_lotes_objetos_itens =defaultdict(list)
for _, linha in df_lote_possui_objeto_agrupar_itens.iterrows():  
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {          
          "subpauta": linha["SUBPAUTA"],
          "zona": linha["ZONA"],
          "nota_serie": linha["NOTA_SERIE"],
          "item": linha["item"],
          "descricao": linha["descricao"],
          "um": linha["um"],
          "qtd_total": linha["QTD_TOTAL"]          
    } 
    dict_lotes_objetos_itens[id_].append(item)
#print(dict_lotes_objetos_itens)
#print(f'---'*40)
#print(f'')

## 04.06  Criar o dict com o dataframe do df_lote_possui_objeto_unificar_rastreamento
print(f'✅Criando dicionário de lotes com objetos postados unificados para rasteamento...') 
dict_lotes_objetos_rastreamento =defaultdict(list)      
for _, linha in df_lote_possui_objeto_unificar_rastreamento.iterrows():  
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {                    
          "subpauta": linha["SUBPAUTA"],
          "zona": linha["ZONA"],
          "nota_serie": linha["NOTA_SERIE"],
          "registro": linha["REGISTRO"],
          "data_expedicao": linha["DATA_EXPEDICAO"]          
    } 
    dict_lotes_objetos_rastreamento[id_].append(item)   
#print(dict_lotes_objetos_rastreamento)
#print(f'---'*40)
#print(f'')

## 04.07 Criar dicionário dos pedidos únicos
print(f'✅Criando dicionário dos pedidos únicos')
dict_pedidos_unicos = defaultdict(list)
for _, linha in df_pedidos_unicos.iterrows():
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
        "dt_trans" : linha["DT_TRANS"],
        "cliente" : linha["CLIENTE"],
        "nome_cli" : linha["NOME_CLI"],
        "mcu" : linha["MCU"],
        "descricao_mcco" : linha["DESCRICAO_MCCO"],
        "fomrato_json" : linha["FORMATO_JSON"]
    }
    dict_pedidos_unicos[id_].append(item)
#print(dict_pedidos_unicos)
#print(f'---'*40)
#print(f'')

# Etapa 05: Gerar estrutura JSON para integração com bot
print(f'⚠️Criando estrutura JSON para cada pedido...')

import json
from datetime import datetime

# Função para converter objetos Timestamp para string ISO
def json_serializer(obj):
    """Serializa objetos Timestamp para string ISO format."""
    if isinstance(obj, pd.Timestamp):
        return obj.isoformat()
    raise TypeError(f"Type {type(obj)} not serializable")

# Função para construir o JSON para um pedido específico
def construir_json_pedido(num_pedido_siscap, dict_pedido_unico, dict_pedidos_novos, 
                         dict_pedidos_cancelados, dict_lotes_tratamento_iii, 
                         dict_lotes_tratamento_i, dict_lotes_objetos_itens, 
                         dict_lotes_objetos_rastreamento):
    
    # Obter dados básicos do pedido
    pedido_info = dict_pedido_unico.get(num_pedido_siscap, [{}])[0]
    
    # Estrutura JSON base conforme especificação
    estrutura_json = {
        "NUM_PEDIDO_SISCAP": num_pedido_siscap,
        "PEDIDO": {
            "DT_TRANS": pedido_info.get("dt_trans"),
            "MCU": pedido_info.get("mcu"),
            "CLIENTE": pedido_info.get("cliente"),
            "NOME_CLI": pedido_info.get("nome_cli"),
            "DESCRICAO_MCCO": pedido_info.get("descricao_mcco")
        },
        "RESERVADO": [],
        "CANCELADO": [],
        "PROCESSAMENTO": [],
        "PAUTA_IMPRESSA": [],
        "ATENDIDO": []  # Lista de objetos, cada um com NOTA_SERIE, SUBPAUTA, ZONA e ITENS
    }
    
    # Preencher RESERVADO (pedidos novos)
    for item in dict_pedidos_novos.get(num_pedido_siscap, []):
        estrutura_json["RESERVADO"].append({
            "ITEM": item.get("item"),
            "DESCRICAO_ITEM": item.get("desc_item"),
            "UN_MEDIDA": item.get("un_medida"),
            "QTDE": item.get("qtd_total")
        })
    
    # Preencher CANCELADO
    for item in dict_pedidos_cancelados.get(num_pedido_siscap, []):
        estrutura_json["CANCELADO"].append({
            "ITEM": item.get("item"),
            "DESCRICAO_ITEM": item.get("desc_item"),
            "UN_MEDIDA": item.get("un_medida"),
            "QTDE": item.get("qtd_total")
        })
    
    # Preencher PROCESSAMENTO (lotes status III)
    for item in dict_lotes_tratamento_iii.get(num_pedido_siscap, []):
        estrutura_json["PROCESSAMENTO"].append({
            "ITEM": item.get("item"),
            "DESCRICAO_ITEM": item.get("descricao"),
            "UN_MEDIDA": item.get("um"),
            "QTDE": item.get("qtd_total")
        })
    
    # Preencher PAUTA_IMPRESSA (lotes status I)
    for item in dict_lotes_tratamento_i.get(num_pedido_siscap, []):
        estrutura_json["PAUTA_IMPRESSA"].append({
            "ITEM": item.get("item"),
            "DESCRICAO_ITEM": item.get("descricao"),
            "UN_MEDIDA": item.get("um"),
            "QTDE": item.get("qtd_total")
        })
    
    # Preencher ATENDIDO conforme especificação
    # Estrutura: Lista de objetos, cada um com:
    # - NOTA_SERIE
    # - SUBPAUTA  
    # - ITENS { SKU: [...], RASTREAMENTO: [...] }
    
    subpauta_agrupadas = {}
    
    # Agrupar itens por nota_serie
    for item in dict_lotes_objetos_itens.get(num_pedido_siscap, []):
        subpauta = item.get("subpauta")
        if subpauta not in subpauta_agrupadas:
            subpauta_agrupadas[subpauta] = {                
                "SUBPAUTA": subpauta,
                "ZONA": item.get("zona"),
                "NOTA_SERIE":item.get("nota_serie"),                
                "ITENS": {
                    "SKU": [],
                    "RASTREAMENTO": []
                }
            }
        
        subpauta_agrupadas[subpauta]["ITENS"]["SKU"].append({
            "ITEM": item.get("item"),
            "DESCRICAO_ITEM": item.get("descricao"),
            "UN_MEDIDA": item.get("um"),
            "QTDE": item.get("qtd_total")
        })
    
    # Adicionar registros de rastreamento
    for rastreamento in dict_lotes_objetos_rastreamento.get(num_pedido_siscap, []):
        subpauta = rastreamento.get("subpauta")
        if subpauta in subpauta_agrupadas:
            subpauta_agrupadas[subpauta]["ITENS"]["RASTREAMENTO"].append({
                "REGISTRO": rastreamento.get("registro")
            })
    
    # Converter dicionário para lista
    estrutura_json["ATENDIDO"] = list(subpauta_agrupadas.values())
    
    return estrutura_json

# Atualizar o dicionário dict_pedidos_unicos com o JSON gerado
for num_pedido_siscap in dict_pedidos_unicos.keys():
    try:
        # Construir o JSON para este pedido
        json_data = construir_json_pedido(
            num_pedido_siscap,
            dict_pedidos_unicos,
            dict_pedidos_novos,
            dict_pedidos_cancelados,
            dict_lotes_tratamento_iii,
            dict_lotes_tratamento_i,
            dict_lotes_objetos_itens,
            dict_lotes_objetos_rastreamento
        )
        
        # Converter para string JSON
        json_str = json.dumps(json_data, default=json_serializer, indent=2, ensure_ascii=False)
        
        # Atualizar o dicionário dict_pedidos_unicos
        if dict_pedidos_unicos[num_pedido_siscap]:
            dict_pedidos_unicos[num_pedido_siscap][0]["fomrato_json"] = json_str
        
        print(f"JSON gerado para pedido: {num_pedido_siscap}")
        
    except Exception as e:
        print(f"Erro ao gerar JSON para pedido {num_pedido_siscap}: {str(e)}")
        if dict_pedidos_unicos[num_pedido_siscap]:
            dict_pedidos_unicos[num_pedido_siscap][0]["fomrato_json"] = "{}"

print(f'✅Processamento JSON concluído!')
print(f'---'*40)
print(f'')

# Etapa 06: Exportar dados do dataframe df_pedido_unico para um arquivo Excel
print(f'⚠️Exportando dados para arquivo Excel...')

# Primeiro, atualizar o DataFrame df_pedidos_unicos com os JSONs gerados
lista_pedidos_atualizados = []
for num_pedido_siscap, itens in dict_pedidos_unicos.items():
    for item in itens:
        # Definir valores padrão
        formato_json_1 = ""
        formato_json_2 = ""
        formato_json_3 = ""
        json_str = item["fomrato_json"]
        json_len = len(json_str) if json_str != "{}" else 0
        if json_str != "{}":
            if json_len <= 32767:
                # Cabe em uma célula
                formato_json_1 = json_str
            elif json_len <= 65535:
                # Precisa de 2 colunas
                formato_json_1 = json_str[:32767]      # Primeiros 32K
                formato_json_2 = json_str[32768:]      # Restante
            else:
                # Precisa de 3 colunas
                formato_json_1 = json_str[:32767]      # Primeiros 32K
                formato_json_2 = json_str[32768:65535] # Próximos 32K
                formato_json_3 = json_str[65536:]      # Restante

        lista_pedidos_atualizados.append({
            "NUM_PEDIDO_SISCAP": num_pedido_siscap,
            "DT_TRANS": item["dt_trans"],
            "CLIENTE": item["cliente"],
            "NOME_CLI": item["nome_cli"],
            "MCU": item["mcu"],
            "MCCO": "",  # Adicionar se necessário
            "DESCRICAO_MCCO": item["descricao_mcco"],
            "FORMATO_JSON_1": formato_json_1,
            "FORMATO_JSON_2": formato_json_2,
            "FORMATO_JSON_3": formato_json_3
        })

# Criar DataFrame atualizado
df_pedidos_unicos_atualizado = pd.DataFrame(lista_pedidos_atualizados)

# Exportar para Excel
nome_arquivo = "pedidos_unicos_com_json.xlsx"
with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
    # Planilha principal com os dados
    df_pedidos_unicos_atualizado.to_excel(writer, sheet_name='PEDIDOS_UNICOS', index=False)
    
    # Planilha com dados dos pedidos novos
    df_pedidos_novos_agrupados.to_excel(writer, sheet_name='PEDIDOS_NOVOS', index=False)
    
    # Planilha com dados dos pedidos cancelados
    df_pedidos_cancelados_agrupados.to_excel(writer, sheet_name='PEDIDOS_CANCELADOS', index=False)
    
    # Planilha com lotes em tratamento
    df_lote_em_tratamento.to_excel(writer, sheet_name='LOTES_TRATAMENTO', index=False)

print(f'Arquivo "{nome_arquivo}" exportado com sucesso!')
print(f'Total de pedidos processados: {len(df_pedidos_unicos_atualizado)}')
print(f'---'*40)
print(f'')

# Exportar também alguns JSONs para arquivos individuais (opcional)
print(f'✅ Exportando JSONs individuais para análise...')

# Criar diretório para os JSONs se não existir
import os
os.makedirs("jsons_gerados", exist_ok=True)

# Exportar os 5 primeiros JSONs para arquivos individuais
for i, (num_pedido_siscap, itens) in enumerate(list(dict_pedidos_unicos.items())[:5]):
    if itens and itens[0]["fomrato_json"] != "{}":
        nome_arquivo_json = f"jsons_gerados/pedido_{num_pedido_siscap}.json"
        with open(nome_arquivo_json, 'w', encoding='utf-8') as f:
            f.write(itens[0]["fomrato_json"])
        print(f"JSON exportado: {nome_arquivo_json}")

print(f'✅ Exportação concluída!')
print(f'='*60)
print(f'')

# Exemplo de como acessar o JSON de um pedido específico
if dict_pedidos_unicos:
    primeiro_pedido = list(dict_pedidos_unicos.keys())[0]
    print(f"✅ Exemplo do JSON para pedido {primeiro_pedido}:")
    print(f"✅ Tamanho do JSON: {len(dict_pedidos_unicos[primeiro_pedido][0]['fomrato_json'])} caracteres")
    
    # Mostrar estrutura do JSON (primeiros 500 caracteres)
    json_exemplo = dict_pedidos_unicos[primeiro_pedido][0]['fomrato_json']
    print(f"✅ Preview do JSON:\n{json_exemplo[:500]}...")



