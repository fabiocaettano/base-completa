import pandas as pd
from collections import defaultdict

# Sere como um script de análise de dados para processar informações de pedidos, lotes e objetos postados a partir de uma planilha Excel.
# Serão 05 etapas principais:
# Etapa 01: Carregar os dados das planilhas Excel
# Etapa 02: Ajustar o nome das colunas chaves nos DataFrames 
# Etapa 03: Configurar os DataFrames para análise
# Etapa 04: Criar dicionários com os dados processados
# Etapa 05: Gerar estrutura JSON para integração com API (futuro)
# Etapa 06: Exportar os dados processados para novos arquivos Excel/CSV (futuro)

# Etapa 01: Carregar os dados das planilhas Excel

## 01.01 Carregar os dados da planilha base_pedidos
df_pedidos = pd.read_excel("001-base_completa.xlsx",
                    sheet_name="base_pedidos",
                    usecols="A:P",nrows=10000)
#print(f"Total de Registros na Planilha BASE_PEDIDOS: {len(df_pedidos)}")

## 01.02 Carregar os dados da planilha base_lotes
df_lotes = pd.read_excel("001-base_completa.xlsx",
                         sheet_name="base_lotes",
                         usecols="A:Q",
                         nrows=10000)
#print(f"Total de Registros na Planilha BASE_LOTES: {len(df_lotes)}")

## 01.03 Carregar os dados da planilha base_objeto_postado
df_objetos = pd.read_excel("001-base_completa.xlsx",
                           sheet_name="base_objeto_postado",
                           usecols="A:L",
                           nrows=10000)
# print(f"Total de Registros na Planilha BASE_OBJSETO_POSTADO: {len(df_objetos)}")


# 01.04 Carregar os dados da planilha base_mcco
df_base_mcco = pd.read_excel("001-base_completa.xlsx",
                           sheet_name="base_mcco",
                           usecols="A:B",
                           nrows=30)
#print(f"Total de Registro da Planilha BASE_MCCO : {len(df_base_mcco) }")
# print(f'---'*40)
#print(f'')

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


#print(f"Total de Pedidos Únicos que começam com 0: {len(df
#print(f"Pedidos que começam com 0: {len(pedidos_unicos)}")
#print(pedidos_unicos.head())
#print(f'---'*40)

##03.02  Criar DataFrame somente com pedidos novos (PROX_STATUS = 540)
df_pedidos_novos = df_pedidos[df_pedidos["PROX_STATUS"] == 540]
#print(f"Total de Pedidos Novos: {len(df_pedidos_novos)}")

df_pedidos_novos_agrupados = df_pedidos_novos.groupby(['NUM_PEDIDO_SISCAP', 'ITEM','DESC_ITEM' ,'UN_MEDIDA'])['QTD'].sum().reset_index().rename(columns={'QTD': 'QTD_TOTAL'})
#print(f"Total de Pedidos Novos Agrupados: {len(df_pedidos_novos_agrupados)}")

#print(df_pedidos_novos_agrupados.head())

#print(f'---'*40)

## 03.03  Criar DataFrame somente com pedidos cancelados
df_pedidos_cancelados = df_pedidos[df_pedidos["ULT_STATUS"].isin([980,982])]
#print(f"Total de Pedidos Cancelados: {len(df_pedidos_cancelados)}")

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
#print(f"Total de Registros de Lotes em Tratamento no WMS com status III : {len(df_lote_em_tratamento_status_iii)}")
#print(df_lote_em_tratamento_status_iii.head())
#print(f'')

# 03.04.06 Dataframe com lotes com impressao de pauta (status I)
df_lote_em_tratamento_status_i = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "I"]
#print(f"Total de Registros de Lotes em Tratamento no WMS com status I : {len(df_lote_em_tratamento_status_i)}")
#print(df_lote_em_tratamento_status_i.head())
#print(f'---'*40)
#print(f'')

# 03.04.07 Dataframe com lotes com etiqueta gerada (status II)
df_lote_em_tratamento_status_ii = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "II"]
#print(f"Total de Registros de Lotes em Tratamento no WMS com status II : {len(df_lote_em_tratamento_status_ii)}")
#print(df_lote_em_tratamento_status_ii.head())
#print(f'---'*40)


## 03.04.08 Concatenar as colunas nota_fiscal e serie em df_objetos
df_objetos['NOTA_SERIE'] = (df_objetos['NOTA_FISCAL'].astype('str') + "0" + df_objetos['SERIE'].astype('str')).astype('float64')

## 03.04.09 Merge do dataframe com Lotes em tratamento com o dataframe de objeotos postados

df_lote_possui_objeto = pd.merge(
    df_lotes_relacionados_com_df_pedidos,
    df_objetos[['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA', 'NOTA_SERIE','REGISTRO', 'DATA_EXPEDICAO']],
    on=['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA'],
    how='inner'
)
#print(f'Total de registros entre lotes e objetos: {len(df_lote_possui_objeto)}')

# 03.04.10 Agrupar os itens do DataFrame de lotes com objetos
df_lote_possui_objeto_agrupar_itens = df_lote_possui_objeto.groupby(['NUM_PEDIDO_SISCAP','NOTA_SERIE','SUBPAUTA','ZONA','item','descricao','um'])['qtde'].sum().reset_index().rename(columns={'qtde': 'QTD_TOTAL'})
#print(f'Total de registros entre lotes e objetos agrupados por NUM_PEDIDO_SISCAP, ZONA e NOTA_SERIE: {len(df_lote_possui_objeto_agrupar_itens)}')
#print(df_lote_possui_objeto_agrupar_itens.head())
#print(f'---'*40)

# 03.04.11 Remover colunas desnecessárias para unificação de rastreamento
df_lote_possui_objeto_unificar_rastreamento = df_lote_possui_objeto.drop(["status_lote","status_impressao","posto","lote_consolidado","cep","roteirizador","cod_cliente","descricao_cliente","se","PEDIDO","TIPO_PEDIDO","item","descricao","um","qtde"], axis=1).drop_duplicates()
#print(f'Total de registros entre lotes e objetos unificados por NUM_PEDIDO_SISCAP, ZONA e NOTA_SERIE: {len(df_lote_possui_objeto_unificar_rastreamento)}')
#print(df_lote_possui_objeto_unificar_rastreamento.head())
#print(f'---'*40)

# Etapa 04: Criar dicionários com os dados processados, onde a chave é o NUM_PEDIDO_SISCAP para toodos os dicionários

## 04.01  Criar o dict com o dataframe dos pedidos novos
print(f'Criando dicionário de pedidos novos...')
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

print(dict_pedidos_novos)
print(f'---'*40)
print(f'')

## 04.02  Criar o dict com o dataframe dos pedidos cancelados
print(f'Criando dicionário de pedidos cancelados...')
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
print(dict_pedidos_cancelados)
print(f'---'*40)
print(f'')

## 04.03  Criar o dict com o dataframe dos lotes em tratamento status III
print(f'Criando dicionário de lotes em tratamento status III...')  
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
print(dict_lotes_tratamento_iii)
print(f'---'*40)    
print(f'')

## 04.04  Criar o dict com o dataframe dos lotes em tratamento status I
print(f'Criando dicionário de lotes em tratamento status I...') 
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
print(dict_lotes_tratamento_i)
print(f'---'*40)    
print(f'') 

## 04.05  Criar o dict com o dataframe do df_lote_possui_objeto_agrupar_iten
print(f'Criando dicionário de lotes com objetos postados com itens agrupados...') 
dict_lotes_objetos_itens =defaultdict(list)
for _, linha in df_lote_possui_objeto_agrupar_itens.iterrows():  
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "nota_serie": linha["NOTA_SERIE"],
          "subpauta": linha["SUBPAUTA"],
          "zona": linha["ZONA"],
          "item": linha["item"],
          "descricao": linha["descricao"],
          "um": linha["um"],
          "qtd_total": linha["QTD_TOTAL"]          
    } 
    dict_lotes_objetos_itens[id_].append(item)
print(dict_lotes_objetos_itens)
print(f'---'*40)
print(f'')

## 04.06  Criar o dict com o dataframe do df_lote_possui_objeto_unificar_rastreamento
print(f'Criando dicionário de lotes com objetos postados unificados para rasteamento...') 
dict_lotes_objetos_rastreamento =defaultdict(list)      
for _, linha in df_lote_possui_objeto_unificar_rastreamento.iterrows():  
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "lote": linha["LOTE"],
          "nota_serie": linha["NOTA_SERIE"],
          "registro": linha["REGISTRO"],
          "data_expedicao": linha["DATA_EXPEDICAO"]          
    } 
    dict_lotes_objetos_rastreamento[id_].append(item)   
print(dict_lotes_objetos_rastreamento)
print(f'---'*40)
print(f'')

## 04.07 Criar dicionário dos pedidos únicos
dict_pedidos_unicos = defaultdict(list)
for _, linha in df_pedidos_unicos.iterrows():
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
        "dt_trans" : linha["DT_TRANS"],
        "cliente" : linha["CLIENTE"],
        "nome_cli" : linha["NOME_CLI"],
        "mcu" : linha["MCU"],
        "fomrato_json" : linha["FORMATO_JSON"]
    }
    dict_pedidos_unicos[id_].append(item)
print(dict_pedidos_unicos)
print(f'---'*40)
print(f'')

# Etapa 05: Gerar estrutura JSON para integração com API
print(df_pedidos_unicos.head())