import pandas as pd
from collections import defaultdict

# 1. Carregar os dados
df_pedidos = pd.read_excel("001-base_completa.xlsx",
                    sheet_name="base_pedidos",
                    usecols="A:P",nrows=10000)
print(f"Total de Registros na Planilha BASE_PEDIDOS: {len(df_pedidos)}")

df_lotes = pd.read_excel("001-base_completa.xlsx",
                         sheet_name="base_lotes",
                         usecols="A:Q",
                         nrows=10000)
print(f"Total de Registros na Planilha BASE_LOTES: {len(df_lotes)}")

df_objetos = pd.read_excel("001-base_completa.xlsx",
                           sheet_name="base_objeto_postado",
                           usecols="A:L",
                           nrows=10000)
print(f"Total de Registros na Planilha BASE_OBJSETO_POSTADO: {len(df_objetos)}")

print(f'---'*40)

# 2. Ajustar o nome das colunas chaves nos DataFrames
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

# 3. Lista de pedidos únicos
# Criar DataFrame com pedidos únicos
pedidos_unicos = df_pedidos.assign(MCU=lambda x: x['NUM_PEDIDO_SISCAP'].astype(str).str[:8])[['NUM_PEDIDO_SISCAP', 'PEDIDO', 'TIPO_PEDIDO','DT_TRANS', 'MCCO','CLIENTE','NOME_CLI', 'MCU']].drop_duplicates()
# Filtrar pedidos que começam com '0' (como NUM_PEDIDO_SISCAP LIKE '0%')
pedidos_unicos = pedidos_unicos[
    pedidos_unicos['NUM_PEDIDO_SISCAP'].astype(str).str.startswith('0')
]

print(f"Pedidos que começam com 0: {len(pedidos_unicos)}")
print(pedidos_unicos.head())
print(f'---'*40)

# 3.Dataframe somente com pedidos novos
df_pedidos_novos = df_pedidos[df_pedidos["PROX_STATUS"] == 540]
print(f"Total de Pedidos Novos: {len(df_pedidos_novos)}")

df_pedidos_novos_agrupados = df_pedidos_novos.groupby(['NUM_PEDIDO_SISCAP', 'ITEM','DESC_ITEM' ,'UN_MEDIDA'])['QTD'].sum().reset_index().rename(columns={'QTD': 'QTD_TOTAL'})
print(f"Total de Pedidos Novos Agrupados: {len(df_pedidos_novos_agrupados)}")

print(df_pedidos_novos_agrupados.head())

print(f'---'*40)

# 4. Dataframe somente com pedidos cancelados
df_pedidos_cancelados = df_pedidos[df_pedidos["ULT_STATUS"].isin([980,982])]
print(f"Total de Pedidos Cancelados: {len(df_pedidos_cancelados)}")

df_pedidos_cancelados_agrupados = df_pedidos_cancelados.groupby(['NUM_PEDIDO_SISCAP', 'ITEM','DESC_ITEM' ,'UN_MEDIDA'])['QTD'].sum().reset_index().rename(columns={'QTD': 'QTD_TOTAL'})
print(f"Total de Pedidos Cancelados Agrupados: {len(df_pedidos_cancelados_agrupados)}")

print(df_pedidos_cancelados_agrupados.head())

print(f'---'*40)

# 5. Dataframe somente com pedidos em tratamento

## Excluir do Dataframe df_lotes o status 291 e tipo ZR
df_lotes = df_lotes.drop(df_lotes[df_lotes["status_lote"] == 291].index)
df_lotes = df_lotes.drop(df_lotes[df_lotes["TIPO_PEDIDO"] == "ZR"].index)

df_lotes_relacionados_com_df_pedidos = pd.merge(
    df_lotes,
    df_pedidos[['PEDIDO', 'TIPO_PEDIDO','NUM_PEDIDO_SISCAP','DT_TRANS']],
    on=['PEDIDO', 'TIPO_PEDIDO'],
    how='inner'
)

print(f"Total de Registros na Planilha BASE_LOTES relacionados com BASE_PEDIDOS antes do tratamento: {len(df_lotes_relacionados_com_df_pedidos)}")
print(df_lotes_relacionados_com_df_pedidos.head())
print(f'---'*40)

## Eilminar duplicidades e ajustar valores NaN na coluna status_impressao
df_lotes_relacionados_com_df_pedidos = df_lotes_relacionados_com_df_pedidos.drop_duplicates()
df_lotes_relacionados_com_df_pedidos = df_lotes_relacionados_com_df_pedidos.replace({'status_impressao': {"": 'III', None: 'III',"NAN": 'III'}})
print(f"Total de Registros na Planilha BASE_LOTES relacionados com BASE_PEDIDOS: {len(df_lotes_relacionados_com_df_pedidos)}")
print(df_lotes_relacionados_com_df_pedidos.head())
print(f'---'*40)

df_lote_em_tratamento = df_lotes_relacionados_com_df_pedidos.groupby(['NUM_PEDIDO_SISCAP','LOTE','status_impressao' ,'item', 'descricao', 'um'])['qtde'].sum().reset_index().rename(columns={'qtde': 'QTD_TOTAL'})
print(f"Total de Registros de Lotes em Tratamento no WMS: {len(df_lote_em_tratamento)}")

# Dataframe com lotes sem impressao de pauta (status III)
df_lote_em_tratamento_status_iii = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "III"]
print(f"Total de Registros de Lotes em Tratamento no WMS com status III : {len(df_lote_em_tratamento_status_iii)}")
print(df_lote_em_tratamento_status_iii.head())

print(f'')

# Dataframe com lotes com impressao de pauta (status I)
df_lote_em_tratamento_status_i = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "I"]
print(f"Total de Registros de Lotes em Tratamento no WMS com status I : {len(df_lote_em_tratamento_status_i)}")
print(df_lote_em_tratamento_status_i.head())

print(f'---'*40)

print(f'')

# Dataframe com lotes com etiqueta gerada (status II)
df_lote_em_tratamento_status_ii = df_lote_em_tratamento[df_lote_em_tratamento["status_impressao"] == "II"]
print(f"Total de Registros de Lotes em Tratamento no WMS com status II : {len(df_lote_em_tratamento_status_ii)}")
print(df_lote_em_tratamento_status_ii.head())


print(f'---'*40)


# Merge do dataframe com Lotes em tratamento com o dataframe de objeotos postados

## Concatenar as colunas nota_fiscal e serie em df_objetos
df_objetos['NOTA_SERIE'] = (df_objetos['NOTA_FISCAL'].astype('str') + "0" + df_objetos['SERIE'].astype('str')).astype('float64')

## Merge entre os dados dos lotes e objetos

df_lote_possui_objeto = pd.merge(
    df_lotes_relacionados_com_df_pedidos,
    df_objetos[['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA', 'NOTA_SERIE','REGISTRO', 'DATA_EXPEDICAO']],
    on=['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA'],
    how='inner'
)
print(f'Total de registros entre lotes e objetos: {len(df_lote_possui_objeto)}')

df_lote_possui_objeto_agrupar_itens = df_lote_possui_objeto.groupby(['NUM_PEDIDO_SISCAP','NOTA_SERIE','ZONA','item','descricao','um'])['qtde'].sum().reset_index().rename(columns={'qtde': 'QTD_TOTAL'})
print(f'Total de registros entre lotes e objetos agrupados por NUM_PEDIDO_SISCAP, ZONA e NOTA_SERIE: {len(df_lote_possui_objeto_agrupar_itens)}')
print(df_lote_possui_objeto_agrupar_itens.head())

# 11. Merge entre os dados dos pedidos novos e lotes com objetos
"""
df_combinado02 = pd.merge(
    df_pedidos,
    df_lote_possui_objeto[['PEDIDO', 'TIPO_PEDIDO', 'NOTA_SERIE','ZONA','REGISTRO', 'DATA_EXPEDICAO']],
    on=['PEDIDO', 'TIPO_PEDIDO', 'NOTA_SERIE'],
    how='inner'
)
"""

# Excluir as duplicidades da combinacao 02
"""
print(f'Total de registros do combinado02 com duplicidade: {len(df_combinado02)}')
df_combinado03 = df_combinado02.drop_duplicates()
print(f'Total de registros do combinado02 sem duplicidade: {len(df_combinado03)}')
"""

# 12. Criar o dict com dados da expedição
"""
dict_expedicao = defaultdict(list)

for _, linha in df_combinado03.iterrows():
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "pedido": linha["PEDIDO"],
          "tipo_pedido": linha["TIPO_PEDIDO"],
          "zona": linha["ZONA"],
          "NOTA_SERIE": linha["NOTA_SERIE"],          
          "registro": linha["REGISTRO"],
          "data_expedicao ": linha["DATA_EXPEDICAO"]          
    } 
    dict_expedicao[id_].append(item)

print(dict_expedicao)
"""