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

# 3.Dataframe somente com pedidos novos
df_pedidos_novos = df_pedidos[df_pedidos["PROX_STATUS"] == 540]
print(f"Total de Pedidos Novos: {len(df_pedidos_novos)}")

# 4. Dataframe somente com pedidos cancelados
df_pedidos_cancelados = df_pedidos[df_pedidos["ULT_STATUS"].isin([980,982,983])]
print(f"Total de Pedidos Cancelados: {len(df_pedidos_cancelados)}")

# 5. Excluir do Dataframe df_lotes o status 291 e tipo ZR
df_lotes = df_lotes.drop(df_lotes[df_lotes["status_lote"] == 291].index)
df_lotes = df_lotes.drop(df_lotes[df_lotes["TIPO_PEDIDO"] == "ZR"].index)
print(f"Total de Registros na Planilha BASE_LOTES após exclusão do status 291: {len(df_lotes)}")

# 6. Dataframe somente com lotes sem tratamento
df_lotes_sem_tratamento = df_lotes[(df_lotes["status_impressao"] != "II")]
print(f"Total de Registros Aguardando Tratamento no WMS: {len(df_lotes_sem_tratamento)}")

# 7. Dataframe com Lote sem pauta gerada
df_lotes_sem_pauta = df_lotes_sem_tratamento[df_lotes_sem_tratamento["status_impressao"] != "I"]
print(f"Total de Registros sem Pauta Gerada: {len(df_lotes_sem_pauta)}")

# 8. Dataframe com Lote com pauta gerada
df_lotes_com_pauta = df_lotes_sem_tratamento[df_lotes_sem_tratamento["status_impressao"] == "I"]
print(f"Total de Registros com Pauta Gerada: {len(df_lotes_com_pauta)}")

# 9. Concatenar as colunas nota_fiscal e serie em df_objetos
df_objetos['NOTA_SERIE'] = (df_objetos['NOTA_FISCAL'].astype('str') + "0" + df_objetos['SERIE'].astype('str')).astype('float64')

# 10.Merge entre os dados dos lotes e objetos
df_lote_possui_objeto = pd.merge(
    df_lotes,
    df_objetos[['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA', 'NOTA_SERIE','REGISTRO', 'DATA_EXPEDICAO']],
    on=['LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'ZONA'],
    how='inner'
)
print(f'Total de registros entre lotes e objetos: {len(df_lote_possui_objeto)}')


# 11. Merge entre os dados dos pedidos novos e lotes com objetos
df_combinado02 = pd.merge(
    df_pedidos,
    df_lote_possui_objeto[['PEDIDO', 'TIPO_PEDIDO', 'NOTA_SERIE','ZONA','REGISTRO', 'DATA_EXPEDICAO']],
    on=['PEDIDO', 'TIPO_PEDIDO', 'NOTA_SERIE'],
    how='inner'
)

# Excluir as duplicidades da combinacao 02
print(f'Total de registros do combinado02 com duplicidade: {len(df_combinado02)}')
df_combinado03 = df_combinado02.drop_duplicates()
print(f'Total de registros do combinado02 sem duplicidade: {len(df_combinado03)}')


# 12. Criar o dict com dados da expedição
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