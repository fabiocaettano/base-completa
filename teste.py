import  pandas as pd
from collections import defaultdict

# Carregar a planilha Excel base_pedidos
df_pedidos = pd.read_excel("001-base_completa.xlsx",
                   sheet_name="base_pedidos",
                   usecols="A:P",nrows=10000)

dados_pedidos = defaultdict(list)

for _, linha in df_pedidos.iterrows():
    id_ = str(linha["NUM_PEDIDO_SISCAP"])
    item = {
          "dr": linha["DR"],
          "dt_trans": linha["DT_TRANS"],
          "ped": linha["PED"],
          "tp_ped": linha["TP_PED"],
          "nota_serie": linha["NOTA_SERIE"],
          "tipo_nota": linha["TIPO_NOTA"],
          "cliente": linha["CLIENTE"],
          "mcco ": linha["MCCO"],
          "nome_cli" : linha["NOME_CLI"],
          "item": linha["ITEM"],
          "desc_item": linha["DESC_ITEM"],
          "qtd": linha["QTD"],
          "un_medida": linha["UN_MEDIDA"],
          "ult_status": linha["ULT_STATUS"],
          "prox_status": linha["PROX_STATUS"]          
    } 
    dados_pedidos[id_].append(item) 
    
    
dados = dict(dados_pedidos)
total = len(dados)
print(f"Total de pedidos únicos: {total}")

print(dados_pedidos)