
# Identificação do cliente e detalhes do pedido
"Olá, " & Topic.GetItem.NOME_CLI & " seu pedido " & Topic.GetItem.NUM_PEDIDO_SISCAP & " foi caotado pelo Centro de Distribuição no dia " & Topic.GetItem.DT_TRANS & ", Segue informações do seu pedido:"

# Pedidos Reservados
If(CountRows(Topic.VarDadosSiscap.RESERVADO) > 0,"✔ Quantidade Reservada: " & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.RESERVADO," ◾ "& DESCRICAO_ITEM & " , " &  QTDE & " " & UN_MEDIDA & Char(13) & Char(10)),Char(13) & Char(13) & Char(10))


# Pedidos Aguardando geração a subpauta
If(CountRows(Topic.VarDadosSiscap.PROCESSAMENTO) > 0,"✔ Aguardando geração da subpauta(s): " & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.PROCESSAMENTO,"◾" & DESCRICAO_ITEM & " , " &  QTDE & " " & UN_MEDIDA & Char(13) & Char(10)),Char(13) & Char(13) & Char(10))

# Pedidos Aguardando separação
If(CountRows(Topic.VarDadosSiscap.PAUTA_IMPRESSA) > 0,"✔ Aguardando a separação: " & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.PAUTA_IMPRESSA," ◾" & DESCRICAO_ITEM & " , " &  QTDE & " " & UN_MEDIDA & Char(13) & Char(10)),Char(13) & Char(13) & Char(10))

# Pedidos Separados

If(CountRows(Topic.VarDadosSiscap.ATENDIDO) > 0,"✔ Atendido" & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.ATENDIDO, "SUBPAUTA: " & SUBPAUTA & " ( nf: " & NOTA_SERIE & ") " & Char(13) & Char(10)), Char(13) & Char(10))

If(CountRows(Topic.VarDadosSiscap.ATENDIDO) > 0,"✔ Atendido" & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.ATENDIDO, "◾" & "SUBPAUTA: " & SUBPAUTA & " ( nf: " & NOTA_SERIE & ") " & Char(13) & Char(10)), Char(13) & Char(10))


If(
CountRows(Topic.VarDadosSiscap.ATENDIDO) > 0,
"✔ Atendido" & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.ATENDIDO, "◾" & "SUBPAUTA: " & SUBPAUTA & " ( nf: " & NOTA_SERIE & ") " & Char(13) & Char(10) &
Concat(ITENS.SKU,DESCRICAO_ITEM & Char(13) & Char(10))
)
,
Char(13) & Char(10)
)



If(
CountRows(Topic.VarDadosSiscap.ATENDIDO) > 0,
"✔ Atendido" & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.ATENDIDO, "◾" & "SUBPAUTA: " & SUBPAUTA & " ( nf: " & NOTA_SERIE & ") " 
& Char(13) & Char(10) &
Concat(ITENS.SKU,DESCRICAO_ITEM & Char(13) & Char(10))
& Char(13) & Char(10) &
Concat(ITENS.RASTREAMENTO,REGISTRO & " ")
)
,
Char(13) & Char(10)
)


If(
CountRows(Topic.VarDadosSiscap.ATENDIDO) > 0,
"✔ Atendido" & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.ATENDIDO, "◾" & "SUBPAUTA: " & SUBPAUTA & " ( nf: " & NOTA_SERIE & ") "
& Char(13) & Char(10) &
Concat(ITENS.SKU,DESCRICAO_ITEM & Char(13) & Char(10))
& Char(13) & Char(10) &
Concat(ITENS.RASTREAMENTO,REGISTRO & " ")
& Char(13) & Char(10)
)
,
Char(13) & Char(10)
)


If(
CountRows(Topic.VarDadosSiscap.ATENDIDO) > 0,
"✔ Atendido" & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.ATENDIDO, "◾" & "SUBPAUTA: " & SUBPAUTA & " ( nf: " & NOTA_SERIE & ") "
& Char(13) & Char(10) &
Concat(ITENS.SKU,DESCRICAO_ITEM & " , " & QTDE & " " & UN_MEDIDA & Char(13) & Char(10))
& Char(13) & Char(10) &
Concat(ITENS.RASTREAMENTO,REGISTRO & " ")
& Char(13) & Char(10)
)
,
Char(13) & Char(10)
)


{
If(
CountRows(Topic.VarDadosSiscap.ATENDIDO) > 0,
"📌 **Atendido :** " & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.ATENDIDO, "📄" & " SUBPAUTA : " & SUBPAUTA & " ( nf: " & NOTA_SERIE & ") "
& Char(13) & Char(10) &
Concat(ITENS.SKU," ✔" &DESCRICAO_ITEM & " , " & QTDE & " " & UN_MEDIDA & Char(13) & Char(10))
& Char(13) & Char(10) &
Concat(ITENS.RASTREAMENTO,REGISTRO & " ")
& Char(13) & Char(10)
)
,
Char(13) & Char(10)
)
}
​


# Pedidos em falta
If(CountRows(Topic.VarDadosSiscap.CANCELADO) > 0,"✔ Cancelados: " & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.CANCELADO," ◾" & DESCRICAO_ITEM & " , " &  QTDE & " " & UN_MEDIDA & Char(13) & Char(10)),Char(13) & Char(13) & Char(10))


If(CountRows(Topic.VarDadosSiscap.CANCELADO) > 0," 📌 Cancelados: " & Char(13) & Char(10) & Concat(Topic.VarDadosSiscap.CANCELADO," ✔ " & DESCRICAO_ITEM & " , " &  QTDE & " " & UN_MEDIDA & Char(13) & Char(10)),Char(13) & Char(13) & Char(10))
