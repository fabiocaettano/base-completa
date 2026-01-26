📦 CLI CD LESTE/SE/SPM
{Concatenate("**Pedido** : ",Topic.VarDadosJson.pedido)}​
Categoria:
{Topic.VarDadosJson.categoria}​
{"----------------------------------------------"}​

🏤{"**Dados do Cliente**"}​
Cadastro Geral: {Concat(Topic.VarDadosJson.cliente, ThisRecord.cadastro_geral)}​
Mcu:
{Concat(Topic.VarDadosJson.cliente,ThisRecord.mcu)}​
Unidade: {Concat(Topic.VarDadosJson.cliente,ThisRecord.cliente)}​
SE:
{Concat(Topic.VarDadosJson.cliente,ThisRecord.se)}​
{"----------------------------------------------"}​

📅 Prazos

{Concat(Topic.VarDadosJson.data_atendimento,"**Solicitação SISCAP** : " & Mid(ThisRecord.solicitacao_inicial,9,2) & "/" & Mid(ThisRecord.solicitacao_inicial,6,2) & Char(10))}​

{Concat(Topic.VarDadosJson.data_atendimento,"**Captado no ERP** : " & Mid(ThisRecord.captado_centro_distribuicao,9,2) & "/" & Mid(ThisRecord.captado_centro_distribuicao,6,2) & Char(10))}​

{Concat(Topic.VarDadosJson.data_atendimento,"**Postagem ERP** : " & Mid(ThisRecord.expedicao_inicial,9,2) & "/" & Mid(ThisRecord.expedicao_inicial,6,2) & Char(10))}​
{"----------------------------------------------"}​

📌
" **Números do Atendimento** "
​
{Concat(Topic.VarDadosJson.numero_atendimento, "**Itens Solicitados** : " & ThisRecord.codigos_solicitados & Char(10))}​

{Concat(Topic.VarDadosJson.numero_atendimento, "**Itens Atendidos** : " & ThisRecord.codigos_atendidos & Char(10))}​

{Concat(Topic.VarDadosJson.numero_atendimento, "**Itens Cancelados** : " & ThisRecord.codigos_cancelados & Char(10))}​

{Concat(Topic.VarDadosJson.numero_atendimento, "**Objeto SRO**  : "  & ThisRecord.objetos_sro & Char(10))}​

{"----------------------------------------------"}​

🛒Atendimento:
{Concat(Topic.VarDadosJson.atendimentos,"✔**SubPauta**: " & ThisRecord.zona & Char(10) & Char(10) & "**Contém**: " & Char(10) & ThisRecord.alias & Char(10) & Char(10) & "**Ratreamento** :" & Char(10) & ThisRecord.rastreamento & Char(10) & Char(10) & "------------------------------" & Char(10))}

