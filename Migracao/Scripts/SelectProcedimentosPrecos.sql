SELECT 
ced001.GRUPO Codigo,
ced001.TABELA Tabela,
ced001.NOME Procedimento_Nome,
ced001.SIMBOLO Abreviacao,
ced001.CODIGO_TUSS TUSS,
ced001.VRVENDA Preco,
ced001.ATIVO Ativo,
ced001.OBS Observacao,
ced001.PARTICULAR Particular,
ced002.CODIGO Codigo_Grupo,
ced002.NOME Nome_Grupo,
ced002.USUARIO Usuario
FROM CED001 ced001
INNER JOIN CED002 ced002 ON CED002.CODIGO = CED001.GRUPO 