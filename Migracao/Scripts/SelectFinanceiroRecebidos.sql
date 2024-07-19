SELECT 
emd101.NOME Nome_Paciente,
bxd111.CGC_CPF CNPJ_CPF,
bxd111.DOCUMENTO Numero_Controle,
bxd111.VALOR Valor_Pago,
bxd111.VR_PARCELA Valor_Parcela,
bxd111.OBS Observacao,
bxd111.BAIXA Data_Baixa,
bxd111.VENCTO Data_Vencimento,
bxd111.TIPO_DOC Tipo_Documento,
bxd111.PARCELA Parcela,
crd013.CODIGO Tipo_Especie,
crd013.Nome Especie_Pagamento,
crd111.VALOR Valor_Devido,
crd111.TIPO_DOC Tipo_Pagamento,
crd111.VALOR_ORIG Valor_Original,
crd111.DUPLICATA,
crd111.VENCTO_ORIG Vencimento_Recebivel,
crd111.SITUACAO,
crd111.NOME_GRUPO,
crd111.ORDEM,
crd111.OBS Pagamento_Observacoes,
crd013.CODIGO Tipo_Especie,
crd013.NOME Especie_Pagamento
FROM BXD111 bxd111
INNER JOIN EMD101 emd101 ON emd101.CGC_CPF = bxd111.CGC_CPF
INNER JOIN CRD111 crd111 ON (crd111.CGC_CPF = bxd111.CGC_CPF AND crd111.DOCUMENTO = bxd111.DOCUMENTO)
INNER JOIN CRD013 crd013 ON crd013.CODIGO = bxd111.TIPO_DOC