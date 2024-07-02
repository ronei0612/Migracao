SELECT 
crd111.CGC_CPF Documento,
emd101.NOME Nome,
crd111.OBS Observacao,
crd111.DOCUMENTO Numero_Controle,
crd111.VALOR Valor_Devido,
crd111.VENCTO Data_Vencimento,
crd111.EMISSAO,
crd111.DUPLICATA ,
crd111.PARCELA Parcelas,
crd111.TIPO_DOC Tipo_Pagamento,
crd111.VALOR_ORIG Valor_Original,
crd111.VENCTO_ORIG Vencimento_Original,
crd111.SITUACAO,
crd111.NOME_GRUPO,
crd111.ORDEM
FROM CRD111 crd111
INNER JOIN EMD101 emd101 ON EMD101.CGC_CPF = crd111.CGC_CPF 