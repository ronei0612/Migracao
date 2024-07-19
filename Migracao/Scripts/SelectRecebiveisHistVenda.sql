SELECT 
cxd555.CNPJ_CPF Paciente_CPF,
emd101.NOME Nome_Paciente,
cxd555.HISTORICO Observacoo_Recebivel,
cxd555.DOCUMENTO Documento_Ref,
cxd555.VALOR Valor_Original,
cxd555.DATA Vencimento,
cxd555.TRANSMISSAO Recebivel
FROM CXD555 cxd555
INNER JOIN EMD101 emd101 ON emd101.CGC_CPF = cxd555.CNPJ_CPF 