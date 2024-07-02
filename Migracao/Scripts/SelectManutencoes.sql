SELECT 
man111.CNPJ_CPF Paciente_CPF,
emd101.NOME Nome_Paciente,
man111.VALOR_ORIG Valor_Original,
man111.NOME_RESP_ATEND Dentista_Nome,
man111.DATA_PAGTO Data_Pagamento,
man111.DOCUMENTO Numero_Controle,
man111.NOME_TIPO Procedimento_Nome,
man111.VALOR Valor_Pagamento,
man111.LANCTO Lancamento,
man111.TIPO_PAGTO Tipo_Pagamento,
man111.VENCTO_ORIG Vencimento,
man111.VALOR_PARCELA Valor_Devido,
man111.DATA_ATEND Data_Atendimento,
man101.OBS_ATEND Procedimentos_Observacao
FROM MAN111 man111
INNER JOIN EMD101 emd101 ON emd101.CGC_CPF = man111.CNPJ_CPF 
INNER JOIN MAN101 man101 ON (man101.CNPJ_CPF = man111.CNPJ_CPF AND man101.LANCTO = man111.LANCTO)