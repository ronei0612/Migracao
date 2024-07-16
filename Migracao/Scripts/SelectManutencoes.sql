SELECT
man111.CNPJ_CPF Paciente_CPF,
emd101.NOME Nome_Paciente,
man111.NOME_RESP_ATEND Dentista_Nome,
man111.DOCUMENTO Numero_Controle,
man111.NOME_TIPO Procedimento_Nome,
man111.VALOR_ORIG Valor,
man111.LANCTO Lancamento,
man101.NUM_MAN Numero_Manutencao,
crd013.NOME Tipo_Pagamento,
man101.OBS_ATEND Procedimentos_Observacao,
man101.DATA_LANC Data_Hora_Inicio,
man111.DATA_ATEND Data_Hora_Termino
FROM MAN111 man111
LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = man111.CNPJ_CPF
LEFT JOIN CRD013 crd013 ON crd013.CODIGO = man111.TIPO_PAGTO
LEFT JOIN MAN101 man101 ON (man101.CNPJ_CPF = man111.CNPJ_CPF AND man101.LANCTO = man111.LANCTO)