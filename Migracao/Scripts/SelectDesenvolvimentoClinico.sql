SELECT 
man101.LANCTO Lancamento,
man101.CNPJ_CPF Paciente_CPF,
emd101.NOME Paciente_Nome,
man101.DATA_RETORNO Data_Retorno,
man101.RESP_ATEND Dentista_Codigo,
man101.NOME_RESP_ATEND Dentista_Nome,
man101.TIPO_ATEND Procedimento_Nome,
man101.DATA_LANC Data_Inicio,
man101.DATA_RETORNO Data_Retorno,
man101.OBS_ATEND Procedimento_Observacao
FROM MAN101 man101 
INNER JOIN EMD101 emd101 ON emd101.CGC_CPF = man101.CNPJ_CPF 