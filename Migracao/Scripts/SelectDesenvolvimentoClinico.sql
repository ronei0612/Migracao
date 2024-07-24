--SELECT
--man101.LANCTO Lancamento,
--man101.CNPJ_CPF Paciente_CPF,
--emd101.NOME Paciente_Nome,
--man101.DATA_RETORNO Data_Retorno,
--man101.RESP_ATEND Dentista_Codigo,
--man101.NOME_RESP_ATEND Dentista_Nome,
--man101.TIPO_ATEND Procedimento_Nome,
--man101.DATA_LANC Data_Inicio,
--man101.DATA_RETORNO Data_Retorno,
--man101.OBS_ATEND Procedimento_Observacao
--FROM MAN101 man101 
--INNER JOIN EMD101 emd101 ON emd101.CGC_CPF = man101.CNPJ_CPF

SELECT
man111.CNPJ_CPF Paciente_CPF,
emd101.NOME Paciente_Nome,
man111.NOME_RESP_ATEND Dentista_Nome,
man101.OBS_ATEND Observacao,
man111.DATA_ATEND Data_Atendimento
FROM MAN111 man111
LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = man111.CNPJ_CPF
LEFT JOIN MAN101 man101 ON (man101.CNPJ_CPF = man111.CNPJ_CPF AND man101.LANCTO = man111.LANCTO)
WHERE man101.TIPO_ATEND IS NULL AND man111.NOME_TIPO IS NULL
UNION 
SELECT
anotacao_clinica.CNPJ_CPF Paciente_CPF,
emd101.NOME Paciente_Nome,
anotacao_clinica.NOME_RESP Dentista_Nome,
anotacao_clinica.ANOTACAO Observacao,
anotacao_clinica."DATA" Data_Atendimento
FROM ANOTACAO_CLINICA anotacao_clinica
LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = anotacao_clinica.CNPJ_CPF