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
AND man111.DATA_ATEND IS NOT NULL
UNION 
SELECT
anotacao_clinica.CNPJ_CPF Paciente_CPF,
emd101.NOME Paciente_Nome,
anotacao_clinica.NOME_RESP Dentista_Nome,
anotacao_clinica.ANOTACAO Observacao,
anotacao_clinica."DATA" Data_Atendimento
FROM ANOTACAO_CLINICA anotacao_clinica
LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = anotacao_clinica.CNPJ_CPF