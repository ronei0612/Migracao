SELECT  
atd222.DOCUMENTO Numero_Controle,
atd222.CNPJ_CPF Paciente_CPF,
emd101.NOME Nome_Paciente,
atd222.NOME_RESP_ATEND Dentista_Nome,
atd222.NUMERO Dente,
atd222.NOME_PRODUTO Nome_Procedimento,
atd222.Valor,
ced003.VRLIQUIDO Valor_Total,
atd222.OBS Observacao,
atd222.DATA Data_Inicio,
atd222.DATA_ATEND Data_Atendimento
FROM ATD222 atd222
INNER JOIN EMD101 emd101 ON emd101.CGC_CPF = atd222.CNPJ_CPF
JOIN CED003 ced003 ON emd101.CGC_CPF = ced003.CGC_CPF AND ced003.DOCUMENTO = atd222.DOCUMENTO