SELECT  
atd222.DOCUMENTO Numero_Controle,
atd222.CNPJ_CPF Paciente_CPF,
emd101.NOME Nome_Paciente,
atd222.NOME_RESP_ATEND Dentista_Nome,
atd222.NUMERO Dente,
atd222.NOME_PRODUTO,
atd222.Valor,
atd222.OBS Observacao,
atd222.DATA Data_Inicio,
atd222.DATA_ATEND Data_Termino,
atd222.DATA_ATEND Data_Atendimento
FROM ATD222 atd222
INNER JOIN EMD101 emd101 ON EMD101.CGC_CPF = atd222.CNPJ_CPF 