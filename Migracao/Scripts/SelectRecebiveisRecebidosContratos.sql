SELECT crd111.DUPLICATA Documento,
       crd111.TIPO_DOC Tipo_Documento,
       crd111.DOCUMENTO Numero_Controle,
       crd111.PARCELA Parcela,
       emd101.NOME Nome_Paciente,
       crd111.CGC_CPF CNPJ_CPF,
       crd111.VALOR_VENDA Valor_Total,
       crd111.VALOR Valor_Devido,
       crd111.VALOR Valor_Original,
       crd111.VENCTO Data_Vencimento,
       crd111.VENCTO Vencimento_Recebivel,
       bxd111.VALOR Valor_Pago,
       bxd111.BAIXA Data_Baixa,
       crd013.NOME Especie_Pagamento,
       crd013.CODIGO Tipo_Especie,
       crd111.OBS Observacao,
       bxd111.OBS Pagamento_Observacoes
  FROM CRD111 crd111
       LEFT JOIN CRD013 crd013 ON crd013.CODIGO = crd111.TIPO_DOC
       LEFT JOIN BXD111 bxd111 ON bxd111.DUPLICATA = crd111.DUPLICATA
                               AND bxd111.CGC_CPF = crd111.CGC_CPF
                               AND bxd111.DOCUMENTO = crd111.DOCUMENTO
       LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = crd111.CGC_CPF
UNION ALL
SELECT man111.DOCUMENTO Documento,
       man111.TIPO_PAGTO Tipo_Documento,
       man111.LANCTO Numero_Controle,
       man111.NUM_MAN Parcela,
       emd101.NOME Nome_Paciente,
       man111.CNPJ_CPF CNPJ_CPF,
       man111.VALOR Valor_Total,
       man111.VALOR_PARCELA Valor_Devido,
       man111.VALOR_PARCELA Valor_Original,
       CAST(man111.VENCTO AS TIMESTAMP) Data_Vencimento,
       CAST(man111.VENCTO AS TIMESTAMP) Vencimento_Recebivel,
       cxd555.VALOR Valor_Pago,
       CAST(man111.DATA_PAGTO AS TIMESTAMP) Data_Baixa,
       crd013.NOME Especie_Pagamento,
       crd013.CODIGO Tipo_Especie,
       man111.MOTIVO_RECEBER Observacao,
       cxd555.HISTORICO Pagamento_Observacoes
  FROM MAN111 man111
       LEFT JOIN CXD555 cxd555 ON cxd555.DOCUMENTO = man111.LANCTO
       LEFT JOIN CRD013 crd013 ON crd013.CODIGO = cxd555.TIPO
       LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = man111.CNPJ_CPF
 WHERE cxd555.HISTORICO NOT LIKE '%ENVELOPE%'

--SELECT crd111.DUPLICATA Documento,
--       crd111.TIPO_DOC Tipo_Documento,
--       crd111.DOCUMENTO Numero_Controle,
--       crd111.PARCELA Parcela,
--       emd101.NOME Nome_Paciente,
--       crd111.CGC_CPF CNPJ_CPF,
--       crd111.VALOR_VENDA Valor_Total,
--       crd111.VALOR Valor_Devido,
--       crd111.VALOR Valor_Original,
--       crd111.VENCTO Data_Vencimento,
--       crd111.VENCTO Vencimento_Recebivel,
--       CASE WHEN cxd555.VALOR IS NOT NULL THEN cxd555.VALOR ELSE bxd111.VALOR END AS Valor_Pago,
--       CASE WHEN cxd555.DATA IS NOT NULL THEN cxd555.DATA ELSE bxd111.BAIXA END AS Data_Baixa,
--       crd013.NOME Especie_Pagamento,
--       crd013.CODIGO Tipo_Especie,
--       crd111.OBS Observacao,
--       CASE WHEN cxd555.HISTORICO IS NOT NULL THEN cxd555.HISTORICO ELSE bxd111.OBS END AS Pagamento_Observacoes
--  FROM CRD111 crd111
--       LEFT JOIN CRD013 crd013 ON crd013.CODIGO = crd111.TIPO_DOC
--       LEFT JOIN BXD111 bxd111 ON bxd111.DUPLICATA = crd111.DUPLICATA
--                               AND bxd111.CGC_CPF = crd111.CGC_CPF
--                               AND bxd111.DOCUMENTO = crd111.DOCUMENTO
--       LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = crd111.CGC_CPF
--       LEFT JOIN CXD555 cxd555 ON cxd555.HISTORICO = 'REC - ' || crd111.DOCUMENTO || ' - ' || emd101.NOME
--UNION ALL
--SELECT man111.DOCUMENTO Documento,
--       man111.TIPO_PAGTO Tipo_Documento,
--       man111.LANCTO Numero_Controle,
--       man111.NUM_MAN Parcela,
--       emd101.NOME Nome_Paciente,
--       man111.CNPJ_CPF CNPJ_CPF,
--       man111.VALOR Valor_Total,
--       man111.VALOR_PARCELA Valor_Devido,
--       man111.VALOR_PARCELA Valor_Original,
--       CAST(man111.VENCTO AS TIMESTAMP) Data_Vencimento,
--       CAST(man111.VENCTO AS TIMESTAMP) Vencimento_Recebivel,
--       cxd555.VALOR Valor_Pago,
--       CAST(man111.DATA_PAGTO AS TIMESTAMP) Data_Baixa,
--       crd013.NOME Especie_Pagamento,
--       crd013.CODIGO Tipo_Especie,
--       man111.MOTIVO_RECEBER Observacao,
--       cxd555.HISTORICO Pagamento_Observacoes
--  FROM MAN111 man111
--       LEFT JOIN CXD555 cxd555 ON cxd555.DOCUMENTO = man111.LANCTO
--       LEFT JOIN CRD013 crd013 ON crd013.CODIGO = cxd555.TIPO
--       LEFT JOIN EMD101 emd101 ON emd101.CGC_CPF = man111.CNPJ_CPF
-- WHERE cxd555.HISTORICO NOT LIKE '%ENVELOPE%'