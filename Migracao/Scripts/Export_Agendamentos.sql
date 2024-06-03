--DECLARE @EstabelecimentoID INT = 999999999;

--SELECT 
--    p.CPF,
--    p.NomeCompleto,
--    *
--FROM Agendamentos a
--LEFT JOIN Pessoas p ON a.PessoaID = p.ID
--WHERE 
--    a.EstabelecimentoID = @EstabelecimentoID;

DECLARE @EstabelecimentoID INT = 999999999;

SELECT 
    *
FROM Agendamentos
WHERE 
    EstabelecimentoID = @EstabelecimentoID;