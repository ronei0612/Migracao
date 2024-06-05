DECLARE @EstabelecimentoID INT = 999999999;

SELECT 
    pt.Nome,
    *
FROM Precos p
LEFT JOIN PrecosTabelas pt 
    ON p.TabelaID = pt.ID 
WHERE 
    pt.EstabelecimentoID = @EstabelecimentoID;