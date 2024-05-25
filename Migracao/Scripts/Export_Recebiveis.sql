DECLARE @EstabelecimentoID INT = 999999999;

SELECT
    *
FROM Recebiveis r
LEFT JOIN FluxoCaixa fc ON fc.RecebivelID = r.ID
WHERE fc.pagovalor IS NULL
AND r.databaixa IS NULL
AND r.ExclusaoData IS NULL
AND r.EstabelecimentoID = @EstabelecimentoID;