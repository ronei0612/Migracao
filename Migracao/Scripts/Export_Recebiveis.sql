DECLARE @EstabelecimentoID INT = 999999999;

SELECT
    *
FROM Recebiveis r
WHERE r.EstabelecimentoID = @EstabelecimentoID;