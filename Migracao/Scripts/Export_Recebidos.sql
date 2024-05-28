DECLARE @EstabelecimentoID INT = 999999999;

SELECT
    r.ID as RecebivelID,
    f.ID as FluxoCaixaID,
    r.DataVencimento,
    r.ValorOriginal,
    r.ExclusaoMotivo,
    f.ConsumidorID,
    f.PagoValor,
    f.[Data]
FROM FluxoCaixa f
LEFT JOIN Recebiveis r ON f.RecebivelID = r.ID
WHERE 
    r.EstabelecimentoID = @EstabelecimentoID
AND TipoID = 1;