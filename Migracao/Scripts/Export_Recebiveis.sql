DECLARE @EstabelecimentoID INT = 999999999;
SELECT
r.ID,
    r.consumidorid,
r.datavencimento,
r.valororiginal,
r.ValorBaixa as pagoValor,
r.databaixa
FROM Recebiveis r
WHERE r.ExclusaoData IS NULL AND r.EstabelecimentoID = @EstabelecimentoID
UNION
SELECT
fc.ID,
    fc.consumidorid,
fc.DataBaseCalculo as datavencimento,
fc.DevidoValor,
fc.PagoValor as pagoValor,
fc.Data
FROM FluxoCaixa fc
WHERE fc.ExclusaoData IS NULL AND fc.EstabelecimentoID = @EstabelecimentoID

--SELECT
    --*
--FROM Recebiveis r
--LEFT JOIN FluxoCaixa fc ON fc.RecebivelID = r.ID
----WHERE fc.pagovalor IS NULL
----AND r.databaixa IS NULL
--WHERE r.ExclusaoData IS NULL
--AND r.EstabelecimentoID = @EstabelecimentoID;