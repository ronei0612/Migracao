CREATE PROCEDURE dbo.MigracaoRecebidos
AS
BEGIN
    INSERT INTO FluxoCaixa (TipoID, Data, TransacaoID, EspecieID, DataBaseCalculo, DevidoValor, PagoValor, FinanceiroID, EstabelecimentoID, LoginID, DataInclusao, FuncionarioID, ConsumidorID, OutroSacadoNome)
    SELECT TipoID, Data, TransacaoID, EspecieID, DataBaseCalculo, DevidoValor, PagoValor, FinanceiroID, EstabelecimentoID, LoginID, DataInclusao, FuncionarioID, ConsumidorID, OutroSacadoNome
    FROM _MigracaoFluxoCaixa_Temp
END;
