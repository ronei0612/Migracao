DECLARE @EstabelecimentoID INT = 999999999;

SELECT 
    p.CPF, 
    p.NomeCompleto, 
    p.ID as PessoaID, 
    f.ID as FuncionarioID, 
    fo.ID as FornecedorID, 
    fo.NomeFantasia, 
    c.ID as ConsumidorID, 
    c.CodigoAntigo 
FROM Pessoas p
LEFT JOIN Consumidores c
    ON c.PessoaID = p.ID AND c.EstabelecimentoID = @EstabelecimentoID
LEFT JOIN Funcionarios f
    ON f.PessoaID = p.ID AND p.EstabelecimentoID = @EstabelecimentoID
LEFT JOIN Fornecedores fo
    ON fo.PessoaID = p.ID AND fo.EstabelecimentoID = @EstabelecimentoID
LEFT JOIN ConsumidorEnderecos ce
    ON ce.ConsumidorID = c.ID
LEFT JOIN PessoaFones pf
    ON pf.PessoaID = p.ID
WHERE 
    c.PessoaID IS NOT NULL OR 
    f.PessoaID IS NOT NULL OR 
    fo.PessoaID IS NOT NULL OR
    pf.PessoaID IS NOT NULL;