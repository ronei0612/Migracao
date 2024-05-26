DECLARE @EstabelecimentoID INT = 999999999;

SELECT 
    p.CPF, 
    p.NomeCompleto, 
    p.ID as PessoaID, 
    f.ID as FuncionarioID, 
    fo.ID as FornecedorID, 
    fo.NomeFantasia, 
    c.ID as ConsumidorID, 
    c.CodigoAntigo,
    ce.Logradouro,
    pf.Telefone 
FROM Pessoas p
LEFT JOIN Consumidores c
    ON c.PessoaID = p.ID
LEFT JOIN Funcionarios f
    ON f.PessoaID = p.ID
LEFT JOIN Fornecedores fo
    ON fo.PessoaID = p.ID
LEFT JOIN ConsumidorEnderecos ce
    ON ce.ConsumidorID = c.ID
LEFT JOIN PessoaFones pf
    ON pf.PessoaID = p.ID
WHERE 
    p.EstabelecimentoID = @EstabelecimentoID;