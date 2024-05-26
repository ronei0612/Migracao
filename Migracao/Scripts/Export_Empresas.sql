DECLARE @EstabelecimentoID INT = 999999999;

SELECT 
    e.NomeFantasia,
    e.RazaoSocial, 
    e.ID as EmpresaID, 
    f.ID as FornecedorID, 
    f.NomeFantasia,
    en.Logradouro,
    ff.Telefone 
FROM Empresas e
LEFT JOIN Fornecedores f
    ON f.EmpresaID = e.ID
LEFT JOIN Enderecos en
    ON en.ParentID = e.ID
LEFT JOIN FornecedorFones ff
    ON ff.FornecedorID = f.ID
WHERE 
    e.EstabelecimentoID = @EstabelecimentoID;