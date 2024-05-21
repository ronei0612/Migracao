DECLARE @EstabelecimentoID INT = 120;

select p.CPF, p.NomeCompleto, p.ID as PessoaID, f.ID as FuncionarioID, fo.ID as FornecedorID, fo.NomeFantasia, c.ID as ConsumidorID, c.CodigoAntigo 
from Pessoas p 
left join Consumidores c on c.PessoaID = p.ID  
left join Funcionarios f on f.PessoaID = p.ID 
left join Fornecedores fo on fo.PessoaID = p.ID  
where c.EstabelecimentoID=@EstabelecimentoID OR p.EstabelecimentoID=@EstabelecimentoID OR fo.EstabelecimentoID=@EstabelecimentoID