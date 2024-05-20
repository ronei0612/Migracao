CREATE PROCEDURE dbo.MigracaoPessoasConsumidores
AS
BEGIN
    SET NOCOUNT ON;
 
    -- Variável para armazenar o ID da pessoa inserida
    DECLARE @PrimeiroID INT;
    -- Tabela temporária para armazenar os IDs das pessoas inseridas
    DECLARE @IDs TABLE (ID INT IDENTITY(1,1), PessoaID INT);
 
    -- Inserção dos dados na tabela Pessoas e armazenamento dos IDs inseridos
    INSERT INTO Pessoas (NomeCompleto, Apelido, CPF, DataInclusao, Email, RG, Sexo, NascimentoData, NascimentoLocal, ProfissaoOutra, EstadoCivilID, LoginID, EstabelecimentoID)
    OUTPUT INSERTED.ID INTO @IDs
    SELECT NomeCompleto, Apelido, CPF, DataInclusao, Email, RG, Sexo, NascimentoData, NascimentoLocal, ProfissaoOutra, EstadoCivilID, LoginID, EstabelecimentoID
    FROM [_MigracaoPessoas_Temp];
 
    -- Pegar o primeiro ID inserido
    SELECT @PrimeiroID = MIN(PessoaID) FROM @IDs;
 
    -- Atualizar a tabela _MigracaoConsumidores_Temp com os números em ordem crescente a partir do primeiro ID gerado
    DECLARE @i INT = 0;
    WHILE @i < (SELECT COUNT(*) FROM [_MigracaoConsumidores_Temp])
    BEGIN

        UPDATE [_MigracaoPessoaFones_Temp]
        SET PessoaID = @PrimeiroID + @i 
        WHERE PessoaID = (SELECT PessoaID FROM [_MigracaoConsumidores_Temp] WHERE ID = @i + 1);
	    
        UPDATE [_MigracaoConsumidores_Temp]
        SET PessoaID = @PrimeiroID + @i
        WHERE ID = @i + 1;
 
        SET @i = @i + 1;
    END
   
   -- Inserção dos dados na tabela Pessoas e armazenamento dos IDs inseridos
    INSERT INTO PessoaFones (PessoaID, FoneTipoID, Telefone, DataInclusao, LoginID)
    SELECT PessoaID, FoneTipoID, Telefone, DataInclusao, LoginID
    FROM [_MigracaoPessoaFones_Temp];
   
   
   -- Variável para armazenar o ID da pessoa inserida
    DECLARE @PrimeiroConsudmidorID INT;
    -- Tabela temporária para armazenar os IDs das pessoas inseridas
    DECLARE @IDConsumidores TABLE (ID INT IDENTITY(1,1), ConsumidorID INT);
   
   -- Inserção dos dados na tabela Pessoas e armazenamento dos IDs inseridos
    INSERT INTO Consumidores (Ativo, DataInclusao, EstabelecimentoID, LGPDSituacaoID, LoginID, PessoaID, CodigoAntigo)
    OUTPUT INSERTED.ID INTO @IDConsumidores
    SELECT Ativo, DataInclusao, EstabelecimentoID, LGPDSituacaoID, LoginID, PessoaID, CodigoAntigo
    FROM [_MigracaoConsumidores_Temp];
   
   -- Pegar o primeiro ID inserido
    SELECT @PrimeiroConsudmidorID = MIN(ConsumidorID) FROM @IDConsumidores;
 
    -- Atualizar a tabela _MigracaoConsumidores_Temp com os números em ordem crescente a partir do primeiro ID gerado
    DECLARE @j INT = 0;
    WHILE @j < (SELECT COUNT(*) FROM [_MigracaoConsumidores_Temp])
    BEGIN
	    UPDATE [_MigracaoConsumidorEnderecos_Temp]
        SET ConsumidorID = @PrimeiroConsudmidorID + @j
        WHERE ConsumidorID = (SELECT ID FROM [_MigracaoConsumidores_Temp] WHERE ID = @j + 1);
 
        SET @j = @j + 1;
    END
    
    INSERT INTO ConsumidorEnderecos (Ativo, ConsumidorID, EnderecoTipoID, LogradouroTipoID, Logradouro, CidadeID, Cep, DataInclusao, Bairro, LogradouroNum, Complemento)
    OUTPUT INSERTED.ID INTO @IDConsumidores
    SELECT Ativo, ConsumidorID, EnderecoTipoID, LogradouroTipoID, Logradouro, CidadeID, Cep, DataInclusao, Bairro, LogradouroNum, Complemento
    FROM [_MigracaoConsumidorEnderecos_Temp];
 
    SET NOCOUNT OFF;
END;