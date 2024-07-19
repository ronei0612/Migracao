CREATE PROCEDURE dbo.MigracaoPessoasConsumidores
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRANSACTION;
    BEGIN TRY
        -- Tabela temporária para armazenar os IDs das pessoas inseridas
        DECLARE @PessoasIDs TABLE (ID INT IDENTITY(1,1), PessoaID INT);
     
        -- Inserção dos dados na tabela Pessoas e armazenamento dos IDs inseridos
        INSERT INTO Pessoas (ConselhoCodigo, NomeCompleto, Apelido, CPF, DataInclusao, Email, RG, Sexo, NascimentoData, NascimentoLocal, ProfissaoOutra, EstadoCivilID, EstabelecimentoID, LoginID, Guid)
        OUTPUT INSERTED.ID INTO @PessoasIDs
        SELECT ConselhoCodigo, NomeCompleto, Apelido, CPF, DataInclusao, Email, RG, Sexo, NascimentoData, NascimentoLocal, ProfissaoOutra, EstadoCivilID, EstabelecimentoID, LoginID, Guid
        FROM [_MigracaoPessoas_Temp];
     
        -- Atualizar a tabela _MigracaoConsumidores_Temp com os números dos IDs gerados
        UPDATE tabelaTemp
        SET PessoaID = pessoas.PessoaID
        FROM [_MigracaoPessoaFones_Temp] tabelaTemp
        JOIN @PessoasIDs pessoas ON tabelaTemp.PessoaID = pessoas.ID;
        
        UPDATE tabelaTemp
        SET PessoaID = pessoas.PessoaID
        FROM [_MigracaoConsumidores_Temp] tabelaTemp
        JOIN @PessoasIDs pessoas ON tabelaTemp.ID = pessoas.ID;
     
        -- Inserção dos dados na tabela Pessoas e armazenamento dos IDs inseridos
        INSERT INTO PessoaFones (PessoaID, FoneTipoID, Telefone, DataInclusao, LoginID)
        SELECT PessoaID, FoneTipoID, Telefone, DataInclusao, LoginID
        FROM [_MigracaoPessoaFones_Temp];
       
       
        -- Tabela temporária para armazenar os IDs dos consumidores inseridos
        DECLARE @ConsumidoresIDs TABLE (ID INT IDENTITY(1,1), ConsumidorID INT);
       
       -- Inserção dos dados na tabela Pessoas e armazenamento dos IDs inseridos
        INSERT INTO Consumidores (Ativo, DataInclusao, EstabelecimentoID, LGPDSituacaoID, LoginID, PessoaID, CodigoAntigo)
        OUTPUT INSERTED.ID INTO @ConsumidoresIDs
        SELECT Ativo, DataInclusao, EstabelecimentoID, LGPDSituacaoID, LoginID, PessoaID, CodigoAntigo
        FROM [_MigracaoConsumidores_Temp];
       
     
        -- Atualizar a tabela _MigracaoConsumidores_Temp com os números dos IDs gerados
        UPDATE tabelaTemp
        SET ConsumidorID = consumidores.ConsumidorID
        FROM [_MigracaoConsumidorEnderecos_Temp] tabelaTemp
        JOIN @ConsumidoresIDs consumidores ON tabelaTemp.ConsumidorID = consumidores.ID;
     
        INSERT INTO ConsumidorEnderecos (Ativo, ConsumidorID, EnderecoTipoID, LogradouroTipoID, Logradouro, CidadeID, Cep, DataInclusao, Bairro, LogradouroNum, Complemento)
        OUTPUT INSERTED.ID INTO @ConsumidoresIDs
        SELECT Ativo, ConsumidorID, EnderecoTipoID, LogradouroTipoID, Logradouro, CidadeID, Cep, DataInclusao, Bairro, LogradouroNum, Complemento
        FROM [_MigracaoConsumidorEnderecos_Temp];
     
        COMMIT TRANSACTION;
       
        TRUNCATE TABLE [_MigracaoPessoas_Temp];
        TRUNCATE TABLE [_MigracaoConsumidores_Temp];
        TRUNCATE TABLE [_MigracaoPessoaFones_Temp];
        TRUNCATE TABLE [_MigracaoConsumidorEnderecos_Temp];
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        THROW;
    END CATCH
    SET NOCOUNT OFF;
END;
