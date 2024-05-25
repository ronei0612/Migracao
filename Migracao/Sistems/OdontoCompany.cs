using ExcelDataReader;
using Migracao.Models;
using Migracao.Utils;
using NPOI.SS.UserModel;

namespace Migracao.Sistems
{
    internal class OdontoCompany
	{
        string arquivoExcelCidades = "Files\\EnderecosCidades.xlsx";
		string arquivoExcelNomesUTF8 = "Files\\NomesUTF8.xlsx";

		public void ImportarRecebiveis(string arquivoExcel, string arquivoExcelConsumidores, int estabelecimentoID, int respFinanceiroPessoaID, int loginID, string arquivoExcelBaixa)
        {
			var dataHoje = DateTime.Now;
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			ISheet sheetConsumidores;
			try
			{
				IWorkbook workbookConsumidores = excelHelper.LerExcel(arquivoExcelConsumidores);
				sheetConsumidores = workbookConsumidores.GetSheetAt(0);
				excelHelper.InitializeDictionary(sheetConsumidores);
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelConsumidores}\": {ex.Message}");
			}

            //var excelRecebidosDict = ExcelRecebidosToDictionary(arquivoExcelBaixa);

			var fluxoCaixas = new List<FluxoCaixa>();
			var recebiveis = new List<Recebivel>();

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					string nomeCompleto = "", cpf = "";
                    string? outroSacadoNome = null, observacoes = null, documento = null;
					int recibo = 0, codigo = 0;
					int? consumidorID = null, fornecedorID = null, colaboradorID = null, funcionarioID = null, clienteID = null;
					decimal pagoValor = 0, valor = 0;
					byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
					DateTime dataPagamento = dataHoje, nascimentoData = dataHoje, dataVencimento = dataHoje, dataInclusao = dataHoje;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "CGC_CPF":
										cpf = celulaValor.ToCPF();
										break;
									case "DOCUMENTO":
										documento = celulaValor;
										break;
									case "VENCTO":
										dataVencimento = celulaValor.ToData();
										break;
									case "EMISSAO":
										dataInclusao = celulaValor.ToData();
										break;
									case "OBS":
                                        observacoes = celulaValor;//.ToTipoPagamento()
										break;
									case "VALOR_VENDA":
										valor = celulaValor.ToMoeda();
										break;
								}
							}
						}
					}

					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf:cpf, codigo: codigo.ToString());
					var fornecedorIDValue = excelHelper.GetFornecedorID(nomeCompleto: nomeCompleto, cpf: cpf);
					var funcionarioIDValue = excelHelper.GetFuncionarioID(nomeCompleto: nomeCompleto, cpf: cpf);

					if (!string.IsNullOrEmpty(consumidorIDValue))
						consumidorID = int.Parse(consumidorIDValue);
                    else if (!string.IsNullOrEmpty(fornecedorIDValue))
						fornecedorID = int.Parse(fornecedorIDValue);
					else if (!string.IsNullOrEmpty(funcionarioIDValue))
						funcionarioID = int.Parse(funcionarioIDValue);
                    else
					    outroSacadoNome = cpf;

                    recebiveis.Add(new Recebivel()
                    {
						ConsumidorID = consumidorID,
					    FornecedorID = fornecedorID,
					    ClienteID = clienteID,
					    ColaboradorID = colaboradorID,
					    SacadoNome = outroSacadoNome,
                        EspecieID = (byte)formaPagamento,
						DataEmissao = dataInclusao,
						ValorOriginal = valor,
						ValorDevido = valor,
                        DataBaseCalculo = dataInclusao,
						DataInclusao = dataInclusao,
                        DataVencimento = dataVencimento,
                        FinanceiroID = respFinanceiroPessoaID,
                        LoginID = loginID,
                        EstabelecimentoID = estabelecimentoID,
                        SituacaoID = (byte)TituloSituacoesID.Normal,
                        Observacoes = observacoes,
                        ExclusaoMotivo = documento
                        //OrcamentoID
						//PlanoContasID
						//Documento = contratoControle
						//ContratoID = contratoID
					});
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "ConsumidorID", recebiveis.ConvertAll(recebivel => (object)recebivel.ConsumidorID).ToArray() },
                    { "FornecedorID", recebiveis.ConvertAll(recebivel => (object)recebivel.FornecedorID).ToArray() },
                    { "ClienteID", recebiveis.ConvertAll(recebivel => (object)recebivel.ClienteID).ToArray() },
                    { "ColaboradorID", recebiveis.ConvertAll(recebivel => (object)recebivel.ColaboradorID).ToArray() },
                    { "SacadoNome", recebiveis.ConvertAll(recebivel => (object)recebivel.SacadoNome).ToArray() },
                    { "EspecieID", recebiveis.ConvertAll(recebivel => (object)recebivel.EspecieID).ToArray() },
                    { "DataEmissao", recebiveis.ConvertAll(recebivel => (object)recebivel.DataEmissao).ToArray() },
                    { "ValorOriginal", recebiveis.ConvertAll(recebivel => (object)recebivel.ValorOriginal).ToArray() },
                    { "ValorDevido", recebiveis.ConvertAll(recebivel => (object)recebivel.ValorDevido).ToArray() },
                    { "DataBaseCalculo", recebiveis.ConvertAll(recebivel => (object)recebivel.DataBaseCalculo).ToArray() },
                    { "DataInclusao", recebiveis.ConvertAll(recebivel => (object)recebivel.DataInclusao).ToArray() },
                    { "DataVencimento", recebiveis.ConvertAll(recebivel => (object)recebivel.DataVencimento).ToArray() },
                    { "FinanceiroID", recebiveis.ConvertAll(recebivel => (object)recebivel.FinanceiroID).ToArray() },
                    { "LoginID", recebiveis.ConvertAll(recebivel => (object)recebivel.LoginID).ToArray() },
                    { "EstabelecimentoID", recebiveis.ConvertAll(recebivel => (object)recebivel.EstabelecimentoID).ToArray() },
                    { "SituacaoID", recebiveis.ConvertAll(recebivel => (object)recebivel.SituacaoID).ToArray() },
                    { "Observacoes", recebiveis.ConvertAll(recebivel => (object)recebivel.Observacoes).ToArray() },
                    { "ExclusaoMotivo", recebiveis.ConvertAll(recebivel => (object)recebivel.ExclusaoMotivo).ToArray() }
				};

				var salvarArquivo = Tools.GerarNomeArquivo($"Recebiveis_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Recebiveis", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
				Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

        public void ImportarRecebidos(string arquivoExcel, string arquivoExcelConsumidores, int estabelecimentoID, int respFinanceiroPessoaID, int loginID, string arquivoExcelRecebiveis)
        {
            //CRD013 Forma de Pagamento
            var dataHoje = DateTime.Now;
            var indiceLinha = 0;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            ISheet sheetConsumidores;
            try
            {
                IWorkbook workbookConsumidores = excelHelper.LerExcel(arquivoExcelConsumidores);
                sheetConsumidores = workbookConsumidores.GetSheetAt(0);
                excelHelper.InitializeDictionary(sheetConsumidores);
            }
            catch (Exception ex)
            {
                throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelConsumidores}\": {ex.Message}");
            }

            var excelRecebidosDict = ExcelRecebiveisToDictionary(arquivoExcelRecebiveis);

            var fluxoCaixas = new List<FluxoCaixa>();

            try
            {
                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;

                    string documento = "";
                    string? observacao = null;
                    decimal pagoValor = 0;
                    byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
                    DateTime dataBaixa = DateTime.Now;

                    foreach (var celula in linha.Cells)
                    {
                        celulaValor = celula.ToString().Trim().Replace("'", "’");
                        tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
                        colunaLetra = excelHelper.GetColumnLetter(celula);

                        switch (tituloColuna)
                        {
                            case "DOCUMENTO":
                                documento = celulaValor;
                                break;
                            case "VALOR":
                                pagoValor = celulaValor.ArredondarValor();
                                break;
                            case "BAIXA":
                                dataBaixa = celulaValor.ToData();
                                break;
                            case "MOTIVO":
                                observacao = celulaValor;
                                break;
                        }
                    }

                    if (!string.IsNullOrEmpty(documento) && excelRecebidosDict.ContainsKey(documento))
                        fluxoCaixas.Add(new FluxoCaixa()
                        {
                            RecebivelID = int.Parse(excelRecebidosDict[documento][0]),
                            ConsumidorID = int.Parse(excelRecebidosDict[documento][1]),
                            SituacaoID = 1,
                            PagoMulta = 0,
                            PagoJuros = 0,
                            TipoID = (byte)TransacaoTiposID.Recebimento,
                            Data = dataBaixa,
                            TransacaoID = (byte)TituloTransacoes.PagamentoAvulso,
                            EspecieID = formaPagamento,
                            DataBaseCalculo = dataBaixa,
                            DevidoValor = pagoValor,
                            PagoValor = pagoValor,
                            EstabelecimentoID = estabelecimentoID,
                            LoginID = loginID,
                            DataInclusao = dataBaixa,
                            FinanceiroID = respFinanceiroPessoaID,
                            Observacoes = observacao
                        });
                }

                indiceLinha = 0;

                var dados = new Dictionary<string, object[]>
                {
                    { "ConsumidorID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.ConsumidorID).ToArray() },
                    { "RecebivelID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.RecebivelID).ToArray() },
                    { "SituacaoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.SituacaoID).ToArray() },
                    { "PagoMulta", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoMulta).ToArray() },
                    { "PagoJuros", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoJuros).ToArray() },
                    { "TipoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TipoID).ToArray() },
                    { "Data", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.Data).ToArray() },
                    { "TransacaoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TransacaoID).ToArray() },
                    { "EspecieID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EspecieID).ToArray() },
                    { "DataBaseCalculo", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataBaseCalculo).ToArray() },
                    { "DevidoValor", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DevidoValor).ToArray() },
                    { "PagoValor", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoValor).ToArray() },
                    { "EstabelecimentoID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EstabelecimentoID).ToArray() },
                    { "LoginID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.LoginID).ToArray() },
                    { "DataInclusao", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataInclusao).ToArray() },
                    { "FinanceiroID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.FinanceiroID).ToArray() },
                    { "Observacoes", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.Observacoes).ToArray() }
                };

                var salvarArquivo = Tools.GerarNomeArquivo($"Recebidos_{estabelecimentoID}_OdontoCompany_Migração");
                sqlHelper.GerarSqlInsert("FluxoCaixa", salvarArquivo, dados);
                excelHelper.GravarExcel(salvarArquivo, dados);
                Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
            }
            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

        public void ImportarFornecedores(string arquivoExcel, string arquivo2, int estabelecimentoID, int loginID)
        {
            var dataHoje = DateTime.Now;
            var indiceLinha = 0;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            if (File.Exists(arquivoExcelCidades))
                try
                {
                    var workbookCidades = excelHelper.LerExcel(arquivoExcelCidades);
                    var sheetCidades = workbookCidades.GetSheetAt(0);
                    excelHelper.InitializeDictionaryCidade(sheetCidades);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelCidades}\": {ex.Message}");
                }

            try
            {
                var linhasCount = excelHelper.linhas.Count;
                //var consumidores = new List<Consumidor>();
                var fornecedores = new List<Fornecedor>();

                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;
                    bool cliente = false, fornecedor = false;
                    DateTime dataNascimento, dataCadastro = dataHoje;
                    int numcadastro;
                    string nomeCompleto = "", cpf = "", rg = "", email = "", apelido = "";
                    bool sexo = true;
                    long telefonePrinc, telefoneAltern, telefoneComercial, telefoneOutro, celular;

                    foreach (var celula in linha.Cells)
                    {
                        if (celula != null)
                        {
                            celulaValor = celula.ToString().Trim().Replace("'", "’");
                            tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
                            colunaLetra = excelHelper.GetColumnLetter(celula);

                            if (!string.IsNullOrWhiteSpace(celulaValor))
                            {
                                switch (tituloColuna)
                                {
                                    case "CLIENTE":
                                        cliente = celulaValor == "S" ? true : false;
                                        break;
                                    case "FORNECEDOR":
                                        fornecedor = celulaValor == "S" ? true : false;
                                        break;
                                    case "NOME":
                                        nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                        apelido = celulaValor.GetPrimeiroNome();
                                        break;
                                    case "CGC_CPF":
                                        cpf = celulaValor.ToCPF();
                                        break;
                                    case "INSC_RG":
                                        rg = celulaValor.GetPrimeirosCaracteres(20);
                                        break;
                                    case "SEXO_M_F":
                                        var sexoLetra = celulaValor.ToLower();
                                        sexo = sexoLetra == "m" || sexoLetra != "f";
                                        break;
                                    case "EMAIL":
                                        email = celulaValor.ToEmail();
                                        break;
                                    case "FONE1":
                                        telefonePrinc = celulaValor.ToFone();
                                        break;
                                    case "FONE2":
                                        telefoneAltern = celulaValor.ToFone();
                                        break;
                                    case "CELULAR":
                                        celular = celulaValor.ToFone();
                                        break;
                                    //case "ENDERECO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "BAIRRO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "NUM_ENDERECO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "CIDADE":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "ESTADO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "CEP":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "OBS1":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    //case "NUM_CONVENIO":
                                    //    nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                    //    break;
                                    case "DT_CADASTRO":
                                        dataCadastro = celulaValor.ToData();
                                        break;
                                    case "DT_NASCIMENTO":
                                        dataNascimento = celulaValor.ToData();
                                        break;
                                }
                            }
                        }
                    }

                    if (fornecedor)
                        fornecedores.Add(new Fornecedor()
                        {
                            Ativo = true,
                            DataInclusao = dataCadastro,
                            EstabelecimentoID = estabelecimentoID,
                            NomeFantasia = nomeCompleto,
                            LoginID = loginID
                        });
                }

                indiceLinha = 0;

                var dados = new Dictionary<string, object[]>
                {
                    { "Ativo", fornecedores.ConvertAll(fornecedor => (object)fornecedor.Ativo).ToArray() },
                    { "DataInclusao", fornecedores.ConvertAll(fornecedor => (object)fornecedor.DataInclusao).ToArray() },
                    { "EstabelecimentoID", fornecedores.ConvertAll(fornecedor => (object)fornecedor.EstabelecimentoID).ToArray() },
                    { "LoginID", fornecedores.ConvertAll(fornecedor => (object)fornecedor.LoginID).ToArray() },
                    { "NomeFantasia", fornecedores.ConvertAll(fornecedor => (object)fornecedor.NomeFantasia).ToArray() }
                };

				var salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_OdontoCompany_Consumidores");
				sqlHelper.GerarSqlInsert("_MigracaoConsumidores_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
				Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
            }

            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

		public void ImportarPacientes(string arquivoExcel, string arquivoPacientesAtuais, int estabelecimentoID, int loginID)
        {
			var indiceLinha = 1;
            var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            DateTime dataHoje = DateTime.Now;
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            if (File.Exists(arquivoExcelCidades))
                try
                {
                    var workbookCidades = excelHelper.LerExcel(arquivoExcelCidades);
                    var sheetCidades = workbookCidades.GetSheetAt(0);
                    excelHelper.InitializeDictionaryCidade(sheetCidades);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelCidades}\": {ex.Message}");
                }

            if (File.Exists(arquivoExcelNomesUTF8))
			    try
			    {
				    var workbookCidades = excelHelper.LerExcel(arquivoExcelNomesUTF8);
				    var sheetCidades = workbookCidades.GetSheetAt(0);
				    excelHelper.InitializeDictionaryNomesUTF8(sheetCidades);
			    }
			    catch (Exception ex)
			    {
				    throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelNomesUTF8}\": {ex.Message}");
			    }

			if (!string.IsNullOrEmpty(arquivoPacientesAtuais))
			    try
			    {
				    var workbook = excelHelper.LerExcel(arquivoPacientesAtuais);
				    var sheet = workbook.GetSheetAt(0);
				    excelHelper.InitializeDictionary(sheet);
			    }
			    catch (Exception ex)
			    {
				    throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoPacientesAtuais}\": {ex.Message}");
			    }


			try
            {
                var linhasCount = excelHelper.linhas.Count;
                var consumidores = new List<Consumidor>();
                var pessoas = new List<Pessoa>();
                var consumidoresEnderecos = new List<ConsumidorEndereco>();
                var pessoaFones = new List<PessoaFone>();
				var fornecedores = new List<Fornecedor>();
				var empresas = new List<Empresa>();
                var empresasEnderecos = new List<Endereco>();
				var fornecedorFones = new List<FornecedorFone>();

				foreach (var linha in excelHelper.linhas)
                {
                    bool cliente = false, fornecedor = false;
                    DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
                    int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
                    long telefonePrinc = 0, telefoneAltern = 0, telefoneComercial = 0, telefoneOutro = 0, celular = 0;
					string nomeCompleto = "null", documento = "null", rg = "null", email = "null", apelido = "null", nascimentoLocal = "null", profissaoOutra = "null", logradouro = "",
						 complemento = "null", bairro = "null", logradouroNum = "null", numcadastro = "null", cidade = "", estado = "null", observacao = "null";

					foreach (var celula in linha.Cells)
                    {
                        if (celula != null)
                        {
                            celulaValor = celula.ToString().Trim().Replace("'", "’");
                            tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
                            colunaLetra = excelHelper.GetColumnLetter(celula);

                            if (!string.IsNullOrWhiteSpace(celulaValor))
                            {
                                switch (tituloColuna)
                                {
                                    case "CLIENTE":
                                        cliente = celulaValor == "S" ? true : false;
                                        break;
                                    case "FORNECEDOR":
                                        fornecedor = celulaValor == "S" ? true : false;
                                        break;
                                    case "NOME":
                                        nomeCompleto = excelHelper.CorrigirNomeUTF8(celulaValor.GetLetras().GetPrimeirosCaracteres(70)).ToNomeCompleto();
                                        apelido = excelHelper.CorrigirNomeUTF8(celulaValor.GetLetras().GetPrimeiroNome()).ToNomeCompleto();
                                        break;
                                    case "CGC_CPF":
										documento = celulaValor.ToCPF();
                                        break;
                                    case "INSC_RG":
                                        rg = celulaValor.GetPrimeirosCaracteres(20);
                                        break;
                                    case "SEXO_M_F":
                                        sexo = celulaValor.ToSexo("m", "f");
                                        break;
                                    case "EMAIL":
                                        email = celulaValor.ToEmail();
                                        break;
                                    case "FONE1":
                                        telefonePrinc = celulaValor.ToFone();
                                        break;
                                    case "FONE2":
                                        telefoneAltern = celulaValor.ToFone();
                                        break;
                                    case "CELULAR":
                                        celular = celulaValor.ToFone();
                                        break;
                                    case "ENDERECO":
                                        logradouro = excelHelper.CorrigirNomeUTF8(celulaValor).ToNomeCompleto();
                                        break;
                                    case "BAIRRO":
                                        bairro = excelHelper.CorrigirNomeUTF8(celulaValor).ToNomeCompleto();
                                        break;
                                    case "NUM_ENDERECO":
                                        logradouroNum = celulaValor;
                                        break;
                                    case "CIDADE":
                                        cidade = excelHelper.CorrigirNomeUTF8(celulaValor);
                                        break;
                                    case "ESTADO":
                                        estado = celulaValor;
                                        break;
                                    case "CEP":
                                        cep = celulaValor.ToNum();
                                        break;
                                    case "OBS1":
                                        observacao = excelHelper.CorrigirNomeUTF8(celulaValor);
										break;
                                    case "NUM_CONVENIO":
                                        break;
                                    case "DT_CADASTRO":
                                        dataCadastro = celulaValor.ToData();
                                        break;
                                    case "DT_NASCIMENTO":
                                        dataNascimento = celulaValor.ToData();
                                        break;
									case "NUM_FICHA":
										numcadastro = celulaValor;
										break;
								}
                            }
                        }
                    }

                    pessoaID = indiceLinha;
					var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento);
					if (!string.IsNullOrEmpty(pessoaIDValue))
						pessoaID = int.Parse(pessoaIDValue);

					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: documento);
                    if (!string.IsNullOrEmpty(consumidorIDValue))
                        consumidorID = int.Parse(consumidorIDValue);

					if (!string.IsNullOrWhiteSpace(nomeCompleto))
                    {
                        if (documento.IsCPF())
                        {
							if (string.IsNullOrEmpty(pessoaIDValue))
							    pessoas.Add(new Pessoa()
                                {
                                    ID = indiceLinha,
                                    NomeCompleto = nomeCompleto,
                                    Apelido = apelido,
                                    CPF = documento,
                                    DataInclusao = dataCadastro,
                                    Email = email,
                                    RG = rg,
                                    Sexo = sexo,
                                    NascimentoData = dataNascimento,
                                    NascimentoLocal = nascimentoLocal,
                                    ProfissaoOutra = profissaoOutra,
                                    EstadoCivilID = estadoCivil,
                                    EstabelecimentoID = estabelecimentoID,
                                    LoginID = loginID,
                                    Guid = new Guid(),
                                    FoneticaApelido = apelido.Fonetizar(),
                                    FoneticaNomeCompleto = nomeCompleto.Fonetizar()
                                });
                        }

                        else if (documento.IsCNPJ_CGC())
                        {
                            empresas.Add(new Empresa()
                            {
                                Ativo = true,
                                CNPJ = documento,
                                DataInclusao = dataCadastro,
                                LoginID = loginID,
                                Marca = "",
                                NomeFantasia = "",
                                RazaoSocial = "",
                                RegimeTribID = 0,
                                InscricaoMunicipal = rg,
                                EstabelecimentoID = estabelecimentoID,
                                Guid = new Guid()
                            });
                        }

                        if (fornecedor)
                        {
                            fornecedores.Add(new Fornecedor()
                            {
                                Ativo = true,
                                DataInclusao = dataCadastro,
                                EstabelecimentoID = estabelecimentoID,
                                LoginID = loginID,
                                Email = email,
                                EmpresaID = indiceLinha,
                                Observacoes = observacao
                            });

                            if (!string.IsNullOrWhiteSpace(cidade) && !string.IsNullOrWhiteSpace(logradouro))
                            {
                                var cidadeID = excelHelper.GetCidadeID(cidade, estado);
                                empresasEnderecos.Add(new Endereco()
                                {
                                    Ativo = true,
                                    Cep = cep,
                                    CidadeID = cidadeID,
                                    DataInclusao = dataCadastro,
                                    EnderecoTipoID = (byte)EnderecoTipos.Residencial,
                                    Logradouro = logradouro,
                                    LogradouroNum = logradouroNum,
                                    Bairro = bairro,
                                    LogradouroTipoID = (int)LogradouroTipos.Outros,
                                    Complemento = complemento,
                                    ParentID = indiceLinha,
                                    TableID = 1,
                                });
                            }

                            if (celular > 0 && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, celular.ToString()))
                                fornecedorFones.Add(new FornecedorFone()
                                {
                                    FornecedorID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Celular,
                                    Telefone = celular,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefonePrinc > 0 && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefonePrinc.ToString()))
								fornecedorFones.Add(new FornecedorFone()
                                {
                                    FornecedorID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Principal,
                                    Telefone = telefonePrinc,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneAltern > 0 && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefoneAltern.ToString()))
								fornecedorFones.Add(new FornecedorFone()
                                {
                                    FornecedorID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Alternativo,
                                    Telefone = telefoneAltern,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneComercial > 0 && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefoneComercial.ToString()))
								fornecedorFones.Add(new FornecedorFone()
                                {
                                    FornecedorID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Comercial,
                                    Telefone = telefoneComercial,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneOutro > 0 && !excelHelper.ConsumidorEnderecoExists(documento, nomeCompleto, telefoneOutro.ToString()))
								fornecedorFones.Add(new FornecedorFone()
                                {
                                    FornecedorID = indiceLinha,
                                    FoneTipoID = (short)FoneTipos.Outros,
                                    Telefone = telefoneOutro,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });
                        }

                        else if (!fornecedor)
                        {
                            if (string.IsNullOrEmpty(arquivoPacientesAtuais) == false)
                            {
                                if (string.IsNullOrEmpty(consumidorIDValue) == false)
                                    consumidores.Add(new Consumidor()
                                    {
                                        Ativo = true,
                                        DataInclusao = dataCadastro,
                                        EstabelecimentoID = estabelecimentoID,
                                        LGPDSituacaoID = 0,
                                        LoginID = loginID,
                                        PessoaID = pessoaID,
                                        CodigoAntigo = numcadastro,
                                        Observacoes = observacao
                                    });
                            }
                            else
                                consumidores.Add(new Consumidor()
                                {
                                    Ativo = true,
                                    DataInclusao = dataCadastro,
                                    EstabelecimentoID = estabelecimentoID,
                                    LGPDSituacaoID = 0,
                                    LoginID = loginID,
                                    PessoaID = pessoaID,
                                    CodigoAntigo = numcadastro,
                                    Observacoes = observacao
                                });

                            if (!string.IsNullOrWhiteSpace(cidade) && !string.IsNullOrWhiteSpace(logradouro))
                            {
                                var cidadeID = excelHelper.GetCidadeID(cidade, estado);

                                if (string.IsNullOrEmpty(arquivoPacientesAtuais) == false)
                                {
                                    if (string.IsNullOrEmpty(consumidorIDValue) == false && cidadeID > 0)
                                        consumidoresEnderecos.Add(new ConsumidorEndereco()
                                        {
                                            Ativo = true,
                                            ConsumidorID = consumidorID,
                                            EnderecoTipoID = (short)EnderecoTipos.Residencial,
                                            LogradouroTipoID = (int)LogradouroTipos.Outros,
                                            Logradouro = logradouro,
                                            CidadeID = cidadeID,
                                            Cep = cep,
                                            DataInclusao = dataCadastro,
                                            Bairro = bairro,
                                            LogradouroNum = logradouroNum,
                                            Complemento = complemento,
                                            LoginID = loginID
                                        });
                                }
                                else
									consumidoresEnderecos.Add(new ConsumidorEndereco()
									{
										Ativo = true,
										ConsumidorID = consumidorID,
										EnderecoTipoID = (short)EnderecoTipos.Residencial,
										LogradouroTipoID = (int)LogradouroTipos.Outros,
										Logradouro = logradouro,
										CidadeID = cidadeID,
										Cep = cep,
										DataInclusao = dataCadastro,
										Bairro = bairro,
										LogradouroNum = logradouroNum,
										Complemento = complemento,
										LoginID = loginID
									});
							}

							if (string.IsNullOrEmpty(arquivoPacientesAtuais) == false)
							    consumidorID++;

                            if (celular > 0)
                                if (string.IsNullOrEmpty(arquivoPacientesAtuais) ||
                                    (!string.IsNullOrEmpty(arquivoPacientesAtuais) && !string.IsNullOrEmpty(pessoaIDValue) && !excelHelper.PessoaFoneExists(documento, nomeCompleto, celular.ToString())))
                                pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = pessoaID,
                                    FoneTipoID = (short)FoneTipos.Celular,
                                    Telefone = celular,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefonePrinc > 0)
								if (string.IsNullOrEmpty(arquivoPacientesAtuais) ||
									(!string.IsNullOrEmpty(arquivoPacientesAtuais) && !string.IsNullOrEmpty(pessoaIDValue) && !excelHelper.PessoaFoneExists(documento, nomeCompleto, telefonePrinc.ToString())))
								pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = pessoaID,
                                    FoneTipoID = (short)FoneTipos.Principal,
                                    Telefone = telefonePrinc,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneAltern > 0)
								if (string.IsNullOrEmpty(arquivoPacientesAtuais) ||
									(!string.IsNullOrEmpty(arquivoPacientesAtuais) && !string.IsNullOrEmpty(pessoaIDValue) && !excelHelper.PessoaFoneExists(documento, nomeCompleto, telefoneAltern.ToString())))
								pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = pessoaID,
                                    FoneTipoID = (short)FoneTipos.Alternativo,
                                    Telefone = telefoneAltern,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneComercial > 0)
								if (string.IsNullOrEmpty(arquivoPacientesAtuais) ||
									(!string.IsNullOrEmpty(arquivoPacientesAtuais) && !string.IsNullOrEmpty(pessoaIDValue) && !excelHelper.PessoaFoneExists(documento, nomeCompleto, telefoneComercial.ToString())))
								pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = pessoaID,
                                    FoneTipoID = (short)FoneTipos.Comercial,
                                    Telefone = telefoneComercial,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });

                            if (telefoneOutro > 0)
								if (string.IsNullOrEmpty(arquivoPacientesAtuais) ||
									(!string.IsNullOrEmpty(arquivoPacientesAtuais) && !string.IsNullOrEmpty(pessoaIDValue) && !excelHelper.PessoaFoneExists(documento, nomeCompleto, telefoneOutro.ToString())))
								pessoaFones.Add(new PessoaFone()
                                {
                                    PessoaID = pessoaID,
                                    FoneTipoID = (short)FoneTipos.Outros,
                                    Telefone = telefoneOutro,
                                    DataInclusao = dataCadastro,
                                    LoginID = loginID
                                });
                        }
                    }

					indiceLinha++;
				}

                indiceLinha = 0;
                var salvarArquivo = "";


				var pessoasDict = new Dictionary<string, object[]>
                {
					//{ "ID", pessoas.ConvertAll(pessoa => (object)pessoa.ID).ToArray() },
					{ "NomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.NomeCompleto).ToArray() },
                    { "Apelido", pessoas.ConvertAll(pessoa => (object)pessoa.Apelido).ToArray() },
                    { "CPF", pessoas.ConvertAll(pessoa => (object)pessoa.CPF).ToArray() },
                    { "DataInclusao", pessoas.ConvertAll(pessoa => (object)pessoa.DataInclusao).ToArray() },
                    { "Email", pessoas.ConvertAll(pessoa => (object)pessoa.Email).ToArray() },
                    { "RG", pessoas.ConvertAll(pessoa => (object)pessoa.RG).ToArray() },
                    { "Sexo", pessoas.ConvertAll(pessoa => (object)pessoa.Sexo).ToArray() },
					{ "NascimentoData", pessoas.ConvertAll(pessoa => (object)pessoa.NascimentoData).ToArray() },
					{ "NascimentoLocal", pessoas.ConvertAll(pessoa => (object)pessoa.NascimentoLocal).ToArray() },
					{ "ProfissaoOutra", pessoas.ConvertAll(pessoa => (object)pessoa.ProfissaoOutra).ToArray() },
					{ "EstadoCivilID", pessoas.ConvertAll(pessoa => (object)pessoa.EstadoCivilID).ToArray() },
					{ "EstabelecimentoID", pessoas.ConvertAll(pessoa => (object)pessoa.EstabelecimentoID).ToArray() },
					{ "LoginID", pessoas.ConvertAll(pessoa => (object)pessoa.LoginID).ToArray() },
					{ "Guid", pessoas.ConvertAll(pessoa => (object)pessoa.Guid).ToArray() },
					{ "FoneticaApelido", pessoas.ConvertAll(pessoa => (object)pessoa.FoneticaApelido).ToArray() },
					{ "FoneticaNomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.FoneticaNomeCompleto).ToArray() }
				};

                salvarArquivo = Tools.GerarNomeArquivo($"Pessoas_{estabelecimentoID}_OdontoCompany_Migração");
                sqlHelper.GerarSqlInsert("_MigracaoPessoas_Temp", salvarArquivo, pessoasDict);
                excelHelper.GravarExcel(salvarArquivo, pessoasDict);
				Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");


				var consumidoresDict = new Dictionary<string, object[]>
				{
					{ "Ativo", consumidores.ConvertAll(consumidor => (object)consumidor.Ativo).ToArray() },
                    { "DataInclusao", consumidores.ConvertAll(consumidor => (object)consumidor.DataInclusao).ToArray() },
                    { "EstabelecimentoID", consumidores.ConvertAll(consumidor => (object)consumidor.EstabelecimentoID).ToArray() },
                    { "LGPDSituacaoID", consumidores.ConvertAll(consumidor => (object)consumidor.LGPDSituacaoID).ToArray() },
                    { "LoginID", consumidores.ConvertAll(consumidor => (object)consumidor.LoginID).ToArray() },
                    { "PessoaID", consumidores.ConvertAll(consumidor => (object)consumidor.PessoaID).ToArray() },
                    { "CodigoAntigo", consumidores.ConvertAll(consumidor => (object)consumidor.CodigoAntigo).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Consumidores_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoConsumidores_Temp", salvarArquivo, consumidoresDict);
				excelHelper.GravarExcel(salvarArquivo, consumidoresDict);


				var consumidoresEnderecosDict = new Dictionary<string, object[]>
				{
					{ "LoginID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.LoginID).ToArray() },
					{ "Ativo", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Ativo).ToArray() },
                    { "ConsumidorID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.ConsumidorID).ToArray() },
                    { "EnderecoTipoID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.EnderecoTipoID).ToArray() },
                    { "LogradouroTipoID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.LogradouroTipoID).ToArray() },
                    { "Logradouro", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Logradouro).ToArray() },
                    { "CidadeID", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.CidadeID).ToArray() },
                    { "Cep", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Cep).ToArray() },
                    { "DataInclusao", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.DataInclusao).ToArray() },
                    { "Bairro", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Bairro).ToArray() },
                    { "LogradouroNum", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.LogradouroNum).ToArray() },
                    { "Complemento", consumidoresEnderecos.ConvertAll(endereco => (object)endereco.Complemento).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"ConsumidorEnderecos_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoConsumidorEnderecos_Temp", salvarArquivo, consumidoresEnderecosDict);
				excelHelper.GravarExcel(salvarArquivo, consumidoresEnderecosDict);
				

				var pessoaFonesDict = new Dictionary<string, object[]>
				{
					{ "PessoaID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.PessoaID).ToArray() },
					{ "FoneTipoID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.FoneTipoID).ToArray() },
					{ "Telefone", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.Telefone).ToArray() },
					{ "DataInclusao", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.DataInclusao).ToArray() },
					{ "LoginID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"PessoaFones_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoPessoaFones_Temp", salvarArquivo, pessoaFonesDict);
				excelHelper.GravarExcel(salvarArquivo, pessoaFonesDict);


				var fornecedorFonesDict = new Dictionary<string, object[]>
				{
					{ "FornecedorID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.FornecedorID).ToArray() },
					{ "FoneTipoID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.FoneTipoID).ToArray() },
					{ "Telefone", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.Telefone).ToArray() },
					{ "DataInclusao", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.DataInclusao).ToArray() },
					{ "LoginID", fornecedorFones.ConvertAll(fornecedorFone => (object)fornecedorFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"FornecedorFones_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoFornecedorFones_Temp", salvarArquivo, fornecedorFonesDict);
				excelHelper.GravarExcel(salvarArquivo, fornecedorFonesDict);


				var empresasDict = new Dictionary<string, object[]>
                {
                    { "Ativo", empresas.ConvertAll(empresa => (object)empresa.Ativo).ToArray() },
                    { "CNPJ", empresas.ConvertAll(empresa => (object)empresa.CNPJ).ToArray() },
                    { "DataInclusao", empresas.ConvertAll(empresa => (object)empresa.DataInclusao).ToArray() },
                    { "LoginID", empresas.ConvertAll(empresa => (object)empresa.LoginID).ToArray() },
                    { "Marca", empresas.ConvertAll(empresa => (object)empresa.Marca).ToArray() },
                    { "NomeFantasia", empresas.ConvertAll(empresa => (object)empresa.NomeFantasia).ToArray() },
                    { "RazaoSocial", empresas.ConvertAll(empresa => (object)empresa.RazaoSocial).ToArray() },
                    { "RegimeTribID", empresas.ConvertAll(empresa => (object)empresa.RegimeTribID).ToArray() },
                    { "InscricaoMunicipal", empresas.ConvertAll(empresa => (object)empresa.InscricaoMunicipal).ToArray() },
                    { "EstabelecimentoID", empresas.ConvertAll(empresa => (object)empresa.EstabelecimentoID).ToArray() },
                    { "Guid", empresas.ConvertAll(empresa => (object)empresa.Guid).ToArray() }
                };

				salvarArquivo = Tools.GerarNomeArquivo($"Empresas_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoEmpresas_Temp", salvarArquivo, empresasDict);
				excelHelper.GravarExcel(salvarArquivo, empresasDict);


				var empresasEnderecosDict = new Dictionary<string, object[]>
				{
	                { "Ativo", empresasEnderecos.ConvertAll(endereco => (object)endereco.Ativo).ToArray() },
	                { "Cep", empresasEnderecos.ConvertAll(endereco => (object)endereco.Cep).ToArray() },
	                { "CidadeID", empresasEnderecos.ConvertAll(endereco => (object)endereco.CidadeID).ToArray() },
	                { "DataInclusao", empresasEnderecos.ConvertAll(endereco => (object)endereco.DataInclusao).ToArray() },
	                { "EnderecoTipoID", empresasEnderecos.ConvertAll(endereco => (object)endereco.EnderecoTipoID).ToArray() },
	                { "Logradouro", empresasEnderecos.ConvertAll(endereco => (object)endereco.Logradouro).ToArray() },
	                { "LogradouroNum", empresasEnderecos.ConvertAll(endereco => (object)endereco.LogradouroNum).ToArray() },
	                { "Bairro", empresasEnderecos.ConvertAll(endereco => (object)endereco.Bairro).ToArray() },
	                { "LogradouroTipoID", empresasEnderecos.ConvertAll(endereco => (object)endereco.LogradouroTipoID).ToArray() },
	                { "Complemento", empresasEnderecos.ConvertAll(endereco => (object)endereco.Complemento).ToArray() },
	                { "ParentID", empresasEnderecos.ConvertAll(endereco => (object)endereco.ParentID).ToArray() },
	                { "TableID", empresasEnderecos.ConvertAll(endereco => (object)endereco.TableID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Enderecos_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoEnderecos_Temp", salvarArquivo, empresasEnderecosDict);
				excelHelper.GravarExcel(salvarArquivo, empresasEnderecosDict);
			}

            catch (Exception error)
            {
				throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

		public static Dictionary<string, string[]> ExcelRecebiveisToDictionary(string arquivoExcel)
		{
			var dataDictionary = new Dictionary<string, string[]>();
			var excelHelper = new ExcelHelper(arquivoExcel);

			try
			{
                foreach (var linha in excelHelper.linhas)
                {
                    string documento = "", recebivelID = "", consumidorID = "";

                    foreach (var celula in linha.Cells)
                    {
                        var celulaValor = celula.ToString().Trim();
                        var tituloColuna = excelHelper.cabecalhos[celula.Address.Column];

                        if (!tituloColuna.Contains("ExclusaoMotivo") && !tituloColuna.Contains("ID") && !tituloColuna.Contains("ConsumidorID"))
							throw new Exception($"Arquivo Excel \"{arquivoExcel}\" não contém as colunas ID, ConsumidorID e/ou ExclusaoMotivo");

						switch (tituloColuna)
                        {
                            case "ExclusaoMotivo":
                                documento = celulaValor;
								break;
							case "ID":
								recebivelID = celulaValor;
								break;
							case "ConsumidorID":
								consumidorID = celulaValor;
								break;
						}
                    }

                    if (!string.IsNullOrEmpty(documento) && !string.IsNullOrEmpty(recebivelID) && !string.IsNullOrEmpty(consumidorID))
					    dataDictionary.Add(documento, { recebivelID, consumidorID });
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

			return dataDictionary;
		}
	}
}
