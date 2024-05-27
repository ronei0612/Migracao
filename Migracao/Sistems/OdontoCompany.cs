using Migracao.Models;
using Migracao.Utils;
using NPOI.SS.UserModel;

namespace Migracao.Sistems
{
    internal class OdontoCompany
	{
        string arquivoExcelCidades = "Files\\EnderecosCidades.xlsx";
		string arquivoExcelNomesUTF8 = "Files\\NomesUTF8.xlsx";

		public void ImportarAgenda(string arquivoExcel, int estabelecimentoID, string arquivoExcelFuncionarios, int loginID)
		{
			var dataHoje = DateTime.Now;
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			ISheet sheet;
			try
			{
				IWorkbook workbook = excelHelper.LerExcel(arquivoExcelFuncionarios);
				sheet = workbook.GetSheetAt(0);
				excelHelper.InitializeDictionary(sheet);
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelFuncionarios}\": {ex.Message}");
			}

			var agendamentos = new List<Agendamento>();

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					string nomeCompleto = "", cpf = "", hora = "", data = "";
					bool faltou = false;
					string? outroSacadoNome = null, observacoes = null, documento = null;
					int recibo = 0, codigo = 0;
					int? consumidorID = null, fornecedorID = null, colaboradorID = null, funcionarioID = null, clienteID = null;
					decimal pagoValor = 0, valor = 0;
					byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
					DateTime dataConsulta = dataHoje;

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
									case "CNPJ_CPF":
										cpf = celulaValor.ToCPF();
										break;
									case "NOME":
										nomeCompleto = celulaValor;
										break;
									case "DATA":
										data = celulaValor;
										break;
									case "HORA":
										hora = celulaValor;
										break;
									case "OBS":
										observacoes = celulaValor;
										break;
									case "FALTOU":
										faltou = celulaValor == "S";
										break;
									case "RESPONSAVEL":
										valor = celulaValor.ToMoeda();
										break;
								}
							}
						}
					}

					dataConsulta = (data + " " + hora).ToData();

					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: cpf, codigo: codigo.ToString());
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

					agendamentos.Add(new Agendamento()
					{
						LoginID = loginID,
						EstabelecimentoID = estabelecimentoID,
						AtendeTipoID = 1,
						DataInicio = dataConsulta,
						DataTermino = dataConsulta.AddMinutes(30),
						ConsumidorID = (int)consumidorID,
						Titulo = observacoes,
						//DataCancelamento = ,
						//
						//AtendimentoValor = ,
						//SecretariaID = ,
						//FuncionarioID = ,
						//SalaID = ,
						DataInclusao = dataConsulta
					});
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "ConsumidorID", agendamentos.ConvertAll(agendamento => (object)agendamento.ConsumidorID).ToArray() }
				};

				var salvarArquivo = Tools.GerarNomeArquivo($"Recebiveis_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Recebiveis", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);

				MessageBox.Show("Sucesso!");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}


		public void ImportarPrecos(string arquivoExcel, int estabelecimentoID, string arquivoExcelGruposProcedimentos)
        {
			//CED001
			//CED002 Categoria
			var dataHoje = DateTime.Now;
			var indiceLinha = 0;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);

			var cabecalhos = new List<string> { "Especialidade", "PROCEDIMENTOS", "PREÇO", "TUSS" };			

			//var categorias = new List<string>();
			//var titulos = new List<string>();
			//var valores = new List<string>();
			//var tuss = new List<string>();

            //var celulas = new List<string>();
            List<List<string>> listaDados = new List<List<string>>();

			var gruposProcedimentosToDictionary = GruposProcedimentosToDictionary(arquivoExcelGruposProcedimentos);

			try
            {
                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;

                    string? titulo = null, tuss = "0";
                    decimal? valor = null;
                    int? grupo = null;
                    byte categoria = (byte)ProcedimentosCategoriasID.Outros;

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
                                    case "NOME":
                                        titulo = celulaValor.PrimeiraLetraMaiuscula();
                                        break;
                                    case "VRVENDA":
										valor = celulaValor.ArredondarValor();
                                        break;
                                    case "GRUPO":
										grupo = int.Parse(celulaValor);
                                        break;
                                }
                            }
                        }
                    }

                    if (grupo != null && !string.IsNullOrEmpty(titulo))
                    {
                        if (gruposProcedimentosToDictionary.ContainsKey((int)grupo))
                            categoria = (byte)gruposProcedimentosToDictionary[(int)grupo];

						listaDados.Add(new List<string> { categoria.ToString(), titulo, valor.ToString(), tuss });
					}
				}

				//var listaDados = new List<List<string>>()
				//{
				//	titulos,
				//	categorias,
				//	valores,
    //                tuss
				//};

				var salvarArquivo = Tools.GerarNomeArquivo($"Precos_{estabelecimentoID}_OdontoCompany_Migração");
				excelHelper.CreateExcelFile(salvarArquivo + ".xlsx", cabecalhos, listaDados);
				//Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarRecebiveis(string arquivoExcel, string arquivoExcelConsumidores, int estabelecimentoID, int respFinanceiroPessoaID, int loginID)
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

				MessageBox.Show("Sucesso!");
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

        public void ImportarRecebidos(string arquivoExcel, int estabelecimentoID, int respFinanceiroPessoaID, int loginID, string arquivoExcelRecebiveis, string arquivoExcelFormaPagamento = "")
        {
            //CRD013 Forma de Pagamento
            var dataHoje = DateTime.Now;
            var indiceLinha = 0;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            var excelRecebidosDict = ExcelRecebiveisToDictionary(arquivoExcelRecebiveis);
			//var excelFormaPagamentoDict = ExcelFormaPagamentoToDictionary(arquivoExcelFormaPagamento);

			var fluxoCaixas = new List<FluxoCaixa>();
			var recebiveis = new List<Recebivel>();

			try
            {
                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;

                    string documento = "";
                    string? pagamento = null;
                    int? tipoPagamento = null;
					string? observacao = null;
                    decimal pagoValor = 0;
                    byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
                    DateTime dataBaixa = DateTime.Now;

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
									//case "TIPO_DOC":
									//tipoPagamento = int.Parse(celulaValor);
									//break;
									case "NOME_GRUPO":
										pagamento = celulaValor;
										break;
								}
							}
						}
					}

                    //if (tipoPagamento != null && excelFormaPagamentoDict.ContainsKey((int)tipoPagamento))
                    //    formaPagamento = (byte)excelFormaPagamentoDict[(int)tipoPagamento];

                    if (!string.IsNullOrEmpty(documento) && excelRecebidosDict.ContainsKey(documento))
                    {
                        fluxoCaixas.Add(new FluxoCaixa()
                        {
                            RecebivelID = int.Parse(excelRecebidosDict[documento][0]),
                            ConsumidorID = int.Parse(excelRecebidosDict[documento][1]),
                            SituacaoID = 1,
                            PagoMulta = 0,
                            PagoJuros = 0,
							PagoDescontos = 0,
                            PagoDespesas = 0,
							TipoID = (byte)TransacaoTiposID.Recebimento,
                            Data = dataBaixa,
                            TransacaoID = (byte)TituloTransacoes.Liquidacao,
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

						recebiveis.Add(new Recebivel()
						{
                            ID = int.Parse(excelRecebidosDict[documento][0]),
                            DataBaixa = dataBaixa,
                            ValorDevido = 0
						});

					}
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


				var dadosRecebivel = new Dictionary<string, object[]>
				{
					{ "ID", recebiveis.ConvertAll(recebivel => (object)recebivel.ID).ToArray() },
					{ "DataBaixa", recebiveis.ConvertAll(recebivel => (object)recebivel.DataBaixa).ToArray() },
					{ "ValorDevido", recebiveis.ConvertAll(recebivel => (object)recebivel.ValorDevido).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Recebidos_{estabelecimentoID}_Update_OdontoCompany_Migração");
				sqlHelper.GerarSqlUpdate("Recebiveis", salvarArquivo, dadosRecebivel);

				MessageBox.Show("Sucesso!");
			}
            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha++, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }
		public void ImportarPessoas(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			if (Path.GetFileNameWithoutExtension(arquivoExcel).Contains("CED006"))
				ImportarPessoasDentistas(arquivoExcel, arquivoPessoasAtuais, estabelecimentoID, loginID);
			else if (Path.GetFileNameWithoutExtension(arquivoExcel).Contains("EMD101"))
				ImportarPessoasClientes(arquivoExcel, arquivoPessoasAtuais, estabelecimentoID, loginID);
		}

		public void ImportarPessoasClientes(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoPessoasAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoPessoasAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoPessoasAtuais}\": {ex.Message}");
				}


			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var pessoas = new List<Pessoa>();

				foreach (var linha in excelHelper.linhas)
				{
					bool cliente = false, fornecedor = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
					long telefonePrinc = 0, telefoneAltern = 0, telefoneComercial = 0, telefoneOutro = 0, celular = 0;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

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
										nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										apelido = nomeCompleto.GetPrimeirosCaracteres(20);
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
										logradouro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "BAIRRO":
										bairro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "NUM_ENDERECO":
										logradouroNum = celulaValor;
										break;
									case "CIDADE":
										cidade = celulaValor;
										break;
									case "ESTADO":
										estado = celulaValor;
										break;
									case "CEP":
										cep = celulaValor.ToNum();
										break;
									case "OBS1":
										observacao = celulaValor;
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
					//if (!string.IsNullOrEmpty(pessoaIDValue))
					//	pessoaID = int.Parse(pessoaIDValue);

					//var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: documento);
					//if (!string.IsNullOrEmpty(consumidorIDValue))
					//	consumidorID = int.Parse(consumidorIDValue);

					if (!fornecedor)
					{
						if ((!string.IsNullOrEmpty(nomeCompleto) && string.IsNullOrEmpty(documento))
							|| (!string.IsNullOrEmpty(documento) && documento.IsCPF()))
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

				salvarArquivo = Tools.GerarNomeArquivo($"PessoasClientes_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Pessoas", salvarArquivo, pessoasDict);
				excelHelper.GravarExcel(salvarArquivo, pessoasDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarEmpresas(string arquivoExcel, string arquivoEmpresasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoEmpresasAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoEmpresasAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoEmpresasAtuais}\": {ex.Message}");
				}


			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var empresas = new List<Empresa>();

				foreach (var linha in excelHelper.linhas)
				{
					bool cliente = false, fornecedor = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
					long telefonePrinc = 0, telefoneAltern = 0, telefoneComercial = 0, telefoneOutro = 0, celular = 0;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

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
										nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										apelido = celulaValor.GetLetras().GetPrimeiroNome().PrimeiraLetraMaiuscula();
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
										logradouro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "BAIRRO":
										bairro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "NUM_ENDERECO":
										logradouroNum = celulaValor;
										break;
									case "CIDADE":
										cidade = celulaValor;
										break;
									case "ESTADO":
										estado = celulaValor;
										break;
									case "CEP":
										cep = celulaValor.ToNum();
										break;
									case "OBS1":
										observacao = celulaValor;
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

					if (string.IsNullOrEmpty(arquivoEmpresasAtuais))
						if (fornecedor && documento.IsCNPJ_CGC())
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
						else
						{
							pessoaID = indiceLinha;
							var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto, cpf: documento);
							if (!string.IsNullOrEmpty(pessoaIDValue))
								pessoaID = int.Parse(pessoaIDValue);

							var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: documento);
							if (!string.IsNullOrEmpty(consumidorIDValue))
								consumidorID = int.Parse(consumidorIDValue);

							if (fornecedor && documento.IsCNPJ_CGC())
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

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";

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
				sqlHelper.GerarSqlInsert("Empresas", salvarArquivo, empresasDict);
				excelHelper.GravarExcel(salvarArquivo, empresasDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarPessoasDentistas(string arquivoExcel, string arquivoPessoasAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoPessoasAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoPessoasAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoPessoasAtuais}\": {ex.Message}");
				}

			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var pessoas = new List<Pessoa>();
				var funcionarios = new List<Funcionario>();
				var enderecos = new List<Endereco>();

				foreach (var linha in excelHelper.linhas)
				{
					bool ativo = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					int? codigo = null;
					long telefonePrinc = 0;
					string? nomeCompleto = null, departamento = null, cro = null, observacao = null, email = null;
					string apelido = "", documento = "";

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
									case "CODIGO":
										codigo = int.Parse(celulaValor);
										break;
									case "NOME":
										nomeCompleto = celulaValor.GetLetras().PrimeiraLetraMaiuscula();
										apelido = nomeCompleto.GetPrimeirosCaracteres(20);
										break;
									case "DEPARTAMENTO":
										departamento = celulaValor;
										break;
									case "OBS":
										observacao = celulaValor;
										break;
									case "ATIVO":
										ativo = celulaValor == "S" ? true : false;
										break;
									case "NOME_COMPLETO":
										//nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										break;
									case "EMAIL":
										email = celulaValor.ToEmail();
										break;
									case "TELEFONE":
										telefonePrinc = celulaValor.ToFone();
										break;
									case "CRO":
										cro = celulaValor;
										break;
									case "MODIFICADO":
										dataCadastro = celulaValor.ToData();
										break;
								}
							}
						}
					}

					if (!string.IsNullOrWhiteSpace(apelido))
					{
						if (string.IsNullOrWhiteSpace(nomeCompleto))
							nomeCompleto = apelido;

						var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto);

						if (string.IsNullOrEmpty(arquivoPessoasAtuais) 
							|| (!string.IsNullOrEmpty(arquivoPessoasAtuais) && string.IsNullOrEmpty(pessoaIDValue)))
						{
							pessoas.Add(new Pessoa()
							{
								ID = indiceLinha,
								NomeCompleto = nomeCompleto,
								Apelido = apelido,
								CPF = "",
								DataInclusao = dataCadastro,
								Email = email,
								NascimentoData = dataNascimento,
								ProfissaoOutra = departamento,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								Guid = new Guid(),
								FoneticaApelido = apelido.Fonetizar(),
								FoneticaNomeCompleto = nomeCompleto.Fonetizar(),
								Sexo = true,
								ConselhoCodigo = cro
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
				sqlHelper.GerarSqlInsert("Pessoas", salvarArquivo, pessoasDict);
				excelHelper.GravarExcel(salvarArquivo, pessoasDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarFuncionarios(string arquivoExcel, string arquivoFuncionariosAtuais, int estabelecimentoID, int loginID)
        {
			var indiceLinha = 1;
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

			if (!string.IsNullOrEmpty(arquivoFuncionariosAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoFuncionariosAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoFuncionariosAtuais}\": {ex.Message}");
				}


			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var funcionarios = new List<Funcionario>();
				var pessoaFones = new List<PessoaFone>();

				foreach (var linha in excelHelper.linhas)
				{
					bool ativo = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					int? codigo = null;
					long telefonePrinc = 0;
					string? nomeCompleto = null, departamento = null, cro = null, observacao = null, email = null;
					string apelido = "";

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
									case "CODIGO":
										codigo = int.Parse(celulaValor);
										break;
									case "NOME":
										nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										apelido = nomeCompleto.GetPrimeirosCaracteres(20);
										break;
									case "DEPARTAMENTO":
										departamento = celulaValor;
										break;
									case "OBS":
										observacao = celulaValor;
										break;
									case "ATIVO":
										ativo = celulaValor == "S" ? true : false;
										break;
									case "NOME_COMPLETO":
										//nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										break;
									case "EMAIL":
										email = celulaValor.ToEmail();
										break;
									case "TELEFONE":
										telefonePrinc = celulaValor.ToFone();
										break;
									case "CRO":
										cro = celulaValor;
										break;
									case "MODIFICADO":
										dataCadastro = celulaValor.ToData();
										break;
								}
							}
						}
					}

					pessoaID = indiceLinha;

					if (!string.IsNullOrWhiteSpace(nomeCompleto))
					{
						if (string.IsNullOrWhiteSpace(apelido))
							apelido = nomeCompleto.GetPrimeirosCaracteres(20);

						var pessoaIDValue = excelHelper.GetPessoaID(nomeCompleto: nomeCompleto);
						var funcionarioIDValue = excelHelper.GetFuncionarioID(nomeCompleto: nomeCompleto);

						if (string.IsNullOrEmpty(arquivoFuncionariosAtuais) 
							|| (!string.IsNullOrEmpty(arquivoFuncionariosAtuais) && !string.IsNullOrEmpty(pessoaIDValue) && string.IsNullOrEmpty(funcionarioIDValue)))
						{
							if (!string.IsNullOrEmpty(pessoaIDValue))
								pessoaID = int.Parse(pessoaIDValue);

							funcionarios.Add(new Funcionario()
							{
								Ativo = ativo,
								DataInclusao = dataCadastro,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								PessoaID = pessoaID,
								Observacoes = observacao,
								CargoID = (byte)CargosID.Dentista,
								PermissaoCoordenacao = false,
								PermissaoGrupoID = 0,
								PermissaoModuloAdmin = false,
								PermissaoModuloAtendimentos = true,
								PermissaoModuloPacientes = true,
								PermissaoModuloEstoque = true,
								PermissaoModuloFinanceiro = false
							});

							if (telefonePrinc > 0)
								if (string.IsNullOrEmpty(arquivoFuncionariosAtuais)
									|| (!string.IsNullOrEmpty(arquivoFuncionariosAtuais) && !excelHelper.PessoaFoneExists(nomeCompleto, telefonePrinc.ToString())))
								{
									pessoaFones.Add(new PessoaFone()
									{
										PessoaID = pessoaID,
										FoneTipoID = (short)FoneTipos.Principal,
										Telefone = telefonePrinc,
										DataInclusao = dataCadastro,
										LoginID = loginID
									});
								}
						}
					}

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";

				var funcionariosDict = new Dictionary<string, object[]>
				{
					{ "CargoID", funcionarios.ConvertAll(funcionario => (object)funcionario.CargoID).ToArray() },
					{ "Ativo", funcionarios.ConvertAll(funcionario => (object)funcionario.Ativo).ToArray() },
					{ "DataInclusao", funcionarios.ConvertAll(funcionario => (object)funcionario.DataInclusao).ToArray() },
					{ "EstabelecimentoID", funcionarios.ConvertAll(funcionario => (object)funcionario.EstabelecimentoID).ToArray() },
					{ "LoginID", funcionarios.ConvertAll(funcionario => (object)funcionario.LoginID).ToArray() },
					{ "PessoaID", funcionarios.ConvertAll(funcionario => (object)funcionario.PessoaID).ToArray() },
					{ "Observacoes", funcionarios.ConvertAll(funcionario => (object)funcionario.Observacoes).ToArray() },
					{ "PermissaoCoordenacao", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoCoordenacao).ToArray() },
					{ "PermissaoGrupoID", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoGrupoID).ToArray() },
					{ "PermissaoModuloAdmin", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloAdmin).ToArray() },
					{ "PermissaoModuloAtendimentos", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloAtendimentos).ToArray() },
					{ "PermissaoModuloPacientes", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloPacientes).ToArray() },
					{ "PermissaoModuloEstoque", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloEstoque).ToArray() },
					{ "PermissaoModuloFinanceiro", funcionarios.ConvertAll(funcionario => (object)funcionario.PermissaoModuloFinanceiro).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Funcionarios_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("Funcionarios", salvarArquivo, funcionariosDict);
				excelHelper.GravarExcel(salvarArquivo, funcionariosDict);


				var pessoaFonesDict = new Dictionary<string, object[]>
				{
					{ "PessoaID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.PessoaID).ToArray() },
					{ "FoneTipoID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.FoneTipoID).ToArray() },
					{ "Telefone", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.Telefone).ToArray() },
					{ "DataInclusao", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.DataInclusao).ToArray() },
					{ "LoginID", pessoaFones.ConvertAll(pessoaFone => (object)pessoaFone.LoginID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"PessoaFones_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("PessoaFones", salvarArquivo, pessoaFonesDict);
				excelHelper.GravarExcel(salvarArquivo, pessoaFonesDict);

				MessageBox.Show("Limpar o Redis" + Environment.NewLine + "redis-cli.exe -h 127.0.0.1 -n 7 del Equipe:" + estabelecimentoID.ToString("D6") + "-Funcionarios", "Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarFornecedores(string arquivoExcel, string arquivoFornecedoresAtuais, int estabelecimentoID, int loginID)
		{
			var indiceLinha = 1;
			var consumidorID = 1;
			var pessoaID = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			if (!string.IsNullOrEmpty(arquivoFornecedoresAtuais))
				try
				{
					var workbook = excelHelper.LerExcel(arquivoFornecedoresAtuais);
					var sheet = workbook.GetSheetAt(0);
					excelHelper.InitializeDictionary(sheet);
				}
				catch (Exception ex)
				{
					throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoFornecedoresAtuais}\": {ex.Message}");
				}

			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var fornecedores = new List<Fornecedor>();
				var empresas = new List<Empresa>();
				var enderecos = new List<Endereco>();
				var fornecedorFones = new List<FornecedorFone>();

				foreach (var linha in excelHelper.linhas)
				{
					bool cliente = false, fornecedor = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
					long telefonePrinc = 0, telefoneAltern = 0, telefoneComercial = 0, telefoneOutro = 0, celular = 0;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

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
										nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
										apelido = celulaValor.GetLetras().GetPrimeiroNome().PrimeiraLetraMaiuscula();
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
										logradouro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "BAIRRO":
										bairro = celulaValor.PrimeiraLetraMaiuscula();
										break;
									case "NUM_ENDERECO":
										logradouroNum = celulaValor;
										break;
									case "CIDADE":
										cidade = celulaValor;
										break;
									case "ESTADO":
										estado = celulaValor;
										break;
									case "CEP":
										cep = celulaValor.ToNum();
										break;
									case "OBS1":
										observacao = celulaValor;
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

					var fornecedorIDValue = excelHelper.GetFornecedorID(nomeCompleto: nomeCompleto, cpf: documento);

					if (!string.IsNullOrWhiteSpace(nomeCompleto))
					{
						if (fornecedor && string.IsNullOrEmpty(fornecedorIDValue))
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
								enderecos.Add(new Endereco()
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
					}

					indiceLinha++;
				}

				indiceLinha = 0;
				var salvarArquivo = "";


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


				var enderecosDict = new Dictionary<string, object[]>
				{
					{ "Ativo", enderecos.ConvertAll(endereco => (object)endereco.Ativo).ToArray() },
					{ "Cep", enderecos.ConvertAll(endereco => (object)endereco.Cep).ToArray() },
					{ "CidadeID", enderecos.ConvertAll(endereco => (object)endereco.CidadeID).ToArray() },
					{ "DataInclusao", enderecos.ConvertAll(endereco => (object)endereco.DataInclusao).ToArray() },
					{ "EnderecoTipoID", enderecos.ConvertAll(endereco => (object)endereco.EnderecoTipoID).ToArray() },
					{ "Logradouro", enderecos.ConvertAll(endereco => (object)endereco.Logradouro).ToArray() },
					{ "LogradouroNum", enderecos.ConvertAll(endereco => (object)endereco.LogradouroNum).ToArray() },
					{ "Bairro", enderecos.ConvertAll(endereco => (object)endereco.Bairro).ToArray() },
					{ "LogradouroTipoID", enderecos.ConvertAll(endereco => (object)endereco.LogradouroTipoID).ToArray() },
					{ "Complemento", enderecos.ConvertAll(endereco => (object)endereco.Complemento).ToArray() },
					{ "ParentID", enderecos.ConvertAll(endereco => (object)endereco.ParentID).ToArray() },
					{ "TableID", enderecos.ConvertAll(endereco => (object)endereco.TableID).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Enderecos_{estabelecimentoID}_OdontoCompany_Migração");
				sqlHelper.GerarSqlInsert("_MigracaoEnderecos_Temp", salvarArquivo, enderecosDict);
				excelHelper.GravarExcel(salvarArquivo, enderecosDict);

				MessageBox.Show("Sucesso!");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
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

       //     if (File.Exists(arquivoExcelNomesUTF8))
			    //try
			    //{
				   // var workbookCidades = excelHelper.LerExcel(arquivoExcelNomesUTF8);
				   // var sheetCidades = workbookCidades.GetSheetAt(0);
				   // excelHelper.InitializeDictionaryNomesUTF8(sheetCidades);
			    //}
			    //catch (Exception ex)
			    //{
				   // throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelNomesUTF8}\": {ex.Message}");
			    //}

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
                var consumidoresEnderecos = new List<ConsumidorEndereco>();
                var pessoaFones = new List<PessoaFone>();
				var fornecedores = new List<Fornecedor>();
				var empresas = new List<Empresa>();
                var enderecos = new List<Endereco>();
				var fornecedorFones = new List<FornecedorFone>();

				foreach (var linha in excelHelper.linhas)
                {
                    bool cliente = false, fornecedor = false;
                    DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
                    int cep = 0;
					byte? estadoCivil = null;
					bool sexo = true;
                    long telefonePrinc = 0, telefoneAltern = 0, telefoneComercial = 0, telefoneOutro = 0, celular = 0;
					string? nomeCompleto = null, documento = null, rg = null, email = null, apelido = null, nascimentoLocal = null, profissaoOutra = null, logradouro = "",
						 complemento = null, bairro = null, logradouroNum = null, numcadastro = null, cidade = "", estado = null, observacao = null;

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
                                        nomeCompleto = celulaValor.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();
                                        apelido = celulaValor.GetLetras().GetPrimeiroNome().PrimeiraLetraMaiuscula();
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
                                        logradouro = celulaValor.PrimeiraLetraMaiuscula();
                                        break;
                                    case "BAIRRO":
                                        bairro = celulaValor.PrimeiraLetraMaiuscula();
                                        break;
                                    case "NUM_ENDERECO":
                                        logradouroNum = celulaValor;
                                        break;
                                    case "CIDADE":
                                        cidade = celulaValor;
                                        break;
                                    case "ESTADO":
                                        estado = celulaValor;
                                        break;
                                    case "CEP":
                                        cep = celulaValor.ToNum();
                                        break;
                                    case "OBS1":
                                        observacao = celulaValor;
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
					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, cpf: documento);

					if (!string.IsNullOrWhiteSpace(nomeCompleto))
                    {
                        if (!fornecedor)
                        {
							if (string.IsNullOrEmpty(pessoaIDValue) == false)
							{
								if (string.IsNullOrEmpty(consumidorIDValue))
								{
									pessoaID = int.Parse(pessoaIDValue);

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
								{
									var cidadeID = excelHelper.GetCidadeID(cidade, estado);

									if (cidadeID > 0)
										consumidoresEnderecos.Add(new ConsumidorEndereco()
										{
											Ativo = true,
											ConsumidorID = int.Parse(consumidorIDValue),
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
                    }

					indiceLinha++;
				}

                indiceLinha = 0;
                var salvarArquivo = "";


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
				sqlHelper.GerarSqlInsert("Consumidores", salvarArquivo, consumidoresDict);
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
				sqlHelper.GerarSqlInsert("ConsumidorEnderecos", salvarArquivo, consumidoresEnderecosDict);
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
				sqlHelper.GerarSqlInsert("PessoaFones", salvarArquivo, pessoaFonesDict);
				excelHelper.GravarExcel(salvarArquivo, pessoaFonesDict);

				MessageBox.Show("Atualize a Base de Pesquisa em: Dados da Clínica => Licença de Uso", "Sucesso!");
			}

            catch (Exception error)
            {
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha + 2, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

		public static Dictionary<string, string[]> ExcelRecebiveisToDictionary(string arquivoExcel)
		{
			var dataDictionary = new Dictionary<string, string[]>();
			var excelHelper = new ExcelHelper(arquivoExcel);

			if (!excelHelper.cabecalhos.Contains("ExclusaoMotivo") && !excelHelper.cabecalhos.Contains("ID") && !excelHelper.cabecalhos.Contains("ConsumidorID"))
				throw new Exception($"Arquivo Excel \"{arquivoExcel}\" não contém as colunas ID, ConsumidorID e/ou ExclusaoMotivo");

			try
			{
                foreach (var linha in excelHelper.linhas)
                {
                    string documento = "", recebivelID = "", consumidorID = "";

                    foreach (var celula in linha.Cells)
                    {
                        var celulaValor = celula.ToString().Trim();
                        var tituloColuna = excelHelper.cabecalhos[celula.Address.Column];							

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

                    if (!string.IsNullOrEmpty(documento) && !string.IsNullOrEmpty(recebivelID) && !string.IsNullOrEmpty(consumidorID) && !dataDictionary.ContainsKey(documento))
					    dataDictionary.Add(documento, new string[] { recebivelID, consumidorID });
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

			return dataDictionary;
		}

		public static Dictionary<int, TitulosEspeciesID> ExcelFormaPagamentoToDictionary(string arquivoExcel)
        {
			var dataDictionary = new Dictionary<int, TitulosEspeciesID>();
			var excelHelper = new ExcelHelper(arquivoExcel);

			if (!excelHelper.cabecalhos.Contains("CODIGO") && !excelHelper.cabecalhos.Contains("NOME"))
				throw new Exception($"Arquivo Excel \"{arquivoExcel}\" não contém as colunas CODIGO e/ou NOME");

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					int? codigo = null;
					string formaPagamento = "";
					TitulosEspeciesID titulosEspeciesID = TitulosEspeciesID.DepositoEmConta;

					foreach (var celula in linha.Cells)
					{
						var celulaValor = celula.ToString().Trim();
						var tituloColuna = excelHelper.cabecalhos[celula.Address.Column];

						switch (tituloColuna)
						{
							case "CODIGO":
								codigo = int.Parse(celulaValor);
								break;
							case "NOME":
								formaPagamento = celulaValor;
								break;
						}
					}

					if (codigo != null && !string.IsNullOrEmpty(formaPagamento))
                    {
                        if (formaPagamento.ToLower().Contains("dinheiro"))
							titulosEspeciesID = TitulosEspeciesID.Dinheiro;
                        else if (formaPagamento.ToLower().Contains("cheque"))
							titulosEspeciesID = TitulosEspeciesID.Cheque;
						else if (formaPagamento.ToLower().Contains("master card") || formaPagamento.ToLower().Contains("visa") || formaPagamento.ToLower().Contains("elo")
                            || formaPagamento.ToLower().Contains("american express") || formaPagamento.ToLower().Contains("hipercard"))
							titulosEspeciesID = TitulosEspeciesID.CartaoCredito;
						else if (formaPagamento.ToLower().Contains("debito"))
							titulosEspeciesID = TitulosEspeciesID.CartaoDebito;
					}

					dataDictionary.Add((int)codigo, titulosEspeciesID);
				}
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

			return dataDictionary;
		}

		public static Dictionary<int, ProcedimentosCategoriasID> GruposProcedimentosToDictionary(string arquivoExcel)
        {
			var dataDictionary = new Dictionary<int, ProcedimentosCategoriasID>();
			var excelHelper = new ExcelHelper(arquivoExcel);

			if (!excelHelper.cabecalhos.Contains("CODIGO") && !excelHelper.cabecalhos.Contains("NOME"))
				throw new Exception($"Arquivo Excel \"{arquivoExcel}\" não contém as colunas CODIGO e/ou NOME");

			var grupoCategoriaDict = new Dictionary<string, ProcedimentosCategoriasID>()
            {
				{ "CIRURGIA", ProcedimentosCategoriasID.Cirurgia },
                { "ENDODONTIA", ProcedimentosCategoriasID.Endodontia },
                { "PERIODONTIA", ProcedimentosCategoriasID.Periodontia },
                { "PROTESE", ProcedimentosCategoriasID.Prótese },
                { "CLINICO", ProcedimentosCategoriasID.Outros }, // Assumindo que "CLINICO" seja uma categoria genérica
                { "MANUTENCAO", ProcedimentosCategoriasID.Prevenção }, // Assumindo que "MANUTENCAO" seja sinônimo de "PREVENÇÃO"
                { "ORTODONTIA", ProcedimentosCategoriasID.Ortodontia },
                { "AMIL", ProcedimentosCategoriasID.Outros }, // Assumindo que "AMIL" seja um tipo de convênio
                { "PREVENÇÃO ODC", ProcedimentosCategoriasID.Prevenção },
                { "INSTITUTO ODONTOCOMPANY", ProcedimentosCategoriasID.Outros }, // Assumindo que "INSTITUTO ODONTOCOMPANY" seja um tipo de convênio
                { "UNIMED", ProcedimentosCategoriasID.Outros }, // Assumindo que "UNIMED" seja um tipo de convênio
                { "PRIMAVIDA", ProcedimentosCategoriasID.Outros }, // Assumindo que "PRIMAVIDA" seja um tipo de convênio
                { "HARMONIZAÇÃO OROFACIAL", ProcedimentosCategoriasID.Orofacial },
                { "ODONTOMAXI", ProcedimentosCategoriasID.Outros }, // Assumindo que "ODONTOMAXI" seja um tipo de convênio
                { "RODRIGUES LEIRA", ProcedimentosCategoriasID.Outros }, // Assumindo que "RODRIGUES LEIRA" seja um tipo de convênio
                { "PORTO SEGURO", ProcedimentosCategoriasID.Outros }, // Assumindo que "PORTO SEGURO" seja um tipo de convênio
                { "INPAO", ProcedimentosCategoriasID.Outros }, // Assumindo que "INPAO" seja um tipo de convênio
                { "DENTAL byteEGRAL", ProcedimentosCategoriasID.Outros }, // Assumindo que "DENTAL byteEGRAL" seja um tipo de convênio
                { "AESP", ProcedimentosCategoriasID.Outros }, // Assumindo que "AESP" seja um tipo de convênio
                { "PROASA", ProcedimentosCategoriasID.Outros }, // Assumindo que "PROASA" seja um tipo de convênio
                { "IDEAL ODONTO", ProcedimentosCategoriasID.Outros }, // Assumindo que "IDEAL ODONTO" seja um tipo de convênio
                { "ODONTOART", ProcedimentosCategoriasID.Outros }, // Assumindo que "ODONTOART" seja um tipo de convênio
                { "BRAZIL DENTAL", ProcedimentosCategoriasID.Outros }, // Assumindo que "BRAZIL DENTAL" seja um tipo de convênio
			};

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
                    int? codigo = null;
                    string procedimentoNome = "";

					foreach (var celula in linha.Cells)
					{
						var celulaValor = celula.ToString().Trim();
						var tituloColuna = excelHelper.cabecalhos[celula.Address.Column];

						switch (tituloColuna)
						{
							case "CODIGO":
								codigo = int.Parse(celulaValor);
								break;
							case "NOME":
								procedimentoNome = celulaValor;
								break;
						}
					}

					if (codigo != null && !string.IsNullOrEmpty(procedimentoNome))
						if (grupoCategoriaDict.ContainsKey(procedimentoNome))
						    dataDictionary.Add((int)codigo, grupoCategoriaDict[procedimentoNome]);
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
