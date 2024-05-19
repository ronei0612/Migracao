using Migração.Models;
using Migração.Utils;
using NPOI.SS.UserModel;
using System.Globalization;

namespace Migração
{
    internal class OdontoCompany
	{
		public void ImportarRecebidos(string arquivoExcel, string arquivoExcelConsumidores, string estabelecimentoID, string respFinanceiroPessoaID, string salvarArquivo)
		{
			var dataHoje = DateTime.Now;
			var indiceLinha = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			ISheet sheetConsumidores;
			try
			{
				IWorkbook workbookConsumidores = excelHelper.LerExcel(arquivoExcelConsumidores);
				sheetConsumidores = workbookConsumidores.GetSheetAt(0);
				excelHelper.InitializeDictionaryConsumidor(sheetConsumidores);
			}
			catch (Exception ex)
			{
				throw new Exception("Erro ao ler o arquivo Excel: " + ex.Message);
			}

			var fluxoCaixas = new List<FluxoCaixa>();

			try
			{
				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					string nomeCompleto = "", outroSacadoNome = "";
					int controle = 0, recibo = 0, codigo = 0, loginID = 1;
					int? consumidorID = 0;
					decimal pagoValor = 0;
					byte titulosEspecies = 0;
					var dataPagamento = dataHoje;
					var nascimentoData = dataHoje;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim();
							tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "paciente":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "numero_registro":
										codigo = int.Parse(celulaValor);
										break;
									case "data_pagamento":
										dataPagamento = celulaValor.ToData();
										break;
									case "forma_pagamento":
										celulaValor.ToTipoPagamento();
										break;
									case "valor":
										pagoValor = celulaValor.ToMoeda();
										break;
								}
							}
						}
					}

					var consumidorIDValue = excelHelper.GetConsumidorID(nomeCompleto: nomeCompleto, codigo: codigo.ToString());
					if (!string.IsNullOrEmpty(consumidorIDValue))
					{
						consumidorID = int.Parse(consumidorIDValue);
						outroSacadoNome = "null";

						fluxoCaixas.Add(new FluxoCaixa()
						{
							ConsumidorID = consumidorID,
							SituacaoID = 1,
							PagoMulta = 0,
							PagoJuros = 0,
							TipoID = (byte)TransacaoTiposID.Recebimento,
							Data = dataPagamento,
							TransacaoID = (byte)TituloTransacoes.PagamentoAvulso,
							EspecieID = titulosEspecies,
							DataBaseCalculo = dataPagamento,
							DevidoValor = pagoValor,
							PagoValor = pagoValor,
							EstabelecimentoID = int.Parse(estabelecimentoID),
							LoginID = 1,
							DataInclusao = dataPagamento,
							FinanceiroID = int.Parse(respFinanceiroPessoaID)
						});
					}
					else
					{
						consumidorID = null;
						outroSacadoNome = nomeCompleto.GetPrimeirosCaracteres(50);

						fluxoCaixas.Add(new FluxoCaixa()
						{
							OutroSacadoNome = nomeCompleto.GetPrimeirosCaracteres(50),
							SituacaoID = 1,
							PagoMulta = 0,
							PagoJuros = 0,
							TipoID = (byte)TransacaoTiposID.Recebimento,
							Data = dataPagamento,
							TransacaoID = (byte)TituloTransacoes.PagamentoAvulso,
							EspecieID = titulosEspecies,
							DataBaseCalculo = dataPagamento,
							DevidoValor = pagoValor,
							PagoValor = pagoValor,
							EstabelecimentoID = int.Parse(estabelecimentoID),
							LoginID = 1,
							DataInclusao = dataPagamento,
							FinanceiroID = int.Parse(respFinanceiroPessoaID)
						});
					}
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "ConsumidorID", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.ConsumidorID).ToArray() },
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
					{ "OutroSacadoNome", fluxoCaixas.ConvertAll(fluxoCaixa => (object)fluxoCaixa.OutroSacadoNome).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo(salvarArquivo);
				sqlHelper.GerarSqlInsert("_MigracaoFluxoCaixa_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
			}
			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

        public void ImportarFornecedores(string arquivoExcel, string arquivoExcelCidades, string estabelecimentoID, string salvarArquivo)
        {
            var dataHoje = DateTime.Now;
            var indiceLinha = 1;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

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
                            celulaValor = celula.ToString().Trim();
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
							EstabelecimentoID = int.Parse(estabelecimentoID),
							NomeFantasia = nomeCompleto,
							LoginID = 1							
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

				salvarArquivo = Tools.GerarNomeArquivo(salvarArquivo);
				sqlHelper.GerarSqlInsert("_MigracaoConsumidores_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
			}

            catch (Exception error)
            {
				throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
        }

        public void ImportarPacientes(string arquivoExcel, string arquivoExcelCidades, string estabelecimentoID, string salvarArquivo)
		{
			var indiceLinha = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

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
				var consumidores = new List<Consumidor>();
				var pessoas = new List<Pessoa>();

				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					bool cliente = false, fornecedor = false;
					DateTime dataNascimento = dataHoje, dataCadastro = dataHoje;
					int numcadastro;
					string nomeCompleto = "", cpf = "", rg = "", email = "", apelido = "";
					bool sexo = true;
					long telefonePrinc, telefoneAltern, telefoneComercial, telefoneOutro, celular;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim();
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
									case "ENDERECO":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "BAIRRO":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "NUM_ENDERECO":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "CIDADE":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "ESTADO":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "CEP":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "OBS1":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
									case "NUM_CONVENIO":
										nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
										break;
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

					if (cliente)
						pessoas.Add(new Pessoa()
						{
							NomeCompleto = nomeCompleto,
							Apelido = apelido,
							CPF = cpf,
							DataInclusao = dataCadastro,
							Email = email,
							RG = rg,
							Sexo = sexo
						});
				}

				indiceLinha = 0;

				var dados = new Dictionary<string, object[]>
				{
					{ "NomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.NomeCompleto).ToArray() },
					{ "Apelido", pessoas.ConvertAll(pessoa => (object)pessoa.Apelido).ToArray() },
					{ "CPF", pessoas.ConvertAll(pessoa => (object)pessoa.CPF).ToArray() },
					{ "DataInclusao", pessoas.ConvertAll(pessoa => (object)pessoa.DataInclusao).ToArray() },
					{ "Email", pessoas.ConvertAll(pessoa => (object)pessoa.Email).ToArray() },
					{ "RG", pessoas.ConvertAll(pessoa => (object)pessoa.RG).ToArray() },
					{ "Sexo", pessoas.ConvertAll(pessoa => (object)pessoa.Sexo).ToArray() }
				};

				salvarArquivo = Tools.GerarNomeArquivo(salvarArquivo);
				sqlHelper.GerarSqlInsert("_MigracaoConsumidores_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(error.Message, indiceLinha, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}
	}
}
