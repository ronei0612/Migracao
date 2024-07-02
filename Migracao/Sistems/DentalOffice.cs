using Migracao.Models;
using Migracao.Models.Context;
using Migracao.Models.Interfaces;
using Migracao.Utils;
using NPOI.SS.UserModel;

namespace Migracao.Sistems
{
    internal class DentalOffice : IDataBaseMigracao
    {
        MySqlContext context;

        public DentalOffice(string dataBaseName, string pathDB = null, string pathDBContratos = null)
        {
            if (!string.IsNullOrEmpty(dataBaseName))
                context = new MySqlContext(dataBaseName);

            var orcamentos = context.Orcamento.ToList();
        }

        public void ImportarPagos(string arquivoExcel, string arquivoExcelFuncionarios, int estabelecimentoID, int responsavelPessoaID, int loginID)
		{
			var indiceLinha = 1;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "", nome = "";
			DateTime dataHoje = DateTime.Now;
			var excelHelper = new ExcelHelper(arquivoExcel);
			var sqlHelper = new SqlHelper();

			string clinicaNome = "Viotto Odontologia";

			try
			{
				var workbookFuncionarios = excelHelper.LerExcel(arquivoExcelFuncionarios);
				var sheetFuncionarios = workbookFuncionarios.GetSheetAt(0);
				excelHelper.InitializeDictionary(sheetFuncionarios);
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcelFuncionarios}\": {ex.Message}");
			}

			try
			{
				var linhasCount = excelHelper.linhas.Count;
				var exigiveis = new List<Exigiveis>();
				var fluxosCaixa = new List<FluxoCaixa>();

				foreach (var linha in excelHelper.linhas)
				{
					indiceLinha++;

					bool clinica = false;
					string dentista = "", categoria = "", codigo = "";
					byte formaPagamento = (byte)TitulosEspeciesID.DepositoEmConta;
					decimal valor = 0, pagoValor = 0;
					DateTime dataVencimento = dataHoje, dataPagamento = dataHoje;

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
									case "clinica":
										if (celulaValor == clinicaNome)
											clinica = true;
										break;
									case "cir_dentista":
										dentista = celulaValor;
										break;
									case "fornecedor":
									case "cpf":
									case "cnpj":
										break;
									case "categoria":
										categoria = celulaValor;
										break;
									case "codigo":
										codigo = celulaValor;
										break;
									case "data_vencimento":
										dataVencimento = celulaValor.ToData();
										break;
									case "valor_pago":
										pagoValor = celulaValor.ToMoeda();
										break;
									case "data_pagamento":
										dataPagamento = celulaValor.ToData();
										break;
									case "forma_pagamento":
										formaPagamento = celulaValor.ToTipoPagamento();
										break;
									case "valor":
										valor = celulaValor.ToMoeda();
										break;
								}
							}
						}
					}

					if (clinica)// && pagoValor > 0)
					{
                        if (pagoValor <= 0)
                        {
							pagoValor = valor;
                            dataPagamento = dataVencimento;
						}

						if (!string.IsNullOrEmpty(dentista))
                        {
							var funcionarioID = excelHelper.GetFuncionarioID(nomeCompleto: dentista);
							if (!string.IsNullOrEmpty(funcionarioID))
								fluxosCaixa.Add(new FluxoCaixa()
								{
									Data = dataPagamento,
									DataBaseCalculo = dataPagamento,
									DataInclusao = dataPagamento,
									DevidoValor = valor,
									EspecieID = formaPagamento,
									FinanceiroID = responsavelPessoaID,
									PagoValor = pagoValor > 0 ? pagoValor : valor,
									TipoID = (byte)TransacaoTiposID.Pagamento,
									TransacaoID = (byte)TituloTransacoes.Liquidacao,
									EstabelecimentoID = estabelecimentoID,
									LoginID = loginID,
									PlanoContasID = 55,
									OutroCedenteNome = null,
                                    ColaboradorID = int.Parse(funcionarioID),
									ConsumidorID = null,
									FornecedorID = null
								});

                            else
								fluxosCaixa.Add(new FluxoCaixa()
								{
									Data = dataPagamento,
									DataBaseCalculo = dataPagamento,
									DataInclusao = dataPagamento,
									DevidoValor = valor,
									EspecieID = formaPagamento,
									FinanceiroID = responsavelPessoaID,
									PagoValor = pagoValor > 0 ? pagoValor : valor,
									TipoID = (byte)TransacaoTiposID.Pagamento,
									TransacaoID = (byte)TituloTransacoes.Liquidacao,
									EstabelecimentoID = estabelecimentoID,
									LoginID = loginID,
									PlanoContasID = 55,
									OutroCedenteNome = dentista,
									ColaboradorID = null,
									ConsumidorID = null,
									FornecedorID = null
								});
						}

                        else
                        {
							fluxosCaixa.Add(new FluxoCaixa()
							{
								Data = dataPagamento,
								DataBaseCalculo = dataPagamento,
								DataInclusao = dataPagamento,
								DevidoValor = valor,
								EspecieID = formaPagamento,
								FinanceiroID = responsavelPessoaID,
								PagoValor = pagoValor > 0 ? pagoValor : valor,
								TipoID = (byte)TransacaoTiposID.Pagamento,
								TransacaoID = (byte)TituloTransacoes.Liquidacao,
								EstabelecimentoID = estabelecimentoID,
								LoginID = loginID,
								PlanoContasID = 55,
								OutroCedenteNome = "Outro",
								ColaboradorID = null,
								ConsumidorID = null,
								FornecedorID = null
							});
						}
					}
				}

				indiceLinha = 0;
				var salvarArquivo = "";

				var fluxosCaixaDict = new Dictionary<string, object[]>
				{
                    { "TipoID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TipoID).ToArray() },
					{ "Data", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.Data).ToArray() },
					{ "OutroCedenteNome", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.OutroCedenteNome).ToArray() },
					{ "ConsumidorID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.ConsumidorID).ToArray() },
					{ "ColaboradorID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.ColaboradorID).ToArray() },
					{ "FornecedorID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.FornecedorID).ToArray() },
					{ "TransacaoID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.TransacaoID).ToArray() },
					{ "EspecieID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EspecieID).ToArray() },
					{ "DataBaseCalculo", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataBaseCalculo).ToArray() },
					{ "DevidoValor", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DevidoValor).ToArray() },
					{ "PagoValor", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PagoValor).ToArray() },
					{ "PlanoContasID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.PlanoContasID).ToArray() },
					{ "FinanceiroID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.FinanceiroID).ToArray() },
					{ "EstabelecimentoID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.EstabelecimentoID).ToArray() },
					{ "LoginID", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.LoginID).ToArray() },
					{ "DataInclusao", fluxosCaixa.ConvertAll(fluxoCaixa => (object)fluxoCaixa.DataInclusao).ToArray() },
				};

				salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_DentalOffice_FluxoCaixa");
				sqlHelper.GerarSqlInsert("_MigracaoFluxoCaixa_Temp", salvarArquivo, fluxosCaixaDict);
				excelHelper.GravarExcel(salvarArquivo, fluxosCaixaDict);
				//Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
			}

			catch (Exception error)
			{
				throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha, colunaLetra, tituloColuna, celulaValor, variaveisValor));
			}
		}

		public void ImportarRecebidos(string arquivoExcel, string arquivoExcelConsumidores, int estabelecimentoID, int respFinanceiroPessoaID, int loginID)
        {
            var dataHoje = DateTime.Now;
            var indiceLinha = 1;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";

            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            ISheet sheetConsumidores;
            if (string.IsNullOrEmpty(arquivoExcelConsumidores) == false)
                try
                {
                    IWorkbook workbookConsumidores = excelHelper.LerExcel(arquivoExcelConsumidores);
                    sheetConsumidores = workbookConsumidores.GetSheetAt(0);
                    excelHelper.InitializeDictionary(sheetConsumidores);
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
                    int controle = 0, recibo = 0, codigo = 0;
                    int? consumidorID = 0;
                    decimal pagoValor = 0;
                    byte titulosEspecies = 0;
                    var dataPagamento = dataHoje;
                    DateTime nascimentoData = dataHoje, data = dataHoje;

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
                                        titulosEspecies = celulaValor.ToTipoPagamento();
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
                            Data = data,
                            TransacaoID = (byte)TituloTransacoes.Liquidacao,
                            EspecieID = titulosEspecies,
                            DataBaseCalculo = data,
                            DevidoValor = pagoValor,
                            PagoValor = pagoValor,
                            EstabelecimentoID = estabelecimentoID,
                            LoginID = loginID,
                            DataInclusao = data,
                            FinanceiroID = respFinanceiroPessoaID
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
                            Data = data,
                            TransacaoID = (byte)TituloTransacoes.Liquidacao,
                            EspecieID = titulosEspecies,
                            DataBaseCalculo = data,
                            DevidoValor = pagoValor,
                            PagoValor = pagoValor,
                            EstabelecimentoID = estabelecimentoID,
                            LoginID = loginID,
                            DataInclusao = data,
                            FinanceiroID = respFinanceiroPessoaID
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

				var salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_DentalOffice_FluxoCaixa");
				sqlHelper.GerarSqlInsert("_MigracaoFluxoCaixa_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
				//Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
            }
            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }
        public void ImportarPacientes(string arquivoExcel, int estabelecimentoID)
        {
            var dataHoje = DateTime.Now;
            var indiceLinha = 1;
            string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";
            var excelHelper = new ExcelHelper(arquivoExcel);
            var sqlHelper = new SqlHelper();

            try
            {
                var linhasCount = excelHelper.linhas.Count;
                var consumidores = new List<Consumidor>();
                var pessoas = new List<Pessoa>();

                foreach (var linha in excelHelper.linhas)
                {
                    indiceLinha++;

                    foreach (var celula in linha.Cells)
                    {
                        if (celula != null)
                        {
                            celulaValor = celula.ToString().Trim().Replace("'", "’");
                            tituloColuna = excelHelper.cabecalhos[celula.Address.Column];
                            colunaLetra = excelHelper.GetColumnLetter(celula);
                            int numcadastro = 0;
                            string nomeCompleto = "", cpf = "";

                            if (!string.IsNullOrWhiteSpace(celulaValor))
                            {
                                //if (!dados.ContainsKey(tituloColuna))
                                //{
                                //	dados[tituloColuna] = new List<object>();
                                //}
                                //dados[tituloColuna].Add(int.Parse(celulaValor));

                                switch (tituloColuna)
                                {
                                    case "numcadastro":
                                        numcadastro = int.Parse(celulaValor);
                                        break;
                                    case "primeironome":
                                        nomeCompleto = celulaValor.GetPrimeirosCaracteres(70);
                                        break;
                                    case "cpf":
                                        cpf = celulaValor.ToCPF();
                                        break;
                                }
                            }
                        }
                    }

                    pessoas.Add(new Pessoa()
                    {
                        NomeCompleto = "",
                        Apelido = "",
                        CPF = "",
                        AssinaturaDigital = "",
                        CNS = "",
                        ConselhoCodigo = "",
                        ConselhoSigla = "",
                        ConselhoUF = "",
                        DataInclusao = dataHoje,
                        Email = "",
                        FalecimentoCausa = "",
                        FoneticaApelido = "",
                        FoneticaNomeCompleto = "",
                        FoneticaNomeSocial = "",
                        ID = 0,
                        Nacionalidade = "",
                        NascimentoLocal = "",
                        NomeSocial = "",
                        Origem = "",
                        ProfissaoOutra = "",
                        ResumoFormacao = "",
                        RG = "",
                        Sexo = false,
                        SkypeNome = "",
                        TipoSangue = ""
                    });
                }

                indiceLinha = 0;

                var dados = new Dictionary<string, object[]>
                {
                    { "NomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.NomeCompleto).ToArray() },
                };

				var salvarArquivo = Tools.GerarNomeArquivo($"Migração_{estabelecimentoID}_DentalOffice_FluxoCaixa");
				sqlHelper.GerarSqlInsert("_MigracaoFluxoCaixa_Temp", salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);
				//Tools.AbrirPastaSelecionandoArquivo(salvarArquivo + ".xlsx");
            }

            catch (Exception error)
            {
                throw new Exception(Tools.TratarMensagemErro(arquivoExcel, error.Message, indiceLinha, colunaLetra, tituloColuna, celulaValor, variaveisValor));
            }
        }

        void IDataBaseMigracao.DataBaseImportacaoProcedimentos()
        {
        }

        void IDataBaseMigracao.DataBaseImportacaoDevClinico()
        {
        }

        void IDataBaseMigracao.DataBaseImportacaoProntuarios()
        {
        }

        void IDataBaseMigracao.DataBaseImportacaoManutencoes()
        {
        }

        void IDataBaseMigracao.DataBaseImportacaoFinanceiroRecebidos()
        {
        }

        void IDataBaseMigracao.DataBaseImportacaoPacientes()
        {
        }

        public void DataBaseImportacaoFinanceiroRecebiveis()
        {
            throw new NotImplementedException();
        }
    }
}
