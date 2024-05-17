using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System.Globalization;
using System.Text.RegularExpressions;
using static NPOI.HSSF.Util.HSSFColor;
using static OfficeOpenXml.ExcelErrorValue;

namespace Migração
{
	internal class DentalOffice
	{
		public void ImportarRecebidos(string arquivoExcel, string arquivoExcelConsumidores, string estabelecimentoID, string respFinanceiroPessoaID, string salvarArquivo)
		{
			DateTime dataMinima = new DateTime(1900, 01, 01), dataMaxima = new DateTime(2079, 06, 06), dataHoje = DateTime.Now;
			var indiceLinha = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";

			var mascaraCPF = "000.000.000-00";
			mascaraCPF = mascaraCPF.Split('.')[0].Replace(".", @"\.").Replace("-", @"\-");
			var mascaraCPFLenth = Regex.Replace(mascaraCPF, "[^0-9]", "").Length.ToString();

			IWorkbook workbook;
			var excelHelper = new ExcelHelper();
			try
			{
				workbook = excelHelper.LerExcel(arquivoExcel);
			}
			catch (Exception ex)
			{
				throw new Exception("Erro ao ler o arquivo Excel: " + ex.Message);
			}

			ISheet sheetConsumidores;
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

			var cabecalhos = excelHelper.GetCabecalhosExcel(workbook);
			var linhas = excelHelper.GetLinhasExcel(workbook);
			var fluxoCaixas = new List<FluxoCaixa>();

			try
			{
				foreach (var linha in linhas)
				{
					indiceLinha++;

					string nomeCompleto = "", dataPagamento, outroSacadoNome = "";
					int controle = 0, recibo = 0, codigo = 0, loginID = 1;
					int? consumidorID = 0;
					decimal pagoValor = 0;
					byte titulosEspecies = 0;
					DateTime nascimentoData = dataHoje, data = dataHoje;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim();
							tituloColuna = cabecalhos[celula.Address.Column];
							colunaLetra = excelHelper.GetColumnLetter(celula);

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "paciente":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "numero_registro":
										codigo = int.Parse(celulaValor);
										break;
									case "data_pagamento":
										if (DateTime.TryParse(celulaValor, out data))
										{
										}
										else if (double.TryParse(celulaValor, out double codigoData))
											data = DateTime.FromOADate(codigoData);
										else
											throw new Exception("Erro na conversão de data");
										if ((data >= dataMinima && data <= dataMaxima) == false)
											data = dataHoje;
										dataPagamento = data.ToString("yyyy-MM-dd HH:mm:ss.f");
										break;
									case "forma_pagamento":
										titulosEspecies = (byte)(celulaValor.ToLower() == "dinheiro" ? TitulosEspeciesID.Dinheiro
											: celulaValor.ToLower() == "cheque" ? TitulosEspeciesID.Cheque
											: celulaValor.ToLower() == "boleto bancário" ? TitulosEspeciesID.BoletoBancario
											: celulaValor.ToLower() == "cartão de crédito" ? TitulosEspeciesID.CartaoCredito
											: celulaValor.ToLower() == "debito" ? TitulosEspeciesID.CartaoDebito
											: celulaValor.ToLower() == "cartão de débito" ? TitulosEspeciesID.CartaoDebito
											: celulaValor.ToLower() == "pix" ? TitulosEspeciesID.CreditoEmConta
											: celulaValor.ToLower() == "débito automático" ? TitulosEspeciesID.CartaoCreditoRecorrente
											: TitulosEspeciesID.DepositoEmConta);
										break;
									case "valor":
										pagoValor = decimal.Parse(celulaValor.Replace(",", "."), CultureInfo.InvariantCulture);
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
							TransacaoID = (byte)TituloTransacoes.PagamentoAvulso,
							EspecieID = titulosEspecies,
							DataBaseCalculo = data,
							DevidoValor = pagoValor,
							PagoValor = pagoValor,
							EstabelecimentoID = int.Parse(estabelecimentoID),
							LoginID = 1,
							DataInclusao = data,
							FinanceiroID = int.Parse(respFinanceiroPessoaID)
						});
					}
					else
					{
						consumidorID = null;
						outroSacadoNome = nomeCompleto.Substring(0, Math.Min(50, nomeCompleto.Length));

						fluxoCaixas.Add(new FluxoCaixa()
						{
							OutroSacadoNome = nomeCompleto.Substring(0, Math.Min(50, nomeCompleto.Length)),
							SituacaoID = 1,
							PagoMulta = 0,
							PagoJuros = 0,
							TipoID = (byte)TransacaoTiposID.Recebimento,
							Data = data,
							TransacaoID = (byte)TituloTransacoes.PagamentoAvulso,
							EspecieID = titulosEspecies,
							DataBaseCalculo = data,
							DevidoValor = pagoValor,
							PagoValor = pagoValor,
							EstabelecimentoID = int.Parse(estabelecimentoID),
							LoginID = 1,
							DataInclusao = data,
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

				var sqlHelper = new SqlHelper();
				var insert = sqlHelper.GerarSqlInsert("_MigracaoFluxoCaixa_Temp", dados);
				File.WriteAllText(salvarArquivo + ".sql", insert);
				excelHelper.GravarExcel(salvarArquivo, dados);
			}
			catch (Exception error)
			{
				var mensagemErro = $"Falha na linha {indiceLinha}, coluna {colunaLetra}, Valor esperado: {tituloColuna}, valor da célula: \"{celulaValor}\": {error.Message}";
				if (!string.IsNullOrWhiteSpace(variaveisValor))
					mensagemErro += Environment.NewLine + "Variáveis" + Environment.NewLine + variaveisValor;

				if (indiceLinha <= 0)
					mensagemErro = error.Message;

				throw new Exception(mensagemErro);
			}
		}
		public void ImportarPacientes(string arquivoExcel, string estabelecimentoID, string salvarArquivo)
		{
			DateTime dataMinima = new DateTime(1900, 01, 01), dataMaxima = new DateTime(2079, 06, 06), dataHoje = DateTime.Now;
			var indiceLinha = 1;
			string tituloColuna = "", colunaLetra = "", celulaValor = "", variaveisValor = "";

			var mascaraCPF = "000.000.000-00";
			mascaraCPF = mascaraCPF.Split('.')[0].Replace(".", @"\.").Replace("-", @"\-");
			var mascaraCPFLenth = Regex.Replace(mascaraCPF, "[^0-9]", "").Length.ToString();

			IWorkbook workbook;
			var excelHelper = new ExcelHelper();
			try
			{
				workbook = excelHelper.LerExcel(arquivoExcel);
			}
			catch (Exception ex)
			{
				throw new Exception("Erro ao ler o arquivo Excel: " + ex.Message);
			}

			var cabecalhos = excelHelper.GetCabecalhosExcel(workbook);
			var linhas = excelHelper.GetLinhasExcel(workbook);

			try
			{
				var linhasCount = linhas.Count;
				var consumidores = new List<Consumidor>();
				var pessoas = new List<Pessoa>();

				foreach (var linha in linhas)
				{
					indiceLinha++;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim();
							tituloColuna = cabecalhos[celula.Address.Column];
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
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "cpf":
										cpf = celulaValor.Contains('.') && celulaValor.Contains('-') && celulaValor.Length <= 14 ? celulaValor 
											: celulaValor.Length == int.Parse(mascaraCPFLenth) ? Convert.ToUInt64(celulaValor).ToString(mascaraCPF) : "";
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

				//dados.Add("numcadastro", numcadastro.Cast<object>().ToArray());
				//dados.Add("nomeCompleto", nomeCompleto.Cast<object>().ToArray());
				//dados.Add("cpf", cpf.Cast<object>().ToArray());

				var dados1 = new Dictionary<string, object[]>
				{
					{ "NomeCompleto", pessoas.ConvertAll(pessoa => (object)pessoa.NomeCompleto).ToArray() },
					//{ "CPF", pessoas.ConvertAll(pessoa => (object)pessoa.CPF).ToArray() },
					//{ "Telefone", pessoas.ConvertAll(pessoa => (object)pessoa.Telefone).ToArray() }
				};

				var sqlHelper = new SqlHelper();

				var insert = sqlHelper.GerarSqlInsert(salvarArquivo, dados1);
				excelHelper.GravarExcel(salvarArquivo, dados1);

				File.WriteAllText(salvarArquivo + ".sql", insert);
			}

			catch (Exception error)
			{
				var mensagemErro = $"Falha na linha {indiceLinha}, coluna {colunaLetra}, Valor esperado: {tituloColuna}, valor da célula: \"{celulaValor}\": {error.Message}";

				if (!string.IsNullOrWhiteSpace(variaveisValor))
					mensagemErro += Environment.NewLine + "Variáveis" + Environment.NewLine + variaveisValor;

				if (indiceLinha <= 0)
					mensagemErro = error.Message;

				throw new Exception(mensagemErro);
			}
		}
	}
}
