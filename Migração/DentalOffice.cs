using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Migração
{
	internal class DentalOffice
	{
		public void ImportarRecebidos(string arquivoExcel, string arquivoExcelConsumidores, string EstabelecimentoID, string RespFinanceiroPessoaID, string salvarArquivo)
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
			}
			catch (Exception ex)
			{
				throw new Exception("Erro ao ler o arquivo Excel: " + ex.Message);
			}

			var cabecalhos = excelHelper.GetCabecalhosExcel(workbook);
			var linhas = excelHelper.GetLinhasExcel(workbook);
			excelHelper.ZerarVariaveis();

			try
			{
				var dados = new Dictionary<string, object[]>();

				var linhasCount = linhas.Count;

				var nomeCompleto = new string[linhasCount];
				var descricao = new string[linhasCount];
				string[] consumidorID = new string[linhasCount];
				string?[] outroSacadoNome = new string?[linhasCount];
				var loginID = new int[linhasCount];
				var planoContasID = new int[linhasCount];
				var codigo = new long[linhasCount];
				var pagoValor = new decimal[linhasCount];
				var pagoMulta = new int[linhasCount];
				var pagoJuros = new int[linhasCount];
				var data = dataHoje;
				var dataPagamento = new string[linhasCount];
				var nascimentoData = new DateTime[linhasCount];
				var titulosEspecies = new byte[linhasCount];
				var transacaoID = new byte[linhasCount];
				var tituloSituacaoID = new byte[linhasCount];
				var tipoID = new byte[linhasCount];
				var respFinanceiroPessoaID = new string[linhasCount];
				var estabelecimentoID = new string[linhasCount];

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

							if (!string.IsNullOrWhiteSpace(celulaValor))
							{
								switch (tituloColuna)
								{
									case "paciente":
										nomeCompleto[indiceLinha - 2] = celulaValor.Substring(0, Math.Min(256, celulaValor.Length));
										break;
									case "numero_registro":
										codigo[indiceLinha - 2] = int.Parse(celulaValor);
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
										dataPagamento[indiceLinha - 2] = data.ToString("yyyy-MM-dd HH:mm:ss.f");
										break;
									case "forma_pagamento":
										titulosEspecies[indiceLinha - 2] = (byte)(celulaValor.ToLower() == "dinheiro" ? TitulosEspeciesID.Dinheiro
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
										pagoValor[indiceLinha - 2] = decimal.Parse(celulaValor.Replace(",", "."), CultureInfo.InvariantCulture);
										break;
								}
							}

							transacaoID[indiceLinha - 2] = (byte)TituloTransacoes.Liquidacao;
							tituloSituacaoID[indiceLinha - 2] = (byte)TituloSituacoesID.Normal;
							tipoID[indiceLinha - 2] = (byte)TransacaoTiposID.Recebimento;
							loginID[indiceLinha - 2] = 1;
							planoContasID[indiceLinha - 2] = 55;
							pagoMulta[indiceLinha - 2] = 0;
							pagoJuros[indiceLinha - 2] = 0;
							respFinanceiroPessoaID[indiceLinha - 2] = RespFinanceiroPessoaID.Trim();
							estabelecimentoID[indiceLinha - 2] = EstabelecimentoID.Trim();

							int consumidorIDValue = excelHelper.GetConsumidorID(sheetConsumidores, nomeCompleto: nomeCompleto[indiceLinha - 2], codigo: codigo[indiceLinha - 2].ToString());
							if (consumidorIDValue > 0)
							{
								consumidorID[indiceLinha - 2] = consumidorIDValue.ToString();
								outroSacadoNome[indiceLinha - 2] = "null";
							}
							else
							{
								consumidorID[indiceLinha - 2] = "null";
								outroSacadoNome[indiceLinha - 2] = nomeCompleto[indiceLinha - 2].Substring(0, Math.Min(50, nomeCompleto[indiceLinha - 2].Length));
							}
						}
					}
				}

				indiceLinha = 0;

				dados.Add("ConsumidorID", consumidorID.Cast<object>().ToArray());
				dados.Add("SituacaoID", tituloSituacaoID.Cast<object>().ToArray());
				dados.Add("PagoMulta", pagoMulta.Cast<object>().ToArray());
				dados.Add("PagoJuros", pagoJuros.Cast<object>().ToArray());
				dados.Add("DevidoValor", pagoValor.Cast<object>().ToArray());
				dados.Add("PagoValor", pagoValor.Cast<object>().ToArray());

				dados.Add("TipoID", tipoID.Cast<object>().ToArray());
				dados.Add("LoginID", loginID.Cast<object>().ToArray());
				dados.Add("PlanoContasID", planoContasID.Cast<object>().ToArray());
				dados.Add("TransacaoID", transacaoID.Cast<object>().ToArray());
				dados.Add("EspecieID", titulosEspecies.Cast<object>().ToArray());
				dados.Add("FinanceiroID", respFinanceiroPessoaID.Cast<object>().ToArray());
				dados.Add("EstabelecimentoID", estabelecimentoID.Cast<object>().ToArray());

				dados.Add("OutroSacadoNome", outroSacadoNome.Cast<object>().ToArray());
				dados.Add("Data", dataPagamento.Cast<object>().ToArray());
				dados.Add("DataBaseCalculo", dataPagamento.Cast<object>().ToArray());
				dados.Add("DataInclusao", dataPagamento.Cast<object>().ToArray());

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
				var dados = new Dictionary<string, object[]>();

				var linhasCount = linhas.Count;

				var nomeCompleto = new string[linhasCount];
				var cpf = new string[linhasCount];
				var numcadastro = new int[linhasCount];
				var consumidorID = new int[linhasCount];
				var codigoAntigo = new int[linhasCount];
				var pessoaID = new int[linhasCount];

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
										numcadastro[indiceLinha - 2] = int.Parse(celulaValor);
										break;
									case "primeironome":
										nomeCompleto[indiceLinha - 2] = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "cpf":
										cpf[indiceLinha - 2] = celulaValor.Contains('.') && celulaValor.Contains('-') && celulaValor.Length <= 14 ? celulaValor 
											: celulaValor.Length == int.Parse(mascaraCPFLenth) ? Convert.ToUInt64(celulaValor).ToString(mascaraCPF) : "";
										break;
								}
							}
						}
					}
				}

				indiceLinha = 0;

				dados.Add("numcadastro", numcadastro.Cast<object>().ToArray());
				dados.Add("nomeCompleto", nomeCompleto.Cast<object>().ToArray());
				dados.Add("cpf", cpf.Cast<object>().ToArray());

				var sqlHelper = new SqlHelper();

				var insert = sqlHelper.GerarSqlInsert(salvarArquivo, dados);
				excelHelper.GravarExcel(salvarArquivo, dados);

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
