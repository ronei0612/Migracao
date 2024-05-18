using Migração.Utils;
using Migração.Models;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Bcpg.OpenPgp;
using System.Globalization;
using System.Runtime.ConstrainedExecution;
using System.Text.RegularExpressions;
using static NPOI.HSSF.Util.HSSFColor;
using static OfficeOpenXml.ExcelErrorValue;
using System.Diagnostics;

namespace Migração
{
	internal class OdontoCompany
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
				excelHelper.InitializeDictionaryConsumidor(sheetConsumidores);
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
		public void ImportarPacientes(string arquivoExcel, string arquivoExcelCidades, string estabelecimentoID, string salvarArquivo)
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
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

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
					bool cliente = false, fornecedor = false;
					DateTime dataNascimento, dataCadastro;
					int numcadastro;
					string nomeCompleto = "", cpf = "", rg = "", email = "", apelido = "";
					bool sexo = true;
					long telefonePrinc, telefoneAltern, telefoneComercial, telefoneOutro, celular;

					foreach (var celula in linha.Cells)
					{
						if (celula != null)
						{
							celulaValor = celula.ToString().Trim().Replace("'", "’");
							tituloColuna = cabecalhos[celula.Address.Column];
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
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										apelido = celulaValor.Contains(" ") ? celulaValor.Split(' ')[0] : celulaValor;
										break;
									case "CGC_CPF":
										cpf = celulaValor.Contains('.') && celulaValor.Contains('-') && celulaValor.Length <= 14 ? celulaValor
											: celulaValor.Length == int.Parse(mascaraCPFLenth) ? Convert.ToUInt64(celulaValor).ToString(mascaraCPF) : "";
										break;
									case "INSC_RG":
										rg = celulaValor.Substring(0, Math.Min(20, celulaValor.Length));
										break;

									case "SEXO_M_F":
										var sexoLetra = celulaValor.ToLower();
										sexo = sexoLetra == "m" || sexoLetra != "f";
										break;
									case "EMAIL":
										email = celulaValor.Contains('@') && celulaValor.Contains('.') ? celulaValor : "";
										break;
									case "FONE1":
										var possivelTel1 = Regex.Replace(celulaValor, "[^0-9]", "");
										if (possivelTel1.Length >= 8 && possivelTel1.Length <= 16)
											telefonePrinc = long.Parse(possivelTel1);
										break;
									case "FONE2":
										var possivelTel2 = Regex.Replace(celulaValor, "[^0-9]", "");
										if (possivelTel2.Length >= 8 && possivelTel2.Length <= 16)
											telefoneAltern = long.Parse(possivelTel2);
										break;
									case "CELULAR":
										var possivelCelular = celulaValor;
										if (celulaValor.Length > 15)
										{
										}
										else
										{
											possivelCelular = Regex.Replace(possivelCelular, "[^0-9]", "");
											if (possivelCelular.Length >= 8 && possivelCelular.Length <= 16)
												celular = long.Parse(possivelCelular);
										}
										break;
									case "ENDERECO":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "BAIRRO":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "NUM_ENDERECO":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "CIDADE":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "ESTADO":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "CEP":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "OBS1":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "NUM_CONVENIO":
										nomeCompleto = celulaValor.Substring(0, Math.Min(70, celulaValor.Length));
										break;
									case "DT_CADASTRO":
										if (DateTime.TryParse(celulaValor, out dataCadastro))
										{
										}
										else if (double.TryParse(celulaValor, out double codigoData))
											dataCadastro = DateTime.FromOADate(codigoData);
										else
											throw new Exception("Erro na conversão de data");
										if ((dataCadastro >= dataMinima && dataCadastro <= dataMaxima) == false)
											dataCadastro = dataHoje;
										break;
									case "DT_NASCIMENTO":
										if (DateTime.TryParse(celulaValor, out dataNascimento))
										{
										}
										else if (double.TryParse(celulaValor, out double codigoData))
											dataNascimento = DateTime.FromOADate(codigoData);
										else
											throw new Exception("Erro na conversão de data");
										if ((dataNascimento >= dataMinima && dataNascimento <= dataMaxima) == false)
											dataNascimento = dataHoje;
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
							DataInclusao = dataHoje,
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

				if (File.Exists($"{salvarArquivo}.xlsx"))
				{
				int count = 1;
				while (File.Exists($"{salvarArquivo} ({count}).xlsx"))
						count++;

					salvarArquivo = $"{salvarArquivo} ({count})";
				}

				var sqlHelper = new SqlHelper();
				var insert = sqlHelper.GerarSqlInsert("_MigracaoConsumidores_Temp", dados);

				File.WriteAllText(salvarArquivo + ".sql", insert);
				excelHelper.GravarExcel(salvarArquivo, dados);

				string argumento = "/select, \"" + salvarArquivo + ".xlsx" + "\"";

				Process.Start("explorer.exe", argumento);
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
