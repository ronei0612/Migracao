﻿using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Migração
{
	internal class DentalOffice
	{
		public int GetConsumidorID(ISheet sheet, string cpf = "", string nomeCompleto = "", string codigo = "")
		{
			//for (int row = 1; row <= sheet.LastRowNum; row++) // começa em 1 para pular o cabeçalho
			//{
			//	if (sheet.GetRow(row) != null) // verifica se a linha não está vazia
			//	{
			//		string cpfCellValue = sheet.GetRow(row).GetCell(2).ToString(); // assumindo que CPF é a terceira coluna
			//		if (cpfCellValue == cpf)
			//		{
			//			return int.Parse(sheet.GetRow(row).GetCell(0).ToString()); // retorna o ConsumidorID se o CPF corresponder
			//		}
			//	}
			//}

			IRow headerRow = sheet.GetRow(0); // assumindo que o cabeçalho está na primeira linha
			int retorno = 0;
			int nomeCompletoColumnIndex = -1;
			int codigoColumnIndex = -1;
			int cpfColumnIndex = -1;

			if (!string.IsNullOrWhiteSpace(cpf))
			{
				for (int column = 0; column < headerRow.LastCellNum; column++)
					if (headerRow.GetCell(column).ToString().Equals("cpf", StringComparison.CurrentCultureIgnoreCase))
					{
						cpfColumnIndex = column;
						break;
					}

				if (cpfColumnIndex == -1)
					throw new Exception("Coluna CPF não encontrada");

				for (int row = 1; row <= sheet.LastRowNum; row++) // começa em 1 para pular o cabeçalho
					if (sheet.GetRow(row) != null) // verifica se a linha não está vazia
					{
						string cpfCellValue = sheet.GetRow(row).GetCell(cpfColumnIndex).ToString();
						if (cpfCellValue == cpf)
							retorno = int.Parse(sheet.GetRow(row).GetCell(0).ToString()); // retorna o ConsumidorID se o CPF corresponder
					}
			}

			if (retorno == 0 && !string.IsNullOrWhiteSpace(nomeCompleto) && !string.IsNullOrWhiteSpace(codigo))
			{
				for (int column = 0; column < headerRow.LastCellNum; column++)
					if (headerRow.GetCell(column).ToString().Equals("nomecompleto", StringComparison.CurrentCultureIgnoreCase))
					{
						nomeCompletoColumnIndex = column;
						break;
					}

				if (nomeCompletoColumnIndex == -1)
					throw new Exception("Coluna NomeCompleto não encontrada");

				for (int column = 0; column < headerRow.LastCellNum; column++)
					if (headerRow.GetCell(column).ToString().Equals("codigoantigo", StringComparison.CurrentCultureIgnoreCase))
					{
						codigoColumnIndex = column;
						break;
					}

				if (codigoColumnIndex == -1)
					throw new Exception("Coluna CodigoAntigo não encontrada");

				for (int row = 1; row <= sheet.LastRowNum; row++) // começa em 1 para pular o cabeçalho
					if (sheet.GetRow(row) != null) // verifica se a linha não está vazia
					{
						string nomeCompletoCellValue = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString();
						string codigoCellValue = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString();

						if (nomeCompletoCellValue == nomeCompleto && codigoCellValue == codigo)
							retorno = int.Parse(sheet.GetRow(row).GetCell(0).ToString()); // retorna o ConsumidorID se o NomeCompleto e codigo corresponderem
					}
			}

			if (retorno == 0 && !string.IsNullOrWhiteSpace(nomeCompleto))
			{
				if (nomeCompletoColumnIndex == -1)
				{
					nomeCompletoColumnIndex = -1;

					for (int column = 0; column < headerRow.LastCellNum; column++)
						if (headerRow.GetCell(column).ToString().Equals("nomecompleto", StringComparison.CurrentCultureIgnoreCase))
						{
							nomeCompletoColumnIndex = column;
							break;
						}

					if (nomeCompletoColumnIndex == -1)
						throw new Exception("Coluna NomeCompleto não encontrada");
				}

				for (int row = 1; row <= sheet.LastRowNum; row++) // começa em 1 para pular o cabeçalho
					if (sheet.GetRow(row) != null) // verifica se a linha não está vazia
					{
						string nomeCompletoCellValue = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString();
						if (nomeCompletoCellValue == cpf)
							retorno = int.Parse(sheet.GetRow(row).GetCell(0).ToString()); // retorna o ConsumidorID se o CPF corresponder
					}
			}

			return retorno;
		}

		public void ImportarRecebidos(string arquivoExcel, string arquivoExcelConsumidores, string estabelecimentoID, string RespFinanceiroPessoaID, string salvarArquivo)
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

			try
			{
				var dados = new Dictionary<string, object[]>();

				var linhasCount = linhas.Count;

				var nomeCompleto = new string[linhasCount];
				var descricao = new string[linhasCount];
				var consumidorID = new int[linhasCount];
				var loginID = new int[linhasCount];
				var planoContasID = new int[linhasCount];
				var codigo = new long[linhasCount];
				var devidoValor = new decimal[linhasCount];
				var titulosEspecies = new byte[linhasCount];
				var dataVencimento = new DateTime[linhasCount];
				var nascimentoData = new DateTime[linhasCount];
				var transacaoID = new TituloTransacoes[linhasCount];
				var tituloSituacaoID = new TituloSituacoesID[linhasCount];

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
									case "paciente":
										nomeCompleto[indiceLinha - 2] = celulaValor.Substring(0, Math.Min(256, celulaValor.Length));
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
										devidoValor[indiceLinha - 2] = decimal.Parse(celulaValor.Replace(",", "."), CultureInfo.InvariantCulture);
										break;
									case "numero_registro":
										codigo[indiceLinha - 2] = int.Parse(celulaValor);
										break;
									case "descricao":
										descricao[indiceLinha - 2] = celulaValor;
										break;
									case "data_vencimento":
										if (DateTime.TryParse(celulaValor, out dataVencimento[indiceLinha - 2]))
										{
										}
										else if (double.TryParse(celulaValor, out double codigoData))
											dataVencimento[indiceLinha - 2] = DateTime.FromOADate(codigoData);
										else
											throw new Exception("Erro na conversão de data");
										if ((dataVencimento[indiceLinha - 2] >= dataMinima && dataVencimento[indiceLinha - 2] <= dataMaxima) == false)
											dataVencimento[indiceLinha - 2] = dataHoje;
										break;
								}
							}

							transacaoID[indiceLinha - 2] = TituloTransacoes.Liquidacao;
							loginID[indiceLinha - 2] = 1;
							planoContasID[indiceLinha - 2] = 1;
							tituloSituacaoID[indiceLinha - 2] = TituloSituacoesID.Normal;

							consumidorID[indiceLinha - 2] = GetConsumidorID(sheetConsumidores, nomeCompleto: nomeCompleto[indiceLinha - 2], codigo: codigo[indiceLinha - 2].ToString());
							//tituloSituacaoID[indiceLinha - 2] = documento > 0 ? documento : (long?)null;
						}
					}
				}

				indiceLinha = 0;

				dados.Add("ConsumidorID", consumidorID.Cast<object>().ToArray());
				dados.Add("EspecieID", titulosEspecies.Cast<object>().ToArray());
				dados.Add("DataEmissao", dataVencimento.Cast<object>().ToArray());
				dados.Add("DataVencimento", dataVencimento.Cast<object>().ToArray());
				dados.Add("DataInclusao", dataVencimento.Cast<object>().ToArray());
				dados.Add("DataBaseCalculo", dataVencimento.Cast<object>().ToArray());
				//dados.Add("DataBaixa", null);
				dados.Add("ValorOriginal", devidoValor.Cast<object>().ToArray());
				dados.Add("ValorDevido", devidoValor.Cast<object>().ToArray());
				dados.Add("FinanceiroID", RespFinanceiroPessoaID.Cast<object>().ToArray());
				dados.Add("SituacaoID", 1);
				dados.Add("LoginID", loginID.Cast<object>().ToArray());
				dados.Add("PlanoContasID", planoContasID.Cast<object>().ToArray());
				dados.Add("TransacaoID", transacaoID.Cast<object>().ToArray());
				dados.Add("Observacoes", descricao.Cast<object>().ToArray());
				//dados.Add("Documento", codigo.Cast<object>().ToArray());
				dados.Add("EstabelecimentoID", estabelecimentoID.Cast<object>().ToArray());
				dados.Add("ExclusaoMotivo", string.IsNullOrWhiteSpace(nomeCompleto) ? null : nomeCompleto.Cast<object>().ToArray());

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
