using NPOI.SS.UserModel;
using System.Text.RegularExpressions;

namespace Migração
{
	internal class DentalOffice
	{
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
