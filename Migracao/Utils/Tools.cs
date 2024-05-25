using Migracao.Models;
using NPOI.SS.UserModel;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Migracao.Utils
{
	internal static class Tools
	{
		public static string mascaraCPF = "000.000.000-00";
		public static string salvarNaPasta = Environment.ExpandEnvironmentVariables("%userprofile%\\Documents");
		public static string ultimaPasta = Environment.ExpandEnvironmentVariables("%userprofile%\\Documents");

		public static string ToCPF(this string possivelCpf)
		{
			if (possivelCpf.Contains('.') && possivelCpf.Contains('-') && possivelCpf.Length <= 14)
				return possivelCpf;

			else
			{
				var mascaraCPFformat = mascaraCPF.Split('.')[0].Replace(".", @"\.").Replace("-", @"\-");
				var mascaraCPFLenth = Regex.Replace(mascaraCPFformat, "[^0-9]", "").Length.ToString();

				if (possivelCpf.Length == int.Parse(mascaraCPFLenth))
					return Convert.ToUInt64(possivelCpf).ToString(mascaraCPFformat).GetPrimeirosCaracteres(14);
			}

			return "";
		}

		public static string GetPrimeiroNome(this string texto)
		{
			return texto.Contains(' ') ? texto.Split(' ')[0] : texto;
		}

		public static string ToEmail(this string email)
		{
			//return texto.Contains('@') && texto.Contains('.') ? texto : "";
			var emailRegex = new Regex(@"^[\w-]+(\.[\w-]+)*@([\w-]+\.)+[a-zA-Z]{2,7}$");
			if (emailRegex.IsMatch(email))
				return email.ToLower();

			return "";
		}

		public static long ToFone(this string telefone)
		{
			var possivelTel = Regex.Replace(telefone, "[^0-9]", "");

			if (string.IsNullOrEmpty(possivelTel))
				return 0;
			else if (possivelTel.Length >= 8 && possivelTel.Length <= 16)
				return long.Parse(possivelTel);
			else
				return long.Parse(possivelTel.GetPrimeirosCaracteres(16));
		}

		public static string GetPrimeirosCaracteres(this string texto, int max)
		{
			return texto.Substring(0, Math.Min(max, texto.Length));
		}

		public static DateTime ToData(this string texto)
		{
			DateTime dataMinima = new(1900, 01, 01), dataMaxima = new(2079, 06, 06), dataHoje = DateTime.Now, data;

			if (DateTime.TryParse(texto, out data))
			{
			}
			else if (double.TryParse(texto, out double codigoData))
				data = DateTime.FromOADate(codigoData);
			else
				throw new Exception("Erro na conversão de data");
			if ((data >= dataMinima && data <= dataMaxima) == false)
				data = dataHoje;

			return data;
		}

		public static int ToNum(this string texto)
		{
			return int.Parse(Regex.Replace(texto, "[^0-9]", ""));
		}

		public static bool ToSexo(this string texto, string masculino, string feminino)
		{
			var sexoLetra = texto.ToLower();

			if (sexoLetra == masculino)
				return true;

			else if (sexoLetra == feminino)
				return false;

			return true;
		}

		public static string ToNomeCompleto(this string texto)
		{
			return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(texto.ToLower());
		}


		public static string GetLetras(this string texto)
		{
			return Regex.Replace(texto, @"[^a-zA-Z\?\s]", "").Trim();
			//return Regex.Replace(texto, @"[^\p{L}\s]", "").Trim();
		}

		public static string GerarNomeArquivo(string nomeArquivo)
		{
			var caminhoDoArquivo = Path.Combine(Tools.salvarNaPasta, nomeArquivo);

			if (File.Exists(caminhoDoArquivo + ".xlsx"))
			{
				int count = 1;
				while (File.Exists($"{caminhoDoArquivo} ({count}).xlsx"))
					count++;

				caminhoDoArquivo = $"{caminhoDoArquivo} ({count++})";
			}

			return caminhoDoArquivo;
		}

		public static void AbrirPastaSelecionandoArquivo(string nomeArquivo)
		{
			string argumento = "/select, \"" + nomeArquivo + "\"";

			Process.Start("explorer.exe", argumento);
		}

		public static void AbrirPastaExplorer(string pasta)
		{
			string argumento = "\"" + pasta + "\"";

			Process.Start("explorer.exe", argumento);
		}

		public static string TratarMensagemErro(string erroMensagem, int indiceLinha, string colunaLetra, string tituloColuna, string celulaValor, string variaveisValor = "")
		{
			var mensagemErro = $"Falha na linha {indiceLinha}, coluna {colunaLetra}, Valor esperado: {tituloColuna}, valor da célula: \"{celulaValor}\": {erroMensagem}";

			if (!string.IsNullOrWhiteSpace(variaveisValor))
				mensagemErro += Environment.NewLine + "Variáveis" + Environment.NewLine + variaveisValor;

			if (indiceLinha <= 0)
				mensagemErro = erroMensagem;

			return mensagemErro;
		}

		public static byte ToTipoPagamento(this string texto)
		{
			return (byte)(texto.Equals("dinheiro", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.Dinheiro
						: texto.Equals("cheque", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.Cheque
						: texto.Equals("boleto bancário", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.BoletoBancario
						: texto.Equals("cartão de crédito", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.CartaoCredito
						: texto.Equals("debito", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.CartaoDebito
						: texto.Equals("cartão de débito", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.CartaoDebito
						: texto.Equals("pix", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.CreditoEmConta
						: texto.Equals("débito automático", StringComparison.CurrentCultureIgnoreCase) ? TitulosEspeciesID.CartaoCreditoRecorrente
						: TitulosEspeciesID.DepositoEmConta);
		}

		public static decimal ToMoeda(this string texto)
		{
			if (texto.Contains(',') && texto.Contains('.'))
				return decimal.Parse(texto.Replace(".", "").Replace(",", "."), CultureInfo.InvariantCulture);
			else
				return decimal.Parse(texto.Replace(",", "."), CultureInfo.InvariantCulture);
		}


		public static string TratarCaracteres(string texto)
		{
			var regex = new Regex("[^a-zA-Z0-9 -_]");
			texto = regex.Replace(texto, "");
			return texto;
		}

		public static bool IsCPF(this string texto)
		{
			texto = Regex.Replace(texto, "[^0-9]", "");
			if (texto.Length == 11)
				return true;

			return false;
		}

		public static bool IsCNPJ_CGC(this string texto)
		{
			texto = Regex.Replace(texto, "[^0-9]", "");
			if (texto.Length == 14)
				return true;

			return false;
		}


		// Função auxiliar para obter valores inteiros de uma célula, tratando células vazias
		public static int? GetIntValueFromCell(ICell cell)
		{
			if (cell == null || cell.CellType == CellType.Blank)
				return null;
			return (int)cell.NumericCellValue;
		}

		// Função auxiliar para obter valores decimais de uma célula, tratando células vazias
		public static decimal? GetDecimalValueFromCell(ICell cell)
		{
			if (cell == null || cell.CellType == CellType.Blank)
				return null;
			if (cell is decimal)
				return (decimal)cell.NumericCellValue;
			else
				return decimal.Parse(cell.StringCellValue);
		}

		// Função auxiliar para obter valores de data/hora de uma célula, tratando células vazias
		public static DateTime? GetDateTimeValueFromCell(ICell cell)
		{
			if (cell == null || cell.CellType == CellType.Blank)
				return DateTime.Now;
			if (cell is DateTime)
				return cell.DateCellValue;
			else
				return DateTime.Parse(cell.ToString());

			//.ToString("yyyy-MM-dd HH:mm:ss.f");
		}

		// Função auxiliar para obter valores de TimeSpan de uma célula, tratando células vazias
		public static TimeSpan? GetTimeSpanValueFromCell(ICell cell)
		{
			if (cell == null || cell.CellType == CellType.Blank)
				return null;

			// Converte o valor da célula (que pode ser um DateTime ou um double) para TimeSpan
			if (cell.CellType == CellType.Numeric)
			{
				// Se for um número, assume que é um valor de tempo em dias (como o Excel armazena)
				return TimeSpan.FromDays(cell.NumericCellValue);
			}
			else
			{
				// Se for um DateTime, converte para TimeSpan
				return TimeSpan.Parse(cell.DateCellValue.ToString());//.TimeOfDay;
			}
		}

		//public static string EncontrarCidadeSemelhante(string textoCidade, string[] cidades)
		//{
		//	textoCidade = RemoverAcentos(textoCidade).ToLower();

		//	string cidadeEncontrada = null;
		//	int maiorSemelhanca = 0;

		//	foreach (string cidade in cidades)
		//	{
		//		string cidadeNormalizada = RemoverAcentos(cidade).ToLower();
		//		int semelhanca = CalcularSemelhanca(textoCidade, cidadeNormalizada);

		//		if (semelhanca > maiorSemelhanca)
		//		{
		//			maiorSemelhanca = semelhanca;
		//			cidadeEncontrada = cidade;
		//		}
		//	}

		//	return cidadeEncontrada;
		//}

		public static string EncontrarCidadeSemelhante(string textoCidade, string[] cidades)
		{
			textoCidade = RemoverAcentos(textoCidade).ToLower();

			string cidadeEncontrada = null;
			int menorDistancia = int.MaxValue;

			foreach (string cidade in cidades)
			{
				string cidadeNormalizada = RemoverAcentos(cidade).ToLower();
				int distancia = DistanciaLevenshtein(textoCidade, cidadeNormalizada);

				if (distancia < menorDistancia)
				{
					menorDistancia = distancia;
					cidadeEncontrada = cidade;
				}
			}

			return cidadeEncontrada;
		}

		private static string RemoverAcentos(string texto)
		{
			return new string(texto
				.Normalize(NormalizationForm.FormD)
				.Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark)
				.ToArray());
		}

		// Função simples para calcular a semelhança entre duas strings
		private static int CalcularSemelhanca(string str1, string str2)
		{
			int contador = 0;
			int tamanhoMenor = Math.Min(str1.Length, str2.Length);

			for (int i = 0; i < tamanhoMenor; i++)
			{
				if (str1[i] == str2[i])
				{
					contador++;
				}
			}

			return contador;
		}

		// Implementação da Distância de Levenshtein
		private static int DistanciaLevenshtein(string s, string t)
		{
			int[,] d = new int[s.Length + 1, t.Length + 1];

			for (int i = 0; i <= s.Length; i++)
			{
				d[i, 0] = i;
			}

			for (int j = 0; j <= t.Length; j++)
			{
				d[0, j] = j;
			}

			for (int j = 1; j <= t.Length; j++)
			{
				for (int i = 1; i <= s.Length; i++)
				{
					int custo = (s[i - 1] == t[j - 1]) ? 0 : 1;

					d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + custo);
				}
			}

			return d[s.Length, t.Length];
		}

		//public static decimal ArredondarValor(this string input)
		//{
		//	input = input.Replace(" ", "").Replace(",", "");

		//	if (decimal.TryParse(input, out decimal valor))
		//		return valor;
		//	else
		//	{
		//		string modificado = input.Substring(0, input.Length - 3);
		//		modificado = modificado.TrimStart('0');
		//		decimal valorResult = decimal.Parse(modificado);

		//		return valorResult / 100;
		//	}
		//}

		public static decimal ArredondarValor(this string valor)
		{
			if (decimal.TryParse(valor, out decimal valorDecimal))
			{
				if (valorDecimal >= 1000000000)
				{
					return Math.Round(valorDecimal / 10000000000, 2);
				}
				else if (valorDecimal >= 1000000)
				{
					return Math.Round(valorDecimal / 1000000, 2);
				}
				else if (valorDecimal >= 1000)
				{
					return Math.Round(valorDecimal / 1000, 2);
				}
				else
				{
					return Math.Round(valorDecimal, 2);
				}
			}
			else
			{
				return 0; // Retorna 0 se a conversão falhar
			}
		}
	}
}