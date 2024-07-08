using Migracao.Models;
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
		public static string ultimoEstabelecimentoID = "";
        public static string ultimoEstabelecimento = "";
        public static string ultimoAntigoSistema = "0";
        public static string ultimoinputDB = "";
        public static string ultimoinputDBContratos = "";
		public static ExcelHelper excelHelper;
        
        private static string arquivoConfig = "config.config";

        public static string? ToCPF(this string possivelCpf)
		{
			if (string.IsNullOrEmpty(possivelCpf))
				return null;

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
			if (string.IsNullOrEmpty(texto))
				return "";

			return texto.Contains(' ') ? texto.Split(' ')[0] : texto;
		}

		public static string ToSN(this bool texto)
		{
			if (texto)
				return "S";

			return "N";
		}

		public static string? ToEmail(this string email)
		{
			if (string.IsNullOrEmpty(email))
				return null;

			//return texto.Contains('@') && texto.Contains('.') ? texto : "";
			var emailRegex = new Regex(@"^[\w-]+(\.[\w-]+)*@([\w-]+\.)+[a-zA-Z]{2,7}$");
			if (emailRegex.IsMatch(email))
				return email.ToLower();

			return null;
		}

		public static long? ToFone(this string telefone)
		{
			if (string.IsNullOrEmpty(telefone))
				return null;
			var possivelTel = Regex.Replace(telefone, "[^0-9]", "");

			if (string.IsNullOrEmpty(possivelTel))
				return null;
			else if (possivelTel.Length >= 8 && possivelTel.Length <= 16)
				return long.Parse(possivelTel);
			else
				return long.Parse(possivelTel.GetPrimeirosCaracteres(16));
		}

		public static string GetPrimeirosCaracteres(this string texto, int max)
		{
			return texto.Substring(0, Math.Min(max, texto.Length));
		}

		public static DateTime? ToDataNull(this string texto)
		{
			if (string.IsNullOrEmpty(texto))
				return null;

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

		public static DateTime ToData(this string texto)
		{

			try
			{
                if (string.IsNullOrEmpty(texto))
                    return DateTime.Now;

                DateTime dataMinima = new(1900, 01, 01), dataMaxima = new(2079, 06, 06), dataHoje = DateTime.Now, data;

                if (DateTime.TryParse(texto, out data))
                {
                }
                else if (double.TryParse(texto, out double codigoData))
                    data = DateTime.FromOADate(codigoData);
                else                    

                if ((data >= dataMinima && data <= dataMaxima) == false)
                    data = dataHoje;

                return data;
            }
			catch (Exception)
			{
                throw new Exception("Erro na conversão de data");
            }
			
		}

		public static long ToNumV2(this string texto)
		{
			if (string.IsNullOrEmpty(texto))
				return 0;

			return long.Parse(Regex.Replace(texto, "[^0-9]", ""));
		}

		public static int ToNum(this string texto)
		{
			if (string.IsNullOrEmpty(texto))
				return 0;

			return int.Parse(Regex.Replace(texto, "[^0-9]", ""));
		}

		public static bool ToSexo(this string texto, string masculino, string feminino)
		{
			if (string.IsNullOrEmpty(texto))
				return true;

			var sexoLetra = texto.ToLower();

			if (sexoLetra == masculino)
				return true;

			else if (sexoLetra == feminino)
				return false;

			return true;
		}

		public static string? PrimeiraLetraMaiuscula(this string texto)
		{
			if (string.IsNullOrEmpty(texto))
				return null;

			return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(texto.ToLower());
		}

		public static string? ToNome(this string texto)
		{
			texto = texto.GetLetras().GetPrimeirosCaracteres(70).PrimeiraLetraMaiuscula();

			if (texto == null) return null;

			return string.Join(" ", texto.Split(' ')
				.Where(parte => parte.Length > 1)
				.ToArray());
		}


		public static string GetLetras(this string texto)
		{
			return Regex.Replace(texto, @"[^a-zA-Z\?\s]", "").Trim();
			//return Regex.Replace(texto, @"[^\p{L}\s]", "").Trim();
		}

		public static string GerarNomeArquivo(string nomeArquivo, string extensao = ".xlsx")
		{
			var caminhoDoArquivo = Path.Combine(salvarNaPasta, nomeArquivo);

			if (File.Exists(caminhoDoArquivo + extensao))
			{
				int count = 1;
				while (File.Exists($"{caminhoDoArquivo} ({count}){extensao}"))
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

		public static string TratarMensagemErro(string arquivo, string erroMensagem, int indiceLinha, string colunaLetra, string tituloColuna, string celulaValor, string variaveisValor = "")
		{
			var mensagemErro = $"\"{arquivo}\" Falha na linha {indiceLinha}, coluna {colunaLetra}: {tituloColuna}, valor esperado: , valor da célula: \"{celulaValor}\": {erroMensagem}";

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

        public static string ToCidade(this string textoCidade, string uf, string cep = "")
        {
            if (excelHelper == null)
			{
                var arquivoExcelCidades = "Files\\EnderecosCidades.xlsx";
                excelHelper = new ExcelHelper();
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
            }

            if (excelHelper.CidadeExists(cep, textoCidade.ToLower(), uf.ToLower()))
                return textoCidade;
			else
			{
                var procurarCidade = excelHelper.EncontrarCidadeSemelhante(textoCidade.ToLower());

                if (excelHelper.CidadeExists(cep, procurarCidade, uf))
                    return textoCidade;
            }
			
            return textoCidade;
        }

		public static string RemoverAcentos(string texto)
		{
			if (string.IsNullOrEmpty(texto))
				return texto;

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

		public static decimal ArredondarValorV2(this string input)
		{
			if (input.Contains('.'))
				input = input.Replace(".", ",");

			if (decimal.TryParse(input, out decimal valorDecimal))
				return decimal.Parse(Math.Round(valorDecimal, 2).ToString("F2"));

			return 0.00m;
		}

		public static decimal ArredondarValor(this string input)
		{
			if (input.Contains('.') && input.Contains(',') == false)
			{
				input = input.Replace(".", "");
				int posicao = input.Length - 4;
				input = input.Insert(posicao, ",");
			}

			if (decimal.TryParse(input, out decimal valorDecimal))
				return Math.Round(valorDecimal, 2);

			return 0;
		}

		public static LogradouroTipos GetLogradouroTipo(this string texto)
		{
			texto = texto.GetPrimeiroNome();

			if (Enum.TryParse<LogradouroTipos>(texto, true, out LogradouroTipos tipo))
				return tipo;

			return LogradouroTipos.Outros;
		}

		public static string RemoverPrimeiroNome(this string texto)
		{
			texto = texto.Trim();

			string[] nameParts = texto.Split(' ');

			return string.Join(" ", nameParts.Skip(1));
		}

        public static void SalvarConfig()
        {
            File.WriteAllText(arquivoConfig,
                    salvarNaPasta + Environment.NewLine +
                    ultimaPasta + Environment.NewLine +
                    ultimoEstabelecimentoID + Environment.NewLine +
                    ultimoEstabelecimento + Environment.NewLine +
                    ultimoAntigoSistema + Environment.NewLine +
                    ultimoinputDB + Environment.NewLine +
                    ultimoinputDBContratos);
        }

        public static string[] LerConfig()
        {
			if (!File.Exists(arquivoConfig))
				SalvarConfig();

            return File.ReadAllLines(arquivoConfig);
        }
    }
}