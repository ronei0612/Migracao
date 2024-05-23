using Migracao.Models;
using NPOI.SS.UserModel;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Migracao.Utils
{
    internal static class Tools
    {
        public static string mascaraCPF = "000.000.000-00";

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
                return email;

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
			var pasta = Environment.ExpandEnvironmentVariables("%userprofile%\\Desktop");
			var caminhoDoArquivo = Path.Combine(pasta, nomeArquivo);

            if (File.Exists(caminhoDoArquivo + ".xlsx")) {
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
	}
}
