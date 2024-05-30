﻿using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Crypto.Operators;
using System.Runtime.ConstrainedExecution;
using System.Text.RegularExpressions;

namespace Migracao.Utils
{
    internal class ExcelHelper
    {
        private ISheet sheet;
		private IWorkbook workbook;
		public List<string> cabecalhos;
		public List<IRow> linhas;

		public static Dictionary<string, string> nomeConsumidorDict = new Dictionary<string, string>();
        public static Dictionary<string, string> cpfConsumidorDict = new Dictionary<string, string>();
        public static Dictionary<string, string> nomeCodConsumidorDict = new Dictionary<string, string>();
		public static Dictionary<string, string> nomePessoaDict = new Dictionary<string, string>();
		public static Dictionary<string, string> cpfPessoaDict = new Dictionary<string, string>();
		public static Dictionary<string, string> cpfFuncionarioDict = new Dictionary<string, string>();
		public static Dictionary<string, string> nomeFuncionarioDict = new Dictionary<string, string>();

		private Dictionary<string, string> cidadeDict = new Dictionary<string, string>();
		private Dictionary<string, string> cidadeEstadoDict = new Dictionary<string, string>();

		private Dictionary<string, string> cpfKeyDict = new Dictionary<string, string>();
		private Dictionary<string, string> nomeKeyDict = new Dictionary<string, string>();
		private Dictionary<string, string> nomesUTF8Dict = new Dictionary<string, string>();

		private Dictionary<string, string> pessoaIDTelefonesDict = new Dictionary<string, string>();
		private Dictionary<string, string> pessoaIDEnderecosDict = new Dictionary<string, string>();
		private Dictionary<string, string> cpfTelefonesDict = new Dictionary<string, string>();
		private Dictionary<string, string> cpfEnderecosDict = new Dictionary<string, string>();
		private Dictionary<string, string> nomeTelefonesDict = new Dictionary<string, string>();
		private Dictionary<string, string> nomeEnderecosDict = new Dictionary<string, string>();

		private Dictionary<string, string> consumidorIDRecebiveisDict = new Dictionary<string, string>();
		private Dictionary<string, string> consumidorIDRecebidosDict = new Dictionary<string, string>();

		public ExcelHelper(string arquivoExcel)
        {
			try
			{
				this.workbook = LerExcel(arquivoExcel);
			}
			catch (Exception ex)
			{
				throw new Exception($"Erro ao ler o arquivo Excel \"{arquivoExcel}\": {ex.Message}");
			}

			this.cabecalhos = GetCabecalhosExcel(workbook);
			this.linhas = GetLinhasExcel(workbook);
		}

		public void InitializeDictionaryRecebiveis(ISheet sheet)
		{
			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);

			int consumidorIDColumnIndex = GetColumnIndex(headerRow, "consumidorid");
			int dataVencimentoColumnIndex = GetColumnIndex(headerRow, "datavencimento");
			int valorOriginalColumnIndex = GetColumnIndex(headerRow, "valororiginal");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string consumidorID = sheet.GetRow(row).GetCell(consumidorIDColumnIndex) != null ? sheet.GetRow(row).GetCell(consumidorIDColumnIndex).ToString() : "";
					string dataVencimento = sheet.GetRow(row).GetCell(dataVencimentoColumnIndex) != null ? sheet.GetRow(row).GetCell(dataVencimentoColumnIndex).ToString().ToLower() : "";
					string valorOriginal = sheet.GetRow(row).GetCell(valorOriginalColumnIndex) != null ? sheet.GetRow(row).GetCell(valorOriginalColumnIndex).ToString().ToLower() : "";

					if (!valorOriginal.Contains('.') && !valorOriginal.Contains(','))
						valorOriginal = valorOriginal.Insert(valorOriginal.Length - 4, ".");

					valorOriginal = Tools.ArredondarValor(valorOriginal).ToString("F2");
					//}

					if (consumidorID == "18283648")
						consumidorID = consumidorID;

					string key = consumidorID + "|" + valorOriginal + "|" + dataVencimento;
					if (!consumidorIDRecebiveisDict.ContainsKey(key))
						consumidorIDRecebiveisDict.Add(key, consumidorID);
				}
			}
		}

		public void InitializeDictionaryRecebidos(ISheet sheet)
		{
			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);

			int consumidorIDColumnIndex = GetColumnIndex(headerRow, "consumidorid");
			int dataColumnIndex = GetColumnIndex(headerRow, "data");
			int pagoValorColumnIndex = GetColumnIndex(headerRow, "pagoValor");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string consumidorID = sheet.GetRow(row).GetCell(consumidorIDColumnIndex) != null ? sheet.GetRow(row).GetCell(consumidorIDColumnIndex).ToString() : "";
					string data = sheet.GetRow(row).GetCell(dataColumnIndex) != null ? sheet.GetRow(row).GetCell(dataColumnIndex).ToString().ToLower() : "";
					string pagoValor = sheet.GetRow(row).GetCell(pagoValorColumnIndex) != null ? sheet.GetRow(row).GetCell(pagoValorColumnIndex).ToString().ToLower() : "";

					if (!pagoValor.Contains('.') && !pagoValor.Contains(','))
					{
						pagoValor = pagoValor.Insert(pagoValor.Length - 4, ".");
						pagoValor = Tools.ArredondarValor(pagoValor).ToString("F2");
					}

					string key = consumidorID + "|" + pagoValor + "|" + data;
					if (!consumidorIDRecebidosDict.ContainsKey(key))
						consumidorIDRecebidosDict.Add(key, consumidorID);
				}
			}
		}

		public void InitializeDictionaryPessoas(ISheet sheet)
		{
			cpfFuncionarioDict.Clear();
			cpfConsumidorDict.Clear();
			cpfPessoaDict.Clear();
			nomeCodConsumidorDict.Clear();
			nomeConsumidorDict.Clear();
			nomePessoaDict.Clear();
			nomeFuncionarioDict.Clear();

			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);

			int cpfColumnIndex = GetColumnIndex(headerRow, "cpf");
			int nomeCompletoColumnIndex = GetColumnIndex(headerRow, "nomecompleto");
			int pessoaidColumnIndex = GetColumnIndex(headerRow, "pessoaid");
			int funcionarioidColumnIndex = GetColumnIndex(headerRow, "funcionarioid");
			int fornecedoridColumnIndex = GetColumnIndex(headerRow, "fornecedorid");
			int nomefantasiaColumnIndex = GetColumnIndex(headerRow, "nomefantasia");
			int consumidoridColumnIndex = GetColumnIndex(headerRow, "consumidorid");
			int codigoantigoColumnIndex = GetColumnIndex(headerRow, "codigoantigo");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string cpf = sheet.GetRow(row).GetCell(cpfColumnIndex) != null ? sheet.GetRow(row).GetCell(cpfColumnIndex).ToString() : "";
					string nomeCompleto = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex) != null ? sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString().ToLower() : "";
					string pessoaid = sheet.GetRow(row).GetCell(pessoaidColumnIndex) != null ? sheet.GetRow(row).GetCell(pessoaidColumnIndex).ToString() : "";
					string funcionarioid = sheet.GetRow(row).GetCell(funcionarioidColumnIndex) != null ? sheet.GetRow(row).GetCell(funcionarioidColumnIndex).ToString() : "";
					string fornecedorid = sheet.GetRow(row).GetCell(fornecedoridColumnIndex) != null ? sheet.GetRow(row).GetCell(fornecedoridColumnIndex).ToString() : "";
					string nomefantasia = sheet.GetRow(row).GetCell(nomefantasiaColumnIndex) != null ? sheet.GetRow(row).GetCell(nomefantasiaColumnIndex).ToString() : "";
					string consumidorid = sheet.GetRow(row).GetCell(consumidoridColumnIndex) != null ? sheet.GetRow(row).GetCell(consumidoridColumnIndex).ToString() : "";
					string codigoantigo = sheet.GetRow(row).GetCell(codigoantigoColumnIndex) != null ? sheet.GetRow(row).GetCell(codigoantigoColumnIndex).ToString() : "";

					nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
					cpf = cpf.Replace(".", "").Replace("-", "");

					string key = cpf;

					if (!cpfConsumidorDict.ContainsKey(key))
						cpfConsumidorDict.Add(key, consumidorid);

					if (!cpfFuncionarioDict.ContainsKey(key))
						cpfFuncionarioDict.Add(key, funcionarioid);

					if (!cpfPessoaDict.ContainsKey(key))
						cpfPessoaDict.Add(key, pessoaid);

					key = nomeCompleto + "|" + codigoantigo;
					if (!nomeCodConsumidorDict.ContainsKey(key))
						nomeCodConsumidorDict.Add(key, consumidorid);

					key = nomeCompleto;

					if (!nomeConsumidorDict.ContainsKey(key))
						nomeConsumidorDict.Add(key, consumidorid);

					if (!nomePessoaDict.ContainsKey(key))
						nomePessoaDict.Add(key, pessoaid);

					if (!nomeFuncionarioDict.ContainsKey(key))
						nomeFuncionarioDict.Add(key, funcionarioid);
				}
			}
		}

		public void InitializeDictionary(ISheet sheet)
		{
			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);

			int cpfColumnIndex = GetColumnIndex(headerRow, "cpf");
			int cepColumnIndex = GetColumnIndex(headerRow, "cep");
			int nomeCompletoColumnIndex = GetColumnIndex(headerRow, "nomecompleto");
			int pessoaidColumnIndex = GetColumnIndex(headerRow, "pessoaid");
			int funcionarioidColumnIndex = GetColumnIndex(headerRow, "funcionarioid");
			int fornecedoridColumnIndex = GetColumnIndex(headerRow, "fornecedorid");
			int nomefantasiaColumnIndex = GetColumnIndex(headerRow, "nomefantasia");
			int consumidoridColumnIndex = GetColumnIndex(headerRow, "consumidorid");
			int codigoantigoColumnIndex = GetColumnIndex(headerRow, "codigoantigo");
			int logradouroColumnIndex = GetColumnIndex(headerRow, "logradouro");
			int telefoneColumnIndex = GetColumnIndex(headerRow, "telefone");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string cpf = sheet.GetRow(row).GetCell(cpfColumnIndex) != null ? sheet.GetRow(row).GetCell(cpfColumnIndex).ToString() : "";
					string cep = sheet.GetRow(row).GetCell(cepColumnIndex) != null ? sheet.GetRow(row).GetCell(cepColumnIndex).ToString() : "";
					string nomeCompleto = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex) != null ? sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString().ToLower() : "";
					string pessoaid = sheet.GetRow(row).GetCell(pessoaidColumnIndex) != null ? sheet.GetRow(row).GetCell(pessoaidColumnIndex).ToString() : "";
					string funcionarioid = sheet.GetRow(row).GetCell(funcionarioidColumnIndex) != null ? sheet.GetRow(row).GetCell(funcionarioidColumnIndex).ToString() : "";
					string fornecedorid = sheet.GetRow(row).GetCell(fornecedoridColumnIndex) != null ? sheet.GetRow(row).GetCell(fornecedoridColumnIndex).ToString() : "";
					string nomefantasia = sheet.GetRow(row).GetCell(nomefantasiaColumnIndex) != null ? sheet.GetRow(row).GetCell(nomefantasiaColumnIndex).ToString() : "";
					string consumidorid = sheet.GetRow(row).GetCell(consumidoridColumnIndex) != null ? sheet.GetRow(row).GetCell(consumidoridColumnIndex).ToString() : "";
					string codigoantigo = sheet.GetRow(row).GetCell(codigoantigoColumnIndex) != null ? sheet.GetRow(row).GetCell(codigoantigoColumnIndex).ToString() : "";
					string logradouro = sheet.GetRow(row).GetCell(logradouroColumnIndex) != null ? sheet.GetRow(row).GetCell(logradouroColumnIndex).ToString() : "";
					string telefone = sheet.GetRow(row).GetCell(telefoneColumnIndex) != null ? sheet.GetRow(row).GetCell(telefoneColumnIndex).ToString() : "";

					nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
					cpf = cpf.Replace(".", "").Replace("-", "");
					logradouro = Tools.RemoverAcentos(logradouro).ToLower();

					string key = cpf;

					if (!cpfConsumidorDict.ContainsKey(key))
						cpfConsumidorDict.Add(key, consumidorid);

					if (!cpfPessoaDict.ContainsKey(key))
						cpfPessoaDict.Add(key, pessoaid);

					if (!cpfFuncionarioDict.ContainsKey(key))
						cpfFuncionarioDict.Add(key, funcionarioid);

					key = cpf + "|" + telefone;
					if (!cpfTelefonesDict.ContainsKey(key))
						cpfTelefonesDict.Add(key, telefone);

					key = pessoaid + "|" + telefone;
					if (!pessoaIDTelefonesDict.ContainsKey(key))
						pessoaIDTelefonesDict.Add(key, telefone);

					key = cpf + "|" + logradouro;
					if (!cpfEnderecosDict.ContainsKey(key))
						cpfEnderecosDict.Add(key, logradouro);

					key = nomeCompleto + "|" + codigoantigo;
					if (!nomeCodConsumidorDict.ContainsKey(key))
						nomeCodConsumidorDict.Add(key, consumidorid);

					key = nomeCompleto;

					if (!nomeConsumidorDict.ContainsKey(key))
						nomeConsumidorDict.Add(key, consumidorid);

					if (!nomePessoaDict.ContainsKey(key))
						nomePessoaDict.Add(key, pessoaid);

					if (!nomeFuncionarioDict.ContainsKey(key))
						nomeFuncionarioDict.Add(key, funcionarioid);

					key = nomeCompleto + "|" + telefone;
					if (!nomeTelefonesDict.ContainsKey(key))
						nomeTelefonesDict.Add(key, funcionarioid);

					key = nomeCompleto + "|" + logradouro;
					if (!nomeEnderecosDict.ContainsKey(key))
						nomeEnderecosDict.Add(key, funcionarioid);

					key = pessoaid + "|" + cep;
					if (!pessoaIDEnderecosDict.ContainsKey(key))
						pessoaIDEnderecosDict.Add(key, consumidorid);
				}
			}
		}

		public void InitializeDictionaryCidade(ISheet sheet)
		{
			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);
			int cidadeIdColumnIndex = GetColumnIndex(headerRow, "id");
			int cidadeColumnIndex = GetColumnIndex(headerRow, "nome");
			int estadoColumnIndex = GetColumnIndex(headerRow, "estado");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string cidadeIdCellValue = sheet.GetRow(row).GetCell(cidadeIdColumnIndex) != null ? sheet.GetRow(row).GetCell(cidadeIdColumnIndex).ToString() : "";
					string cidadeCellValue = sheet.GetRow(row).GetCell(cidadeColumnIndex) != null ? sheet.GetRow(row).GetCell(cidadeColumnIndex).ToString() : "";
					string estadoCellValue = sheet.GetRow(row).GetCell(estadoColumnIndex) != null ? sheet.GetRow(row).GetCell(estadoColumnIndex).ToString() : "";

					string key = Tools.RemoverAcentos(cidadeCellValue).ToLower() + "|" + estadoCellValue.ToLower();
					if (!cidadeEstadoDict.ContainsKey(key))
						cidadeEstadoDict.Add(key, cidadeIdCellValue);
				}
			}
		}

		public void InitializeDictionaryNomesUTF8(ISheet sheet)
		{
			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);
			int nomeErradoColumnIndex = GetColumnIndex(headerRow, "nomeErrado");
			int nomeCorrigidoColumnIndex = GetColumnIndex(headerRow, "nomeCorrigido");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string nomeErradoCellValue = sheet.GetRow(row).GetCell(nomeErradoColumnIndex) != null ? sheet.GetRow(row).GetCell(nomeErradoColumnIndex).ToString() : "";
					string nomeCorrigidoCellValue = sheet.GetRow(row).GetCell(nomeCorrigidoColumnIndex) != null ? sheet.GetRow(row).GetCell(nomeCorrigidoColumnIndex).ToString() : "";

					string key = nomeErradoCellValue.ToLower();
					if (!nomesUTF8Dict.ContainsKey(key))
						nomesUTF8Dict.Add(key, nomeCorrigidoCellValue);
				}
			}
		}

		private int GetColumnIndex(IRow headerRow, string columnName)
        {
            for (int column = 0; column < headerRow.LastCellNum; column++)
            {
                if (headerRow.GetCell(column).ToString().Equals(columnName, StringComparison.CurrentCultureIgnoreCase))
                {
                    return column;
                }
            }
            throw new Exception($"Coluna {columnName} não encontrada");
        }

		public int GetCidadeID(string cidade, string estado)
		{
			cidade = Tools.RemoverAcentos(cidade).ToLower();

			if (!string.IsNullOrWhiteSpace(cidade))
            {
				string key = cidade + "|" + estado.ToLower();
                if (cidadeEstadoDict.ContainsKey(key))
                    return int.Parse(cidadeEstadoDict[key]);

				key = cidade.EncontrarCidadeSemelhante().ToLower() + "|" + estado.ToLower();
				if (cidadeEstadoDict.ContainsKey(key))
					return int.Parse(cidadeEstadoDict[key]);
				
				//key = cidade.ToLower();
				//if (cidadeDict.ContainsKey(key))
				//    return int.Parse(cidadeDict[key]);
			}

			return 0;
		}

		public string GetPessoaID(string cpf = "", string nomeCompleto = "")
		{
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return "";

			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
			cpf = cpf.Replace(".", "").Replace("-", "");

			if (!string.IsNullOrWhiteSpace(cpf))
				if (cpfPessoaDict.ContainsKey(cpf))
					return cpfPessoaDict[cpf];

			if (!string.IsNullOrWhiteSpace(nomeCompleto))
				if (nomePessoaDict.ContainsKey(nomeCompleto))
					return nomePessoaDict[nomeCompleto];

			return "";
		}

		public string GetConsumidorID(string cpf = "", string nomeCompleto = "", string codigo = "")
        {
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return "";

			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
			cpf = cpf.Replace(".", "").Replace("-", "");

			if (!string.IsNullOrWhiteSpace(cpf))
				if (cpfConsumidorDict.ContainsKey(cpf))
					return cpfConsumidorDict[cpf];

			if (!string.IsNullOrWhiteSpace(nomeCompleto) && !string.IsNullOrWhiteSpace(codigo))
			{
				string key = nomeCompleto + "|" + codigo;
				if (nomeCodConsumidorDict.ContainsKey(key))
					return nomeCodConsumidorDict[key];
			}

			if (!string.IsNullOrWhiteSpace(nomeCompleto))
				if (nomeConsumidorDict.ContainsKey(nomeCompleto))
					return nomeConsumidorDict[nomeCompleto];

            return "";
        }

		public bool RecebidoExists(int consumidorID, decimal pagoValor, DateTime data)
		{
			if (consumidorID <= 0)
				return true;

			pagoValor = Tools.ArredondarValor(pagoValor.ToString("F2"));

			string key = consumidorID + "|" + pagoValor + "|" + data.ToString("yyyy-MM-dd HH:mm:ss.fff");

			if (!string.IsNullOrWhiteSpace(key))
				if (consumidorIDRecebidosDict.ContainsKey(key))
					return true;

			return false;
		}

		public bool RecebivelExists(int consumidorID, decimal valorOriginal, DateTime dataVencimento)
		{
			if (consumidorID == 18283648)
				consumidorID = consumidorID;

			string key = consumidorID + "|" + valorOriginal.ToString("F2") + "|" + dataVencimento.ToString("yyyy-MM-dd HH:mm:ss.fff");
			if (!string.IsNullOrWhiteSpace(key))
				if (consumidorIDRecebiveisDict.ContainsKey(key))
					return true;

			return false;
		}

		public bool PessoaFoneExists(int pessoaID, string telefone)
		{
			if (pessoaID <= 0)
				return true;

			string key = pessoaID + "|" + telefone;
			if (!string.IsNullOrWhiteSpace(key))
				if (pessoaIDTelefonesDict.ContainsKey(key))
					return true;

			return false;
		}

		public bool PessoaFoneExists(string cpf = "", string nomeCompleto = "", string telefone = "")
		{
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return true;
			
			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
			cpf = cpf.Replace(".", "").Replace("-", "");

			string key = cpf + "|" + telefone;
			if (!string.IsNullOrWhiteSpace(key))
				if (cpfTelefonesDict.ContainsKey(key))
					return true;

			if (!string.IsNullOrWhiteSpace(nomeCompleto) && !string.IsNullOrWhiteSpace(telefone))
			{
				key = nomeCompleto + "|" + telefone;
				if (nomeTelefonesDict.ContainsKey(key))
					return true;
			}

			return false;
		}

		public bool ConsumidorEnderecoExists(int pessoaID, int cep)
		{
			if (pessoaID <= 0 && cep <= 0)
				return true;

			string key = pessoaID + "|" + cep;
			if (!string.IsNullOrWhiteSpace(key))
				if (pessoaIDEnderecosDict.ContainsKey(key))
					return true;

			return false;
		}

		public bool ConsumidorEnderecoExists(string cpf = "", string nomeCompleto = "", string logradouro = "")
		{
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return true;

			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
			cpf = cpf.Replace(".", "").Replace("-", "");
			logradouro = Tools.RemoverAcentos(logradouro).ToLower();

			string key = cpf + "|" + logradouro;
			if (!string.IsNullOrWhiteSpace(cpf))
				if (cpfEnderecosDict.ContainsKey(key))
					return true;

			if (!string.IsNullOrWhiteSpace(nomeCompleto) && !string.IsNullOrWhiteSpace(logradouro))
			{
				key = nomeCompleto + "|" + logradouro;
				if (nomeEnderecosDict.ContainsKey(key))
					return true;
			}

			return false;
		}

		public string GetFuncionarioID(string cpf = "", string nomeCompleto = "")
		{
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return "";

			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
			cpf = cpf.Replace(".", "").Replace("-", "");

			string key = cpf;
			if (!string.IsNullOrEmpty(cpf))
				if (cpfFuncionarioDict.ContainsKey(key))
					return cpfFuncionarioDict[key];

			key = nomeCompleto;
			if (!string.IsNullOrEmpty(nomeCompleto))
				if (nomeFuncionarioDict.ContainsKey(key))
					return nomeFuncionarioDict[key];

			return "";
		}

		public string GetFornecedorID(string cpf = "", string nomeCompleto = "", string codigo = "")
		{
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return "";

			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
			cpf = cpf.Replace(".", "").Replace("-", "");

			string key = cpf;
			if (cpfConsumidorDict.ContainsKey(key))
				return cpfConsumidorDict[key];

			key = nomeCompleto + "|" + codigo;
			if (nomeCodConsumidorDict.ContainsKey(key))
				return nomeCodConsumidorDict[key];

			key = nomeCompleto;
			if (nomeConsumidorDict.ContainsKey(key))
				return nomeConsumidorDict[key];

			return "";
		}

		public IWorkbook LerExcel(string filePath)
        {
            IWorkbook workbook;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }
            return workbook;
        }

        public List<string> GetCabecalhosExcel(IWorkbook workbook)
        {
            ISheet sheet1 = workbook.GetSheetAt(0);
            IRow headerRow = sheet1.GetRow(0);

            List<string> titulos = new List<string>();
            foreach (ICell cell in headerRow.Cells)
            {
                titulos.Add(cell.ToString());
            }

            return titulos;
        }

        public List<IRow> GetLinhasExcel(IWorkbook workbook)
        {
            ISheet sheet1 = workbook.GetSheetAt(0);
            List<IRow> linhas = new List<IRow>();

            for (int i = 1; i <= sheet1.LastRowNum; i++)
            {
                IRow row = sheet1.GetRow(i);
                if (row != null)
                {
                    linhas.Add(row);
                }
            }

            return linhas;
        }

        public string GetColumnLetter(ICell cell)
        {
            int columnIndex = cell.ColumnIndex;
            int dividend = columnIndex + 1;
            string columnLetter = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnLetter = Convert.ToChar(65 + modulo) + columnLetter;
                dividend = (dividend - modulo) / 26;
            }

            return columnLetter;
        }

        public void GravarExcel(string nomeArquivo, Dictionary<string, object[]> linhas)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Planilha1");

            // Cria a linha de cabeçalho e insere os nomes das colunas
            IRow row = sheet1.CreateRow(0);
            int coluna = 0;
            foreach (var item in linhas)
            {
                ICell cell = row.CreateCell(coluna);
                cell.SetCellValue(item.Key);
                coluna++;
            }

            // Insere os dados nas linhas para cada coluna
            coluna = 0;
            foreach (var item in linhas)
            {
                int linha = 1;
                foreach (var valor in item.Value)
                {
                    row = sheet1.GetRow(linha) ?? sheet1.CreateRow(linha);
                    ICell cell = row.CreateCell(coluna);
                    if (valor == null)
						cell.SetCellValue("null");
                    else
					    cell.SetCellValue(valor.ToString());
                    linha++;
                }
                coluna++;
            }

            FileStream sw = File.Create(nomeArquivo + ".xlsx");
            workbook.Write(sw);
            sw.Close();
        }

		public string CorrigirNomeUTF8(string nome)
		{
			if (nome.Contains('?'))
			{
				//string[] palavras = nome.Split(' ');
				string[] palavras = Regex.Split(nome, @"[\s\(\)/:\-,]");
				for (int i = 0; i < palavras.Length; i++)
				{
					if (nomesUTF8Dict.ContainsKey(palavras[i].ToLower()))
						palavras[i] = nomesUTF8Dict[palavras[i].ToLower()];
				}
				return string.Join(" ", palavras);
			}

			return nome;
		}

		public void CreateExcelFile(string salvarArquivo, List<string> cabecalhos, List<List<string>> dados)
		{
			// Crie um novo livro de trabalho
			IWorkbook workbook = new XSSFWorkbook();

			// Crie uma nova planilha
			ISheet sheet = workbook.CreateSheet("Planilha1");

			// Crie a linha de cabeçalho
			IRow headerRow = sheet.CreateRow(0);
			for (int i = 0; i < cabecalhos.Count; i++)
			{
				ICell headerCell = headerRow.CreateCell(i);
				headerCell.SetCellValue(cabecalhos[i]);
			}

			// Adicione os dados
			for (int i = 0; i < dados.Count; i++)
			{
				IRow dataRow = sheet.CreateRow(i + 1);
				for (int j = 0; j < dados[i].Count; j++)
				{
					ICell dataCell = dataRow.CreateCell(j);
					dataCell.SetCellValue(dados[i][j]);
				}
			}

			// Salve o arquivo Excel
			using (var fileStream = new FileStream(salvarArquivo, FileMode.Create, FileAccess.Write))
			{
				workbook.Write(fileStream);
			}
		}

		public static List<string[]> LerCSV(string filePath, char separador)
		{
			List<string[]> linhas = new List<string[]>();
			using (var reader = new StreamReader(filePath))
			{
				string linha;
				while ((linha = reader.ReadLine()) != null)
				{
					string[] valores = linha.Split(separador); // Assumindo que o separador é ';'
					linhas.Add(valores);
				}
			}
			return linhas;
		}

		// Método para obter os cabeçalhos do CSV
		public static List<string> GetCabecalhosCSV(string filePath, char separador)
		{
			List<string[]> linhas = LerCSV(filePath, separador);
			if (linhas.Count > 0)
			{
				return linhas[0].ToList(); // Primeira linha é o cabeçalho
			}
			return new List<string>();
		}

		public static char DetectarSeparadorCSV(string filePath)
		{
			char[] separadores = { ',', ';', '\t', '|' }; // Separadores comuns

			using (var reader = new StreamReader(filePath))
			{
				string primeiraLinha = reader.ReadLine();

				// Verifica qual separador tem o maior número de ocorrências
				char separadorMaisFrequente = separadores.OrderByDescending(s => primeiraLinha.Count(c => c == s)).First();

				return separadorMaisFrequente;
			}
		}

	}
}
