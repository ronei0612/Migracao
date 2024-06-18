using Migracao.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Text;
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
		public static Dictionary<string, string> nomeNascimentoPessoaDict = new Dictionary<string, string>();
		public static Dictionary<string, string> cpfPessoaDict = new Dictionary<string, string>();
		public static Dictionary<string, string> cpfFuncionarioDict = new Dictionary<string, string>();
		public static Dictionary<string, string> nomeFuncionarioDict = new Dictionary<string, string>();

		private Dictionary<string, string> cidadeDict = new Dictionary<string, string>();
		private Dictionary<string, string> cepEstadoDict = new Dictionary<string, string>();
		private Dictionary<string, string> cidadeEstadoDict = new Dictionary<string, string>();

		private Dictionary<string, string> cpfKeyDict = new Dictionary<string, string>();
		private Dictionary<string, string> nomeKeyDict = new Dictionary<string, string>();
		private Dictionary<string, string> nomesUTF8Dict = new Dictionary<string, string>();

		public static Dictionary<string, string> pessoaIDTelefonesDict = new Dictionary<string, string>();
		public static Dictionary<string, string> pessoaIDEnderecosDict = new Dictionary<string, string>();
		public static Dictionary<string, string> cpfTelefonesDict = new Dictionary<string, string>();
		public static Dictionary<string, string> cpfEnderecosDict = new Dictionary<string, string>();
		public static Dictionary<string, string> nomeTelefonesDict = new Dictionary<string, string>();
		public static Dictionary<string, string> nomeEnderecosDict = new Dictionary<string, string>();

		public static Dictionary<string, string> consumidorIDRecebiveisDict = new Dictionary<string, string>();
		public static Dictionary<string, string> consumidorIDRecebidosDict = new Dictionary<string, string>();

		private Dictionary<string, string> pessoaIDDataAgendaDict = new Dictionary<string, string>();
		private Dictionary<string, string> tituloDataAgendaDict = new Dictionary<string, string>();
		private Dictionary<string, string> nomeDataAgendaDict = new Dictionary<string, string>();

		public ExcelHelper(string? arquivoExcel = null)
        {
			if (!string.IsNullOrEmpty(arquivoExcel))
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
		}

		public void InitializeDictionaryRecebiveis(ISheet sheet)
		{
			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);

			int consumidorIDColumnIndex = GetColumnIndex(headerRow, "consumidorid");
			int dataVencimentoColumnIndex = GetColumnIndex(headerRow, "datavencimento");
			int valorOriginalColumnIndex = GetColumnIndex(headerRow, "valororiginal");
			int pagoValorColumnIndex = GetColumnIndex(headerRow, "pagoValor");
			int dataBaixaColumnIndex = GetColumnIndex(headerRow, "databaixa");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string consumidorID = sheet.GetRow(row).GetCell(consumidorIDColumnIndex) != null ? sheet.GetRow(row).GetCell(consumidorIDColumnIndex).ToString() : "";
					string dataVencimento = sheet.GetRow(row).GetCell(dataVencimentoColumnIndex) != null ? sheet.GetRow(row).GetCell(dataVencimentoColumnIndex).ToString().ToLower() : "";
					string valorOriginal = sheet.GetRow(row).GetCell(valorOriginalColumnIndex) != null ? sheet.GetRow(row).GetCell(valorOriginalColumnIndex).ToString().ToLower() : "";
					string pagoValor = sheet.GetRow(row).GetCell(pagoValorColumnIndex) != null ? sheet.GetRow(row).GetCell(pagoValorColumnIndex).ToString().ToLower() : "";
					string dataBaixa = sheet.GetRow(row).GetCell(dataBaixaColumnIndex) != null ? sheet.GetRow(row).GetCell(dataBaixaColumnIndex).ToString().ToLower() : "";

					if (!valorOriginal.Contains('.') && !valorOriginal.Contains(','))
						valorOriginal = valorOriginal.Insert(valorOriginal.Length - 4, ".");

					valorOriginal = Tools.ArredondarValorV2(valorOriginal).ToString("F2");
					pagoValor = Tools.ArredondarValorV2(pagoValor).ToString("F2");

					string key = consumidorID + "|" + valorOriginal + "|" + dataVencimento;
					if (!consumidorIDRecebiveisDict.ContainsKey(key))
						consumidorIDRecebiveisDict.Add(key, consumidorID);

					key = consumidorID + "|" + pagoValor + "|" + dataBaixa;
					if (!consumidorIDRecebidosDict.ContainsKey(key))
						consumidorIDRecebidosDict.Add(key, consumidorID);
				}
			}
		}

		public void InitializeDictionaryAgendamentos(ISheet sheet)
		{
			this.sheet = sheet;
			IRow headerRow = sheet.GetRow(0);

			int dataInicioColumnIndex = GetColumnIndex(headerRow, "datainicio");
			int tituloColumnIndex = GetColumnIndex(headerRow, "titulo");
			int consumidoridColumnIndex = GetColumnIndex(headerRow, "consumidorid");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string dataInicio = sheet.GetRow(row).GetCell(dataInicioColumnIndex) != null ? sheet.GetRow(row).GetCell(dataInicioColumnIndex).ToString() : "";
					string titulo = sheet.GetRow(row).GetCell(tituloColumnIndex) != null ? sheet.GetRow(row).GetCell(tituloColumnIndex).ToString().ToLower() : "";
					string consumidorid = sheet.GetRow(row).GetCell(consumidoridColumnIndex) != null ? sheet.GetRow(row).GetCell(consumidoridColumnIndex).ToString() : "";

					string key = titulo + "|" + dataInicio;
					if (!tituloDataAgendaDict.ContainsKey(key))
						tituloDataAgendaDict.Add(key, consumidorid);

					if (!string.IsNullOrEmpty(consumidorid))
					{
						key = consumidorid + "|" + dataInicio;
						if (!pessoaIDDataAgendaDict.ContainsKey(key))
							pessoaIDDataAgendaDict.Add(key, consumidorid);
					}
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
			cpfTelefonesDict.Clear();
			nomePessoaDict.Clear();
			nomeFuncionarioDict.Clear();
			nomeNascimentoPessoaDict.Clear();
			pessoaIDTelefonesDict.Clear();
			cpfEnderecosDict.Clear();
			nomeTelefonesDict.Clear();
			nomeEnderecosDict.Clear();
			pessoaIDEnderecosDict.Clear();

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
			int nascimentoDataColumnIndex = GetColumnIndex(headerRow, "nascimentodata");

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
					string nascimentoData = sheet.GetRow(row).GetCell(nascimentoDataColumnIndex) != null ? sheet.GetRow(row).GetCell(nascimentoDataColumnIndex).ToString().ToLower() : "";

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

					key = nomeCompleto + "|" + cep;
					if (!nomeEnderecosDict.ContainsKey(key))
						nomeEnderecosDict.Add(key, funcionarioid);

					key = pessoaid + "|" + cep;
					if (!pessoaIDEnderecosDict.ContainsKey(key))
						pessoaIDEnderecosDict.Add(key, consumidorid);

					key = nomeCompleto + "|" + nascimentoData;
					if (!nomeNascimentoPessoaDict.ContainsKey(key))
						nomeNascimentoPessoaDict.Add(key, pessoaid);
				}
			}
		}

		public string ProcurarCelula(ISheet sheet, string coluna, string texto, string colunaRetorno)
		{
			if (sheet == null)
				return "";

			int columnIndex = sheet.GetRow(0)
				.Cells
				.FirstOrDefault(c => c.StringCellValue.Equals(coluna, StringComparison.OrdinalIgnoreCase))
				?.ColumnIndex ?? -1;

			int columnRetorno = sheet.GetRow(0)
				.Cells
				.FirstOrDefault(c => c.StringCellValue.Equals(colunaRetorno, StringComparison.OrdinalIgnoreCase))
				?.ColumnIndex ?? -1;

			if (columnIndex == -1)
				throw new Exception($"Coluna {coluna} não encontrada");

			if (columnRetorno == -1)
				throw new Exception($"Coluna {colunaRetorno} não encontrada");

			for (int rowIdx = 1; rowIdx <= sheet.LastRowNum; rowIdx++)
			{
				IRow row = sheet.GetRow(rowIdx);
				ICell cell = row.GetCell(columnIndex);
				ICell cellRetorno = row.GetCell(columnRetorno);

				if (cell != null && cell.CellType != CellType.Blank && cell.StringCellValue.Equals(texto, StringComparison.OrdinalIgnoreCase))
					return cellRetorno.ToString();
			}

			return "";
		}

		public bool ExisteTexto(ISheet sheet, string coluna, string texto)
		{
			if (sheet == null)
				return false;

			int columnIndex = sheet.GetRow(0)
				.Cells
				.FirstOrDefault(c => c.StringCellValue.Equals(coluna, StringComparison.OrdinalIgnoreCase))
				?.ColumnIndex ?? -1;

			if (columnIndex == -1)
				throw new Exception($"Coluna {coluna} não encontrada");

			for (int rowIdx = 1; rowIdx <= sheet.LastRowNum; rowIdx++)
			{
				IRow row = sheet.GetRow(rowIdx);
				ICell cell = row.GetCell(columnIndex);

				if (cell != null && cell.CellType != CellType.Blank && cell.StringCellValue.Equals(texto, StringComparison.OrdinalIgnoreCase))
					return true;
			}

			return false;
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

					key = nomeCompleto + "|" + cep;
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
			int cepColumnIndex = GetColumnIndex(headerRow, "cep");

			for (int row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null)
				{
					string cidadeIdCellValue = sheet.GetRow(row).GetCell(cidadeIdColumnIndex) != null ? sheet.GetRow(row).GetCell(cidadeIdColumnIndex).ToString() : "";
					string cidadeCellValue = sheet.GetRow(row).GetCell(cidadeColumnIndex) != null ? sheet.GetRow(row).GetCell(cidadeColumnIndex).ToString() : "";
					string estadoCellValue = sheet.GetRow(row).GetCell(estadoColumnIndex) != null ? sheet.GetRow(row).GetCell(estadoColumnIndex).ToString() : "";
					string cepCellValue = sheet.GetRow(row).GetCell(cepColumnIndex) != null ? sheet.GetRow(row).GetCell(cepColumnIndex).ToString() : "";

					string key = Tools.RemoverAcentos(cidadeCellValue).ToLower() + "|" + estadoCellValue.ToLower();
					if (!cidadeEstadoDict.ContainsKey(key))
						cidadeEstadoDict.Add(key, cidadeIdCellValue);

					if (cepCellValue != "0" && !string.IsNullOrEmpty(cepCellValue))
					{
						key = cepCellValue + "|" + estadoCellValue.ToLower();
						if (!cepEstadoDict.ContainsKey(key))
							cepEstadoDict.Add(key, cidadeIdCellValue);
					}
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

		public bool CidadeExists(string cidade, string estado)
		{
			if (string.IsNullOrEmpty(cidade))
				return true;

			cidade = Tools.RemoverAcentos(cidade).ToLower();

			if (!string.IsNullOrWhiteSpace(cidade))
			{
				string key = cidade + "|" + estado.ToLower();
				if (cidadeEstadoDict.ContainsKey(key))
					return true;
			}

			return false;
		}

		public bool AgendamentoExists(string titulo, DateTime dataConsulta, int? consumidorID = null)
		{
			if (consumidorID == null && string.IsNullOrWhiteSpace(titulo))
				return false;

			if (consumidorID > 0)
			{
				string key = consumidorID + "|" + dataConsulta.ToString("yyyy-MM-dd HH:mm:ss.fff");
				if (pessoaIDDataAgendaDict.ContainsKey(key))
					return true;
			}

			if (!string.IsNullOrWhiteSpace(titulo))
			{
				string key = titulo.ToLower() + "|" + dataConsulta.ToString("yyyy-MM-dd HH:mm:ss.fff");
				if (tituloDataAgendaDict.ContainsKey(key))
					return true;
			}

			return false;
		}

		public int GetCidadeID(string cidade, string estado, int cep = 0)
		{
			cidade = Tools.RemoverAcentos(cidade).ToLower();

			if (!string.IsNullOrWhiteSpace(cidade))
            {
				string key = "";

				if (!string.IsNullOrEmpty(estado) && cep > 0) {
					string cepFormat = cep.ToString().Substring(0, 5) + "000";

					key = cepFormat + "|" + estado.ToLower();
					if (cepEstadoDict.ContainsKey(key))
						return int.Parse(cepEstadoDict[key]);
				}

				if (string.IsNullOrEmpty(estado))
				{
					if (cidadeDict.ContainsKey(cidade))
						return int.Parse(cidadeDict[cidade]);
					else
						return 0;
				}

				key = cidade + "|" + estado.ToLower();
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

		public string GetPessoaID(string cpf = "", string nomeCompleto = "", DateTime? nascimentoData = null)
		{
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return "";

			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();

			if (!string.IsNullOrWhiteSpace(cpf))
			{
				cpf = cpf.Replace(".", "").Replace("-", "");
				if (cpfPessoaDict.ContainsKey(cpf))
					return cpfPessoaDict[cpf];
			}

			if (!string.IsNullOrWhiteSpace(nomeCompleto))
				if (nomePessoaDict.ContainsKey(nomeCompleto))
					return nomePessoaDict[nomeCompleto];

			if (!string.IsNullOrWhiteSpace(nomeCompleto) && nascimentoData != null)
			{
				var nascimentoD = (DateTime)nascimentoData;
				string key = nomeCompleto + "|" + nascimentoD.ToString("yyyy-MM-dd HH:mm:ss.fff");
				if (nomeNascimentoPessoaDict.ContainsKey(key))
					return nomeNascimentoPessoaDict[key];
			}

			return "";
		}

		public string GetConsumidorID(string cpf = "", string nomeCompleto = "", string codigo = "")
        {
			if (string.IsNullOrWhiteSpace(cpf) && string.IsNullOrWhiteSpace(nomeCompleto))
				return "";

			if (!string.IsNullOrWhiteSpace(cpf))
			{
				cpf = cpf.Replace(".", "").Replace("-", "");
				if (cpfConsumidorDict.ContainsKey(cpf))
					return cpfConsumidorDict[cpf];
			}

			if (!string.IsNullOrWhiteSpace(nomeCompleto) && !string.IsNullOrWhiteSpace(codigo))
			{
				nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();

				string key = nomeCompleto + "|" + codigo;
				if (nomeCodConsumidorDict.ContainsKey(key))
					return nomeCodConsumidorDict[key];
			}

			if (!string.IsNullOrWhiteSpace(nomeCompleto))
			{
				nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();
				if (nomeConsumidorDict.ContainsKey(nomeCompleto))
					return nomeConsumidorDict[nomeCompleto];
			}

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

		public bool RecebivelExists(int consumidorID, decimal valorOriginal, DateTime dataVencimento, decimal pagoValor = 0, DateTime? dataBaixa = null)
		{
			if (pagoValor > 0 && dataBaixa != null)
			{
				var dataBaixa_ = (DateTime)dataBaixa;
				string key = consumidorID + "|" + pagoValor.ToString("F2") + "|" + dataBaixa_.ToString("yyyy-MM-dd HH:mm:ss.fff");
				if (!string.IsNullOrWhiteSpace(key))
					if (consumidorIDRecebidosDict.ContainsKey(key))
						return true;
			}

			else
			{
				string key = consumidorID + "|" + valorOriginal.ToString("F2") + "|" + dataVencimento.ToString("yyyy-MM-dd HH:mm:ss.fff");
				if (!string.IsNullOrWhiteSpace(key))
					if (consumidorIDRecebiveisDict.ContainsKey(key))
						return true;
			}

			return false;
		}

		public bool PessoaFoneExists(int pessoaID, string telefone)
		{
			if (pessoaID <= 0)
				return false;

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

		public bool ConsumidorEnderecoExists(string nomeCompleto = "", int? cep = null)
		{
			if (cep == null && string.IsNullOrWhiteSpace(nomeCompleto))
				return true;

			nomeCompleto = Tools.RemoverAcentos(nomeCompleto).ToLower();

			if (!string.IsNullOrWhiteSpace(nomeCompleto) && cep != null)
			{
				string key = nomeCompleto + "|" + cep.ToString();
				if (nomeEnderecosDict.ContainsKey(key))
					return true;
			}

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
			if (Path.GetExtension(filePath).ToLower() == ".csv")
				return LerExcelCSV(filePath);

			IWorkbook workbook;
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }
            return workbook;
        }

		public static IWorkbook LerExcelCSV(string caminhoCSV)
		{
			var separador = DetectarSeparadorCSV(caminhoCSV);
			List<string> cabecalhosCSV = GetCabecalhosCSV(caminhoCSV, separador, Encoding.UTF8);
			List<string[]> linhasCSV = GetLinhasCSV(caminhoCSV, separador, cabecalhosCSV.Count(), Encoding.UTF8);

			// Criar o workbook e a planilha
			IWorkbook workbook = new XSSFWorkbook();
			ISheet sheet = workbook.CreateSheet("Planilha1");

			// Adicionar os cabeçalhos na primeira linha da planilha
			int linha = 0;
			IRow headerRow = sheet.CreateRow(linha);
			for (int coluna = 0; coluna < cabecalhosCSV.Count; coluna++)
			{
				ICell cell = headerRow.CreateCell(coluna);
				cell.SetCellValue(cabecalhosCSV[coluna]);
			}

			// Adicionar os dados do CSV na planilha
			linha++; // Avançar para a próxima linha para os dados
			int numeroColunas = cabecalhosCSV.Count();

			foreach (var registro in linhasCSV)
			{
				IRow row = sheet.CreateRow(linha);
				for (int coluna = 0; coluna < numeroColunas; coluna++)
				{
					ICell cell = row.CreateCell(coluna);
					cell.SetCellValue(registro[coluna]);
				}
				linha++;
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

		public void CriarExcelArquivo(string nomeArquivo, DataTable dataTable)
		{
			IWorkbook workbook = new XSSFWorkbook();
			ISheet sheet = workbook.CreateSheet("Planilha1");

			// Adiciona os nomes das colunas ao arquivo Excel
			IRow headerRow = sheet.CreateRow(0);
			for (int j = 0; j < dataTable.Columns.Count; j++)
			{
				ICell cell = headerRow.CreateCell(j);
				cell.SetCellValue(dataTable.Columns[j].ColumnName);
			}

            // Adiciona o DataTable ao arquivo Excel
            for (int i = 0; i < dataTable.Rows.Count; i++)
			{
				IRow row = sheet.CreateRow(i + 1); // Começa na segunda linha, pois a primeira linha é para os nomes das colunas
				for (int j = 0; j < dataTable.Columns.Count; j++)
				{
					ICell cell = row.CreateCell(j);
					cell.SetCellValue(dataTable.Rows[i][j].ToString());
				}
			}

			FileStream sw = File.Create(nomeArquivo);
			workbook.Write(sw);
			sw.Close();
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

		public static List<string[]> LerCSV(string filePath, char separador, Encoding encoding)
		{
			var linhas = new List<string[]>();
			using (var reader = new StreamReader(filePath, encoding))
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

		public static List<string[]> GetLinhasCSV(string filePath, char separador, int cabecalhos, Encoding encoding)
		{
			var linhas = new List<string[]>();

			using (var reader = new StreamReader(filePath, encoding))
			{
				// Ignora a primeira linha (cabeçalho)
				reader.ReadLine();

				string linha;
				List<string> valoresTemp = new List<string>();

				while ((linha = reader.ReadLine()) != null)
				{
					var valores = linha.Split(separador);

					//Remover o primeiro elemento quando for quebra de linha
					if (valoresTemp.Count() > 0)
						valores = valores.Skip(1).ToArray();

					// Remover aspas duplas de cada valor na linha
					for (int i = 0; i < valores.Length; i++)
						valores[i] = valores[i].Replace("\"", "");

					valoresTemp.AddRange(valores);

					// Se a quantidade de valores for igual à quantidade de cabeçalhos, adicione à lista de linhas
					if (valoresTemp.Count >= cabecalhos)
					{
						linhas.Add(valoresTemp.ToArray());
						valoresTemp.Clear();
					}
				}
			}
			return linhas;
		}


		// Método para obter os cabeçalhos do CSV
		public static List<string> GetCabecalhosCSV(string filePath, char separador, Encoding encoding)
		{
			List<string[]> linhas = LerCSV(filePath, separador, encoding);
			if (linhas.Count > 0)
			{
				return linhas[0].Select(cabecalho => cabecalho.Replace("\"", "")).ToList(); // Remove aspas duplas e pega a Primeira linha que é o cabeçalho
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

		public static string GerarSqlScriptPessoas(List<Pessoa> pessoas, List<Consumidor> consumidores, List<Endereco> enderecos)
		{
			StringBuilder sb = new StringBuilder();

			// Iterate through the list of people and generate INSERT statements for Pessoas table
			foreach (Pessoa pessoa in pessoas)
			{
				sb.AppendLine($"INSERT INTO Pessoas (NomeCompleto, Apelido) VALUES ('{pessoa.NomeCompleto}', '{pessoa.Apelido}');");
				sb.AppendLine("DECLARE @PessoaID int;");
				sb.AppendLine("SELECT @PessoaID = SCOPE_IDENTITY();");
			}

			// Iterate through the list of consumers and generate INSERT statements for Consumidores table
			foreach (Consumidor consumidor in consumidores)
			{
				sb.AppendLine($"INSERT INTO Consumidores (Ativo, DataInclusao, EstabelecimentoID, LGPDSituacaoID, LoginID, PessoaID) VALUES ('{consumidor.Ativo}', '{consumidor.DataInclusao:yyyy-MM-dd HH:mm:ss.fff}', '{consumidor.EstabelecimentoID}', '{consumidor.LGPDSituacaoID}', '{consumidor.LoginID}', @PessoaID{consumidores.IndexOf(consumidor) + 1});");
			}

			return sb.ToString();
		}
	}
}
