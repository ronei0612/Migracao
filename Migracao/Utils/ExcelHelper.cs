using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Migracao.Utils
{
    internal class ExcelHelper
    {
        private ISheet sheet;
		private IWorkbook workbook;
		public List<string> cabecalhos;
		public List<IRow> linhas;

		private Dictionary<string, string> nomeDict = new Dictionary<string, string>();
        private Dictionary<string, string> cpfDict = new Dictionary<string, string>();
        private Dictionary<string, string> nomeCodDict = new Dictionary<string, string>();

		private Dictionary<string, string> cidadeDict = new Dictionary<string, string>();
		private Dictionary<string, string> cidadeEstadoDict = new Dictionary<string, string>();

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

					string key = cidadeCellValue;
					if (!cidadeDict.ContainsKey(key))
						cidadeDict.Add(key, cidadeIdCellValue);

					key = cidadeCellValue + "|" + estadoCellValue;
					if (!cidadeEstadoDict.ContainsKey(key))
						cidadeEstadoDict.Add(key, cidadeIdCellValue);
				}
			}
		}

		public void InitializeDictionaryConsumidor(ISheet sheet)
        {
            this.sheet = sheet;
            IRow headerRow = sheet.GetRow(0);
            int cpfColumnIndex = GetColumnIndex(headerRow, "cpf");
            int nomeCompletoColumnIndex = GetColumnIndex(headerRow, "nomecompleto");
            int codigoColumnIndex = GetColumnIndex(headerRow, "codigoantigo");
            int consumidorColumnIndex = GetColumnIndex(headerRow, "consumidorid");

            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                if (sheet.GetRow(row) != null)
                {
                    string cpfCellValue = sheet.GetRow(row).GetCell(cpfColumnIndex) != null ? sheet.GetRow(row).GetCell(cpfColumnIndex).ToString() : "";
                    string nomeCompletoCellValue = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex) != null ? sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString() : "";
                    string codigoCellValue = sheet.GetRow(row).GetCell(codigoColumnIndex) != null ? sheet.GetRow(row).GetCell(codigoColumnIndex).ToString() : "";
                    string consumidorIdCellValue = sheet.GetRow(row).GetCell(consumidorColumnIndex) != null ? sheet.GetRow(row).GetCell(consumidorColumnIndex).ToString() : "";

                    string key = cpfCellValue;
                    if (!cpfDict.ContainsKey(key))
                        cpfDict.Add(key, consumidorIdCellValue);

                    key = nomeCompletoCellValue + "|" + codigoCellValue;
                    if (!nomeCodDict.ContainsKey(key))
                        nomeCodDict.Add(key, consumidorIdCellValue);

                    key = nomeCompletoCellValue;
                    if (!nomeDict.ContainsKey(key))
                        nomeDict.Add(key, consumidorIdCellValue);
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

		public string GetCidadeID(string cidade, string estado = "")
		{
			string key = cidade + "|" + estado;
			if (cidadeEstadoDict.ContainsKey(key))
				return cidadeEstadoDict[key];

			key = cidade;
			if (cidadeDict.ContainsKey(key))
				return cidadeDict[key];

			return "";
		}

		public string GetPessoaID(string cpf = "", string nomeCompleto = "")
		{
			string key = cpf;
			if (cpfDict.ContainsKey(key))
				return cpfDict[key];

			key = nomeCompleto;
			if (nomeDict.ContainsKey(key))
				return nomeDict[key];

			return "";
		}

		public string GetConsumidorID(string cpf = "", string nomeCompleto = "", string codigo = "")
        {
            string key = cpf;
            if (cpfDict.ContainsKey(key))
                return cpfDict[key];

            key = nomeCompleto + "|" + codigo;
            if (nomeCodDict.ContainsKey(key))
                return nomeCodDict[key];

            key = nomeCompleto;
            if (nomeDict.ContainsKey(key))
                return nomeDict[key];

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
                    if (valor != null)
                        cell.SetCellValue(valor.ToString());
                    linha++;
                }
                coluna++;
            }

            FileStream sw = File.Create(nomeArquivo + ".xlsx");
            workbook.Write(sw);
            sw.Close();
        }
    }
}
