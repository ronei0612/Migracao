using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Migração
{
	internal class ExcelHelper
	{
		int nomeCompletoColumnIndex = -1;
		int codigoColumnIndex = -1;
		int cpfColumnIndex = -1;
		int consumidorColumnIndex = -1;

		public void ZerarVariaveis()
		{
			nomeCompletoColumnIndex = -1;
			codigoColumnIndex = -1;
			cpfColumnIndex = -1;
			consumidorColumnIndex = -1;
		}

		public int GetConsumidorID(ISheet sheet, string cpf = "", string nomeCompleto = "", string codigo = "")
		{
			IRow headerRow = sheet.GetRow(0); // assumindo que o cabeçalho está na primeira linha
			int retorno = 0;

			if (consumidorColumnIndex == -1)
			{
				for (int column = 0; column < headerRow.LastCellNum; column++)
					if (headerRow.GetCell(column).ToString().Equals("consumidorid", StringComparison.CurrentCultureIgnoreCase))
					{
						consumidorColumnIndex = column;
						break;
					}

				if (consumidorColumnIndex == -1)
					throw new Exception("Coluna ConsumidorID não encontrada");
			}

			if (!string.IsNullOrWhiteSpace(cpf))
			{
				if (cpfColumnIndex == -1)
				{
					for (int column = 0; column < headerRow.LastCellNum; column++)
						if (headerRow.GetCell(column).ToString().Equals("cpf", StringComparison.CurrentCultureIgnoreCase))
						{
							cpfColumnIndex = column;
							break;
						}

					if (cpfColumnIndex == -1)
						throw new Exception("Coluna CPF não encontrada");
				}

				for (int row = 1; row <= sheet.LastRowNum; row++) // começa em 1 para pular o cabeçalho
					if (sheet.GetRow(row) != null) // verifica se a linha não está vazia
					{
						string cpfCellValue = sheet.GetRow(row).GetCell(cpfColumnIndex).ToString();
						if (cpfCellValue == cpf)
							retorno = int.Parse(sheet.GetRow(row).GetCell(consumidorColumnIndex).ToString()); // retorna o ConsumidorID se o CPF corresponder
					}
			}

			if (retorno == 0 && !string.IsNullOrWhiteSpace(nomeCompleto) && !string.IsNullOrWhiteSpace(codigo))
			{
				if (nomeCompletoColumnIndex == -1)
				{
					for (int column = 0; column < headerRow.LastCellNum; column++)
						if (headerRow.GetCell(column).ToString().Equals("nomecompleto", StringComparison.CurrentCultureIgnoreCase))
						{
							nomeCompletoColumnIndex = column;
							break;
						}

					if (nomeCompletoColumnIndex == -1)
						throw new Exception("Coluna NomeCompleto não encontrada");
				}

				if (codigoColumnIndex == -1)
				{
					for (int column = 0; column < headerRow.LastCellNum; column++)
						if (headerRow.GetCell(column).ToString().Equals("codigoantigo", StringComparison.CurrentCultureIgnoreCase))
						{
							codigoColumnIndex = column;
							break;
						}

					if (codigoColumnIndex == -1)
						throw new Exception("Coluna CodigoAntigo não encontrada");
				}

				for (int row = 1; row <= sheet.LastRowNum; row++) // começa em 1 para pular o cabeçalho
					if (sheet.GetRow(row) != null) // verifica se a linha não está vazia
					{
						string nomeCompletoCellValue = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString();
						string codigoCellValue = sheet.GetRow(row).GetCell(nomeCompletoColumnIndex).ToString();

						if (nomeCompletoCellValue == nomeCompleto && codigoCellValue == codigo)
							retorno = int.Parse(sheet.GetRow(row).GetCell(consumidorColumnIndex).ToString()); // retorna o ConsumidorID se o NomeCompleto e codigo corresponderem
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
							retorno = int.Parse(sheet.GetRow(row).GetCell(consumidorColumnIndex).ToString()); // retorna o ConsumidorID se o CPF corresponder
					}
			}

			return retorno;
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
		//IWorkbook workbook = LerExcel(filePath);
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

		public void GravarExcel1(string nomeArquivo, Dictionary<string, object[]> linhas)
		{
			// Criando um novo arquivo Excel
			IWorkbook workbook = new XSSFWorkbook();
			ISheet sheet = workbook.CreateSheet("Dados");

			// Escrevendo cabeçalhos
			IRow headerRow = sheet.CreateRow(0);
			//for (int i = 0; i < cabecalhos.Count; i++)
			//{
			//	headerRow.CreateCell(i).SetCellValue(cabecalhos[i]);
			//}

			var cabecalhos = new List<string>(linhas.Keys);
			for (int i = 0; i < cabecalhos.Count; i++)
			{
				headerRow.CreateCell(i).SetCellValue(cabecalhos[i]);
			}

			// Escrevendo dados
			int rowIndex = 1;
			foreach (var linha in linhas)
			{
				IRow row = sheet.CreateRow(rowIndex++);
				for (int i = 0; i < linha.Value.Length; i++)
				{
					if (linha.Value[i] != null)
						row.CreateCell(i).SetCellValue(linha.Value[i].ToString());
				}
			}

			using (FileStream stream = new FileStream(nomeArquivo + ".xlsx", FileMode.Create, FileAccess.Write))
			{
				workbook.Write(stream);
			}
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
