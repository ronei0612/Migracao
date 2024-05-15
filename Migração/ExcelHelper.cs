using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Migração
{
	internal class ExcelHelper
	{
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

			// Salvando o arquivo
			using (FileStream stream = new FileStream(nomeArquivo + ".xlsx", FileMode.Create, FileAccess.Write))
			{
				workbook.Write(stream);
			}
		}

		public void GravarExcel(string nomeArquivo, Dictionary<string, object[]> linhas)
		{
			IWorkbook workbook = new XSSFWorkbook();  // Cria um novo workbook
			ISheet sheet1 = workbook.CreateSheet("Planilha1");  // Cria uma nova planilha

			// Insere os dados do dicionário na planilha
			int coluna = 0;
			foreach (var item in linhas)
			{
				// Insere o nome da coluna
				IRow row = sheet1.CreateRow(0);
				ICell cell = row.CreateCell(coluna);
				cell.SetCellValue(item.Key);

				// Insere os dados da coluna
				int linha = 1;
				foreach (var valor in item.Value)
				{
					row = sheet1.GetRow(linha) ?? sheet1.CreateRow(linha);
					cell = row.CreateCell(coluna);
					if (valor != null)
						cell.SetCellValue(valor.ToString());
					linha++;
				}

				coluna++;
			}

			// Salva o arquivo Excel
			FileStream sw = File.Create(nomeArquivo + ".xlsx");
			workbook.Write(sw);
			sw.Close();
		}
	}
}
