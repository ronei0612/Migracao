using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Migração.Helpers
{
    internal class ExcelHelper
    {
        private ISheet sheet;
        //private Dictionary<string, string> consumidorIdDict = new Dictionary<string, string>();
        private Dictionary<string, string> nomeDict = new Dictionary<string, string>();
        private Dictionary<string, string> cpfDict = new Dictionary<string, string>();
        private Dictionary<string, string> nomeCodDict = new Dictionary<string, string>();

        public void InitializeDictionary(ISheet sheet)
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

                    //string key = cpfCellValue + "|" + nomeCompletoCellValue + "|" + codigoCellValue;
                    //if (!consumidorIdDict.ContainsKey(key))
                    //	consumidorIdDict.Add(key, consumidorIdCellValue);

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
