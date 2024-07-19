using ClosedXML.Excel;
using Migracao.Models;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Text;

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

                    //string key = Tools.RemoverAcentos(cidadeCellValue).ToLower() + "|" + estadoCellValue.ToLower();
                    string key = cidadeCellValue.ToLower() + "|" + estadoCellValue.ToLower();
                    if (!cidadeEstadoDict.ContainsKey(key))
                        cidadeEstadoDict.Add(key, cidadeIdCellValue);

                    key = Tools.RemoverAcentos(cidadeCellValue).ToLower();
                    if (!cidadeDict.ContainsKey(key))
                        cidadeDict.Add(key, cidadeCellValue);

                    if (cepCellValue != "0" && !string.IsNullOrEmpty(cepCellValue))
                    {
                        key = cepCellValue.Substring(0, 5) + "|" + estadoCellValue.ToLower();
                        if (!cepEstadoDict.ContainsKey(key))
                            cepEstadoDict.Add(key, cidadeIdCellValue);
                    }
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

        public bool CidadeExists(string cep = "", string cidade = "", string estado = "")
        {
            string key = "";

            if (cep.Length >= 5 && !string.IsNullOrEmpty(estado))
            {
                key = cep.Substring(0, 5) + "|" + estado.ToLower();
                if (cepEstadoDict.ContainsKey(key))
                    return true;
            }

            if (string.IsNullOrEmpty(cidade))
                return true;

            cidade = Tools.RemoverAcentos(cidade).ToLower();

            if (!string.IsNullOrWhiteSpace(cidade))
            {
                key = cidade + "|" + estado.ToLower();
                if (cidadeEstadoDict.ContainsKey(key))
                    return true;
            }

            return false;
        }

        public string BuscarCidade(string cidade = "")
        {
            string key = "";

            if (string.IsNullOrEmpty(cidade))
                return cidade;

            cidade = Tools.RemoverAcentos(cidade).ToLower();

            if (!string.IsNullOrWhiteSpace(cidade))
            {
                if (cidadeDict.ContainsKey(cidade))
                    return cidadeDict[cidade];
            }

            return cidade;
        }

        public string EncontrarCidadeSemelhante(string textoCidade)
        {
            textoCidade = textoCidade.ToLower();

            string cidadeEncontrada = null;
            int menorDistancia = int.MaxValue;

            foreach (string cidade in cidadeDict.Keys)
            {
                string cidadeNormalizada = Tools.RemoverAcentos(cidade).ToLower();
                int distancia = DistanciaLevenshtein(textoCidade, cidadeNormalizada);

                if (distancia < menorDistancia)
                {
                    menorDistancia = distancia;
                    cidadeEncontrada = cidade;
                }
            }

            return cidadeEncontrada;
        }
        private int DistanciaLevenshtein(string s, string t)
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

        public void CriarExcelArquivo(string nomeArquivo, DataTable dataTable)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Planilha1");

            int numberOfColumns = dataTable.Columns.Count;

            sheet.SetAutoFilter(new CellRangeAddress(0, 0, 0, numberOfColumns - 1));

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

            ICellStyle headerStyle = workbook.CreateCellStyle();
            headerStyle.FillForegroundColor = NPOI.SS.UserModel.IndexedColors.LightCornflowerBlue.Index;
            headerStyle.FillPattern = FillPattern.SolidForeground;

            // Aplicar o estilo de fundo ao cabeçalho
            headerRow.RowStyle = headerStyle;
            FileStream sw = File.Create(nomeArquivo);
            workbook.Write(sw);
            sw.Close();
        }

        public void CriarExcelArquivoV2(string nomeArquivo, DataTable dataTable)
        {
            var culture = new CultureInfo("pt-BR");

            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Planilha1");

            // Adicionar os nomes das colunas ao arquivo Excel
            var headerRow = worksheet.Row(1); // Começa na primeira linha
            for (int j = 0; j < dataTable.Columns.Count; j++)
            {
                headerRow.Cell(j + 1).Value = dataTable.Columns[j].ColumnName;
            }

            // Adicionar o DataTable ao arquivo Excel
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                var row = worksheet.Row(i + 2); // Começa na segunda linha, pois a primeira linha é para os nomes das colunas
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    var cellValue = dataTable.Rows[i][j].ToString();

                    if (Decimal.TryParse(cellValue, out var moeda))
                        row.Cell(j + 1).Value = moeda;
                    else if (DateTime.TryParseExact(cellValue, "dd/MM/yyyy HH:mm:ss", culture, DateTimeStyles.None, out DateTime data) ||
                        DateTime.TryParseExact(cellValue, "dd/MM/yyyy", culture, DateTimeStyles.None, out data))
                        row.Cell(j + 1).Value = data;
                    else
                        row.Cell(j + 1).Value = cellValue;
                }
            }

            // Fixar a primeira linha (cabeçalho)
            worksheet.SheetView.FreezeRows(1);

            // Aplicar o estilo de fundo ao cabeçalho
            headerRow.Style.Fill.BackgroundColor = XLColor.LightCornflowerBlue;

            // Aplicar o filtro automático
            worksheet.Range(1, 1, 1, dataTable.Columns.Count).SetAutoFilter();

            // Ajustar o tamanho das colunas automaticamente
            worksheet.Columns().AdjustToContents();

            // Salvar o arquivo
            workbook.SaveAs(nomeArquivo);
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

        public static string GetEspecieIDFromFormaPagamentoEntidades(string formaPagamento)
        {

            switch (formaPagamento)
            {
                case string a when a.Equals("17"):
                    return TitulosEspeciesID.Carne.ToString();

                case string b when b.Contains("4"):
                    return TitulosEspeciesID.Cheque.ToString();

                case string b when b.Contains("1"):
                    return TitulosEspeciesID.Dinheiro.ToString();

                case string b when b.Contains("22"):
                    return TitulosEspeciesID.TransferenciaBancaria.ToString();

                case string b when b.Contains("31") || b.Contains("PIX"):
                    return TitulosEspeciesID.DepositoEmConta.ToString();

                case string b when b.Contains("8") || b.Contains("MASTER") || b.Contains("VISA"):
                    return TitulosEspeciesID.CartaoCredito.ToString();

                default:
                    return TitulosEspeciesID.Dinheiro.ToString();
            }
        }

        public static DataTable ConversorEntidadeParaDataTable<T>(List<T> entidadeDTO) where T : class
        {
            {
                var dataTable = new DataTable();

                try
                {
                    foreach (PropertyDescriptor prop in TypeDescriptor.GetProperties(typeof(T)))
                    {
                        dataTable.Columns.Add(prop.DisplayName, prop.PropertyType);
                    }                    

                    foreach (var entidade in entidadeDTO)
                    {
                        DataRow row = dataTable.NewRow();

                        // Preenche as células da linha com os valores das propriedades da entidade
                        foreach (PropertyDescriptor prop in TypeDescriptor.GetProperties(typeof(T)))
                        {
                            row[prop.DisplayName] = prop.GetValue(entidade);
                        }

                        // Adiciona a linha ao DataTable de forma thread-safe
                        lock (dataTable)
                        {
                            dataTable.Rows.Add(row);
                        }
                    }
                }
                catch (Exception error)
                {
                    throw new Exception($"Erro na conversão da entidade para datatable: {error.Message}");
                }
                return dataTable;
            }
        }
    }
}
