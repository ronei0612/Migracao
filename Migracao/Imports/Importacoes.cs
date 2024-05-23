using Migracao.Utils;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;

namespace Migracao.Imports
{
	internal class Importacoes
	{
		public void Atendimentos(string filePath)
		{
			// Carrega o arquivo Excel
			IWorkbook workbook;
			using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				if (Path.GetExtension(filePath) == ".xls")
				{
					workbook = new HSSFWorkbook(file);
				}
				else
				{
					workbook = new XSSFWorkbook(file);
				}
			}

			// Lê a primeira planilha
			ISheet sheet = workbook.GetSheetAt(0);

			// Cria um DataTable para armazenar os dados
			DataTable dataTable = new DataTable();
			dataTable.Columns.Add("LoginID", typeof(int));
			dataTable.Columns.Add("EstabelecimentoID", typeof(int));
			dataTable.Columns.Add("AtendeTipoID", typeof(int));
			dataTable.Columns.Add("AtendimentoTipoCustomID", typeof(int));
			dataTable.Columns.Add("DataChegada", typeof(DateTime));
			dataTable.Columns.Add("DataInicio", typeof(DateTime));
			dataTable.Columns.Add("DataTermino", typeof(DateTime));
			dataTable.Columns.Add("DataCancelamento", typeof(DateTime));
			dataTable.Columns.Add("ConsumidorID", typeof(int));
			dataTable.Columns.Add("AtendimentoValor", typeof(decimal));
			dataTable.Columns.Add("SecretariaID", typeof(int));
			dataTable.Columns.Add("FuncionarioID", typeof(int));
			dataTable.Columns.Add("SalaID", typeof(int));
			dataTable.Columns.Add("ConvenioID", typeof(int));
			dataTable.Columns.Add("TempoAtraso", typeof(TimeSpan));
			dataTable.Columns.Add("TempoSalaEspera", typeof(TimeSpan));
			dataTable.Columns.Add("TempoAtendimento", typeof(TimeSpan));
			dataTable.Columns.Add("Observacoes", typeof(string));
			dataTable.Columns.Add("DataInclusao", typeof(DateTime));
			dataTable.Columns.Add("DataUltAlteracao", typeof(DateTime));
			dataTable.Columns.Add("AtendimentoIndex", typeof(int));
			dataTable.Columns.Add("EncaminhadoPorMedicoPessoaID", typeof(int));
			dataTable.Columns.Add("DiagnosticoID", typeof(int));

			// Lê os dados das linhas, ignorando o cabeçalho
			for (var row = 1; row <= sheet.LastRowNum; row++)
			{
				if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
				{
					DataRow dataRow = dataTable.NewRow();

					dataRow["LoginID"] = sheet.GetRow(row).GetCell(0) == null || string.IsNullOrEmpty(sheet.GetRow(row).GetCell(0).ToString())
						? DBNull.Value : Convert.ToInt32(sheet.GetRow(row).GetCell(0).NumericCellValue);
					dataRow["EstabelecimentoID"] = sheet.GetRow(row).GetCell(1) == null ? DBNull.Value : Convert.ToInt32(sheet.GetRow(row).GetCell(1).NumericCellValue);
					dataRow["AtendeTipoID"] = sheet.GetRow(row).GetCell(2) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(2));
					dataRow["AtendimentoTipoCustomID"] = sheet.GetRow(row).GetCell(3) == null || string.IsNullOrEmpty(sheet.GetRow(row).GetCell(3).ToString())
						? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(3));
					dataRow["DataChegada"] = sheet.GetRow(row).GetCell(4) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(4));
					dataRow["DataInicio"] = sheet.GetRow(row).GetCell(5) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(5));
					dataRow["DataTermino"] = sheet.GetRow(row).GetCell(6) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(6));
					dataRow["DataCancelamento"] = sheet.GetRow(row).GetCell(7) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(7));
					dataRow["ConsumidorID"] = sheet.GetRow(row).GetCell(8) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(8));
					dataRow["AtendimentoValor"] = sheet.GetRow(row).GetCell(9) == null ? DBNull.Value : Tools.GetDecimalValueFromCell(sheet.GetRow(row).GetCell(9));
					dataRow["SecretariaID"] = sheet.GetRow(row).GetCell(10) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(10));
					dataRow["FuncionarioID"] = sheet.GetRow(row).GetCell(11) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(11));
					dataRow["SalaID"] = sheet.GetRow(row).GetCell(12) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(12));
					dataRow["ConvenioID"] = sheet.GetRow(row).GetCell(13) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(13));
					dataRow["TempoAtraso"] = sheet.GetRow(row).GetCell(14) == null ? DBNull.Value : Tools.GetTimeSpanValueFromCell(sheet.GetRow(row).GetCell(14));
					dataRow["TempoSalaEspera"] = sheet.GetRow(row).GetCell(15) == null ? DBNull.Value : Tools.GetTimeSpanValueFromCell(sheet.GetRow(row).GetCell(15));
					dataRow["TempoAtendimento"] = sheet.GetRow(row).GetCell(16) == null ? DBNull.Value : Tools.GetTimeSpanValueFromCell(sheet.GetRow(row).GetCell(16));
					dataRow["Observacoes"] = sheet.GetRow(row).GetCell(17) == null ? DBNull.Value : sheet.GetRow(row).GetCell(17).StringCellValue;
					dataRow["DataInclusao"] = sheet.GetRow(row).GetCell(18) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(18));
					dataRow["DataUltAlteracao"] = sheet.GetRow(row).GetCell(19) == null ? DBNull.Value : Tools.GetDateTimeValueFromCell(sheet.GetRow(row).GetCell(19));
					dataRow["AtendimentoIndex"] = sheet.GetRow(row).GetCell(20) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(20));
					dataRow["EncaminhadoPorMedicoPessoaID"] = sheet.GetRow(row).GetCell(21) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(21));
					dataRow["DiagnosticoID"] = sheet.GetRow(row).GetCell(22) == null ? DBNull.Value : Tools.GetIntValueFromCell(sheet.GetRow(row).GetCell(22));

					dataTable.Rows.Add(dataRow);
				}
			}

			// Gera o script SQL de inserção
			//string sql = GerarSqlInsert(dataTable);

			//File.WriteAllText("asdf.sql", sql);
		}
	}
}
