using System.Data;
using System.Text;

namespace Migracao.Utils
{
    internal class SqlHelper
    {
		public void GerarSqlInsert(string tableName, string salvarArquivo, Dictionary<string, object[]> dataDict)
		{
			var sql = new StringBuilder($"INSERT INTO {tableName} (");

			// Adiciona os nomes das colunas
			foreach (var key in dataDict.Keys)
				sql.Append($"{key}, ");

			// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
			sql.Remove(sql.Length - 2, 2).Append(") VALUES " + Environment.NewLine);

			// Adiciona os valores das colunas para cada linha
			int count = 0;
			for (int i = 0; i < dataDict.Values.First().Length; i++)
			{
				sql.Append('(');
				foreach (var valueArray in dataDict.Values)
				{
					try
					{
						if (valueArray[i] == null)
							sql.Append($"NULL, ");
						else if (valueArray[i].ToString().Equals("null", StringComparison.CurrentCultureIgnoreCase))
							sql.Append($"NULL, ");
						else if (valueArray[i] is decimal)
							sql.Append($"'{valueArray[i].ToString().Replace(',', '.')}', ");
						else
							sql.Append($"'{VerificarSeDateTime(valueArray[i])}', ");
					}
					catch
					{
						sql.Append($"NULL, ");
					}
				}
				sql.Remove(sql.Length - 2, 2).Append("), " + Environment.NewLine);

				count++;
				if (count == 200)
				{
					sql.Remove(sql.Length - 4, 4).Append(';');
					sql.Append(Environment.NewLine + $"INSERT INTO {tableName} (");
					foreach (var key in dataDict.Keys)
						sql.Append($"{key}, ");
					sql.Remove(sql.Length - 2, 2).Append(") VALUES " + Environment.NewLine);
					count = 0;
				}
			}

			// Remove a última quebra de linha e vírgula e espaço e adiciona um ponto e vírgula
			sql.Remove(sql.Length - 4, 4).Append(';');

			File.WriteAllText(salvarArquivo + ".sql", sql.ToString());
		}

		private void GerarSqlInsert(string tableName, string salvarArquivo, DataTable dataTable)
		{
			var sql = new StringBuilder();
			var values = new StringBuilder();
			var insertCount = 0;

			sql.AppendLine($"INSERT INTO {tableName} ({string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}]"))}) VALUES ");

			foreach (DataRow row in dataTable.Rows)
			{
				var rowValues = new List<string>();

				foreach (DataColumn column in dataTable.Columns)
				{
					object value = row[column];
					if (value == null || (value is DBNull))
					{
						rowValues.Add("NULL");
					}
					else if (column.DataType == typeof(string) || column.DataType == typeof(DateTime) || column.DataType == typeof(TimeSpan))
					{
						if (column.DataType == typeof(DateTime))
							rowValues.Add($"'{((DateTime)value).ToString("yyyy-MM-dd HH:mm:ss").Replace("'", "''")}'");
						else if (column.DataType == typeof(TimeSpan))
							rowValues.Add($"'{((TimeSpan)value).TotalSeconds.ToString().Split(',')[0].Replace("'", "''")}'");
						else
							rowValues.Add($"'{value.ToString().Replace("'", "''")}'");
					}
					else
					{
						rowValues.Add(value.ToString());
					}
				}

				values.AppendLine($"({string.Join(", ", rowValues)}),");
				insertCount++;

				// A cada 1000 inserts, adiciona um novo bloco INSERT INTO
				if (insertCount % 200 == 0)
				{
					sql.Append(values.ToString().TrimEnd(',', '\r', '\n') + ";"); // Remove a última virgula e quebra de linha
					sql.AppendLine();
					sql.AppendLine($"INSERT INTO {tableName} ({string.Join(", ", dataTable.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}]"))}) VALUES ");
					values.Clear(); // Limpa o StringBuilder de values
				}
			}

			// Adiciona o último bloco de inserts, caso haja algum
			if (values.Length > 0)
			{
				sql.Append(values.ToString().TrimEnd(',', '\r', '\n') + ";");
			}

			File.WriteAllText(salvarArquivo + ".sql", sql.ToString());
		}

		public object VerificarSeDateTime(object input)
		{
			if (input is DateTime dateTime)
				return dateTime.ToString("yyyy-MM-dd HH:mm:ss.f");

			return input;
		}
	}
}
