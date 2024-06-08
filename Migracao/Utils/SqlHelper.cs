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

		public string GerarSqlInsert(int index, Dictionary<string, object> pessoaDict, int pessoaID, Dictionary<string, object[]> pessoaFonesDict, Dictionary<string, object> consumidorDict, int consumidorID, Dictionary<string, object> consumidorEnderecoDict)
		{
			var sql = new StringBuilder();

			if (pessoaDict != null)
			{
				sql.AppendLine("INSERT INTO Pessoas (");

				foreach (var key in pessoaDict.Keys)
					sql.Append($"{key}, ");

				// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
				sql.Remove(sql.Length - 2, 2).Append(") VALUES " + Environment.NewLine);

				sql.Append('(');
				foreach (var value in pessoaDict.Values)
				{
					try
					{
						if (value == null)
							sql.Append("NULL, ");
						else if (value.ToString().Equals("null", StringComparison.CurrentCultureIgnoreCase))
							sql.Append("NULL, ");
						else if (value is decimal)
							sql.Append($"'{value.ToString().Replace(',', '.')}', ");
						else
							sql.Append($"'{VerificarSeDateTime(value)}', ");
					}
					catch
					{
						sql.Append("NULL, ");
					}
				}
				sql.Remove(sql.Length - 2, 2).Append("); " + Environment.NewLine);

				// Obtendo ID da Pessoa inserida
				sql.AppendLine($"DECLARE @PessoaID{index} int;");
				sql.AppendLine($"SELECT @PessoaID{index} = SCOPE_IDENTITY();");
			}

			if (consumidorDict != null)
			{
				sql.AppendLine("INSERT INTO Consumidores (");

				foreach (var key in consumidorDict.Keys)
					sql.Append($"{key}, ");

				// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
				sql.Remove(sql.Length - 2, 2).Append(", PessoaID) VALUES " + Environment.NewLine);

				sql.Append('(');
				foreach (var value in consumidorDict.Values)
				{
					try
					{
						if (value == null)
							sql.Append("NULL, ");
						else if (value.ToString().Equals("null", StringComparison.CurrentCultureIgnoreCase))
							sql.Append("NULL, ");
						else if (value is decimal)
							sql.Append($"'{value.ToString().Replace(',', '.')}', ");
						else
							sql.Append($"'{VerificarSeDateTime(value)}', ");
					}
					catch
					{
						sql.Append("NULL, ");
					}
				}
				sql.Remove(sql.Length - 2, 2).Append($", @PessoaID{index}); " + Environment.NewLine);
			}


			if (pessoaFonesDict != null)
			{
				sql = new StringBuilder($"INSERT INTO PessoaFones (");

				// Adiciona os nomes das colunas
				foreach (var key in pessoaFonesDict.Keys)
					sql.Append($"{key}, ");

				// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
				sql.Remove(sql.Length - 2, 2).Append(", PessoaID) VALUES " + Environment.NewLine);

				// Adiciona os valores das colunas para cada linha
				for (int i = 0; i < pessoaFonesDict.Values.First().Length; i++)
				{
					sql.Append('(');
					foreach (var valueArray in pessoaFonesDict.Values)
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

					if (consumidorID <= 0)
						sql.Remove(sql.Length - 2, 2).Append($", @PessoaID{index}); " + Environment.NewLine);
					else
						sql.Remove(sql.Length - 2, 2).Append($", {pessoaID}); " + Environment.NewLine);

					//sql.Remove(sql.Length - 4, 4).Append(';');
					//sql.Append(Environment.NewLine + $"INSERT INTO {tableName} (");
					//foreach (var key in dataDict.Keys)
					//	sql.Append($"{key}, ");
					//sql.Remove(sql.Length - 2, 2).Append(") VALUES " + Environment.NewLine);
				}
			}


			if (consumidorEnderecoDict != null)
			{
				// Obtendo ID da Pessoa inserida caso não tenha consumidor para adicionar endereço
				if (consumidorID <= 0)
				{
					sql.AppendLine($"DECLARE @ConsumidorID{index} int;");
					sql.AppendLine($"SELECT @ConsumidorID{index} = SCOPE_IDENTITY();");
				}

				sql.AppendLine("INSERT INTO ConsumidorEnderecos (");

				foreach (var key in consumidorEnderecoDict.Keys)
					sql.Append($"{key}, ");

				// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
				sql.Remove(sql.Length - 2, 2).Append(", ConsumidorID) VALUES " + Environment.NewLine);

				sql.Append('(');
				foreach (var value in consumidorEnderecoDict.Values)
				{
					try
					{
						if (value == null)
							sql.Append("NULL, ");
						else if (value.ToString().Equals("null", StringComparison.CurrentCultureIgnoreCase))
							sql.Append("NULL, ");
						else if (value is decimal)
							sql.Append($"'{value.ToString().Replace(',', '.')}', ");
						else
							sql.Append($"'{VerificarSeDateTime(value)}', ");
					}
					catch
					{
						sql.Append("NULL, ");
					}
				}
				if (consumidorID <= 0)
					sql.Remove(sql.Length - 2, 2).Append($", @ConsumidorID{index}); " + Environment.NewLine);
				else
					sql.Remove(sql.Length - 2, 2).Append($", {consumidorID}); " + Environment.NewLine);
			}


			// Remove a última quebra de linha e vírgula e espaço e adiciona um ponto e vírgula
			sql.Remove(sql.Length - 4, 4).Append(';' + Environment.NewLine);

			return sql.ToString();
		}

		public void GerarSqlUpdate(string tableName, string salvarArquivo, Dictionary<string, object[]> dataDict)
		{
			var sql = new StringBuilder();

			// Adiciona os nomes das colunas e valores para cada linha
			int count = 0;
			for (int i = 0; i < dataDict.Values.First().Length; i++)
			{
				sql.Append($"UPDATE {tableName} SET ");

				// Adiciona as colunas e seus valores, ignorando ID
				int columnCount = 0;
				foreach (var key in dataDict.Keys)
				{
					if (key.Equals("ID", StringComparison.OrdinalIgnoreCase))
					{
						continue; // Ignora a coluna ID
					}

					try
					{
						if (dataDict[key][i] == null)
							sql.Append($"{key} = NULL, ");
						else if (dataDict[key][i].ToString().Equals("null", StringComparison.CurrentCultureIgnoreCase))
							sql.Append($"{key} = NULL, ");
						else if (dataDict[key][i] is decimal)
							sql.Append($"{key} = '{dataDict[key][i].ToString().Replace(',', '.')}', ");
						else
							sql.Append($"{key} = '{VerificarSeDateTime(dataDict[key][i])}', ");
					}
					catch
					{
						sql.Append($"{key} = NULL, ");
					}

					columnCount++;
				}

				// Remove a última vírgula e espaço
				if (columnCount > 0) // Verifica se alguma coluna foi atualizada
				{
					sql.Remove(sql.Length - 2, 2);

					// Adiciona a cláusula WHERE com a coluna de chave primária
					sql.Append($" WHERE ID = '{dataDict["ID"][i]}';" + Environment.NewLine);
				}
				else
				{
					// Se nenhuma coluna foi atualizada, ignora a linha
					sql.Clear();
				}

				count++;
				if (count == 200)
				{
					count = 0;
				}
			}

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
