using Migracao.Models;
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

		private string FormatValue(object value)
		{
			//if (value == null || value.ToString().Equals("null", StringComparison.OrdinalIgnoreCase))
			//	return "NULL";

			//if (value is decimal)
			//	return $"'{value.ToString().Replace(',', '.')}'";

			//return $"'{value}'";

			try
			{
				if (value == null)
					return "NULL";
				else if (value.ToString().Equals("null", StringComparison.CurrentCultureIgnoreCase))
					return "NULL";
				else if (value is decimal)
					return $"'{value.ToString().Replace(',', '.')}'";
				else
					return $"'{VerificarSeDateTime(value)}'";
			}
			catch
			{
				return "NULL";
			}
		}

		public string GerarSqlInsertPessoas(int index, Dictionary<string, object> pessoaDict, int pessoaID, Dictionary<string, object[]> pessoaFonesDict, Dictionary<string, object> consumidorDict, int consumidorID, Dictionary<string, object> consumidorEnderecoDict, Dictionary<string, object> funcionarioDict, Dictionary<string, object> enderecoDict)
		{
			var sql = new StringBuilder();

			if (pessoaDict != null)
			{
				sql.AppendLine($"INSERT INTO Pessoas ({string.Join(", ", pessoaDict.Keys)}) " +
					$"VALUES ({string.Join(", ", pessoaDict.Values.Select(FormatValue))});");
				sql.AppendLine($"DECLARE @PessoaID{index} int;");
				sql.AppendLine($"SELECT @PessoaID{index} = SCOPE_IDENTITY();");
			}

			if (consumidorDict != null)
			{
				sql.AppendLine($"INSERT INTO Consumidores ({string.Join(", ", consumidorDict.Keys)}, PessoaID) " +
					$"VALUES ({string.Join(", ", consumidorDict.Values.Select(FormatValue))}, {(pessoaID > 0 ? pessoaID.ToString() : $"@PessoaID{index}")});");
				sql.AppendLine($"DECLARE @ConsumidorID{index} int;");
				sql.AppendLine($"SELECT @ConsumidorID{index} = SCOPE_IDENTITY();");
			}

			if (funcionarioDict != null)
			{
				sql.AppendLine($"INSERT INTO Funcionarios ({string.Join(", ", funcionarioDict.Keys)}, PessoaID) " +
					$"VALUES ({string.Join(", ", funcionarioDict.Values.Select(FormatValue))}, {(pessoaID > 0 ? pessoaID.ToString() : $"@PessoaID{index}")});");
			}

			if (pessoaFonesDict != null)
			{
				sql.AppendLine($"INSERT INTO PessoaFones ({string.Join(", ", pessoaFonesDict.Keys)}, PessoaID) VALUES (");

				foreach (var valueArray in pessoaFonesDict.Values)
					sql.Append($"{FormatValue(valueArray[0])}, ");

				sql.Append($"{(consumidorID > 0 ? consumidorID.ToString() : $"@PessoaID{index}")});");
			}

			if (consumidorEnderecoDict != null)
			{
				sql.AppendLine($"INSERT INTO ConsumidorEnderecos ({string.Join(", ", consumidorEnderecoDict.Keys)}, ConsumidorID) " +
					$"VALUES ({string.Join(", ", consumidorEnderecoDict.Values.Select(FormatValue))}, {(consumidorID > 0 ? consumidorID.ToString() : $"@ConsumidorID{index}")});");
			}

			if (enderecoDict != null)
			{
				sql.AppendLine($"INSERT INTO Enderecos ({string.Join(", ", enderecoDict.Keys)}, PessoaID) " +
					$"VALUES ({string.Join(", ", enderecoDict.Values.Select(FormatValue))}, {(pessoaID > 0 ? pessoaID.ToString() : $"@PessoaID{index}")});");
			}

			return sql.ToString().TrimEnd(';');
		}

		public string GerarSqlInsertRecebiveis(int index, Dictionary<string, object> recebivelDict, Dictionary<string, object> fluxoCaixaDict)
		{
			var sql = new StringBuilder();

			if (recebivelDict != null)
			{
				sql.AppendLine($"INSERT INTO Recebiveis ({string.Join(", ", recebivelDict.Keys)}) " +
					$"VALUES ({string.Join(", ", recebivelDict.Values.Select(FormatValue))});");
				sql.AppendLine($"DECLARE @RecebivelID{index} int;");
				sql.AppendLine($"SELECT @RecebivelID{index} = SCOPE_IDENTITY();");
			}

			if (fluxoCaixaDict != null)
			{
				sql.AppendLine($"INSERT INTO FluxoCaixa ({string.Join(", ", fluxoCaixaDict.Keys)}, RecebivelID) " + 
					$"VALUES ({string.Join(", ", fluxoCaixaDict.Values.Select(FormatValue))}, {$"@RecebivelID{index}"});");
			}

			return sql.ToString().TrimEnd(';');
		}

		public string GerarSqlInsertPrecos(int index, Dictionary<string, object> tabelaDict, long tabelaID, Dictionary<string, object> precoDict)
		{
			var sql = new StringBuilder();

			if (tabelaDict != null)
			{
				sql.AppendLine("INSERT INTO PrecosTabelas (");

				foreach (var key in tabelaDict.Keys)
					sql.Append($"{key}, ");

				// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
				sql.Remove(sql.Length - 2, 2).Append(") VALUES " + Environment.NewLine);

				sql.Append('(');
				foreach (var value in tabelaDict.Values)
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

				// Obtendo ID da Recebivel inserido
				sql.AppendLine($"DECLARE @TabelaID{index} int;");
				sql.AppendLine($"SELECT @TabelaID{index} = SCOPE_IDENTITY();");
			}

			if (precoDict != null)
			{
				sql.AppendLine("INSERT INTO Precos (");

				foreach (var key in precoDict.Keys)
					sql.Append($"{key}, ");

				// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
				sql.Remove(sql.Length - 2, 2).Append(", TabelaID) VALUES " + Environment.NewLine);

				sql.Append('(');
				foreach (var value in precoDict.Values)
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
				if (tabelaID <= 0)
					sql.Remove(sql.Length - 2, 2).Append($", @TabelaID{index}); " + Environment.NewLine);
				else
					sql.Remove(sql.Length - 2, 2).Append($", {tabelaID}); " + Environment.NewLine);
				
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
