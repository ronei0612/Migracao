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
				if (count == 1000)
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

		public object VerificarSeDateTime(object input)
		{
			if (input is DateTime dateTime)
				return dateTime.ToString("yyyy-MM-dd HH:mm:ss.f");

			return input;
		}
	}
}
