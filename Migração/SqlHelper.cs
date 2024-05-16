using System.Diagnostics;
using System.Text;

namespace Migração
{
	internal class SqlHelper
	{
		public string GerarSqlInsert(string tableName, Dictionary<string, object[]> dataDict)
		{
			var sql = new StringBuilder($"INSERT INTO {tableName} (");

			// Adiciona os nomes das colunas
			foreach (var key in dataDict.Keys)
			{
				sql.Append($"{key}, ");
			}

			// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
			sql.Remove(sql.Length - 2, 2).Append(") VALUES " + Environment.NewLine);

			// Adiciona os valores das colunas para cada linha
			for (int i = 0; i < dataDict.Values.First().Length; i++)
			{
				sql.Append('(');
				foreach (var valueArray in dataDict.Values)
				{
					Debug.WriteLine("i=" + i);
					Debug.WriteLine("valueArray[i]=" + valueArray[i]);
					try
					{
						sql.Append($"'{valueArray[i]}', ");
					} catch
					{
						sql.Append($"NULL, ");
					}
				}
				sql.Remove(sql.Length - 2, 2).Append("), " + Environment.NewLine);
			}

			// Remove a última quebra de linha e vírgula e espaço e adiciona um ponto e vírgula
			sql.Remove(sql.Length - 4, 4).Append(';');

			return sql.ToString();
		}
	}
}
