using System.Text;

namespace Migração.Utils
{
    internal class SqlHelper
    {
        public string GerarSqlInsert(string tableName, Dictionary<string, object[]> dataDict)
        {
			var sqlList = new List<string>();
			var sql = new StringBuilder($"INSERT INTO {tableName} (");

			// Adiciona os nomes das colunas
			foreach (var key in dataDict.Keys)
				sql.Append($"{key}, ");

			// Remove a última vírgula e espaço e adiciona um parêntese de fechamento e a palavra VALUES
			sql.Remove(sql.Length - 2, 2).Append(") VALUES ");

			// Adiciona os valores das colunas para cada linha
			for (int i = 0; i < dataDict.Values.First().Length; i++)
			{
				if (i != 0 && i % 1000 == 0)
				{
					sql.Remove(sql.Length - 2, 2).Append(';');
					sqlList.Add(sql.ToString());
					sql.Clear().Append($"INSERT INTO {tableName} (");

					foreach (var key in dataDict.Keys)
						sql.Append($"{key}, ");

					sql.Remove(sql.Length - 2, 2).Append(") VALUES ");
				}

				sql.Append('(');
				foreach (var valueArray in dataDict.Values)
				{
					try
					{
						if (valueArray[i] == null)
							sql.Append($"NULL, ");
						else
							sql.Append($"'{VerificarSeDateTime(valueArray[i])}', ");
					}
					catch
					{
						sql.Append($"NULL, ");
					}
				}
				sql.Remove(sql.Length - 2, 2).Append("), ");
			}

			// Remove a última vírgula e espaço e adiciona um ponto e vírgula
			sql.Remove(sql.Length - 2, 2).Append(';');
			sqlList.Add(sql.ToString());

			return string.Join(Environment.NewLine, sqlList);
        }

        public object VerificarSeDateTime(object input)
        {
            if (input is DateTime dateTime)
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss.f");

            return input;
        }
    }
}
