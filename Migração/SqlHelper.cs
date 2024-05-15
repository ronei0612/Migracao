using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
			sql.Remove(sql.Length - 2, 2).Append(") VALUES ");

			// Adiciona os valores das colunas para cada linha
			for (int i = 0; i < dataDict.Values.First().Length; i++)
			{
				sql.Append('(');
				foreach (var valueArray in dataDict.Values)
				{
					sql.Append($"'{valueArray[i]}', ");
				}
				sql.Remove(sql.Length - 2, 2).Append("), ");
			}

			// Remove a última vírgula e espaço e adiciona um ponto e vírgula
			sql.Remove(sql.Length - 2, 2).Append(';');

			return sql.ToString();
		}
	}
}
