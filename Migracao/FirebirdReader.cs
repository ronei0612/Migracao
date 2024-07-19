using System;
using System.Data;
using FirebirdSql.Data.FirebirdClient;

namespace Migracao
{
    public class FirebirdReader
    {
        private string _connectionString;

        public FirebirdReader(string connectionString)
        {
            _connectionString = connectionString;
        }

        public DataTable ReadData(string sqlQuery)
        {
            DataTable dataTable = new DataTable();

            using (FbConnection connection = new FbConnection(_connectionString))
            {
                using (FbCommand command = new FbCommand(sqlQuery, connection))
                {
                    connection.Open();
                    using (FbDataReader reader = command.ExecuteReader())
                    {
                        // Add columns to the DataTable based on the reader's schema
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            dataTable.Columns.Add(reader.GetName(i), reader.GetFieldType(i));
                        }

                        // Read data and add rows to the DataTable
                        while (reader.Read())
                        {
                            DataRow row = dataTable.NewRow();
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                row[i] = reader.GetValue(i);
                            }
                            dataTable.Rows.Add(row);
                        }
                    }
                }
            }

            return dataTable;
        }

        public void ExecuteNonQuery(string sqlQuery)
        {
            using (FbConnection connection = new FbConnection(_connectionString))
            {
                using (FbCommand command = new FbCommand(sqlQuery, connection))
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}