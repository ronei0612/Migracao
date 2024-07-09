using Dapper;
using FirebirdSql.Data.FirebirdClient;
using MySqlConnector;
using NPOI.OpenXmlFormats.Shared;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Migracao.Models.Context
{
    public class FireBirdContext<T>
    {
        string _connectionString = @"Server=localhost;Database={0};User=SYSDBA;Password=masterkey;Charset=ISO8859_1";

        public FireBirdContext(string dataBase)
        {
            _connectionString = string.Format(_connectionString, dataBase);
        }

        public long Add(T entity)
        {

            using (IDbConnection db = new FbConnection(_connectionString))
            {

                try
                {
                    return db.Insert(entity).Value;

                }
                catch (Exception ex)
                {
                    return 0;
                }
            }
        }
        public T GetById(long id)
        {

            using (IDbConnection db = new FbConnection(_connectionString))
            {

                return db.Get<T>(id);
            }
        }

        public List<T> GetAll()
        {

            using (IDbConnection db = new FbConnection(_connectionString))
            {

                var tipo = typeof(T).CustomAttributes.Where(x => x.AttributeType.Name == "TableAttribute").FirstOrDefault().ConstructorArguments[0].Value;
                var list = db.Query<T>($"SELECT * FROM  {tipo} ").ToList();

                return list;//ToList();
            }
        }

        public List<T> RetornaItensBancoPorQuery(string arquivoSql)
        {            
            string sqlScript = File.ReadAllText(arquivoSql, Encoding.UTF8);
            try
            {
                using (IDbConnection db = new FbConnection(_connectionString))
                {
                    var list = db.Query<T>(sqlScript).ToList();

                    return list;
                }
            }
            catch (Exception error)
            {
                throw new Exception($"Erro ao fazer a consulta no banco Firebird: {error.Message}");
            }
            
        }
    }
}
