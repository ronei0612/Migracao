using Dapper;
using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.Context
{
    public class FireBirdContext<T>
    {
        string _connectionString = @"Server=localhost;Database={0};User=SYSDBA;Password=masterkey;";

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

                    var aa = 1;
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
    }
}
