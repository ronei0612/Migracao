using Microsoft.EntityFrameworkCore;
using Migracao.Models.DentalOffice;

namespace Migracao.Models.Context
{
    public class MySqlContext : DbContext
    {
        private string _connectionString = @"server=127.0.0.1;database={0};uid=admin;pwd=admin;port=3306";

        public MySqlContext(string databaseName)
        {
            _connectionString = string.Format(_connectionString, databaseName);
        }

        public DbSet<Orcamento> Orcamento { get; set; }
        public DbSet<Pacientes> Paciente { get; set; }
        public DbSet<Funcionario> Funcionario { get; set; }
        public DbSet<Tratamento> Tratamento { get; set; }
        public DbSet<Procedimentos> Procedimentos { get; set; }
        public DbSet<Preco> Precos { get; set; }
        public DbSet<DentalOffice.Recebiveis> Recebiveis { get; set; }
        public DbSet<Credits> Credits { get; set; }
        public DbSet<FluxoCaixa> FluxoCaixa { get; set; }
        public DbSet<Exigivel> Exigivel { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseMySql(_connectionString, ServerVersion.AutoDetect(_connectionString));
        }
    }
}
