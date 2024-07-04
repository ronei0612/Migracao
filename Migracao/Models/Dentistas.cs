using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class Dentistas
    {
        public string? Codigo { get; set; }
        public string? NOME { get; set; }
        public string? Departamento { get; set; }
        public string? Observacoes { get; set; }
        public string? Ativo { get; set; }
        public string? Nome_Completo { get; set; }
        public string? Email { get; set; }
        public string? Telefone { get; set; }
        public string? CRO { get; set; }
        public DateTime? Data_Cadastro { get; set; }
    }
}
