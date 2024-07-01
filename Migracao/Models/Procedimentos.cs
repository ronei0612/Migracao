using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class Procedimentos
    {
        public string? Numero_Controle { get; set; }
        public string? Paciente_CPF { get; set; }
        public string? Nome_Paciente { get; set; }
        public string? Dentista_CPF { get; set; }
        public string? Dentista_Nome { get; set; }
        public string? Dente { get; set; }
        public string? NOME_PRODUTO { get; set; }
        public string? Valor { get; set; }
        public string? Observacao { get; set; }
        public string? Data_Inicio { get; set; }
        public string? Data_Termino { get; set; }
        public string? Data_Atendimento { get; set; }
    }
}
