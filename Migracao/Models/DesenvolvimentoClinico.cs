using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class DesenvolvimentoClinico
    {
        public string? Lancamento { get; set; }
        public string? Paciente_CPF { get; set; }
        public string? Paciente_Nome { get; set; } // Nome do Paciente
        public DateTime? Data_Retorno { get; set; } // Data de Retorno (pode ser nulo)
        public string? Dentista_Codigo { get; set; } // Código do Dentista
        public string? Dentista_Nome { get; set; } // Nome do Dentista
        public string? Procedimento_Nome { get; set; } // Nome do Procedimento
        public DateTime? Data_Inicio { get; set; } // Data de Início (obrigatória)
        public string? Procedimento_Observacao { get; set; } // Observação do Procedimento
    }
}
