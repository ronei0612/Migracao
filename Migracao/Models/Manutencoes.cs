using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class Manutencoes
    {
        public string? NumeroControle { get; set; }
        public string? PacienteCPF { get; set; }
        public string? PacienteNome { get; set; }
        public string? DentistaNome { get; set; }
        public string? ProcedimentoNome { get; set; }
        public string? ProcedimentoValor { get; set; }
        public string? ValorOriginal { get; set; }
        public string? ValorPagamento { get; set; }
        public string? DataPagamento { get; set; }
        public string? Dente { get; set; }
        public string? ProcedimentoObservacao { get; set; }
        public string? QuantidadeOrto { get; set; }
        public string? TipoPagamento { get; set; }
        public string? Vencimento { get; set; }
        public string? ValorDevido { get; set; }
        public string? ValorTotal { get; set; }
        public string? DataAtendimento { get; set; }
    }
}
