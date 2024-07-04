using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class Manutencoes
    {
        public string? Lancamento { get; set; }
        public string? Paciente_CPF { get; set; }
        public string? Nome_Paciente { get; set; }
        public decimal? Valor_Original { get; set; }
        public string? Dentista_Nome { get; set; }
        public DateTime? Data_Pagamento { get; set; }
        public string? Numero_Controle { get; set; }
        public string? Procedimento_Nome { get; set; }
        public decimal? Valor_Pagamento { get; set; }       
        public string? Tipo_Pagamento { get; set; }
        public DateTime? Vencimento { get; set; }
        public string? Valor_Devido { get; set; }
        public DateTime? Data_Atendimento { get; set; }
        public string? Procedimentos_Observacao { get; set; }
        public string? Procedimento_Valor { get; set; }
        public string? Dente { get; set; }
        public string? Quantidade_Orto { get; set; }
        public string? Valor_Total { get; set; }

        
    }
}
