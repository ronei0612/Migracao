using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class ManutencoesDTO
    {
        public ManutencoesDTO()
        {
                
        }
        public ManutencoesDTO(List<Manutencoes> manutencoes)
        {

        }

        [DisplayName("Número do Controle")]
        public string? NumeroControle { get; set; }
        [DisplayName("Paciente CPF")]
        public string? PacienteCPF { get; set; }
        [DisplayName("Paciente Nome")]
        public string? PacienteNome { get; set; }
        [DisplayName("Dentista Nome")]
        public string? DentistaNome { get; set; }
        [DisplayName("Procedimento Nome")]
        public string? ProcedimentoNome { get; set; }
        [DisplayName("Procedimento Valor")]
        public string? ProcedimentoValor { get; set; }
        [DisplayName("Valor Original")]
        public string? ValorOriginal { get; set; }
        [DisplayName("Valor Pagamento")]
        public string? ValorPagamento { get; set; }
        [DisplayName("Data do Pagamento")]
        public string? DataPagamento { get; set; }
        [DisplayName("Dente")]
        public string? Dente { get; set; }
        [DisplayName("Procedimento Observação")]
        public string? ProcedimentoObservacao { get; set; }
        [DisplayName("Quantidade Orto")]
        public string? QuantidadeOrto { get; set; }
        [DisplayName("Tipo Pagamento")]
        public string? TipoPagamento { get; set; }
        [DisplayName("Vencimento")]
        public string? Vencimento { get; set; }
        [DisplayName("Valor Devido")]
        public string? ValorDevido { get; set; }
        [DisplayName("Valor Total")]
        public string? ValorTotal { get; set; }
        [DisplayName("Data Atendimento")]
        public string? DataAtendimento { get; set; }
    }
}
