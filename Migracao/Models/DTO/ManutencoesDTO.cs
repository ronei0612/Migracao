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

        [DisplayName("Numero do Controle")]
        public string Numero_Controle { get; set; }

        [DisplayName("Paciente CPF")]
        public string Paciente_CPF { get; set; }

        [DisplayName("Paciente Nome")]
        public string Paciente_Nome { get; set; }

        [DisplayName("Dentista Nome")]
        public string Dentista_Nome { get; set; }

        [DisplayName("Procedimento Nome")]
        public string Procedimento_Nome { get; set; }

        [DisplayName("Procedimento Valor")]
        public decimal Procedimento_Valor { get; set; }

        [DisplayName("Valor Original")]
        public decimal Valor_Original { get; set; }

        [DisplayName("Valor do Pagamento")]
        public decimal Valor_Pagamento { get; set; }

        [DisplayName("Data do Pagamento")]
        public DateTime? Data_Pagamento { get; set; }

        [DisplayName("Dente")]
        public string Dente { get; set; }

        [DisplayName("Procedimento Observação")]
        public string Procedimento_Observacao { get; set; }

        [DisplayName("Quantidade Orto")]
        public int Quantidade_Orto { get; set; }

        [DisplayName("Tipo Pagamento")]
        public string Tipo_Pagamento { get; set; }

        [DisplayName("Vencimento")]
        public DateTime Vencimento { get; set; }

        [DisplayName("Valor Devido")]
        public decimal Valor_Devido { get; set; }

        [DisplayName("Valor Total")]
        public decimal Valor_Total { get; set; }

        [DisplayName("Data Atendimento")]
        public DateTime Data_Atendimento { get; set; }
    }
}
