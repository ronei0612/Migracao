using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class FinanceiroRecebiveisDTO
    {
        public FinanceiroRecebiveisDTO()
        {
                
        }

        public FinanceiroRecebiveisDTO(List<Recebiveis> recebiveis)
        {

        }

        [DisplayName("Paciente CPF")]
        public string? Paciente_CPF { get; set; }

        [DisplayName("Nome")]
        public string? Nome { get; set; }

        [DisplayName("Número do Controle")]
        public string? Numero_Controle { get; set; }

        [DisplayName("Recebível Exigível(R/E)")]
        public string? Recebivel_Exigivel { get; set; }

        [DisplayName("Valor Devido")]
        public string? Valor_Devido { get; set; }

        [DisplayName("Valor Pago")]
        public string? Valor_Pago { get; set; }

        [DisplayName("Prazo")]
        public string? Prazo { get; set; }

        [DisplayName("Data Vencimento")]
        public string? Data_Vencimento { get; set; }

        [DisplayName("Data do Pagamento")]
        public string? Data_Pagamento { get; set; }

        [DisplayName("Emissão")]
        public string? Emissao { get; set; }

        [DisplayName("Observação Recebível")]
        public string Observacao_Recebivel { get; set; }

        [DisplayName("Observação Recebido")]
        public string? Observacao_Recebido { get; set; }

        [DisplayName("Tipo Pagamento")]
        public string? Tipo_Pagamento { get; set; }

        [DisplayName("Valor Original")]
        public string? Valor_Original { get; set; }

        [DisplayName("Vencimento Recebível")]
        public string? Vencimento_Recebivel { get; set; }

        [DisplayName("Duplicata")]
        public string? Duplicata { get; set; }

        [DisplayName("Parcela")]
        public string? Parcela { get; set; }

        [DisplayName("Situação")]
        public string? Situacao { get; set; }

        [DisplayName("Nome grupo")]
        public string? Nome_Grupo { get; set; }

        [DisplayName("Ordem")]
        public string? Ordem { get; set; }
    }
}
