using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class ProcedimentosManutencaoDTO
    {
        public ProcedimentosManutencaoDTO()
        {
            
        }

        public ProcedimentosManutencaoDTO(List<Procedimentos> procedimentos, List<Manutencoes> manutencoes)
        {

        }
        [DisplayName("Número do Controle")]
        public string? Numero_Controle { get; set; }

        [DisplayName("Paciente CPF")]
        public string? Paciente_CPF { get; set; }

        [DisplayName("Paciente Nome")]
        public string? Paciente_Nome { get; set; }

        [DisplayName("Dentista CPF")]
        public string? Dentista_CPF { get; set; }

        [DisplayName("Dentista Nome")]
        public string? Dentista_Nome { get; set; }

        [DisplayName("Dente")]
        public string? Dente { get; set; }

        [DisplayName("Procedimento Nome")]
        public string? Procedimento_Nome { get; set; }

        [DisplayName("Procedimento Valor")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Procedimento_Valor { get; set; }

        [DisplayName("Procedimento Observação")]
        public string? Procedimento_Observacao { get; set; }

        [DisplayName("Data Hora Início")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Inicio { get; set; }

        [DisplayName("Data Hora Término")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Termino { get; set; }

        [DisplayName("Data Atendimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Atendimento { get; set; }

        [DisplayName("Quantidade Orto")]
        public string? Quantidade_Orto { get; set; }

        [DisplayName("Tipo Pagamento")]
        public string? Tipo_Pagamento { get; set; }

        [DisplayName("Vencimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Vencimento { get; set; }

        [DisplayName("Valor Devido")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Devido { get; set; }

        [DisplayName("Valor Total")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Total { get; set; }

        [DisplayName("Valor Pago")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Pago { get; set; }

        [DisplayName("Valor Original")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Original { get; set; }

        [DisplayName("Data Pagamento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Pagamento { get; set; }

        [DisplayName("Observação Recebido")]
        public string? Observacao_Recebido { get; set; }
    }
}
