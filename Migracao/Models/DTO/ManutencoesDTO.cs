using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

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
        public string? Numero_Controle { get; set; }

        [DisplayName("Paciente CPF")]
        public string? Paciente_CPF { get; set; }

        [DisplayName("Paciente Nome")]
        public string? Paciente_Nome { get; set; }

        [DisplayName("Dentista Nome")]
        public string? Dentista_Nome { get; set; }

        [DisplayName("Manutenção Nome")]
        public string? Procedimento_Nome { get; set; }

        [DisplayName("Valor")]
        public string? Procedimento_Valor { get; set; }

        [DisplayName("Observação")]
        public string? Procedimento_Observacao { get; set; }

        [DisplayName("Quantidade Orto")]
        public string? Quantidade_Orto { get; set; }

        [DisplayName("Valor Total")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Total { get; set; }

        [DisplayName("Data Início")]
        public string? Data_Inicio { get; set; }

        [DisplayName("Data Atendimento")]
        public string? Data_Atendimento { get; set; }
    }
}
