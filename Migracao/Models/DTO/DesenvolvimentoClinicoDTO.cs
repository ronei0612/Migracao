using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

namespace Migracao.Models.DTO
{
    public class DesenvolvimentoClinicoDTO
    {
        public DesenvolvimentoClinicoDTO()
        {
            
        }

        public DesenvolvimentoClinicoDTO(List<DesenvolvimentoClinico> desenvolvimentoClinico)
        {

        }

        [DisplayName("Paciente Nome")]
        public string? Nome_Completo { get; set; }

        [DisplayName("CPF")]
        public string? CPF { get; set; }

        [DisplayName("Telefone")]
        public string? Telefone { get; set; }

        [DisplayName("Dentista Nome")]
        public string? Dentista { get; set; }

        [DisplayName("Data Hora Início")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Hora_Inicio { get; set; }

        [DisplayName("Data Hora Término")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Hora_Termino { get; set; }

        [DisplayName("Data Hora Atendimento Início")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Hora_Atendimento_Inicio { get; set; }

        [DisplayName("Data Hora Atendimento Término")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Hora_Atendimento_Termino { get; set; }

        [DisplayName("Observação Desenvolvimento Clínico")]
        public string? Desenvolvimento_Clinico { get; set; }
    }
}
