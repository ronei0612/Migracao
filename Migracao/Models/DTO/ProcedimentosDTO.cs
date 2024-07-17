using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Migracao.Models.DTO
{
    public class ProcedimentosDTO
    {
        public ProcedimentosDTO()
        {
                
        }

        public ProcedimentosDTO(List<Procedimentos> procedimentos)
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
        public string? Procedimento_Valor { get; set; }

        [DisplayName("Procedimento Observação")]
        public string? Procedimento_Observacao { get; set; }

        [DisplayName("Data Início")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Inicio { get; set; }

        //[DisplayName("Data Término")]
        //[DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        //public string? Data_Termino { get; set; }

        [DisplayName("Data Atendimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Atendimento { get; set; }
    }
}
