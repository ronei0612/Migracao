using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        [DisplayName("Paciente CPF")]
        public string? Paciente_CPF { get; set; }

        [DisplayName("Paciente Nome")]
        public string? Paciente_Nome { get; set; }

        [DisplayName("Dentista Nome")]
        public string? Dentista_Nome { get; set; }

        [DisplayName("Dentista Codigo")]
        public string? Dentista_Codigo { get; set; }

        [DisplayName("Procedimento Nome")]
        public string? Procedimento_Nome { get; set; }

        [DisplayName("Data Atendimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Atendimento { get; set; }

        [DisplayName("Data Início")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Inicio { get; set; }

        [DisplayName("Data Retorno")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Retorno { get; set; }

        [DisplayName("Procedimento Observação")]
        public string? Procedimento_Observacao { get; set; }

        [DisplayName("Lancamento")]
        public string? Lancamento { get; set; }
    }
}
