using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class AgendamentosDTO
    {
        [DisplayName("ID")]
        public string? Lancamento { get; set; }

        [DisplayName("CPF")]
        public string? CPF { get; set; }

        [DisplayName("Nome Completo")]
        public string? Nome_Completo { get; set; }

        [DisplayName("Telefone")]
        public string? Telefone { get; set; }

        [DisplayName("Data Início")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public DateTime Data_Inicio { get; set; }

        [DisplayName("Data Término")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public DateTime? Data_Termino { get; set; }

        [DisplayName("Data Inclusão")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy}", ApplyFormatInEditMode = true)]
        public DateTime? Data_Inclusao { get; set; }

        [DisplayName("Nome Completo Dentista")]
        public string? Nome_Completo_Dentista { get; set; }

        [DisplayName("Observação")]
        public string? Observacao { get; set; }
    }
}
