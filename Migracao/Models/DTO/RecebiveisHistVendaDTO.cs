using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class RecebiveisHistVendaDTO
    {
        public RecebiveisHistVendaDTO()
        {
                
        }

        public RecebiveisHistVendaDTO(List<RecebiveisHistVenda> recebiveisHistVendas)
        {

        }

        [DisplayName("CPF")]
        public string? CPF { get; set; }

        [DisplayName("Nome")]
        public string? Nome { get; set; }

        [DisplayName("Observação Recebível")]
        public string? Observacao_Recebivel { get; set; }

        [DisplayName("Documento Ref")]
        public string? Documento_Ref { get; set; }

        [DisplayName("Valor Original")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public string? Valor_Original { get; set; }

        [DisplayName("Vencimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Vencimento { get; set; }

        [DisplayName("Emissão")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Emissao { get; set; }

        [DisplayName("Recebível Exigível(R/E)")]
        public string? Recebivel_Exigivel { get; set; }
    }
}
