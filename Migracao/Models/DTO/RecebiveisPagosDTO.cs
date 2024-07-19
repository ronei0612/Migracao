using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class RecebiveisPagosDTO
    {
        [DisplayName("Documento")]
        public string? Documento { get; set; }

        [DisplayName("Parcela")]
        public string? Parcela { get; set; }

        [DisplayName("CPF")]
        public string? CPF { get; set; }

        [DisplayName("Nome Completo")]
        public string? Nome_Completo { get; set; }

        [DisplayName("Tipo do Título")]
        public string? Tipo_Titulo { get; set; }

        [DisplayName("Tipo do Pagamento")]
        public string? Tipo_Pagamento { get; set; }

        [DisplayName("Data de Emissão")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Emissao { get; set; }

        [DisplayName("Data Vencimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Vencimento { get; set; }

        [DisplayName("Valor Original")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Original { get; set; }

        [DisplayName("Valor Devido")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Devido { get; set; }

        [DisplayName("Observação")]
        public string? Observacao { get; set; }

        [DisplayName("Data Pagamento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Pagamento { get; set; }

        [DisplayName("Valor Pago")]
        [DisplayFormat(DataFormatString = "{0:C}", ApplyFormatInEditMode = true)]
        public string? Valor_Pago { get; set; }

        [DisplayName("Pagamento Observações")]
        public string? Pagamento_Observacoes { get; set; }
    }
}
