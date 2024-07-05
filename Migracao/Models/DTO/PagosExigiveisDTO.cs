using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class PagosExigiveisDTO
    {
        public PagosExigiveisDTO()
        {
                
        }

        public PagosExigiveisDTO(List<Recebidos> recebidos)
        {

        }

        //Documento	Parcela	CPF	Nome Completo	Tipo do Título	Tipo do Pagamento	
        //Data de Emissão	Data Vencimento	Valor Original	Valor Devido	Observação	
        //Data Pagamento	Valor Pago	Pagamento Observações

        [DisplayName("Documento")]
        public string? Numero_Controle { get; set; }

        [DisplayName("Parcela")]
        public string? Parcela { get; set; }

        [DisplayName("CPF")]
        public string? CPF { get; set; }

        [DisplayName("Nome Completo")]
        public string? Nome { get; set; }


        [DisplayName("Tipo do Título")]
        public string? Recebivel_Exigivel { get; set; }


        [DisplayName("Tipo do Pagamento")]
        public string? Tipo_Pagamento { get; set; }


        [DisplayName("Data de Emissão")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Emissao { get; set; }


        [DisplayName("Data Vencimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Vencimento { get; set; }


        [DisplayName("Valor Original")]
        public string? Valor_Original { get; set; }


        [DisplayName("Valor Devido")]
        public string? Valor_Devido { get; set; }


        [DisplayName("Observação Recebido")]
        [MaxLength(512)]
        public string? Observacao_Recebido { get; set; }

        [DisplayName("Data do Pagamento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Pagamento { get; set; }


        [DisplayName("Valor Pago")]
        public string? Valor_Pago { get; set; }

        [DisplayName("Pagamento Observações")]
        public string? Pagamento_Observacoes { get; set; }


        //[DisplayName("Prazo")]
        //public string? Prazo { get; set; }
       
        //[DisplayName("Vencimento Recebível")]
        //[DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        //public string? Vencimento_Recebivel { get; set; }

        //[DisplayName("Duplicata")]
        //public string? Duplicata { get; set; }
       

        //[DisplayName("Tipo Espécie Pagamento")]
        //public string? Tipo_Especie_Pagamento { get; set; }

        //[DisplayName("Espécie Pagamento")]
        //public string? Especie_Pagamento { get; set; }
    }
}
