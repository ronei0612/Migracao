using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class Recebivel
    {
        public string CNPJ_CPF { get; set; }
        public string Nome { get; set; }
        public string Observacao { get; set; }
        public string Numero_Controle { get; set; }
        public decimal Valor_Devido { get; set; }
        public DateTime Data_Vencimento { get; set; }
        public DateTime Emissao { get; set; }
        public string Duplicata { get; set; }
        public int Parcela { get; set; }
        public string Tipo_Pagamento { get; set; }
        public decimal Valor_Original { get; set; }
        public DateTime Vencimento_Original { get; set; }
        public string Situacao { get; set; }
        public string Nome_Grupo { get; set; }
        public int Ordem { get; set; }
    }
}
