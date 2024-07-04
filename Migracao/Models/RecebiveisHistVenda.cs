using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class RecebiveisHistVenda
    {
        public string? Paciente_CPF { get; set; }
        public string? Nome_Paciente { get; set; }
        public string? Observacao_Recebivel { get; set; }
        public string? Documento_Ref { get; set; }
        public decimal? Valor_Original { get; set; }
        public DateTime? Vencimento { get; set; }
        public string? Recebivel { get; set; }
        public string? Emissao { get; set; }
        public string? Recebivel_Exigivel { get; set; }
    }
}
