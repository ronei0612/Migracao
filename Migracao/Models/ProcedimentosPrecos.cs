using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class ProcedimentosPrecos
    {
        public string? Codigo { get; set; }            
        public string? Procedimento_Nome { get; set; }
        public string? Tabela { get; set; }
        public string? Abreviacao { get; set; }        
        public string? TUSS { get; set; }              
        public decimal? Preco { get; set; }            
        public string? Ativo { get; set; }               
        public string? OBS { get; set; }               
        public string? Particular { get; set; }
        public string? Codigo_Grupo { get; set; }
        public string? Nome_Grupo { get; set; }  
        public string? Usuario { get; set; }
    }
}
