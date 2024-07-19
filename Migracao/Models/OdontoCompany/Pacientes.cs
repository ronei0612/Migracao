using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{

	[Table("ATIVOS")]
	public class Pacientes
	{

        public int ID { get; set; }
        public string CNPJ_CPF { get; set; }
        public string NOME { get; set; }
        public DateTime DATA { get; set; }
        public string ATIVO { get; set; }
        public string MOTIVO { get; set; }
        public string HORA { get; set; }        
    }
}
