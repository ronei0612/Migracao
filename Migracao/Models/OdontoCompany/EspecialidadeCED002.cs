using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("CED002")]
	public class EspecialidadeCED002
	{
		 public string? CODIGO { get; set; }
		public string? NOME { get; set; }
		public string? CAMPOX { get; set; }
		public DateTime? TRANSMISSAO { get; set; }
		public string? LOJA { get; set; }
		public string? USUARIO { get; set; }
		public DateTime? MODIFICADO { get; set; }
		public string? ATIVO { get; set; }
		public DateTime DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
	}
}
