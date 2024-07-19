using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{

	[Table("MAN101")]
	public class Manuntencao101
	{
		public string? LANCTO { get; set; }
		public int? NUM_MAN { get; set; }
		public string? CNPJ_CPF { get; set; }
		public int? CONTROLE { get; set; }
		public DateTime? DATA_PAGO { get; set; }
		public DateTime? DATA_RETORNO { get; set; }
		public string? RESP_ATEND { get; set; }
		public string? NOME_RESP_ATEND { get; set; }
		public string? OBS_ATEND { get; set; }
		public string? TIPO_ATEND { get; set; }
		public DateTime? DATA_LANC { get; set; }
		public DateTime? DATA_MANUT { get; set; }
		public DateTime? DATA_MODIFICADO { get; set; }
	}
}
