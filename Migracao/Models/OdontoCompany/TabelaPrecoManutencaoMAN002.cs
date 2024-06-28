using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("MAN002")]
	public class TabelaPrecoManutencaoMAN002
	{
		public string? CODIGO { get; set; }
		public string? NOME { get; set; }
		public int? NUM_PARCELAS { get; set; }
		public decimal? VALOR { get; set; }
		public int? ORDEM { get; set; }
		public string? CONVENIO { get; set; }
		public int? CONTROLE { get; set; }
		public decimal? MULTA { get; set; }
		public decimal? JUROS { get; set; }
		public int? MAX_VENCTO { get; set; }
		public string? COD_VENDA_ODC { get; set; }
		public string? ATIVO { get; set; }
		public DateTime? DATA_MODIFICADO { get; set; }
		public int? INTERVALO { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }

	}
}
