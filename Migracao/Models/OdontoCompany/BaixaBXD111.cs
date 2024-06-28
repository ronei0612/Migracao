using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("BXD111")]
	public  class BaixaBXD111
	{
		public string? LANCTO { get; set; }
		public string? CGC_CPF { get; set; }
		public string? DOCUMENTO { get; set; }
		public float? VR_PARCELA { get; set; }
		public float? VALOR { get; set; }
		public DateTime? VENCTO { get; set; }
		public DateTime? BAIXA { get; set; }
		public string? CAMPOX { get; set; }
		public DateTime? TRANSMISSAO { get; set; }
		public string? OBS { get; set; }
		public string? TIPO_DOC { get; set; }
		public string? DUPLICATA { get; set; }
		public string? PARCELA { get; set; }
		public string? RESPONSAVEL { get; set; }
		public string? CONTA_CORRENTE { get; set; }
		public string? CONTA_DOCUMENTO { get; set; }
		public string? LOJA { get; set; }
		public string? USUARIO { get; set; }
		public DateTime MODIFICADO { get; set; }
		public string? TERMINAL { get; set; }
		public decimal? VR_CALCULADO { get; set; }
		public string? MOTIVO { get; set; }
		public string? GRUPO { get; set; }
		public string? NOME_GRUPO { get; set; }
		public string? NUM_BANCO { get; set; }
		public string? COD_CAIXA { get; set; }
		public DateTime? DATA_REMESSA { get; set; }
		public short? AGUARDANDO_VINCULO { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
		public int? ID_BAIXAPLANOS { get; set; }
	}

}
