using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("ATD333")]
	public class AtendimentoComplemento333
	{
		public int? LANCTO { get; set; }
		public string? DOCUMENTO { get; set; }
		public string? CNPJ_CPF { get; set; }
		public int? CONTROLE { get; set; }
		public string? PRODUTO { get; set; }
		public string? NOME_PRODUTO { get; set; }
		public int? NUMERO { get; set; }
		public string? OBS { get; set; }
		public DateTime? DATA_ATEND { get; set; }
		public string? RESP_ATEND { get; set; }
		public string? NOME_RESP_ATEND { get; set; }
		public string? NUM_SESSAO_S { get; set; }
		public int? NUM_SESSAO_I { get; set; }
		public string? TERMINAL { get; set; }
		public string? USUARIO { get; set; }
		public string? PARCELA { get; set; }
		public DateTime? DATA_LANC { get; set; }
		public string? HORA { get; set; }
		public string? RETORNO { get; set; }
		public int? LANCTO_ATD444 { get; set; }
		public int? NUM_SESSAO_T { get; set; }
		public int? CODIGO_FASE_PT { get; set; }
		public string? NOME_FASE_PT { get; set; }
		public DateTime? DATA_ENVIO_PT { get; set; }
		public string? USUARIO_ENVIO_PT { get; set; }
		public DateTime? DATA_RETORNO_PT { get; set; }
		public string? USUARIO_RETORNO_PT { get; set; }
		public DateTime? DATA_PAGO_PT { get; set; }
		public string? USUARIO_PAGO_PT { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }

	}

}

