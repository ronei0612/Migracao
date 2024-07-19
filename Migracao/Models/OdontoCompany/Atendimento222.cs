using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{

	[Table("ATD222")]
	public class Atendimento222
	{
		public string? DOCUMENTO { get; set; }
		public string? CNPJ_CPF { get; set; }
		public string? PRODUTO { get; set; }
		public string? TIPO { get; set; }
		public string? NOME_PRODUTO { get; set; }
		public DateTime? DATA { get; set; }
		public decimal? VALOR { get; set; }
		public string? OBS { get; set; }
		public string? CAMPOX { get; set; }
		public string? RESPONSAVEL { get; set; }
		public DateTime? DATA_ATEND { get; set; }
		public string? RESP_ATEND { get; set; }
		public string? NOME_RESP_ATEND { get; set; }
		public int? NUMERO { get; set; }
		public int? CONTROLE { get; set; }
		public string? CBARRA { get; set; }
		public string? TERMINAL { get; set; }
		public string? USUARIO { get; set; }
		public int? QTDE_SESSAO { get; set; }
		public int? QTDE_SESSAO_ORIG { get; set; }
		public DateTime? MODIFICADO { get; set; }
		public DateTime? DATA_CANCELADO { get; set; }
		public string? USUARIO_CANCELADO { get; set; }
		public DateTime? DATA_INCLUIDO { get; set; }
		public string? USUARIO_INCLUIDO { get; set; }
		public string? FORMA_MEDIDA { get; set; }
		public decimal? QTDE_MEDIDA { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
	}
}
