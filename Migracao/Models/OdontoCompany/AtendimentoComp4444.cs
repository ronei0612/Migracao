using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("ATD444")]
	public class AtendimentoComplemento444
	{
		public int? LANCTO { get; set; }
		public string? CNPJ_CPF { get; set; }
		public string? NOME { get; set; }
		public DateTime? DATA_LANC { get; set; }
		public string? HORA { get; set; }
		public DateTime? DATA_ATEND { get; set; }
		public string? RESP_ATEND { get; set; }
		public string? NOME_RESP_ATEND { get; set; }
		public string? OBS { get; set; }
		public string? TERMINAL { get; set; }
		public string? USUARIO { get; set; }
		public string? HORA_ATEND { get; set; }
		public string? HORA_MARCADA { get; set; }
		public string? RESP_MARCADO { get; set; }
		public string? SENHA_ATEND { get; set; }
		public string? EM_ATENDIMENTO { get; set; }
		public string? NOME_ATEND { get; set; }
		public string? HORA_FINAL { get; set; }
		public int? UNIDADE { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
	}
}
