using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{

	[Table("CED006")]
	public class FuncionariosCED006
	{
	
    public string? CODIGO { get; set; }

		public string? NOME { get; set; }

		public string? DEPARTAMENTO { get; set; }

		public string? OBS { get; set; }

		public string? CAMPOX { get; set; }

		public decimal? COMISSAO { get; set; }

		public decimal? COTA_MENSAL { get; set; }

		public string? LOJA { get; set; }

		public string? USUARIO { get; set; }

		public DateTime? MODIFICADO { get; set; }

		public string? ATIVO { get; set; }

		public string? AGENDA { get; set; }

		public string? SENHA { get; set; }

		public decimal? COMISSAO_ATEND { get; set; }

		public decimal? HORA_MES { get; set; }

		public decimal? ATEND_HORA { get; set; }

		public string? NOME_COMPLETO { get; set; }

		public string? EMAIL { get; set; }

		public string? TELEFONE { get; set; }

		public string? CRO { get; set; }

		public string? RESP_TECNICO { get; set; }

		public string? ORCAMENTISTA { get; set; }

		public DateTime? ATUALIZADO { get; set; }

		public DateTime? DT_AXON { get; set; }

		public string? AXON_ID { get; set; }
	}

}
