using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("MAN001")]
	public class Manuntencao001
	{

    public string? CNPJ_CPF { get; set; }

		public int? QTDE_MANUT { get; set; }

		public string? OBS { get; set; }

		public string? ATIVO { get; set; }

		public DateTime? INICIO_MANUT { get; set; }

		public DateTime? FINAL_MANUT { get; set; }

		public DateTime? MODIFICADO { get; set; }

		public string? USUARIO { get; set; }

		public DateTime? DATA_ATIVO { get; set; }

		public string? USU_ATIVO { get; set; }

		public string? USU_INSTALACAO { get; set; }

		public string? USU_RETIRADA { get; set; }

		public string? MOTIVO_ATIVO { get; set; }

		public string? DIAGNOSTICO { get; set; }

		public string? PROGNOSTICO { get; set; }

		public string? OBS_CLASSE { get; set; }

		public string? CLASSE { get; set; }

		public DateTime? DATA_ALTERACAO { get; set; }

		public string? HORA_ALTERACAO { get; set; }

		public string? USUARIO_ALTERACAO { get; set; }

		public DateTime? DATA_MODIFICADO { get; set; }

		public DateTime? DT_AXON { get; set; }

		public string? AXON_ID { get; set; }

	}
}
