using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{

	[Table("CED001")]
	public class TabelaPreco  {
		public string CODIGO { get; set; }
		public string? NOME { get; set; }
		public string? UNIDADE { get; set; }
		public string? GRUPO { get; set; }
		public string? LOCAL { get; set; }
		public decimal? VRVENDA { get; set; }
		public decimal? VRCUSTO { get; set; }
		public decimal? VRMEDIO { get; set; }
		public decimal? QTMAX { get; set; }
		public decimal? QTMIN { get; set; }
		public string? OBS { get; set; }
		public string? FAMILIA { get; set; }
		public DateTime? TRANSMISSAO { get; set; }
		public string? FORNEC1 { get; set; }
		public string? FORNEC2 { get; set; }
		public string? FORNEC3 { get; set; }
		public decimal? DESCONTO { get; set; }
		public string? TABELA { get; set; }
		public string? LOTE { get; set; }
		public DateTime? VALIDADE { get; set; }
		public decimal? PERC_FRETE { get; set; }
		public decimal? PERC_IPI { get; set; }
		public decimal? COMISSAO { get; set; }
		public string? ISENTO { get; set; }
		public decimal? VR_ATACADO { get; set; }
		public decimal? ICMS_DENTRO { get; set; }
		public decimal? ICMS_FORA { get; set; }
		public string? OBS2 { get; set; }
		public string? OBS3 { get; set; }
		public decimal? VR_ESPECIAL { get; set; }
		public string? IDENTIFICACAO { get; set; }
		public string? LOJA { get; set; }
		public string? USUARIO { get; set; }
		public DateTime? MODIFICADO { get; set; }
		public DateTime? DT_CADASTRO { get; set; }
		public decimal? ACRESCIMO { get; set; }
		public string? A_VISTA { get; set; }
		public string? COR { get; set; }
		public string? MODELO { get; set; }
		public int? ORCAMENTO { get; set; }
		public decimal? DESCONTO_VISTA { get; set; }
		public string? TIPO_SIMBOLO { get; set; }
		public string? SIMBOLO { get; set; }
		public string? CONVENIO { get; set; }
		public string? COD_ISENTO { get; set; }
		public string? CFOP { get; set; }
		public string? NCM { get; set; }
		public string? CST { get; set; }
		public string? ATIVO { get; set; }
		public string? PARTICULAR { get; set; }
		public string? UNIPLAN { get; set; }
		public string? CSOSN { get; set; }
		public decimal? PIS { get; set; }
		public decimal? COFINS { get; set; }
		public string? PARTMED { get; set; }
		public string? ESTOQUE { get; set; }
		public decimal? VALOR_ANTERIOR { get; set; }
		public string? PROD_MEDICINA { get; set; }
		public string? TIPO_ATEND { get; set; }
		public int? SESSOES { get; set; }
		public string? VENDA_VISTA { get; set; }
		public string? CONSULTA_PARTMED { get; set; }
		public int? CODIGO_PROTETICO { get; set; }
		public int? VALIDADE_DIAS { get; set; }
		public string? VALIDADE_SUGESTAO { get; set; }
		public string? FORMA_MEDIDA { get; set; }
		public string? CODIGO_TUSS { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
	}
}
