using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{

	[Table("CRD013")]
	public class EspecieCRD013
	{
		public string? CODIGO { get; set; }
		public string? NOME { get; set; }
		public string? CAMPOX { get; set; }
		public string? HISTORICO { get; set; }
		public string? GRAVAR_CAIXA { get; set; }
		public string? GRAVAR_DUPLICATA { get; set; }
		public string? BAIXAR_AUTOMATICO { get; set; }
		public int? DIAS_PRAZO { get; set; }
		public string? CONTA { get; set; }
		public string? SCUSTO { get; set; }
		public string? BAIXAR_GRAVAR_CAIXA { get; set; }
		public int? CARENCIA { get; set; }
		public float? PERC_ACRESCIMO { get; set; }
		public float? PERC_DESCONTO { get; set; }
		public float? PERC_JUROS { get; set; }
		public float? MULTA { get; set; }
		public int? QTDE_PARCELAS { get; set; }
		public string? CODIGO_CAIXA { get; set; }
		public float? PRO_RATA { get; set; }
		public float? TAXA_CARTAO { get; set; }
		public string? RESUMIR_BOBINA { get; set; }
		public float? TARIFA { get; set; }
		public string? RESUMIR_F767 { get; set; }
		public string? ECF { get; set; }
		public int? DIAS_CC { get; set; }
		public string? CODIGO_ENTRADA { get; set; }
		public string? CARNE { get; set; }
		public string? DUPLICATA { get; set; }
		public string? INICIAIS_FXD111 { get; set; }
		public string? GRAVAR_CC_NA_EMISSAO { get; set; }
		public string? CODIGO_RETIRADA { get; set; }
		public decimal? ENCARGOS { get; set; }
		public decimal? PERC_COMISSAO { get; set; }
		public decimal? DESCONTO_MAXIMO_VENDA { get; set; }
		public string? LOJA { get; set; }
		public string? USUARIO { get; set; }
		public DateTime? MODIFICADO { get; set; }
		public string? ATIVO { get; set; }
		public string? OBRIGAR_DIAS_ADICIONAIS { get; set; }
		public int? MAX_DIAS_PRIMEIRA { get; set; }
		public decimal? ACRESCIMO_MAXIMO_VENDA { get; set; }
		public string? OBRIGAR_ACRESCIMO { get; set; }
		public string? OBRIGAR_DESCONTO { get; set; }
		public string? OBRIGAR_ENTRADA { get; set; }
		public string? A_VISTA { get; set; }
		public string? NOME_BOBINA { get; set; }
		public string? GRAVAR_PAGAR { get; set; }
		public string? GRUPO { get; set; }
		public string? NOME_GRUPO { get; set; }
		public string? BAIXAR_RECEBER { get; set; }
		public string? BAIXAR_PAGAR { get; set; }
		public string? GRAVAR_MANUTENCAO { get; set; }
		public string? GRAVAR_ESTOQUE { get; set; }
		public string? VER_CONVENIO { get; set; }
		public string? FRANQUIA { get; set; }
		public int? ORDEM { get; set; }
		public string? BAIXA_BANCO { get; set; }
		public string? MANUTENCAO_BANCO { get; set; }
		public string? UNIPLAN { get; set; }
		public string? VENDA_CONVENIO { get; set; }
		public decimal? VLR_TITULAR { get; set; }
		public decimal? VLR_DEPENDENTE { get; set; }
		public string? COD_VENDA_ODC { get; set; }
		public string? ORTODONTIA { get; set; }
		public string? PARTMED { get; set; }
		public string? BAIXA_COBRADORA { get; set; }
		public string? TIPO_CONTRATO { get; set; }
		public string? TAXA_ADESAO { get; set; }
		public string? RENOVACAO { get; set; }
		public string? BAIXA_CARTAO { get; set; }
		public string? BAIXA_ORTO { get; set; }
		public string? CONSULTA_PARTMED { get; set; }
		public int? QTDE_CONSULTAS { get; set; }
		public string? BAIXAR_CODIGO { get; set; }
		public string? GERAR_CHEQUE { get; set; }
		public string? GERAR_TAXA_ADESAO { get; set; }
		public string? ESTORNO_ROYALTIES { get; set; }
		public short? FORMA_PAGAMENTO { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
		public int? GRUPO_FORMA_PAGTO { get; set; }
	}
}
