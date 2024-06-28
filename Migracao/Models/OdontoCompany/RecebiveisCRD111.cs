using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("CRD111")]
	public  class RecebiveisCRD111
	{
		
		public string? CGC_CPF { get; set; }
		public string? EMITENTE { get; set; }
		public string? DOCUMENTO { get; set; }
		public float? VALOR { get; set; }
		public int? PRAZO { get; set; }
		public DateTime? VENCTO { get; set; }
		public DateTime? EMISSAO { get; set; }
		public string? CONTA { get; set; }
		public string? BANCO { get; set; }
		public string? AGENCIA { get; set; }
		public string? TIPO_DOC { get; set; }
		public string? PERIODO { get; set; }
		public string? FILIAL { get; set; }
		public string? OBS { get; set; }
		public string? CAMPOX { get; set; }
		public string? BANDA1 { get; set; }
		public string? BANDA2 { get; set; }
		public string? BANDA3 { get; set; }
		public DateTime? TRANSMISSAO { get; set; }
		public string? SITUACAO { get; set; }
		public string? GEROU_TRANSMISSAO { get; set; }
		public string? RECEBEU_TRANSMISSAO { get; set; }
		public string? ALINEA { get; set; }
		public DateTime? DEVOLUCAO { get; set; }
		public DateTime? REAPRESENTOU { get; set; }
		public int? SEQ_ALINEA11 { get; set; }
		public string? LOTE { get; set; }
		public DateTime? BAIXA { get; set; }
		public string? CHEQUE_BAIXA { get; set; }
		public float? DESCONTOS { get; set; }
		public float? JUROS { get; set; }
		public string? NOSSONUMERO { get; set; }
		public string? RESPONSAVEL { get; set; }
		public float? TOTAL { get; set; }
		public float? MULTA { get; set; }
		public string? DUPLICATA { get; set; }
		public string? PARCELA { get; set; }
		public decimal? ENCARGOS { get; set; }
		public decimal? VALOR_VENDA { get; set; }
		public string? LOJA { get; set; }
		public string? USUARIO { get; set; }
		public DateTime? MODIFICADO { get; set; }
		public DateTime? DATA_ENV_CART { get; set; }
		public DateTime? DATA_RET_CART { get; set; }
		public DateTime? DATA_ENV_SCPC { get; set; }
		public DateTime? DATA_RET_SCPC { get; set; }
		public string? TERMINAL { get; set; }
		public DateTime VENCTO_ORIG { get; set; }
		public float? VALOR_ORIG { get; set; }
		public string? CHEQUE { get; set; }
		public string? TITULO { get; set; }
		public string? GRUPO { get; set; }
		public string? NOME_GRUPO { get; set; }
		public string? MOTIVO { get; set; }
		public string? REMESSA { get; set; }
		public string? NUM_BANCO { get; set; }
		public string? TIPO_COBRANCA { get; set; }
		public float? ORDEM { get; set; }
		public string? LOCAL { get; set; }
		public string? NOME_LOCAL { get; set; }
		public short? CALC_JUROS { get; set; }
		public DateTime? COBRADORA { get; set; }
		public string? COBRANCA { get; set; }
		public DateTime? DATA_REMESSA { get; set; }
		public string? SITUACAO_REMESSA { get; set; }
		public string? NSU_TRANSACAO { get; set; }
		public int? CONTROLE_CARTAO { get; set; }
		public decimal? DESCONTO_BOLETO { get; set; }
		public string? ID_PIX { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
		public string? CODIGO_TUSS { get; set; }
		public int? RecebivelID { get; set; }
		public int? ConsumidorId { get; set; }
		
	}
}
