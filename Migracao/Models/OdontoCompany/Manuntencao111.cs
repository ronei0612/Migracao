using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
	[Table("MAN111")]
	public class Manuntencao111 {
		public string? LANCTO { get; set; }

		public string? CNPJ_CPF { get; set; }

		public DateTime? DATA_PAGTO { get; set; }

		public string? TIPO_PAGTO { get; set; }

		public decimal? VALOR { get; set; }

		public string? MES_ANO { get; set; }

		public string? CAMPOX { get; set; }

		public DateTime? DIA_MES_ANO { get; set; }

		public string? RESPONSAVEL { get; set; }

		public DateTime? DATA_ATEND { get; set; }

		public string? RESP_ATEND { get; set; }

		public int? NUM_MAN { get; set; }

		public string? NOME_RESP_ATEND { get; set; }

		public string? OBS { get; set; }

		public DateTime? RETORNO { get; set; }

		public DateTime? DATA_LANC { get; set; }

		public string? HORA { get; set; }

		public DateTime? VENCTO { get; set; }

		public decimal? VALOR_PARCELA { get; set; }

		public string? NOSSO_NUMERO { get; set; }

		public string? TIPO_MAN { get; set; }

		public DateTime? VENCTO_ORIG { get; set; }

		public decimal? VALOR_ORIG { get; set; }

		public string? MOTIVO_ALTERAR { get; set; }

		public string? AUTORIZA_ALTERAR { get; set; }

		public string? MOTIVO_INCLUIR { get; set; }

		public string? AUTORIZA_INCLUIR { get; set; }

		public decimal? VALOR_CALCULADO { get; set; }

		public string? AUTORIZA_RECEBER { get; set; }

		public string? MOTIVO_RECEBER { get; set; }

		public DateTime? EMISSAO { get; set; }

		public string? BANCO { get; set; }

		public DateTime? DATA_ALTERACAO { get; set; }

		public string? OPERACAO { get; set; }

		public string? DOCUMENTO { get; set; }

		public string? NOME_TIPO { get; set; }

		public string? TIPO_COBRANCA { get; set; }

		public string? TERMINAL { get; set; }

		public string? REMESSA { get; set; }

		public string? SENHA_ATEND { get; set; }

		public string? USUARIO { get; set; }

		public string? EM_ATENDIMENTO { get; set; }

		public string? NOME_ATEND { get; set; }

		public short? CALC_JUROS { get; set; }

		public string? USU_VENDA { get; set; }

		public DateTime? COBRADORA { get; set; }

		public DateTime? COBRANCA { get; set; }

		public DateTime? DATA_REMESSA { get; set; }

		public string? SITUACAO_REMESSA { get; set; }

		public DateTime? DATA_MODIFICADO { get; set; }

		public string? CONTA { get; set; }

		public string? NSU_TRANSACAO { get; set; }

		public string? COBRANCA_ORTO { get; set; }

		public DateTime? DATA_REMESSA_CARTAO { get; set; }

		public int? CONTROLE_CARTAO { get; set; }

		public int? UNIDADE { get; set; }

		public string? HORA_ATEND { get; set; }

		public string? HORA_FINAL { get; set; }

		public decimal? DESCONTO_BOLETO { get; set; }

		public short? AGUARDANDO_VINCULO { get; set; }

		public string? ID_PIX { get; set; }

		public DateTime? DT_AXON { get; set; }

		public string? AXON_ID { get; set; }
	}}

