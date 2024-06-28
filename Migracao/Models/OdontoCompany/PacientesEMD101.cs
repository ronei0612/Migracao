using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{

	[Table("EMD101")]
	public  class PacientesEMD101
	{
		public string? CODIGO { get; set; }
		public string? NOME { get; set; }
		public string? F_OU_J { get; set; }
		public string? CGC_CPF { get; set; }
		public string? INSC_RG { get; set; }
		public string? CLIENTE { get; set; }
		public string? FORNECEDOR { get; set; }
		public string? PRESTADOR { get; set; }
		public DateTime? DT_CADASTRO { get; set; }
		public DateTime? DT_ULTMOV { get; set; }
		public DateTime? DT_NASCIMENTO { get; set; }
		public float? VR_LIMITE { get; set; }
		public string? OBS1 { get; set; }
		public string? BANCO { get; set; }
		public string? AGENCIA { get; set; }
		public string? CONTA { get; set; }
		public string? ENDERECO { get; set; }
		public string? BAIRRO { get; set; }
		public string? CEP { get; set; }
		public string? CIDADE { get; set; }
		public string? ESTADO { get; set; }
		public string? FONE1 { get; set; }
		public string? FONE2 { get; set; }
		public string? FAX { get; set; }
		public string? CELULAR { get; set; }
		public string? EMAIL { get; set; }
		public string? BANCO2 { get; set; }
		public string? AGENCIA2 { get; set; }
		public string? CONTA2 { get; set; }
		public DateTime? TRANSMISSAO { get; set; }
		public string? ONDE_TRABALHA { get; set; }
		public string? FUNCAO { get; set; }
		public DateTime? ADMISSAO { get; set; }
		public float? RENDA_MES { get; set; }
		public string? CAIXA_POSTAL { get; set; }
		public string? CLASSE { get; set; }
		public string? PAI { get; set; }
		public string? MAE { get; set; }
		public string? CONJUGE { get; set; }
		public string? QTDE_DEPENDENTES { get; set; }
		public string? END_TRAB { get; set; }
		public string? FONE_TRAB { get; set; }
		public string? NOME_REF_1 { get; set; }
		public string? FONE_REF_1 { get; set; }
		public string? NOME_REF_2 { get; set; }
		public string? FONE_REF_2 { get; set; }
		public string? NOME_FIA { get; set; }
		public string? PARENTESCO_FIA { get; set; }
		public string? CPF_FIA { get; set; }
		public string? RG_FIA { get; set; }
		public string? FONE_FIA { get; set; }
		public string? ENDERECO_FIA { get; set; }
		public string? BAIRRO_FIA { get; set; }
		public string? CIDADE_FIA { get; set; }
		public string? CEP_FIA { get; set; }
		public string? ESTADO_FIA { get; set; }
		public string? PROFISSAO_FIA { get; set; }
		public decimal? RENDA_FIA { get; set; }
		public DateTime? MODIFICADO { get; set; }
		public string? USUARIO { get; set; }
		public string? LOJA { get; set; }
		public string? NUM_FICHA { get; set; }
		public string? DEPENDENTE { get; set; }
		public string? TITULAR { get; set; }
		public string? NUM_CONVENIO { get; set; }
		public string? NAO_AUTORIZADO { get; set; }
		public DateTime? DT_NASC_FIA { get; set; }
		public string? NUM_BLOQUEIO { get; set; }
		public DateTime? DATA_BLOQUEIO { get; set; }
		public string? USU_BLOQUEIO { get; set; }
		public string? NUM_ENDERECO { get; set; }
		public string? HIST_BLOQUEIO { get; set; }
		public string? COD_MUNICIPIO { get; set; }
		public string? COD_UF { get; set; }
		public int? COD_VENDEDOR { get; set; }
		public string? CLIENTE_PRAMELHOR { get; set; }
		public string? USU_CADASTRO { get; set; }
		public string? COD_PRAMELHOR { get; set; }
		public string? CNPJ_CPF_VALIDO { get; set; }
		public string? INDICACAO { get; set; }
		public string? PROFISSAO { get; set; }
		public string? SEXO_M_F { get; set; }
		public string? CODIGO_CLIENTE { get; set; }
		public DateTime? DATA_DEP_EXCLUIDO { get; set; }
		public string? TITULAR_DEP_EXCLUIDO { get; set; }
		public DateTime? DT_NASCIMENTO_DEP { get; set; }
		public string? CODIGO_VALIDADE { get; set; }
		public string? NOME_VALIDADE { get; set; }
		public DateTime? DATA_VALIDADE { get; set; }
		public string? USUARIO_VALIDADE { get; set; }
		public string? OBS_VALIDADE { get; set; }
		public int? CODIGO_INDICACAO { get; set; }
		public string? CPF_INDICACAO { get; set; }
		public short? CATEGORIA { get; set; }
		public string? COMPLEMENTO { get; set; }
		public short? INSTITUTO_ODC { get; set; }
		public string? LGPD_USUARIO { get; set; }
		public DateTime? LGPD_DATA_HORA { get; set; }
		public short? LGPD_CPF { get; set; }
		public short? LGPD_TELEFONE { get; set; }
		public short? LGPD_MENSAGEM { get; set; }
		public short? LGPD_IMAGEM { get; set; }
		public string? PROTETICO { get; set; }
		public string? PROTETICO_ATIVO { get; set; }
		public DateTime? DT_AXON { get; set; }
		public string? AXON_ID { get; set; }
		public string? ID_DRCASH { get; set; }
		public DateTime? DATA_APROVACAO_DRCASH { get; set; }
		public decimal? VALOR_MAXIMO_DRCASH { get; set; }
	}
}
