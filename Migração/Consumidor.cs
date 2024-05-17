namespace Migração
{
	internal class Consumidor
	{
		public int ID { get; set; }
		public required int PessoaID { get; set; }
		public Nullable<int> Controle { get; set; }
		public Nullable<int> IndicacaoTipoID { get; set; }
		public string IndicacaoTexto { get; set; }
		public Nullable<int> ConvenioID { get; set; }
		public string ConvenioCartao { get; set; }
		public string ConvenioObservacoes { get; set; }
		public required int EstabelecimentoID { get; set; }
		public Nullable<int> ClienteID { get; set; }
		public Nullable<int> EmpresaID { get; set; }
		public required bool Ativo { get; set; }
		public Nullable<short> SolucaoID { get; set; }
		public required int LoginID { get; set; }
		public required System.DateTime DataInclusao { get; set; }
		public Nullable<System.DateTime> DataUltAlteracao { get; set; }
		public Nullable<System.DateTime> ExclusaoData { get; set; }
		public string ExclusaoMotivo { get; set; }
		public Nullable<int> CampanhaEventoID { get; set; }
		public Nullable<decimal> RendaMensal { get; set; }
		public string TrabalhoEmpresaNome { get; set; }
		public Nullable<bool> AnaliseCreditoAprovacao { get; set; }
		public Nullable<System.DateTime> AnaliseCreditoData { get; set; }
		public Nullable<decimal> AnaliseCreditoValor { get; set; }
		public string AnaliseCreditoResultado { get; set; }
		public Nullable<int> AnaliseCreditoFuncionarioID { get; set; }
		public string CodigoAntigo { get; set; }
		public Nullable<int> PreferenciaMedicoID { get; set; }
		public Nullable<int> RedeID { get; set; }
		public Nullable<int> CadastramentoFuncionarioID { get; set; }
		public Nullable<decimal> SaldoBalancaFinanceira { get; set; }
		public Nullable<int> IndicacaoPessoaID { get; set; }
		public Nullable<int> IndicacaoFuncionarioID { get; set; }
		public Nullable<int> SaldoPontos { get; set; }
		public string BloqueioTexto { get; set; }
		public Nullable<int> BloqueioPessoaID { get; set; }
		public Nullable<byte> BloqueioID { get; set; }
		public Nullable<System.DateTime> BloqueioData { get; set; }
		public Nullable<int> FornecedorID { get; set; }
		public Nullable<int> CidadeID { get; set; }
		public string Observacoes { get; set; }
		public Nullable<int> ContaBancariaID { get; set; }
		public Nullable<byte> TipoID { get; set; }
		public required Nullable<short> LGPDSituacaoID { get; set; }
	}
}
