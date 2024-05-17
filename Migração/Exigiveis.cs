namespace Migração
{
	internal class Exigiveis
	{
		public int ID { get; set; }
		public Nullable<int> Documento { get; set; }
		public Nullable<short> Sequencia { get; set; }
		public Nullable<short> Parcelas { get; set; }
		public required int EspecieID { get; set; }
		public Nullable<int> FornecedorID { get; set; }
		public Nullable<int> ConsumidorID { get; set; }
		public Nullable<int> ColaboradorID { get; set; }
		public string OutroCedenteNome { get; set; }
		public required System.DateTime DataEmissao { get; set; }
		public required System.DateTime DataVencimento { get; set; }
		public required System.DateTime DataBaseCalculo { get; set; }
		public Nullable<System.DateTime> DataBaixa { get; set; }
		public required decimal ValorOriginal { get; set; }
		public required decimal ValorDevido { get; set; }
		public Nullable<short> BancoID { get; set; }
		public Nullable<int> PlanoContasID { get; set; }
		public Nullable<int> ContaBancariaID { get; set; }
		public required bool DespesaFixa { get; set; }
		public required int FinanceiroID { get; set; }
		public required short SituacaoID { get; set; }
		public Nullable<int> EstabelecimentoID { get; set; }
		public Nullable<int> EmpresaID { get; set; }
		public Nullable<int> NotaFiscalID { get; set; }
		public string Observacoes { get; set; }
		public required int LoginID { get; set; }
		public required System.DateTime DataInclusao { get; set; }
		public Nullable<System.DateTime> DataUltAlteracao { get; set; }
		public Nullable<int> CompraID { get; set; }
		public Nullable<int> ClienteID { get; set; }
		public Nullable<int> ProteseID { get; set; }
		public Nullable<System.DateTime> ExclusaoData { get; set; }
		public string ExclusaoMotivo { get; set; }
		public Nullable<decimal> ValorBaixa { get; set; }
		public Nullable<int> BaixaID { get; set; }
		public Nullable<int> NotaFiscalNumero { get; set; }
	}
}
