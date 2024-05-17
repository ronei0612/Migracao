namespace Migração
{
	internal class Recebiveis
	{
		public int ID { get; set; }
		public Nullable<long> Documento { get; set; }
		public Nullable<short> Sequencia { get; set; }
		public Nullable<short> Parcelas { get; set; }
		public required int EspecieID { get; set; }
		public Nullable<int> ConsumidorID { get; set; }
		public Nullable<int> FornecedorID { get; set; }
		public string SacadoNome { get; set; }
		public Nullable<int> FuncionarioID { get; set; }
		public required System.DateTime DataEmissao { get; set; }
		public required System.DateTime DataVencimento { get; set; }
		public required System.DateTime DataBaseCalculo { get; set; }
		public Nullable<System.DateTime> DataBaixa { get; set; }
		public required decimal ValorOriginal { get; set; }
		public required decimal ValorDevido { get; set; }
		public Nullable<short> BancoID { get; set; }
		public Nullable<int> PlanoContasID { get; set; }
		public Nullable<int> ContaBancariaID { get; set; }
		public Nullable<int> BoletoID { get; set; }
		public required int FinanceiroID { get; set; }
		public Nullable<bool> Ortodontia { get; set; }
		public Nullable<bool> Manutencao { get; set; }
		public Nullable<int> AtendimentoID { get; set; }
		public required short SituacaoID { get; set; }
		public Nullable<int> EstabelecimentoID { get; set; }
		public Nullable<int> EmpresaID { get; set; }
		public Nullable<int> ContratoID { get; set; }
		public Nullable<int> NotaFiscalID { get; set; }
		public string Observacoes { get; set; }
		public required int LoginID { get; set; }
		public required System.DateTime DataInclusao { get; set; }
		public Nullable<System.DateTime> DataUltAlteracao { get; set; }
		public Nullable<int> ConvenioID { get; set; }
		public Nullable<int> ClienteID { get; set; }
		public Nullable<int> ColaboradorID { get; set; }
		public Nullable<int> BaixaID { get; set; }
		public Nullable<decimal> ValorBaixa { get; set; }
		public Nullable<System.DateTime> ConsolidacaoComissaoData { get; set; }
		public Nullable<int> ConsolidacaoComissaoFuncionarioID { get; set; }
		public Nullable<int> ProteseID { get; set; }
		public Nullable<int> ProdutosSaidaID { get; set; }
		public Nullable<int> SMSContratadoID { get; set; }
		public Nullable<int> LicencaID { get; set; }
		public Nullable<decimal> ValorDesconto { get; set; }
		public Nullable<System.DateTime> ExclusaoData { get; set; }
		public string ExclusaoMotivo { get; set; }
		public Nullable<int> VendedorID { get; set; }
		public Nullable<decimal> ValorDevidoReajustado { get; set; }
		public Nullable<byte> CobrancasEnviadasQtde { get; set; }
		public Nullable<System.DateTime> CobrancasEnviadasUltData { get; set; }
		public Nullable<System.DateTime> RemessaData { get; set; }
		public Nullable<int> RemessaID { get; set; }
		public Nullable<int> ContaBoletoID { get; set; }
		public Nullable<int> ContaBoletoNossoNumero { get; set; }
		public Nullable<int> OrcamentoID { get; set; }
		public Nullable<decimal> DescontoPontualidade { get; set; }
		public Nullable<int> TissLoteID { get; set; }
		public Nullable<System.DateTime> UltimoReajusteData { get; set; }
		public Nullable<int> UltimoReajustePessoaID { get; set; }
		public Nullable<System.DateTime> RegistroData { get; set; }
		public Nullable<System.DateTime> ImpressaoBoletoUltimaData { get; set; }
		public Nullable<int> UltimoNossoNumeroCB { get; set; }
		public Nullable<int> DebitoAutoID { get; set; }
		public Nullable<int> RemessaOutrosID { get; set; }
		public Nullable<int> BotContratadoID { get; set; }
		public Nullable<bool> Protestado { get; set; }
		public string IdentificadorPEFIN { get; set; }
		public string IdentificadorRemocaoPEFIN { get; set; }
		public Nullable<System.Guid> TokenRecorrencia { get; set; }
	}
}
