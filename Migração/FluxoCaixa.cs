namespace Migração
{
	internal class FluxoCaixa
	{
		public int ID { get; set; }
		public required byte TipoID { get; set; }
		public required System.DateTime Data { get; set; }
		public Nullable<int> RecebivelID { get; set; }
		public Nullable<int> ExigivelID { get; set; }
		public Nullable<int> FornecedorID { get; set; }
		public Nullable<int> ConsumidorID { get; set; }
		public Nullable<int> ColaboradorID { get; set; }
		public string OutroCedenteNome { get; set; }
		public string OutroSacadoNome { get; set; }
		public required short TransacaoID { get; set; }
		public required int EspecieID { get; set; }
		public Nullable<long> Referencia { get; set; }
		public Nullable<short> EspecieParcelas { get; set; }
		public Nullable<System.DateTime> ReferenciaData { get; set; }
		public required System.DateTime DataBaseCalculo { get; set; }
		public required decimal DevidoValor { get; set; }
		public Nullable<decimal> PagoMulta { get; set; }
		public Nullable<decimal> PagoJuros { get; set; }
		public Nullable<decimal> PagoDespesas { get; set; }
		public Nullable<decimal> PagoDescontos { get; set; }
		public Nullable<decimal> PagoTarifaBancaria { get; set; }
		public required decimal PagoValor { get; set; }
		public Nullable<short> SituacaoID { get; set; }
		public Nullable<int> ContaBancariaID { get; set; }
		public Nullable<int> PlanoContasID { get; set; }
		public Nullable<int> FuncionarioID { get; set; }
		public Nullable<bool> Ortodontia { get; set; }
		public Nullable<bool> Manutencao { get; set; }
		public Nullable<int> ContratoID { get; set; }
		public Nullable<int> BoletoID { get; set; }
		public Nullable<short> BancoID { get; set; }
		public Nullable<int> NotaFiscal { get; set; }
		public string Observacoes { get; set; }
		public required int FinanceiroID { get; set; }
		public required int EstabelecimentoID { get; set; }
		public Nullable<int> EmpresaID { get; set; }
		public required int LoginID { get; set; }
		public required System.DateTime DataInclusao { get; set; }
		public Nullable<System.DateTime> DataUltAlteracao { get; set; }
		public Nullable<decimal> ValorTarifa { get; set; }
		public Nullable<System.DateTime> ConsolidacaoComissaoData { get; set; }
		public Nullable<int> ConsolidacaoComissaoFuncionarioID { get; set; }
		public Nullable<int> ClienteID { get; set; }
		public Nullable<System.DateTime> EspeciePreDatado { get; set; }
		public Nullable<int> CaixaFechamentoID { get; set; }
		public string EspecieSacadoNome { get; set; }
		public Nullable<int> TarifaID { get; set; }
		public Nullable<short> EspecieParcela { get; set; }
		public Nullable<int> ReferenciaFluxoID { get; set; }
		public Nullable<decimal> PagoTotalParcelas { get; set; }
		public Nullable<int> NotaFiscalID { get; set; }
		public Nullable<int> ServicoID { get; set; }
		public Nullable<int> ConvenioID { get; set; }
		public Nullable<System.DateTime> DataOcorrencia { get; set; }
		public Nullable<bool> DespesaFixa { get; set; }
		public Nullable<byte> CartaoTransacaoStatusID { get; set; }
		public Nullable<short> BandeiraID { get; set; }
		public Nullable<int> VendedorID { get; set; }
		public string CartaoTransacaoTID { get; set; }
		public Nullable<System.DateTime> ExclusaoData { get; set; }
		public string ExclusaoMotivo { get; set; }
		public Nullable<int> DebitoAutoID { get; set; }
		public Nullable<System.DateTime> CompensacaoData { get; set; }
		public Nullable<decimal> CompensacaoValor { get; set; }
		public Nullable<int> CashbackUtilizadoID { get; set; }
	}
}
