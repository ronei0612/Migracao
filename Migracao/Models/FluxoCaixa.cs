namespace Migracao.Models
{
    public class FluxoCaixa
    {
        public string? Nome { get; set; }
        public int ID { get; set; }
        public required byte TipoID { get; set; }
        public required DateTime Data { get; set; }
        public int? RecebivelID { get; set; }
        public int? ExigivelID { get; set; }
        public int? FornecedorID { get; set; }
        public int? ConsumidorID { get; set; }
        public int? ColaboradorID { get; set; }
        public string OutroCedenteNome { get; set; }
        public string OutroSacadoNome { get; set; }
        public required short TransacaoID { get; set; }
        public required int EspecieID { get; set; }
        public long? Referencia { get; set; }
        public short? EspecieParcelas { get; set; }
        public DateTime? ReferenciaData { get; set; }
        public required DateTime DataBaseCalculo { get; set; }
        public required decimal DevidoValor { get; set; }
        public decimal? PagoMulta { get; set; }
        public decimal? PagoJuros { get; set; }
        public decimal? PagoDespesas { get; set; }
        public decimal? PagoDescontos { get; set; }
        public decimal? PagoTarifaBancaria { get; set; }
        public required decimal PagoValor { get; set; }
        public short? SituacaoID { get; set; }
        public int? ContaBancariaID { get; set; }
        public int? PlanoContasID { get; set; }
        public int? FuncionarioID { get; set; }
        public bool? Ortodontia { get; set; }
        public bool? Manutencao { get; set; }
        public int? ContratoID { get; set; }
        public int? BoletoID { get; set; }
        public short? BancoID { get; set; }
        public int? NotaFiscal { get; set; }
        public string Observacoes { get; set; }
        public required int FinanceiroID { get; set; }
        public required int EstabelecimentoID { get; set; }
        public int? EmpresaID { get; set; }
        public required int LoginID { get; set; }
        public required DateTime DataInclusao { get; set; }
        public DateTime? DataUltAlteracao { get; set; }
        public decimal? ValorTarifa { get; set; }
        public DateTime? ConsolidacaoComissaoData { get; set; }
        public int? ConsolidacaoComissaoFuncionarioID { get; set; }
        public int? ClienteID { get; set; }
        public DateTime? EspeciePreDatado { get; set; }
        public int? CaixaFechamentoID { get; set; }
        public string EspecieSacadoNome { get; set; }
        public int? TarifaID { get; set; }
        public short? EspecieParcela { get; set; }
        public int? ReferenciaFluxoID { get; set; }
        public decimal? PagoTotalParcelas { get; set; }
        public int? NotaFiscalID { get; set; }
        public int? ServicoID { get; set; }
        public int? ConvenioID { get; set; }
        public DateTime? DataOcorrencia { get; set; }
        public bool? DespesaFixa { get; set; }
        public byte? CartaoTransacaoStatusID { get; set; }
        public short? BandeiraID { get; set; }
        public int? VendedorID { get; set; }
        public string CartaoTransacaoTID { get; set; }
        public DateTime? ExclusaoData { get; set; }
        public string ExclusaoMotivo { get; set; }
        public int? DebitoAutoID { get; set; }
        public DateTime? CompensacaoData { get; set; }
        public decimal? CompensacaoValor { get; set; }
        public int? CashbackUtilizadoID { get; set; }
    }
}
