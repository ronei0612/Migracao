namespace Migracao.Models
{
    internal class Exigiveis
    {
        public int ID { get; set; }
        public int? Documento { get; set; }
        public short? Sequencia { get; set; }
        public short? Parcelas { get; set; }
        public required int EspecieID { get; set; }
        public int? FornecedorID { get; set; }
        public int? ConsumidorID { get; set; }
        public int? ColaboradorID { get; set; }
        public string OutroCedenteNome { get; set; }
        public required DateTime DataEmissao { get; set; }
        public required DateTime DataVencimento { get; set; }
        public required DateTime DataBaseCalculo { get; set; }
        public DateTime? DataBaixa { get; set; }
        public required decimal ValorOriginal { get; set; }
        public required decimal ValorDevido { get; set; }
        public short? BancoID { get; set; }
        public int? PlanoContasID { get; set; }
        public int? ContaBancariaID { get; set; }
        public required bool DespesaFixa { get; set; }
        public required int FinanceiroID { get; set; }
        public required short SituacaoID { get; set; }
        public int? EstabelecimentoID { get; set; }
        public int? EmpresaID { get; set; }
        public int? NotaFiscalID { get; set; }
        public string Observacoes { get; set; }
        public required int LoginID { get; set; }
        public required DateTime DataInclusao { get; set; }
        public DateTime? DataUltAlteracao { get; set; }
        public int? CompraID { get; set; }
        public int? ClienteID { get; set; }
        public int? ProteseID { get; set; }
        public DateTime? ExclusaoData { get; set; }
        public string ExclusaoMotivo { get; set; }
        public decimal? ValorBaixa { get; set; }
        public int? BaixaID { get; set; }
        public int? NotaFiscalNumero { get; set; }
    }
}
