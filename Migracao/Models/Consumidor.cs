namespace Migracao.Models
{
    internal class Consumidor
    {
        public int ID { get; set; }
        public required int PessoaID { get; set; }
        public int? Controle { get; set; }
        public int? IndicacaoTipoID { get; set; }
        public string IndicacaoTexto { get; set; }
        public int? ConvenioID { get; set; }
        public string ConvenioCartao { get; set; }
        public string ConvenioObservacoes { get; set; }
        public required int EstabelecimentoID { get; set; }
        public int? ClienteID { get; set; }
        public int? EmpresaID { get; set; }
        public required bool Ativo { get; set; }
        public short? SolucaoID { get; set; }
        public required int LoginID { get; set; }
        public required DateTime DataInclusao { get; set; }
        public DateTime? DataUltAlteracao { get; set; }
        public DateTime? ExclusaoData { get; set; }
        public string ExclusaoMotivo { get; set; }
        public int? CampanhaEventoID { get; set; }
        public decimal? RendaMensal { get; set; }
        public string TrabalhoEmpresaNome { get; set; }
        public bool? AnaliseCreditoAprovacao { get; set; }
        public DateTime? AnaliseCreditoData { get; set; }
        public decimal? AnaliseCreditoValor { get; set; }
        public string AnaliseCreditoResultado { get; set; }
        public int? AnaliseCreditoFuncionarioID { get; set; }
        public string CodigoAntigo { get; set; }
        public int? PreferenciaMedicoID { get; set; }
        public int? RedeID { get; set; }
        public int? CadastramentoFuncionarioID { get; set; }
        public decimal? SaldoBalancaFinanceira { get; set; }
        public int? IndicacaoPessoaID { get; set; }
        public int? IndicacaoFuncionarioID { get; set; }
        public int? SaldoPontos { get; set; }
        public string BloqueioTexto { get; set; }
        public int? BloqueioPessoaID { get; set; }
        public byte? BloqueioID { get; set; }
        public DateTime? BloqueioData { get; set; }
        public int? FornecedorID { get; set; }
        public int? CidadeID { get; set; }
        public string Observacoes { get; set; }
        public int? ContaBancariaID { get; set; }
        public byte? TipoID { get; set; }
        public required short? LGPDSituacaoID { get; set; }
    }
}
