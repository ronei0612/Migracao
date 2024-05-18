﻿namespace Migração.Models
{
    internal class Recebiveis
    {
        public int ID { get; set; }
        public long? Documento { get; set; }
        public short? Sequencia { get; set; }
        public short? Parcelas { get; set; }
        public required int EspecieID { get; set; }
        public int? ConsumidorID { get; set; }
        public int? FornecedorID { get; set; }
        public string SacadoNome { get; set; }
        public int? FuncionarioID { get; set; }
        public required DateTime DataEmissao { get; set; }
        public required DateTime DataVencimento { get; set; }
        public required DateTime DataBaseCalculo { get; set; }
        public DateTime? DataBaixa { get; set; }
        public required decimal ValorOriginal { get; set; }
        public required decimal ValorDevido { get; set; }
        public short? BancoID { get; set; }
        public int? PlanoContasID { get; set; }
        public int? ContaBancariaID { get; set; }
        public int? BoletoID { get; set; }
        public required int FinanceiroID { get; set; }
        public bool? Ortodontia { get; set; }
        public bool? Manutencao { get; set; }
        public int? AtendimentoID { get; set; }
        public required short SituacaoID { get; set; }
        public int? EstabelecimentoID { get; set; }
        public int? EmpresaID { get; set; }
        public int? ContratoID { get; set; }
        public int? NotaFiscalID { get; set; }
        public string Observacoes { get; set; }
        public required int LoginID { get; set; }
        public required DateTime DataInclusao { get; set; }
        public DateTime? DataUltAlteracao { get; set; }
        public int? ConvenioID { get; set; }
        public int? ClienteID { get; set; }
        public int? ColaboradorID { get; set; }
        public int? BaixaID { get; set; }
        public decimal? ValorBaixa { get; set; }
        public DateTime? ConsolidacaoComissaoData { get; set; }
        public int? ConsolidacaoComissaoFuncionarioID { get; set; }
        public int? ProteseID { get; set; }
        public int? ProdutosSaidaID { get; set; }
        public int? SMSContratadoID { get; set; }
        public int? LicencaID { get; set; }
        public decimal? ValorDesconto { get; set; }
        public DateTime? ExclusaoData { get; set; }
        public string ExclusaoMotivo { get; set; }
        public int? VendedorID { get; set; }
        public decimal? ValorDevidoReajustado { get; set; }
        public byte? CobrancasEnviadasQtde { get; set; }
        public DateTime? CobrancasEnviadasUltData { get; set; }
        public DateTime? RemessaData { get; set; }
        public int? RemessaID { get; set; }
        public int? ContaBoletoID { get; set; }
        public int? ContaBoletoNossoNumero { get; set; }
        public int? OrcamentoID { get; set; }
        public decimal? DescontoPontualidade { get; set; }
        public int? TissLoteID { get; set; }
        public DateTime? UltimoReajusteData { get; set; }
        public int? UltimoReajustePessoaID { get; set; }
        public DateTime? RegistroData { get; set; }
        public DateTime? ImpressaoBoletoUltimaData { get; set; }
        public int? UltimoNossoNumeroCB { get; set; }
        public int? DebitoAutoID { get; set; }
        public int? RemessaOutrosID { get; set; }
        public int? BotContratadoID { get; set; }
        public bool? Protestado { get; set; }
        public string IdentificadorPEFIN { get; set; }
        public string IdentificadorRemocaoPEFIN { get; set; }
        public Guid? TokenRecorrencia { get; set; }
    }
}
