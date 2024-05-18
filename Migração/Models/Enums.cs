using System.ComponentModel;

namespace Migração.Models
{
    public enum TitulosEspeciesID
    {
        [Description("Dinheiro")]
        Dinheiro = 1,
        [Description("Boleto Bancário")]
        BoletoBancario = 2,
        [Description("Cheque")]
        Cheque = 4,
        [Description("Cheque pré-datado")]
        ChequePreDatado = 5,
        [Description("Crédito em conta")]
        CreditoEmConta = 6,
        [Description("Cartão de crédito")]
        CartaoCredito = 8,
        [Description("Carteira")]
        Carteira = 11,
        [Description("Cartão de crédito rotativo")]
        CartaoCreditoRotativo = 13,
        [Description("Cartão de crédito")]
        CartaoCreditoParcelado = 14,
        [Description("Carnê")]
        Carne = 17,
        [Description("Depósito em Conta")]
        DepositoEmConta = 31,
        [Description("Cartão em Recorrência")]
        CartaoCreditoRecorrente = 33,
        [Description("Cartão de débito")]
        CartaoDebito = 15,
        [Description("Compensação de boleto")]
        CompesacaoBoleto = 18,
        [Description("Cheque de terceiros")]
        ChequeTerceiros = 20,
        [Description("Liquidação de boleto")]
        LiquidacaoBoleto = 21,
        [Description("Transferência bancária")]
        TransferenciaBancaria = 22,
        [Description("Convênio")]
        Convenio = 30,
        [Description("Débito em conta")]
        DebitoEmConta = 32,
        [Description("Caixa administrativo")]
        CaixaAdmin = 100,
        [Description("CashBack ControleBoletos")]
        CashBackCB = 115
    }

    public enum TituloTransacoes
    {
        Liquidacao = 1,
        PagamentoParcial = 2,
        EncaminhadoProtesto = 3,
        Protestado = 4,
        CobrancaExtraJudicial = 5,
        CobrancaJudicial = 6,
        PagamentoAvulso = 9,
        BaixaDevolução = 11,
        BaixaAcordo = 12,
        BaixaPerda = 13,
        Cancelamento = 90
    }

    public enum TituloSituacoesID
    {
        Normal = 1,
        EncaminhadoProtesto = 3,
        Protestado = 4,
        CobrancaExtraJudicial = 5,
        CobrancaJudicial = 6,
        Cancelamento = 90
    }
    public enum TransacaoTiposID
    {
        All = 0,
        Recebimento = 1,
        Pagamento = 2
    }

}
