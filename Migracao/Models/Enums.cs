﻿using System.ComponentModel;

namespace Migracao.Models
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

	public enum EnderecoTipos
	{
		Residencial = 1,
		Principal = 2,
		Comercial = 3,
		Cobranca = 4,
		Correspondencia = 8,
		Outro = 99
	}

	public enum EstadoCivilIDs
	{
		Solteiro = 1,
		Casado = 2,
		Separado = 3,
		Divorciado = 4,
		Viuvo = 7
	}

	public enum FoneTipos
	{
		Residencial = 1,
		Principal = 2,
		Celular = 3,
		Nextel = 5,
		Alternativo = 6,
		Comercial = 7,
		Cobrança = 8,
		Fax = 9,
		Zero800 = 80,
		Outros = 99
	}

	public enum LogradouroTipos
	{
		Acesso = 1,
		Adro = 2,
		Alameda = 4,
		Alto = 5,
		Atalho = 7,
		Av = 8,
		Balneário = 9,
		Belvedere = 10,
		Beco = 11,
		Bloco = 12,
		Bosque = 13,
		Boulevard = 14,
		Baixa = 15,
		Cais = 16,
		Caminho = 17,
		Chapadão = 19,
		Conjunto = 20,
		Colônia = 21,
		Corredor = 22,
		Campo = 23,
		Córrego = 24,
		Desvio = 27,
		Distrito = 28,
		Escada = 30,
		Estrada = 31,
		Estação = 32,
		Estádio = 33,
		Favela = 36,
		Fazenda = 37,
		Ferrovia = 38,
		Fonte = 39,
		Feira = 40,
		Forte = 43,
		Galeria = 45,
		Granja = 46,
		Ilha = 50,
		Jardim = 52,
		Ladeira = 53,
		Largo = 54,
		Lagoa = 55,
		Loteamento = 56,
		Morro = 59,
		Monte = 60,
		Paralela = 62,
		Passeio = 63,
		Pátio = 64,
		Praça = 65,
		Parada = 67,
		Praia = 70,
		Prolongamento = 71,
		Parque = 72,
		Passarela = 73,
		Passagem = 74,
		Ponte = 76,
		Quadra = 77,
		Quinta = 79,
		Rua = 81,
		Ramal = 82,
		Recanto = 87,
		Retiro = 88,
		Reta = 89,
		RodoviaFederal = 90,
		Retorno = 91,
		Sítio = 92,
		Servidão = 94,
		Setor = 95,
		Subida = 96,
		Trincheira = 97,
		Terminal = 98,
		Trevo = 99,
		Travessa = 100,
		Via = 101,
		Viaduto = 103,
		Vila = 104,
		Viela = 105,
		Vale = 106,
		ZigueZague = 108,
		Linha = 187,
		Povoado = 188,
		Trecho = 452,
		Vereda = 453,
		Artéria = 465,
		Elevada = 468,
		Porto = 469,
		Balão = 470,
		Paradouro = 471,
		Área = 472,
		Jardinete = 473,
		Esplanada = 474,
		Quintas = 475,
		Rotula = 476,
		Marina = 477,
		Descida = 478,
		Circular = 479,
		Unidade = 480,
		Chácara = 481,
		Rampa = 482,
		Ponta = 483,
		ViaDePedestre = 484,
		Condomínio = 485,
		Habitacional = 486,
		Residencial = 487,
		Canal = 495,
		Buraco = 496,
		Módulo = 497,
		Estância = 498,
		Lago = 499,
		Núcleo = 500,
		Aeroporto = 501,
		PassagemSubterrânea = 502,
		ComplexoViário = 503,
		PraçaDeEsportes = 504,
		ViaElevada = 505,
		Rotatória = 506,
		PrimeiraTravessa = 507,
		SegundaTravessa = 508,
		TerceiraTravessa = 509,
		QuartaTravessa = 510,
		QuintaTravessa = 511,
		SextaTravessa = 512,
		SétimaTravessa = 513,
		OitavaTravessa = 514,
		NonaTravessa = 515,
		DécimaTravessa = 516,
		DécimaPrimeiraTravessa = 517,
		DécimaSegundaTravessa = 518,
		DécimaTerceiraTravessa = 519,
		DécimaQuartaTravessa = 520,
		DécimaQuintaTravessa = 521,
		DécimaSextaTravessa = 522,
		PrimeiroAlto = 523,
		SegundoAlto = 524,
		TerceiroAlto = 525,
		QuartoAlto = 526,
		QuintoAlto = 527,
		PrimeiroBeco = 528,
		SegundoBeco = 529,
		TerceiroBeco = 530,
		QuartoBeco = 531,
		QuintoBeco = 532,
		PrimeiraParalela = 533,
		SegundaParalela = 534,
		TerceiraParalela = 535,
		QuartaParalela = 536,
		QuintaParalela = 537,
		PrimeiraSubida = 538,
		SegundaSubida = 539,
		TerceiraSubida = 540,
		QuartaSubida = 541,
		QuintaSubida = 542,
		SextaSubida = 543,
		PrimeiraVila = 544,
		SegundaVila = 545,
		TerceiraVila = 546,
		QuartaVila = 547,
		QuintaVila = 548,
		PrimeiroParque = 549,
		SegundoParque = 550,
		TerceiroParque = 551,
		PrimeiraRua = 552,
		SegundaRua = 553,
		TerceiraRua = 554,
		QuartaRua = 555,
		QuintaRua = 556,
		SextaRua = 557,
		SétimaRua = 558,
		OitavaRua = 559,
		NonaRua = 560,
		DécimaRua = 561,
		DécimaPrimeiraRua = 562,
		DécimaSegundaRua = 563,
		Estacionamento = 564,
		Vala = 565,
		RuaDePedestre = 566,
		Túnel = 567,
		Variante = 568,
		RodoAnel = 569,
		TravessaParticular = 570,
		Calçada = 571,
		ViaDeAcesso = 572,
		EntradaParticular = 573,
		Acampamento = 645,
		ViaExpressa = 646,
		EstradaMunicipal = 650,
		AvenidaContorno = 651,
		EntreQuadra = 652,
		RuaDeLigação = 653,
		ÁreaEspecial = 654,
		RodoviaEstadual = 655,
		Outros = 656
	}

}
