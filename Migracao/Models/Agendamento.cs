namespace Migracao.Models
{
	internal class Agendamento
	{
		public int ID { get; set; }
		public short AgendaTipoID { get; set; }
		public short AtendeTipoID { get; set; }
		public string? Titulo { get; set; }
		public string Descricao { get; set; }
		public DateTime DataInicio { get; set; }
		public DateTime DataTermino { get; set; }
		public DateTime ConfirmadoData { get; set; }
		public int ConfirmadoPessoaID { get; set; }
		public DateTime DataCancelamento { get; set; }
		public string CancelamentoMotivo { get; set; }
		public bool CancelamentoClinica { get; set; }
		public int CancelamentoPessoaID { get; set; }
		public bool DiaTodo { get; set; }
		public DateTime DataRealizado { get; set; }
		public int PessoaID { get; set; }
		public int ClienteID { get; set; }
		public int? ConsumidorID { get; set; }
		public int ContatoEmpresaID { get; set; }
		public int ConsumidorPessoaID { get; set; }
		public string ConsumidorPessoaNome { get; set; }
		public long ConsumidorPessoaFone1 { get; set; }
		public bool ConsumidorPessoaFone1SMS { get; set; }
		public long ConsumidorPessoaFone2 { get; set; }
		public bool ConsumidorPessoaFone2SMS { get; set; }
		public int ConvenioID { get; set; }
		public decimal AtendimentoValor { get; set; }
		public int? FuncionarioID { get; set; }
		public int EstabelecimentoID { get; set; }
		public int EmpresaID { get; set; }
		public int SecretariaID { get; set; }
		public int ReagendamentoID { get; set; }
		public int AtendimentoID { get; set; }
		public short SolucaoID { get; set; }
		public int LoginID { get; set; }
		public DateTime DataInclusao { get; set; }
		public DateTime DataUltAlteracao { get; set; }
		public string ConsumidorPessoaEmail { get; set; }
		public int RetornoID { get; set; }
		public string ConvenioCartao { get; set; }
		public DateTime ExclusaoData { get; set; }
		public int ExclusaoPessoaID { get; set; }
		public int SalaID { get; set; }
		public int CampanhaEventoID { get; set; }
		public DateTime ConsumidorPessoaDataNascimento { get; set; }
		public bool ClienteNovo { get; set; }
		public string NotificadoSMSDia { get; set; }
		public string NotificadoSMSDiaAnterior { get; set; }
		public string NotificadoWhatsAppDia { get; set; }
		public string NotificadoWhatsAppDiaAnterior { get; set; }
		public int AtendimentoTipoCustomID { get; set; }
		public int EsperaID { get; set; }
	}
}
