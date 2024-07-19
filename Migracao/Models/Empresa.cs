namespace Migracao.Models
{
    internal class Empresa
    {
		public int ID { get; set; }
		public required string RazaoSocial { get; set; }
		public required string NomeFantasia { get; set; }
		public required string Marca { get; set; }
		public required string CNPJ { get; set; }
		public string IE { get; set; }
		public int EnderecoID { get; set; }
		public string WebSite { get; set; }
		public string Email { get; set; }
		public DateTime DataFundacao { get; set; }
		public required byte RegimeTribID { get; set; }
		public int FotoID { get; set; }
		public int EstabelecimentoID { get; set; }
		public int SolucaoID { get; set; }
		public Guid Guid { get; set; }
		public required bool Ativo { get; set; }
		public required int LoginID { get; set; }
		public required DateTime DataInclusao { get; set; }
		public DateTime DataUltAlteracao { get; set; }
		public Guid EnotasEmpresaId { get; set; }
		public string InscricaoMunicipal { get; set; }
	}
}
