namespace Migracao.Models
{
    internal class Endereco
    {
		public int ID { get; set; }
		public required int TableID { get; set; }
		public int ParentID { get; set; }
		public required short EnderecoTipoID { get; set; }
		public required int LogradouroTipoID { get; set; }
		public required string Logradouro { get; set; }
		public string LogradouroNum { get; set; }
		public string Complemento { get; set; }
		public string Bairro { get; set; }
		public required int CidadeID { get; set; }
		public required int Cep { get; set; }
		public string CaixaPostal { get; set; }
		public string Fone { get; set; }
		public required bool Ativo { get; set; }
		public int LoginID { get; set; }
		public required DateTime DataInclusao { get; set; }
		public DateTime DataUltAlteracao { get; set; }
	}
}
