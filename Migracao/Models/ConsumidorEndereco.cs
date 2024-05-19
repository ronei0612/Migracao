namespace Migracao.Models
{
	internal class ConsumidorEndereco
	{
		public int ID { get; set; }
		public required int ConsumidorID { get; set; }
		public required short EnderecoTipoID { get; set; }
		public required int LogradouroTipoID { get; set; }
		public required string Logradouro { get; set; }
		public string LogradouroNum { get; set; }
		public string Complemento { get; set; }
		public string Bairro { get; set; }
		public required int CidadeID { get; set; }
		public required int Cep { get; set; }
		public string CaixaPostal { get; set; }
		public Nullable<decimal> Latitude { get; set; }
		public Nullable<decimal> Longitude { get; set; }
		public required bool Ativo { get; set; }
		public Nullable<int> LoginID { get; set; }
		public required System.DateTime DataInclusao { get; set; }
		public Nullable<System.DateTime> DataUltAlteracao { get; set; }
	}
}
