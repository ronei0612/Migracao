namespace Migração.Models
{
	internal class PessoaFone
	{
		public int ID { get; set; }
		public required int PessoaID { get; set; }
		public required short FoneTipoID { get; set; }
		public required long Telefone { get; set; }
		public string Extensao { get; set; }
		public Nullable<int> LoginID { get; set; }
		public required System.DateTime DataInclusao { get; set; }
		public Nullable<System.DateTime> DataUltAlteracao { get; set; }
		public Nullable<bool> WhatsAppOptIn { get; set; }
	}
}
