namespace Migracao.Models
{
	internal class FornecedorFone
	{
		public int ID { get; set; }
		public required int FornecedorID { get; set; }
		public required short FoneTipoID { get; set; }
		public required long Telefone { get; set; }
		public string Extensao { get; set; }
		public string ContatoNome { get; set; }
		public int LoginID { get; set; }
		public required DateTime DataInclusao { get; set; }
		public DateTime DataUltAlteracao { get; set; }
	}
}
