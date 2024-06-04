namespace Migracao.Models
{
	internal class PrecosTabela
	{
		public int ID { get; set; }
		public string Nome { get; set; }
		public string Descricao { get; set; }
		public int ConvenioID { get; set; }
		public int OperadoraID { get; set; }
		public string Codigo { get; set; }
		public int PlanoID { get; set; }
		public int EstabelecimentoID { get; set; }
		public required short SeguimentoID { get; set; }
		public required bool Ativo { get; set; }
		public required short SolucaoID { get; set; }
		public required int LoginID { get; set; }
		public required DateTime DataInclusao { get; set; }
		public DateTime DataUltAlteracao { get; set; }
		public DateTime ExclusaoData { get; set; }
		public int ExclusaoPessoaID { get; set; }
	}
}
