namespace Migracao.Models
{
	public class Preco
	{
		public int ID { get; set; }
		public int TabelaID { get; set; }
		public string Codigo { get; set; }
		public required string Titulo { get; set; }
		public string Atalho { get; set; }
		public int ProcedimentoID { get; set; }
		public required short CategoriaID { get; set; }
		public required decimal Valor { get; set; }
		public string Observacoes { get; set; }
		public required int LoginID { get; set; }
		public required DateTime DataInclusao { get; set; }
		public DateTime DataUltAlteracao { get; set; }
		public short ComissaoTipoID { get; set; }
		public decimal ComissaoPercentual { get; set; }
		public decimal ComissaoValor { get; set; }
		public bool Conjunto { get; set; }
		public int ConjuntoID { get; set; }
		public decimal Minimo { get; set; }
		public decimal Maximo { get; set; }
		public long CodigoTISS { get; set; }
		public decimal CustoValor { get; set; }
		public bool RestringirMargens { get; set; }
		public required bool Ativo { get; set; }
		public decimal ValorReajustado { get; set; }
		public DateTime ExclusaoData { get; set; }
		public DateTime ExclusaoPessoaID { get; set; }
	}
}
