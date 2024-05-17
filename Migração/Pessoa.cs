namespace Migração
{
	internal class Pessoa
	{
		public int ID { get; set; }
		public Nullable<byte> Digito { get; set; }
		public required string Apelido { get; set; }
		public required string NomeCompleto { get; set; }
		public required bool Sexo { get; set; }
		public Nullable<int> EnderecoID { get; set; }
		public string CPF { get; set; }
		public string RG { get; set; }
		public Nullable<System.DateTime> NascimentoData { get; set; }
		public string NascimentoLocal { get; set; }
		public Nullable<int> NascimentoCidadeID { get; set; }
		public string Nacionalidade { get; set; }
		public Nullable<System.DateTime> FalecimentoData { get; set; }
		public string FalecimentoCausa { get; set; }
		public string Email { get; set; }
		public string TipoSangue { get; set; }
		public Nullable<byte> CorPele { get; set; }
		public Nullable<byte> CorOlhos { get; set; }
		public Nullable<byte> CorCabelo { get; set; }
		public Nullable<byte> ArcadaTipoID { get; set; }
		public Nullable<byte> EstadoCivilID { get; set; }
		public Nullable<short> EscolaridadeID { get; set; }
		public Nullable<int> ProfissaoID { get; set; }
		public string ProfissaoOutra { get; set; }
		public Nullable<int> EspecialidadeID { get; set; }
		public string ConselhoCodigo { get; set; }
		public Nullable<int> FiliacaoPaiID { get; set; }
		public Nullable<int> FiliacaoMaeID { get; set; }
		public Nullable<int> ResponsavelID { get; set; }
		public Nullable<int> PromocionalID { get; set; }
		public string Origem { get; set; }
		public Nullable<int> PessoaFotoID { get; set; }
		public Nullable<int> FotoID { get; set; }
		public Nullable<int> EstabelecimentoID { get; set; }
		public Nullable<short> SolucaoID { get; set; }
		public Nullable<System.Guid> Guid { get; set; }
		public Nullable<int> LoginID { get; set; }
		public required System.DateTime DataInclusao { get; set; }
		public Nullable<System.DateTime> DataUltAlteracao { get; set; }
		public Nullable<byte> FormatoRosto { get; set; }
		public string FoneticaApelido { get; set; }
		public string FoneticaNomeCompleto { get; set; }
		public string ConselhoSigla { get; set; }
		public string ConselhoUF { get; set; }
		public Nullable<int> CadastramentoFuncionarioID { get; set; }
		public string SkypeNome { get; set; }
		public Nullable<long> PIS { get; set; }
		public string CNS { get; set; }
		public string ResumoFormacao { get; set; }
		public Nullable<System.DateTime> ExclusaoData { get; set; }
		public Nullable<int> ExclusaoFuncionarioID { get; set; }
		public string NomeSocial { get; set; }
		public string FoneticaNomeSocial { get; set; }
		public Nullable<int> UsuarioID { get; set; }
		public string AssinaturaDigital { get; set; }
	}
}
