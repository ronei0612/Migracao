namespace Migracao.Models
{
    internal class Pessoa
    {
        public int ID { get; set; }
        public byte? Digito { get; set; }
        public required string Apelido { get; set; }
        public required string NomeCompleto { get; set; }
        public required bool Sexo { get; set; }
        public int? EnderecoID { get; set; }
        public string CPF { get; set; }
        public string RG { get; set; }
        public DateTime? NascimentoData { get; set; }
        public string NascimentoLocal { get; set; }
        public int? NascimentoCidadeID { get; set; }
        public string Nacionalidade { get; set; }
        public DateTime? FalecimentoData { get; set; }
        public string FalecimentoCausa { get; set; }
        public string Email { get; set; }
        public string TipoSangue { get; set; }
        public byte? CorPele { get; set; }
        public byte? CorOlhos { get; set; }
        public byte? CorCabelo { get; set; }
        public byte? ArcadaTipoID { get; set; }
        public byte? EstadoCivilID { get; set; }
        public short? EscolaridadeID { get; set; }
        public int? ProfissaoID { get; set; }
        public string ProfissaoOutra { get; set; }
        public int? EspecialidadeID { get; set; }
        public string ConselhoCodigo { get; set; }
        public int? FiliacaoPaiID { get; set; }
        public int? FiliacaoMaeID { get; set; }
        public int? ResponsavelID { get; set; }
        public int? PromocionalID { get; set; }
        public string Origem { get; set; }
        public int? PessoaFotoID { get; set; }
        public int? FotoID { get; set; }
        public int? EstabelecimentoID { get; set; }
        public short? SolucaoID { get; set; }
        public Guid? Guid { get; set; }
        public int? LoginID { get; set; }
        public required DateTime DataInclusao { get; set; }
        public DateTime? DataUltAlteracao { get; set; }
        public byte? FormatoRosto { get; set; }
        public string FoneticaApelido { get; set; }
        public string FoneticaNomeCompleto { get; set; }
        public string ConselhoSigla { get; set; }
        public string ConselhoUF { get; set; }
        public int? CadastramentoFuncionarioID { get; set; }
        public string SkypeNome { get; set; }
        public long? PIS { get; set; }
        public string CNS { get; set; }
        public string ResumoFormacao { get; set; }
        public DateTime? ExclusaoData { get; set; }
        public int? ExclusaoFuncionarioID { get; set; }
        public string NomeSocial { get; set; }
        public string FoneticaNomeSocial { get; set; }
        public int? UsuarioID { get; set; }
        public string AssinaturaDigital { get; set; }
    }
}
