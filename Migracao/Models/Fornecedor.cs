namespace Migracao.Models
{
    internal class Fornecedor
    {
        public int ID { get; set; }
        public int EmpresaID { get; set; }
        public int PessoaID { get; set; }
        public string Observacoes { get; set; }
        public required bool Ativo { get; set; }
        public required int EstabelecimentoID { get; set; }
        public required int LoginID { get; set; }
        public required System.DateTime DataInclusao { get; set; }
        public System.DateTime DataUltAlteracao { get; set; }
        public System.DateTime ExclusaoData { get; set; }
        public string ExclusaoMotivo { get; set; }
        public string NomeFantasia { get; set; }
        public string Email { get; set; }
        public int ConvenioID { get; set; }
        public int ExclusaoPessoaID { get; set; }
        public int ContaBancariaID { get; set; }
    }
}
