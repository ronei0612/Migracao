namespace Migração.Models
{
    internal class Funcionario
    {
        public int ID { get; set; }
        public required int PessoaID { get; set; }
        public string Matricula { get; set; }
        public System.DateTime DataAdmissao { get; set; }
        public System.DateTime DataDemissao { get; set; }
        public decimal Comissao { get; set; }
        public short ComissaoTipoID { get; set; }
        public decimal SalarioMes { get; set; }
        public decimal SalarioHora { get; set; }
        public decimal SalarioDia { get; set; }
        public bool SalarioCLT { get; set; }
        public byte CargaHorariaSemanal { get; set; }
        public required int CargoID { get; set; }
        public string Email { get; set; }
        public byte AgendaIntervalo { get; set; }
        public required int PermissaoGrupoID { get; set; }
        public required bool PermissaoCoordenacao { get; set; }
        public required bool PermissaoModuloAdmin { get; set; }
        public required bool PermissaoModuloAtendimentos { get; set; }
        public required bool PermissaoModuloPacientes { get; set; }
        public required bool PermissaoModuloFinanceiro { get; set; }
        public int EstabelecimentoID { get; set; }
        public int ClienteID { get; set; }
        public int EmpresaID { get; set; }
        public short SolucaoID { get; set; }
        public required bool Ativo { get; set; }
        public required int LoginID { get; set; }
        public required System.DateTime DataInclusao { get; set; }
        public System.DateTime DataUltAlteracao { get; set; }   
        public System.DateTime ExclusaoData { get; set; }
        public string ExclusaoMotivo { get; set; }
        public required bool PermissaoModuloEstoque { get; set; }
        public decimal ParticipacaoProteses { get; set; }
        public bool ComissionarNaoConcluidos { get; set; }
        public int CodigoCBO { get; set; }
        public int CargoID2 { get; set; }
        public string Observacoes { get; set; }
        public bool PermissaoEngenharia { get; set; }
        public byte SituacaoID { get; set; }
        public short AtivacaoMonthCount { get; set; }
        public System.DateTime AtivacaoUltima { get; set; }
        public int ContaBancariaID { get; set; }
        public decimal ComissaoVendas { get; set; }
    }
}
