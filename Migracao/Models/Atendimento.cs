using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class Atendimento
    {
        public int ID { get; set; }
        [Required(ErrorMessage = "O campo é obrigatório")]
        public short AtendeTipoID { get; set; }
        [Required(ErrorMessage = "O campo é obrigatório")]
        public System.DateTime DataChegada { get; set; }
        public Nullable<System.DateTime> DataInicio { get; set; }
        public Nullable<System.DateTime> DataTermino { get; set; }
        public Nullable<System.DateTime> DataCancelamento { get; set; }
        [Required(ErrorMessage = "O campo é obrigatório")]
        public int ConsumidorID { get; set; }
        public Nullable<int> ResponsavelID { get; set; }
        public Nullable<decimal> AtendimentoValor { get; set; }
        public Nullable<int> FuncionarioID { get; set; }
        public Nullable<int> FuncionarioIDConclusao { get; set; }
        [Required(ErrorMessage = "O campo é obrigatório")]
        public int EstabelecimentoID { get; set; }
        [Required(ErrorMessage = "O campo é obrigatório")]
        public int SecretariaID { get; set; }
        public Nullable<int> AgendaID { get; set; }
        public Nullable<int> ConvenioID { get; set; }
        public Nullable<short> SolucaoID { get; set; }
        [Required(ErrorMessage = "O campo é obrigatório")]
        public int LoginID { get; set; }

        [StringLength(2147483647)]
        public string Observacoes { get; set; }
        [Required(ErrorMessage = "O campo é obrigatório")]
        public System.DateTime DataInclusao { get; set; }
        public Nullable<System.DateTime> DataUltAlteracao { get; set; }
        public Nullable<int> SalaID { get; set; }
        public Nullable<System.DateTime> DataChamadaMedico { get; set; }
        public Nullable<System.DateTime> DataChamadaAtendente { get; set; }
        public Nullable<short> TempoAtraso { get; set; }
        public Nullable<short> TempoSalaEspera { get; set; }
        public Nullable<short> TempoAtendimento { get; set; }
        public Nullable<decimal> AtendimentoIndex { get; set; }
        public Nullable<int> EncaminhadoPorMedicoPessoaID { get; set; }
        public Nullable<int> DiagnosticoID { get; set; }
        public Nullable<System.DateTime> ExclusaoData { get; set; }
        public Nullable<int> ExclusaoPessoaID { get; set; }
        public Nullable<int> AtendimentoTipoCustomID { get; set; }
    }
}
