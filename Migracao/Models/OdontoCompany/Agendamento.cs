using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.OdontoCompany
{
    public class Agendamento
    {
        [Table("AGENDA")]
        public class Agenda
        {
            public int? LANCTO { get; set; }
            public string? CNPJ_CPF { get; set; }
            public string? NOME { get; set; }
            public DateTime? DATA { get; set; }
            public string? HORA { get; set; }
            public string? CODIGO_RESP { get; set; }
            public string? RESPONSAVEL { get; set; }
            public string? FALTOU { get; set; }
            public string? OBS { get; set; }
            public string? TERMINAL { get; set; }
            public string? USUARIO { get; set; }
            public int? LINHA { get; set; }
            public int? COLUNA { get; set; }
            public string? FICHA { get; set; }
            public string? AVISO { get; set; }
            public string? DEPARTAMENTO { get; set; }
            public DateTime? DATA_EXC { get; set; }
            public string? USUARIO_EXC { get; set; }
            public string? ENCAIXE { get; set; }
            public string? FONE_1 { get; set; }
            public string? FONE_2 { get; set; }
            public int? ID_AGENDA { get; set; }
            public DateTime? MODIFICADO { get; set; }
            public int? CODIGO_ESPECIALIDADE { get; set; }
            public DateTime? CONFIRMACAO_DATA_HORA { get; set; }
            public string? CONFIRMACAO_USUARIO { get; set; }

            public string? CONFIRMACAO_OBSERVACAO { get; set; }

            public DateTime? DT_AXON { get; set; }

            public string? AXON_ID { get; set; }


        }
    }
}
