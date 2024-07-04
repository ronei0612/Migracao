using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class GruposProcedimentosDTO
    {

        public GruposProcedimentosDTO()
        {
                
        }

        public GruposProcedimentosDTO(List<GruposProcedimentos> gruposProcedimentos)
        {

        }

        [DisplayName("Nome Tabela")]
        public string? NomeTabela { get; set; }

        [DisplayName("Especialidade")]
        public string? Especialidade { get; set; }

        [DisplayName("Ativo (Sim/Não)")]
        public string? Ativo { get; set; }

        [DisplayName("Nome do Procedimento")]
        public string? NomeProcedimento { get; set; }

        [DisplayName("Abreviação")]
        public string? Abreviacao { get; set; }

        [DisplayName("Preço")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public string? Preco { get; set; }

        [DisplayName("TUSS")]
        public string? TUSS { get; set; }

        [DisplayName("Especialidade Código")]
        public string? EspecialidadeCodigo { get; set; }
    }
}
