using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Migracao.Models.DentalOffice;

namespace Migracao.Models.DTO
{
    public class ProcedimentosPrecosDTO
    {

        public ProcedimentosPrecosDTO()
        {
                
        }

        public ProcedimentosPrecosDTO(List<ProcedimentosPrecos> gruposProcedimentos)
        {

        }

        [DisplayName("Especialidade")]
        public string? Especialidade { get; set; }

        [DisplayName("Nome do Procedimento")]
        public string? NomeProcedimento { get; set; }

        [DisplayName("Abreviação")]
        public string? Abreviacao { get; set; }

        [DisplayName("Preço")]
        [DisplayFormat(DataFormatString = "{0:N2}")]
        public string? Preco { get; set; }

        [DisplayName("TUSS")]
        public string? TUSS { get; set; }
    }
}
