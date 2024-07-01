using Migracao.Models.DentalOffice;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class ProcedimentosDTO
    {
        public ProcedimentosDTO()
        {
                
        }

        public ProcedimentosDTO(List<Procedimentos> procedimentos)
        {

        }

        [DisplayName("Número do Controle")]
        public string? NumeroControle { get; set; }

        [DisplayName("CPF do Paciente")]
        public string? PacienteCPF { get; set; }

        [DisplayName("Nome do Paciente")]
        public string? PacienteNome { get; set; }

        [DisplayName("CPF do Dentista")]
        public string? DentistaCPF { get; set; }

        [DisplayName("Nome do Dentista")]
        public string? DentistaNome { get; set; }

        [DisplayName("Dente")]
        public string? Dente { get; set; }

        [DisplayName("Nome do Procedimento")]
        public string? ProcedimentoNome { get; set; }

        [DisplayName("Valor do Procedimento")]
        public string? ProcedimentoValor { get; set; }

        [DisplayName("Observação do Procedimento")]
        public string? ProcedimentoObservacao { get; set; }

        [DisplayName("Data de Início")]
        public string? DataInicio { get; set; }

        [DisplayName("Data de Término")]
        public string? DataTermino { get; set; }

        [DisplayName("Data de Atendimento")]
        public string? DataAtendimento { get; set; }
    }
}
