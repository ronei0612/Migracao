using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models.DTO
{
    public class DentistasDTO
    {
        public DentistasDTO() { }

        public DentistasDTO(List<Dentistas> dentistas) { }

        [DisplayName("Código")]
        public string? Codigo { get; set; }

        [DisplayName("Ativo(S/N)")]
        public string? Ativo { get; set; }

        [DisplayName("Nome Completo")]
        public string? Nome_Completo { get; set; }

        [DisplayName("Nome Social")]
        public string? NomeSocial { get; set; }

        [DisplayName("Apelido")]
        public string? Apelido { get; set; }

        [DisplayName("Data de Cadastro")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Cadastro { get; set; }

        [DisplayName("Observações")]
        public string? Observacoes { get; set; }

        [DisplayName("Email")]
        public string? Email { get; set; }

        [DisplayName("Local de Nascimento")]
        public string? Nascimento_Local { get; set; }

        [DisplayName("Estado Civil")]
        public string? Estado_Civil { get; set; }

        [DisplayName("Profissão")]
        public string? Profissao { get; set; }

        [DisplayName("Cargo na Clínica")]
        public string? Cargo_Clinica { get; set; }

        [DisplayName("Dentista(S/N)")]
        public string? Dentista { get; set; }

        [DisplayName("Conselho Código")]
        public string? Conselho_Codigo { get; set; }

        [DisplayName("Paciente(S/N)")]
        public string? Paciente { get; set; }

        [DisplayName("Funcionário(S/N)")]
        public string? Funcionario { get; set; }

        [DisplayName("Fornecedor(S/N)")]
        public string? Fornecedor { get; set; }
    }
}
