using Migracao.Models;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Migracao.DTO
{
    public class PacientesDentistasDTO
    {

        public PacientesDentistasDTO() { }

        public PacientesDentistasDTO(List<Pacientes> pacientes)
        {

        }

        [DisplayName("Nº Prontuário")]
        public string? Numero_Prontuario { get; set; }

        [DisplayName("Nome Completo")]
        public string? Nome_Completo { get; set; }

        [DisplayName("Nome Social")]
        public string? Nome_Social { get; set; }

        [DisplayName("Apelido")]
        public string? Apelido { get; set; }

        [DisplayName("CPF")]
        public string? CPF { get; set; }

        [DisplayName("RG")]
        public string? RG { get; set; }

        [DisplayName("E-mail")]
        public string? Email { get; set; }

        [DisplayName("Sexo")]
        public string? Sexo { get; set; }

        [DisplayName("Data de Nascimento")]
        [DisplayFormat(DataFormatString = "{0:dd/MM/yyyy HH:mm}", ApplyFormatInEditMode = true)]
        public string? Data_Nascimento { get; set; }

        [DisplayName("Cidade de Nascimento")]
        public string? Cidade_Nascimento { get; set; }

        [DisplayName("Estado Civil")]
        public string? Estado_Civil { get; set; }

        [DisplayName("Profissão")]
        public string? Profissao { get; set; }

        [DisplayName("Telefone Principal")]
        public string? Telefone_Principal { get; set; }

        [DisplayName("Celular")]
        public string? Celular { get; set; }

        [DisplayName("Telefone Alternativo")]
        public string? Telefone_Alternativo { get; set; }

        [DisplayName("Logradouro")]
        public string? Logradouro { get; set; }

        [DisplayName("Número")]
        public string? Numero { get; set; }

        [DisplayName("Complemento")]
        public string? Complemento { get; set; }

        [DisplayName("Bairro")]
        public string? Bairro { get; set; }

        [DisplayName("Cidade")]
        public string? Cidade { get; set; }

        [DisplayName("UF")]
        public string? UF { get; set; }

        [DisplayName("CEP")]
        public string? CEP { get; set; }

        [DisplayName("Observações")]
        public string? Observacoes { get; set; }
    }
}