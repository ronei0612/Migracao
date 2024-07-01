using DocumentFormat.OpenXml.Wordprocessing;
using Migracao.Models;
using Migracao.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.DTO
{
    public class PacientesDTO
    {

        public PacientesDTO() { }

        public PacientesDTO(List<Pacientes> pacientes)
        {

        }
        [DisplayName("Código")]
        public string? Codigo { get; set; }
        [DisplayName("Ativo")]
        public string? Ativo { get; set; }
        [DisplayName("Nome Completo")]
        public string? NomeCompleto { get; set; }
        [DisplayName("Nome Social")]
        public string? NomeSocial { get; set; }
        [DisplayName("Apelido")]
        public string? Apelido { get; set; }
        [DisplayName("Documento")]
        public string? Documento { get; set; }
        [DisplayName("Data Cadastro")]
        public string? DataCadastro { get; set; }
        [DisplayName("Observacões")]
        public string? Observacoes { get; set; }
        [DisplayName("Email")]
        public string? Email { get; set; }
        [DisplayName("RG")]
        public string? RG { get; set; }
        [DisplayName("Sexo")]
        public string? Sexo { get; set; }
        [DisplayName("Nascimento Data")]
        public string NascimentoData { get; set; }
        [DisplayName("Nascimento Local")]
        public string? NascimentoLocal { get; set; }
        [DisplayName("Estado Civil")]
        public string? EstadoCivil { get; set; }
        [DisplayName("Profissão")]
        public string? Profissao { get; set; }
        [DisplayName("Cargo na Clinica")]
        public string? CargoNaClinica { get; set; }
        [DisplayName("Dentista")]
        public string? Dentista { get; set; }
        [DisplayName("Conselho Codigo")]
        public string? ConselhoCodigo { get; set; }
        [DisplayName("Paciente")]
        public string? Paciente { get; set; }
        [DisplayName("Funcionário")]
        public string? Funcionario { get; set; }
        [DisplayName("Fornecedor")]
        public string? Fornecedor { get; set; }
        [DisplayName("Telefone Principal")]
        public string TelefonePrincipal { get; set; }
        [DisplayName("Celular")]
        public string? Celular { get; set; }
        [DisplayName("Telefone Alternativo")]
        public string? TelefoneAlternativo { get; set; }
        [DisplayName("Logradouro")]
        public string? Logradouro { get; set; }
        [DisplayName("LogradouroNum")]
        public string? LogradouroNum { get; set; }
        [DisplayName("Complemento")]
        public string? Complemento { get; set; }
        [DisplayName("Bairro")]
        public string? Bairro { get; set; }
        [DisplayName("Cidade")]
        public string? Cidade { get; set; }
        [DisplayName("Estado")]
        public string? Estado { get; set; }
        [DisplayName("CEP")]
        public string? CEP { get; set; }
    }
}