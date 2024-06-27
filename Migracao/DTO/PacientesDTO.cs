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
        [Description("Codigo")]
        public string? Codigo { get; set; }
        [Description("Ativo")]
        public string? Ativo { get; set; }
        [Description("Nome Completo")]
        public string? NomeCompleto { get; set; }
        [Description("Nome Social")]
        public string? NomeSocial { get; set; }
        [Description("Apelido")]
        public string? Apelido { get; set; }
        [Description("Documento")]
        public string? Documento { get; set; }
        [Description("Data Cadastro")]
        public string? DataCadastro { get; set; }
        [Description("Observacões")]
        public string? Observacoes { get; set; }
        [Description("Email")]
        public string? Email { get; set; }
        [Description("RG")]
        public string? RG { get; set; }
        [Description("Sexo")]
        public string? Sexo { get; set; }
        [Description("Nascimento Data")]
        public string NascimentoData { get; set; }
        [Description("Nascimento Local")]
        public string? NascimentoLocal { get; set; }
        [Description("Estado Civil")]
        public string? EstadoCivil { get; set; }
        [Description("Profissão")]
        public string? Profissao { get; set; }
        [Description("Cargo na Clinica")]
        public string? CargoNaClinica { get; set; }
        [Description("Dentista")]
        public string? Dentista { get; set; }
        [Description("Conselho Codigo")]
        public string? ConselhoCodigo { get; set; }
        [Description("Paciente")]
        public string? Paciente { get; set; }
        [Description("Funcionário")]
        public string? Funcionario { get; set; }
        [Description("Fornecedor")]
        public string? Fornecedor { get; set; }
        [Description("Telefone Principal")]
        public string TelefonePrincipal { get; set; }
        [Description("Celular")]
        public string? Celular { get; set; }
        [Description("Telefone Alternativo")]
        public string? TelefoneAlternativo { get; set; }
        [Description("Logradouro")]
        public string? Logradouro { get; set; }
        [Description("LogradouroNum")]
        public string? LogradouroNum { get; set; }
        [Description("Complemento")]
        public string? Complemento { get; set; }
        [Description("Bairro")]
        public string? Bairro { get; set; }
        [Description("Cidade")]
        public string? Cidade { get; set; }
        [Description("Estado")]
        public string? Estado { get; set; }
        [Description("CEP")]
        public string? CEP { get; set; }
    }
}
