using DocumentFormat.OpenXml.Wordprocessing;
using Migracao.Models;
using Migracao.Utils;
using System;
using System.Collections.Generic;
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

        public string? Codigo { get; set; }
        public string? Ativo { get; set; }
        public string? NomeCompleto { get; set; }
        public string? NomeSocial { get; set; }
        public string? Apelido { get; set; }
        public string? Documento { get; set; }
        public DateTime? DataCadastro { get; set; }
        public string? Observacoes { get; set; }
        public string? Email { get; set; }
        public string? RG { get; set; }
        public string? Sexo { get; set; }
        public DateTime NascimentoData { get; set; }
        public string? NascimentoLocal { get; set; }
        public string? EstadoCivil { get; set; }
        public string? Profissao { get; set; }
        public string? CargoNaClinica { get; set; }
        public string? Dentista { get; set; }
        public string? ConselhoCodigo { get; set; }
        public string? Paciente { get; set; }
        public string? Funcionario { get; set; }
        public string? Fornecedor { get; set; }
        public string TelefonePrincipal { get; set; }
        public string? Celular { get; set; }
        public string? TelefoneAlternativo { get; set; }
        public string? Logradouro { get; set; }
        public string? LogradouroNum { get; set; }
        public string? Complemento { get; set; }
        public string? Bairro { get; set; }
        public string? Cidade { get; set; }
        public string? Estado { get; set; }
        public string? CEP { get; set; }
    }
}
