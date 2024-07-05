using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Migracao.Models
{
    public class Pacientes
    {
        public string? Nome_Paciente { get; set; }
        public string? Numero_Prontuario { get; set; }
        public string? Observacoes { get; set; }
        public string? E_mail { get; set; }
        public string? Telefone_Principal { get; set; }
        public string? CPF { get; set; }
        public string? RG { get; set; }
        public string? Sexo { get; set; }
        public string? Data_de_Nascimento { get; set; }
        public string? Celular { get; set; }
        public string? Telefone_Alternativo { get; set; }
        public string? Logradouro { get; set; }
        public string? Numero { get; set; }
        public string? Complemento { get; set; }
        public string? Bairro { get; set; }
        public string? Cidade { get; set; }
        public string? UF { get; set; }
        public string? CEP { get; set; }
    }
}