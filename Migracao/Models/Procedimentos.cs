namespace Migracao.Models
{
    public class Procedimentos
    {
        public string? Numero_Controle { get; set; }
        public string? Paciente_CPF { get; set; }
        public string? Nome_Paciente { get; set; }
        public string? Dentista_CPF { get; set; }
        public string? Dentista_Nome { get; set; }
        public string? Dente { get; set; }
        public string? Nome_Procedimento { get; set; }
        public decimal? Valor { get; set; }
        public string? Observacao { get; set; }
        public DateTime? Data_Inicio { get; set; }
        public DateTime? Data_Atendimento { get; set; }
    }
}
