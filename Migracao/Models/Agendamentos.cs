namespace Migracao.Models
{
    public class Agendamentos
    {
        public string? Lancamento { get; set; }
        public string? Paciente_CPF { get; set; }
        public string? Nome { get; set; }
        public DateTime Data { get; set; }
        public string? Hora { get; set; }
        public string? Codigo_Responsavel { get; set; }
        public string? Nome_Dentista { get; set; }
        public string? Telefone { get; set; }
        public DateTime? Data_Inclusao { get; set; }
        public string? Observacao { get; set; }
    }
}
