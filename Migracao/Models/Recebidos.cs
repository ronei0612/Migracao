namespace Migracao.Models
{
    public class Recebidos
    {
        public string Nome_Paciente { get; set; }
        public string CNPJ_CPF { get; set; }
        public string? Numero_Controle { get; set; }
        public int? Documento { get; set; }
        public decimal? Valor_Pago { get; set; }
        public decimal? Valor_Parcela { get; set; }
        public string Observacao { get; set; }
        public DateTime? Data_Baixa { get; set; }
        public DateTime? Data_Vencimento { get; set; }
        public string Tipo_Documento { get; set; }
        public int? Parcela { get; set; }
        public string Tipo_Especie { get; set; }
        public string Especie_Pagamento { get; set; }
        public decimal? Valor_Devido { get; set; }
        public string Tipo_Pagamento { get; set; }
        public decimal? Valor_Original { get; set; }
        public string Duplicata { get; set; }
        public DateTime? Vencimento_Recebivel { get; set; }
        public string? Situacao { get; set; }
        public string? Nome_Grupo { get; set; }
        public string? Ordem { get; set; }
        public string? Pagamento_Observacoes { get; set; }
    }
}
