using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    [Table("Payments")]
    public class Recebiveis
    {
        [Key]
        public int id { get; set; }
        public string description { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public int? estimate_id { get; set; }
        //[ForeignKey("estimate_id")]
        //public virtual Orcamento Orcamento{ get; set; }
        public int? customer_id { get; set; }
        //[ForeignKey("customer_id")]
        //public virtual Pacientes? Paciente { get; set; }
        public int? payment_type { get; set; }
        public int? treatment_procedure_id { get; set; }
        public int? payment_method_id { get; set; }
        public int? dental_insurance_id { get; set; }
        public int? clinic_id { get; set; }
        public int? bank_slip_id { get; set; }
        public decimal? adjustment { get; set; }
        public decimal? amount { get; set; }
        public decimal? balance { get; set; }
        public decimal? payment_amount { get; set; }
        public DateTime? expiration_date { get; set; }
        public DateTime? payment_date { get; set; }
        public bool? paid { get; set; }
        public bool? available { get; set; }
        public bool? down_payment { get; set; }
        public int? RecebivelID { get; set; }
        public int? FluxoCaixaID { get; set; }


    


    }
}
