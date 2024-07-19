using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    [Table("dentist_credits")]
    public class ComissoesPaga
    {
        [Key]
        public int id { get; set; }
        public string split_status { get; set; }
        public DateTime? commission_authorized_date { get; set; }
        public DateTime? commission_paid_date { get; set; }
        public DateTime? commission_unauthorized_date { get; set; }
        public DateTime? commission_unpaid_date { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? paid_date { get; set; }
        public DateTime? updated_at { get; set; }
        public int? clinic_id { get; set; }
        public int? commission_authorized_by_id { get; set; }
        public int? commission_outgoing_id { get; set; }
        public int? commission_paid_by_id { get; set; }
        public int? commission_unauthorized_by_id { get; set; }
        public int? commission_unpaid_by_id { get; set; }
        public int? credit_id { get; set; }
        [ForeignKey("credit_id")]
        public virtual Credits Credits { get; set; }
        public int? customer_id { get; set; }
        public int? dentist_id { get; set; }
        public int? treatment_procedure_id { get; set; }
        public double? amount { get; set; }
        public double? charge_amount { get; set; }
        public double? commission_percent { get; set; }
        public double? paid_amount { get; set; }
        public double? prosthetic_expense { get; set; }
        public double? taxe_in_percentage { get; set; }
        public bool? commission_authorized { get; set; }
        public bool? commission_paid { get; set; }
        public bool? split { get; set; }
    }
}
