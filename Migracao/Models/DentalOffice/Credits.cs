using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    public class Credits
    {
        [Key]
        public int id { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public int? customer_id { get; set; }
        public int? payment_id { get; set; }

        //[ForeignKey("payment_id")]
        //public virtual Payments Payments { get; set; }
        public int? payment_method_id { get; set; }
        public int? paycheck_id { get; set; }
        public int? card_id { get; set; }
        public int? cash_flow_id { get; set; }
        public int? receipt_id { get; set; }
        public double? amount { get; set; }
        public double? interest_amount { get; set; }
        public DateTime? payment_date { get; set; }
    }
}
