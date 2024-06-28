using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{

    [Table("cash_flow_items")]
    public class FluxoCaixa 
    {
        public int id { get; set; }
        public string transaction_type { get; set; }
        public string description { get; set; }
        public string resource_type { get; set; }
        public DateTime? transaction_date { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public int? clinic_id { get; set; }
        public int? io { get; set; }
        public int? user_id { get; set; }
        public int? payment_method_id { get; set; }
        public int? resource_id { get; set; }
        public int? cash_flow_id { get; set; }
        public decimal? amount { get; set; }
    }
}
