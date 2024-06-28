using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    [Table("outgoings")]
    public class Exigivel
    {
        [Key]
        public int? id { get; set; }
        public string description { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public int? clinic_id { get; set; }
        public int? payment_method_id { get; set; }
        public int? chart_of_account_id { get; set; }
        public int? dentist_id { get; set; }
        public int? supplier_id { get; set; }
        public int? partial_from_id { get; set; }
        public double? amount { get; set; }
        public double? payment_amount { get; set; }
        public double? balance { get; set; }
        public DateTime? expiration_date { get; set; }
        public DateTime? payment_date { get; set; }
        public bool? paid { get; set; }
        public int? ExigivelID{ get; set; }
        public int? FluxoCaixaID{ get; set; }

    }

}
