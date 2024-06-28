using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dal.DetalOffice
{
    public class Payments
    {
        [Key]
        public int id { get; set; }
        public bool? paid { get; set; }
        public string description { get; set; }        
        public int? bank_slip_id { get; set; }
        public int? clinic_id { get; set; }
        public int? customer_id { get; set; }
        public int? dental_insurance_id { get; set; }
        public int? estimate_id { get; set; }
        public int? payment_method_id { get; set; }
        public int? payment_type { get; set; }
        public int? treatment_procedure_id { get; set; }
        public double? adjustment { get; set; }
        public double? amount { get; set; }
        public double? balance { get; set; }
        public double? payment_amount { get; set; }
        public DateTime? expiration_date { get; set; }
        public DateTime? payment_date { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public bool? available { get; set; }
        public bool? down_payment { get; set; }
        public int? RecebivelID { get; set; }
        public int? FluxoCaixaID { get; set; }
        
    }
}
