using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{

        [Table("treatment_procedures")]
        public class Procedimentos
        {
            [Key]
            public int? id { get; set; }
            public string? faces { get; set; }
            public string? type { get; set; }
            public string? color { get; set; }
            public string? product { get; set; }
            public string? lot_number { get; set; }
            public string? dilution_volume { get; set; }
            public DateTime? deleted_at { get; set; }
            public DateTime? created_at { get; set; }
            public DateTime? updated_at { get; set; }
            public DateTime? paid_date { get; set; }
            public DateTime? commission_authorized_date { get; set; }
            public DateTime? commission_paid_date { get; set; }
            public DateTime? commission_unauthorized_date { get; set; }
            public DateTime? commission_unpaid_date { get; set; }
            public string? notes { get; set; }
            public string? description { get; set; }
            public int? treatment_id { get; set; }

            [ForeignKey("treatment_id")]
            public virtual Tratamento Tratamento { get; set; }
            public int? procedure_id { get; set; }
            public int? dentist_id { get; set; }
            public int? situation { get; set; }
            public int? qtde { get; set; }
            public int? teeth_start { get; set; }
            public int? teeth_end { get; set; }
            public int? commission_authorized_by_id { get; set; }
            public int? commission_paid_by_id { get; set; }
            public int? commission_unauthorized_by_id { get; set; }
            public int? commission_unpaid_by_id { get; set; }
            public int? commission_outgoing_id { get; set; }
            public int? face_type { get; set; }
            public double? amount { get; set; }
            public double? discount { get; set; }
            public double? prosthetic_expense { get; set; }
            public double? commission_amount { get; set; }
            public double? commission_percent { get; set; }
            public double? commission_total { get; set; }
            public double? commission_taxe_in_percentage { get; set; }
            public double? commission_charge_in_percentage { get; set; }
            public double? commission_prosthetic_amount { get; set; }
            public double? original_amount { get; set; }
            public DateTime? procedure_date { get; set; }
            public DateTime? finalized_date { get; set; }
            public DateTime? application_date { get; set; }
            public DateTime? expiration_date { get; set; }
            public bool? general { get; set; }
            public bool? estimate_included { get; set; }
            public bool? paid { get; set; }
            public bool? commission_authorized { get; set; }
            public bool? commission_paid { get; set; }
            public bool? multiplied { get; set; }
            public bool? maintenance { get; set; }
            public bool? for_approval { get; set; }
            public int? OrcamentoItemID { get; set; }
            public int? ContratoItemID { get; set; }

    }
}
