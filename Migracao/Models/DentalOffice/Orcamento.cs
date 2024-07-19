using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    //estimates

    [Table("estimates")]
     public class Orcamento
    {
        [Key]
        public int id { get; set; } //fixo
        public string? type { get; set; }
        public DateTime? deleted_at { get; set; }

        [Column(name: "created_at")]
        public DateTime? DataCriacao { get; set; }/// Fixo
        public DateTime? updated_at { get; set; }
        public DateTime? approved_at { get; set; }

        [Column(name: "additional_text")]
        public string? Observacoes { get; set; } //Fixo
        public string? email_additional_text { get; set; }
        public string? notes { get; set; }
        public string? internal_notes { get; set; }
        
        public int? treatment_id { get; set; }
        //[ForeignKey("treatment_id ")]
        //public virtual Tratamento Tratamento { get; set; }

        public int? discount_type { get; set; }
        public int? payment_type { get; set; }
        public int? portion_type { get; set; }
        public int? customer_id { get; set; }
        //[ForeignKey("customer_id")]
        //public virtual Pacientes Paciente { get; set; }
        public int? payment_method_id { get; set; }
        public int? signature_status { get; set; }

        [Column(name: "dentist_id")]
        public int? dentist_id { get; set; } //Fixo

        //[ForeignKey("dentist_id ")]
        //public virtual Funcionario Funcionario { get; set; }

        public double? discount { get; set; }
        public double? adjustment { get; set; }

        [Column(name:"total")]
        public double? ValorTotalProcedimentos { get; set; }//Fixo
        public double? monthly_adjustment { get; set; }
        public double? original_total { get; set; }
        public bool? approved { get; set; }
        public int? OrcamentoID { get; set; }
        public int? ContratoID { get; set; }

    }
}
