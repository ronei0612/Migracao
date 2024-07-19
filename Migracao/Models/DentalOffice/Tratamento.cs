using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{


    [Table("treatments")]
    public  class Tratamento
    {

        public Tratamento()
        {
            
        }

        [Key]
        public int? id { get; set; }
        public string? title { get; set; }
        public string? evolution_file_name { get; set; }
        public string? evolution_content_type { get; set; }
        public string? patient_planning_approval_signature_file_name { get; set; }
        public string? patient_planning_approval_signature_content_type { get; set; }
        public string? patient_close_approval_signature_file_name { get; set; }
        public string? patient_close_approval_signature_content_type { get; set; }
        public string? type { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? opening_hour { get; set; }
        public DateTime? closing_hour { get; set; }
        public DateTime? evolution_updated_at { get; set; }
        public DateTime? planning_approved_at { get; set; }
        public DateTime? close_approved_at { get; set; }
        public DateTime? patient_planning_approval_signature_updated_at { get; set; }
        public DateTime? patient_planning_approved_at { get; set; }
        public DateTime? patient_close_approval_signature_updated_at { get; set; }
        public DateTime? patient_close_approved_at { get; set; }
        public string? additional_text { get; set; }
        public string? obs { get; set; }
        public int? evolution_file_size { get; set; }
        public int? patient_planning_approval_signature_file_size { get; set; }
        public int? patient_close_approval_signature_file_size { get; set; }
        public int? treatment_situation_id { get; set; }
        public int? customer_id { get; set; }
        public int? dental_insurance_id { get; set; }
        //public int? treatment_type { get; set; }
        public int? title_index { get; set; }
        public int? title_count { get; set; }
        public int? clinic_id { get; set; }
        public int? dentist_id { get; set; }
        public int? text_id { get; set; }
        public int? planning_approved_by_id { get; set; }
        public int? close_approved_by_id { get; set; }
        public int? face_type { get; set; }
        public DateTime? opening_date { get; set; }
        public DateTime? reopening_date { get; set; }
        public DateTime? closing_date { get; set; }
        public bool? migrated { get; set; }
        public bool? planning_approved { get; set; }
        public bool? close_approved { get; set; }
        public bool? patient_planning_approved { get; set; }
        public bool? patient_close_approved { get; set; }

        public ICollection<Procedimentos> Procedimentos { get; set; } = new List<Procedimentos>()!; // Required reference navigation to principal
    }
}

