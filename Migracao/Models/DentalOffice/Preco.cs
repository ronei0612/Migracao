using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    [Table("procedures")]
    public class Preco {
        [Key]
        public int id { get; set; }
        public int? dental_insurance_id { get; set; }
        public int? category_id { get; set; }
        [ForeignKey("category_id")]
        public virtual Categories Categories { get; set; }

        public string? name { get; set; }
        public decimal? price { get; set; }
        public string? code { get; set; }
        public int? procedure_time { get; set; }
        public decimal? commissioning { get; set; }
        public string? obs { get; set; }
        public int? icon_type { get; set; }
        public int? face_material { get; set; }
        public string? icon { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public DateTime? deleted_at { get; set; }
        public int? used_count { get; set; }
        public short? procedure_type { get; set; }
        public short? orofacial { get; set; }        
        public int? PrecoID { get; set; }
    }
}
