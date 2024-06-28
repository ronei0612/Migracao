using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
   public class Categories
    {
        [Key]
        public int id { get; set; }
        public string name { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? updated_at { get; set; }
        public int? dental_insurance_id { get; set; }
    }
}
