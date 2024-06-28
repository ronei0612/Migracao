using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    [Table("procedures",Schema ="rn")]
    public class PrecoTemp {
        [Key]
        public int id { get; set; }
        public int? dental_insurance_id { get; set; }
        public int? category_id { get; set; }
        public string name { get; set; }
   
        public int? PRECOID { get; set; }
    }
}
