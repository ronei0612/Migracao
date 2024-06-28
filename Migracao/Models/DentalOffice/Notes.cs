using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    [Table("customer_notes")]
    public class Anotacoes
    {
        [Key]
        public int? id { get; set; }
        public string name { get; set; }
        public string content { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public int? customer_id { get; set; }
        [ForeignKey("customer_id ")]
        public  Pacientes Paciente { get; set; }
      

        public int? AnotacaoID { get; set; }
        
    }
}
