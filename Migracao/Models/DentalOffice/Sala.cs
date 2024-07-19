using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{
    [Table("chairs")]
    public class Sala
    {
        public int id { get; set; }
        public string name { get; set; }
        public int SalaID{ get; set; }
    }
}
