using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{

    [Table("dentists")]
    public  class Funcionario
    {
       
            [Key]
            public int? id { get; set; }
            public string? name { get; set; }
            public string? rg { get; set; }
            public string? cpf { get; set; }
            public string? cr_type { get; set; }
            public int? cr_number { get; set; }
            public string? cr_uf { get; set; }
            public string? register_number { get; set; }
            public string? memed_id { get; set; }
            public string? memed_token { get; set; }
            public string? cnes { get; set; }
            public string? cbos { get; set; }
            public string? memed_external_id { get; set; }
            public DateTime? deleted_at { get; set; }
            public DateTime? created_at { get; set; }
            public DateTime? updated_at { get; set; }
            public int? specialty_id { get; set; }
            public int? duration { get; set; }
            public int? commission_type { get; set; }
            public int? dentist_type { get; set; }
            public int? semester { get; set; }
            public int? gender { get; set; }
            public int? user_id { get; set; }
            public double? commission_percent { get; set; }
            public double? commission_prosthetic_percent { get; set; }
            public DateTime? birth_date { get; set; }
            public bool? commission_prosthetic { get; set; }
            public bool? hired_tiss { get; set; }
            public bool? active { get; set; }
            public bool? imported { get; set; }
            public int FuncionarioID { get; set; }

    }
}
