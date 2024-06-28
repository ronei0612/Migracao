using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{

    [Table("customers")]
    public class Pacientes
    {
        [Key]
        public int id { get; set; }
        public string? title { get; set; }
        public string? name { get; set; }
        public string? photo_file_name { get; set; }
        public string? photo_content_type { get; set; }
        public long? photo_file_size { get; set; }
        public DateTime? photo_updated_at { get; set; }
        public string? record_number { get; set; }
        public DateTime? record_date { get; set; }
        public int? gender { get; set; }
        public string? dental_arcade { get; set; }
        public DateTime? birth_date { get; set; }
        public string? birth_place { get; set; }
        public int? marital_status { get; set; }
        public int? dentist_id { get; set; }
        public int? document_id { get; set; }
        public short? is_holder { get; set; }
        public int? holder_id { get; set; }
        public int? company_id { get; set; }
        public string? department { get; set; }
        public int? profession_id { get; set; }
        public string? father_name { get; set; }
        public int? father_profession_id { get; set; }
        public string? mother_name { get; set; }
        public int? mother_profession_id { get; set; }
        public int? sponsor_id { get; set; }
        public int? indicator_id { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? updated_at { get; set; }
        public int? customer_group_id { get; set; }
        public int? customer_situation_id { get; set; }
        public short? is_sponsor { get; set; }
        public string? indicator_type { get; set; }
        public int? clinic_id { get; set; }
        public decimal? payment_credit { get; set; }
        public string? father_rg { get; set; }
        public string? father_cpf { get; set; }
        public string? father_phone { get; set; }
        public string? mother_rg { get; set; }
        public string? mother_cpf { get; set; }
        public string? mother_phone { get; set; }
        public string? obs { get; set; }
        public int? nota_facil_id { get; set; }
        public short? active { get; set; }
        public short? asaas_sms_notifications { get; set; }
        public short? imported { get; set; }     
        public int? PessoaID { get; set; }
        public int? ConsumidorID { get; set; }
        public string? NomeErrado { get; set; }



    }
}
