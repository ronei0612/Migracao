using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Migracao.Models.DentalOffice
{

    [Table("schedules")]
    public class Agendamento
    {

        [Key]
        public int id { get; set; }
        public string? cellphone { get; set; }
        public string? description { get; set; }
        public string? email { get; set; }
        public string? job_id { get; set; }
        public string? phone { get; set; }
        public DateTime? arrived_at { get; set; }
        public DateTime? created_at { get; set; }
        public DateTime? deleted_at { get; set; }
        public DateTime? schedule_end { get; set; }
        public DateTime? schedule_start { get; set; }
        public DateTime? served_at { get; set; }
        public DateTime? updated_at { get; set; }
        public string? notes { get; set; }
        public int? chair_id { get; set; }
        [ForeignKey("chair_id")]
        public Sala Sala { get; set; }

        public int? clinic_id { get; set; }
        public int? customer_id { get; set; }
        [ForeignKey("customer_id ")]
        public virtual Pacientes Paciente { get; set; }
        public int? dentist_id { get; set; }
        public int? discipline_id { get; set; }
        public int? duration { get; set; }
        public int? marketing_action_id { get; set; }
        public int? periodic_schedule_id { get; set; }
        public int? schedule_integration_id { get; set; }
        public int? schedule_reason_id { get; set; }
        public int? schedule_situation_id { get; set; }
        public int? schedule_type_id { get; set; }
        public int? sms_confirmation { get; set; }
        public bool? confirm_by_sms { get; set; }
        public bool? confirm_by_whatsapp { get; set; }
        public bool? confirmed_by_whatsapp { get; set; }
        public bool? imported { get; set; }
        public bool? periodic { get; set; }
        public bool? personal { get; set; }
        //public bool? public {get;set;}
        public bool? teleconsultation { get; set; }

    }

}
