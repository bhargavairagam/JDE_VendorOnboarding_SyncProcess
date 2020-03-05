using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ryan.VendorOnboarding.Domain.Entities
{
    [Table("JDESTATUSES")]
    public class JdeStatuses
    {

        [Key]
        public int ID { get; set; }
        public string JDESTATUSTEXT { get; set; }
        public string JDESTATUSVALUE { get; set; }
    }
}
