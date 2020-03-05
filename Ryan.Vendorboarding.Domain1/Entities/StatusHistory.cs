using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ryan.VendorOnboarding.Domain.Entities
{
    [Table("StatusHistory")]
    public class StatusHistory
    {
        [Key]
        public int ID { get; set; }
        public int VID { get; set; }

        public string VStatus { get; set; }
        public DateTime UpdatedTime { get; set; }
        public string UpdatedBy { get; set; }
        public string Description { get; set; }
    }
}
