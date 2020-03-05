using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ryan.VendorOnboarding.Domain.Entities
{
    [Table("VStatus")]
    public class Status
    {
        [Key]
        public int ID { get; set; }
        public string StatusText { get; set; }
        public string StatusValue { get; set; }

    }
}
