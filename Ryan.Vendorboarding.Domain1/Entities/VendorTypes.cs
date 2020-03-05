using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ryan.VendorOnboarding.Domain.Entities
{
    [Table("VENDORTYPES")]
    public class VendorTypes
    {

        [Key]
        public int ID { get; set; }
        public string VENDORTYPETEXT { get; set; }
        public string VENDORTYPEVALUE { get; set; }
    }
}
