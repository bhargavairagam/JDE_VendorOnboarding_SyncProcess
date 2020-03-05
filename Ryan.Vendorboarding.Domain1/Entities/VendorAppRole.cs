using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;


namespace Ryan.VendorOnboarding.Domain.Entities
{
    [Table("VendorAppRole")]
    public class VendorAppRole
    {
        [Key]
        public int ID { get; set; }
        public string Username { get; set; }
        public string UserRole { get; set; }
        public string UserType { get; set; }


    }
}
