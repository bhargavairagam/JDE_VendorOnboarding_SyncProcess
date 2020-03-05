using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Ryan.VendorOnboarding.Domain.Entities
{
    [Table("Packages")]
    public class Packages
    {
        [Key]
        public int ID { get; set; }
        public string MRI1 { get; set; }
        public string MRI2 { get; set; }
        public string MRI3 { get; set; }
        public string MRI4 { get; set; }
        public string YARDI5 { get; set; }
        public string YARDI6 { get; set; }
        public string YARDI7 { get; set; }
        public string YARDI8 { get; set; }

        public string JDE { get; set; }
        public string PROCORE { get; set; }
        public string PROLIANCE { get; set; }

    }
}
