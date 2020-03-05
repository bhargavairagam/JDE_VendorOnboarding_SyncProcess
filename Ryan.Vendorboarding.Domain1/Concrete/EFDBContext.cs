using System;
using System.Data.Entity;
using Ryan.VendorOnboarding.Domain.Entities;

namespace Ryan.VendorOnboarding.Domain.Concrete
{
    public class EFDbContext : DbContext
    {
        public EFDbContext() : base("VendorOnboardingConn")
        {
            Database.SetInitializer<EFDbContext>(null);
        }

        public DbSet<VendorProfile> VendorProfiles { get; set; }
        public DbSet<Status> Statuses { get; set; }
        public DbSet<StatusHistory> StatusHistory { get; set; }

        public DbSet<VendorAddressContacts> VendorAddressContacts { get; set; }

        public DbSet<Packages>  VendorPackages{ get; set; }

        public DbSet<VendorTypes> VendorTypes { get; set; }
        public DbSet<JdeStatuses> JdeStatuses { get; set; }
        public DbSet<VendorAppRole> VendorAppRoles { get; set; }

    }
}
