using Ryan.VendorOnboarding.Domain.Abstract;
using Ryan.VendorOnboarding.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ryan.VendorOnboarding.Domain.Concrete
{
    public class EFVendorAppRoleRepositary : IVendorAppRoleRepositary
    {
        private EFDbContext context;

        public EFVendorAppRoleRepositary()
        {
            context = new EFDbContext();
        }

        public async Task<VendorAppRole> GetVendorRole(string username)
        {
            return await Task<VendorAppRole>.Factory.StartNew(() =>
            {
                return context.VendorAppRoles.Where(x => x.Username == username).FirstOrDefault();
            });
           
           
        }

        public async Task<IEnumerable<VendorAppRole>> GetAllVendorRoles()
        {
            return await Task<IEnumerable<VendorAppRole>>.Factory.StartNew(() =>
            {
                return context.VendorAppRoles.Distinct();
            });

        }
    }
}
