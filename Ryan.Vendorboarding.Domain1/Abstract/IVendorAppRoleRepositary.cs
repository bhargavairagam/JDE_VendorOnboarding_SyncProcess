using System.Collections.Generic;
using System.Threading.Tasks;
using Ryan.VendorOnboarding.Domain.Entities;

namespace Ryan.VendorOnboarding.Domain.Abstract
{
    public interface IVendorAppRoleRepositary
    {
        Task<VendorAppRole> GetVendorRole(string username);
        Task<IEnumerable<VendorAppRole>> GetAllVendorRoles();
    }
}
