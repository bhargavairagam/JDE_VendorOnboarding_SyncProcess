using System.Collections.Generic;
using System.Threading.Tasks;
using Ryan.VendorOnboarding.Domain;
using Ryan.VendorOnboarding.Domain.Entities;

namespace Ryan.VendorOnboarding.Domain.Abstract
{
    public interface IVendorTypesRepositary
    {
        Task<IEnumerable<VendorTypes>> GetVendorTypes();
    }
}
