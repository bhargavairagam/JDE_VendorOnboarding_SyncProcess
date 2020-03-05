using Ryan.VendorOnboarding.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ryan.VendorOnboarding.Domain.Abstract
{
    public interface IVendorAddressContactsRepositary
    {
        VendorAddressContacts GetVendorAddressContacts(string username);
        IEnumerable<VendorAddressContacts> GetAllVendorAddressContacts();
        IDbActionResult SaveVendorAddressContacts(VendorAddressContacts vendorProfile);

    }
}
