
using Ryan.VendorOnboarding.Domain;
using Ryan.VendorOnboarding.Domain.Abstract;
using Ryan.VendorOnboarding.Domain.Concrete;
using Ryan.VendorOnboarding.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ryan.VendorOnboarding.Domain.Concrete
{
    public class EFVendorAddressContactsRespositary : IVendorAddressContactsRepositary
    {
        private EFDbContext context = new EFDbContext();


        public IEnumerable<VendorAddressContacts> GetAllVendorAddressContacts()
        {

            IEnumerable<VendorAddressContacts> res = context.VendorAddressContacts.Distinct().ToList();
            return res;

        }

        public VendorAddressContacts GetVendorAddressContacts(string EIN)
        {
            VendorAddressContacts res = new VendorAddressContacts();
            res = context.VendorAddressContacts.FirstOrDefault(x => x.VendorEIN == EIN);


            return res;
        }

        public IDbActionResult SaveVendorAddressContacts(VendorAddressContacts vendorProfile)
        {
            throw new NotImplementedException();
        }
    }
}
