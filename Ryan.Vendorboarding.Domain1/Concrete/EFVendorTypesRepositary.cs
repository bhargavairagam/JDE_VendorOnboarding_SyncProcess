using Ryan.VendorOnboarding.Domain.Abstract;
using Ryan.VendorOnboarding.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ryan.VendorOnboarding.Domain.Concrete
{
    public class EFVendorTypesRepositary: IVendorTypesRepositary
    {
        private EFDbContext context = new EFDbContext();


        public async Task<IEnumerable<VendorTypes>> GetVendorTypes()
        {

            return await Task<IEnumerable<VendorTypes>>.Factory.StartNew(() =>
            {
                return context.VendorTypes.Distinct();
            });

        }
    }
}
