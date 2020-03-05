using Ryan.VendorOnboarding.Domain.Abstract;
using Ryan.VendorOnboarding.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ryan.VendorOnboarding.Domain.Concrete
{
    public class EFJdeStatusesRepositary : IJdeStatusesRepositary
    {
        private EFDbContext context = new EFDbContext();

      
        public async Task<IEnumerable<JdeStatuses>> GetJdeStatuses()
        {

            return await Task<IEnumerable<JdeStatuses>>.Factory.StartNew(() =>
            {
                return context.JdeStatuses.Distinct();
            });

        }
    }
}
