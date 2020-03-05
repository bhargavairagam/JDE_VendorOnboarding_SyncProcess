using System;
using System.Collections.Generic;
using System.Text;
using Ryan.VendorOnboarding.Domain.Abstract;
using Ryan.VendorOnboarding.Domain.Entities;
using System.Threading.Tasks;
using Ryan.VendorOnboarding.Domain;
using System.Linq;
using System.Data.Entity.Migrations;

namespace Ryan.VendorOnboarding.Domain.Concrete
{
    public class EFVendorStatusHistory : IStatusHistoryRepositary
    {
        private EFDbContext context = new EFDbContext();


        public async Task<IEnumerable<StatusHistory>> GetStatusHistoryByID(int ein)
        {
            return await Task<IEnumerable<StatusHistory>>.Factory.StartNew(() =>
            {
                return context.StatusHistory.Where(a => a.VID == ein).OrderByDescending(x=> x.UpdatedTime).Distinct().ToList();
            });

          
        }

        public async Task<StatusHistory> GetStatusHistoryDetailsByID(int ID)
        {
            return await Task<StatusHistory>.Factory.StartNew(() =>
            {
                return context.StatusHistory.FirstOrDefault(a => a.ID == ID);
            });


        }

        public async Task<EFDbActionResult> SaveStatusHistoryByID(StatusHistory statushistory)
        {
            EFDbActionResult results = new EFDbActionResult();
            List<string> errors = new List<string>();

            try
            {
                statushistory.UpdatedTime = DateTime.Now;
                context.StatusHistory.Add(statushistory);
                results.SavedID = context.SaveChanges();
            }
            catch (Exception ex)
            {
                errors.Add("Cannot find Status History table to add record");
                // errors.AddRange(ErrorHelper.GetAllExceptionMessages(ex));
            }
            results.Errors = errors;


            return await Task<EFDbActionResult>.Factory.StartNew(() =>
            {
                return results;
            });
        }
    }
}
