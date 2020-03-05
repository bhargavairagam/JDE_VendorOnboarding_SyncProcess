using System;
using System.Collections.Generic;
using System.Text;
using Ryan.VendorOnboarding.Domain.Abstract;
using Ryan.VendorOnboarding.Domain.Entities;
using System.Threading.Tasks;
using Ryan.VendorOnboarding.Domain;
using System.Linq;
using System.Data.Entity.Migrations;
using System.Data.Entity;

namespace Ryan.VendorOnboarding.Domain.Concrete
{
    public class EFVendorProfileRepositary : IVendorProfileRepositary
    {
        private EFDbContext context = new EFDbContext();

        public IEnumerable<VendorProfile> GetAllVendorDetails()
        {

            IEnumerable<VendorProfile> res = context.VendorProfiles.Distinct();
            return res;

        }

        public async Task<IEnumerable<VendorProfile>> GetVendorProfiles()
        {

            return await Task<IEnumerable<VendorProfile>>.Factory.StartNew(() =>
            {
                return context.VendorProfiles.Distinct();
            });

        }

        public async Task<IEnumerable<VendorProfile>> GetVendorProfiles(string name, string status, string package)
        {
            

            return await Task<IEnumerable<VendorProfile>>.Factory.StartNew(() =>
            {
                return context.VendorProfiles.Distinct().ToList();
            });

        }

        public async Task<VendorProfile> GetVendorDetialsByJdeID(string jdeID)
        {

            return await Task<VendorProfile>.Factory.StartNew(() =>
            {
                return context.VendorProfiles.Where(a => a.JDEVendorID == jdeID).FirstOrDefault();
            });

        }

        public async Task<IEnumerable<VendorProfile>> GetVendorListByName(string name)
        {

            return await Task<IEnumerable<VendorProfile>>.Factory.StartNew(() =>
            {
                return context.VendorProfiles.Where(a => (a.VendorDBAName.ToLower().Contains(name)));
            });

        }


        public async Task<IEnumerable<VendorProfile>> GetVendorListByNameAndDate(string name , DateTime? SubmittedDate, string jdeid)
        {
            
            IEnumerable<VendorProfile> tak = context.VendorProfiles.Distinct();

            if (!(SubmittedDate is null))
            {
                tak = context.VendorProfiles.Where(a => (DbFunctions.TruncateTime(a.SubmittedTime) == DbFunctions.TruncateTime(SubmittedDate.Value))).ToList();
            }
            if (!string.IsNullOrEmpty(name))
            {
                tak = tak.Where(a => (a.VendorDBAName.ToLower().Contains(name))).ToList();
            }

           

            if(!string.IsNullOrEmpty(jdeid))
            {
                tak = tak.Where(a => (a.JDEVendorID == jdeid)).ToList();
            }

            return await Task<IEnumerable<VendorProfile>>.Factory.StartNew(() =>
            {
                return tak;
            });

        }

        public  VendorProfile GetVendorDetialsByEin(string ein)
        {
            VendorProfile v = new VendorProfile();
            try
            {
                v = context.VendorProfiles.Where(a => a.VendorEIN == ein).FirstOrDefault(); 
            }
            catch(Exception ex)
            {

            }

            return v;

        }


        public async Task<EFDbActionResult> SaveVendorDetials(VendorProfile vendorProfile)
        {
            EFDbActionResult results = new EFDbActionResult();
            List<string> errors = new List<string>();

            try
            {
                vendorProfile.LastUpdatedTime = DateTime.Now;
                context.VendorProfiles.AddOrUpdate(vendorProfile);
                context.SaveChanges();
                results.SavedID = vendorProfile.ID;
            }
            catch (Exception ex)
            {
                errors.Add("Cannot find custompage to update");
                // errors.AddRange(ErrorHelper.GetAllExceptionMessages(ex));
               
            }
            results.Errors = errors;


            return await Task<EFDbActionResult>.Factory.StartNew(() =>
            {
                return results;
            });


        }


        public IEnumerable<Status> GetVendorStatuses()
        {
                return context.Statuses.ToList();
        }

    }
}
