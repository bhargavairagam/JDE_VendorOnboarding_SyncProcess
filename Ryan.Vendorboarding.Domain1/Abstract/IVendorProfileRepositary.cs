
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Ryan.VendorOnboarding.Domain.Entities;

namespace Ryan.VendorOnboarding.Domain.Abstract
{
    public interface IVendorProfileRepositary
    {
        Task<IEnumerable<VendorProfile>> GetVendorProfiles();
        Task<VendorProfile> GetVendorDetialsByJdeID(string ein);
        Task<EFDbActionResult> SaveVendorDetials(VendorProfile vendorProfile);
        Task<IEnumerable<VendorProfile>> GetVendorListByNameAndDate(string name, DateTime? SubmittedDate, string jdeid);
        Task<IEnumerable<VendorProfile>> GetVendorListByName(string name);

    }

    public interface IStatusHistoryRepositary
    {
        
        Task<IEnumerable<StatusHistory>> GetStatusHistoryByID(int vid);
        Task<EFDbActionResult> SaveStatusHistoryByID(StatusHistory statushistory);


    }

    public interface IVendorPackagesRepositary
    {

        Task<IEnumerable<StatusHistory>> GetVendorPackagesByID(int vid);
        Task<EFDbActionResult> SaveStatusHistoryByID(StatusHistory statushistory);


    }
}
