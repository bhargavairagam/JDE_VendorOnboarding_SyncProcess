using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ryan.VendorOnboarding.Domain
{
    public interface IDbActionResult
    {
        IEnumerable<string> Errors { get; set; }
        bool Succeeded { get; }
    }
}
