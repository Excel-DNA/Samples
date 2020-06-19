using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Registration;

namespace GenericHost
{
    public class DefaultFunctionRegistration : IFunctionRegistration
    {
        public IEnumerable<ExcelFunctionRegistration> GetFunctionRegistrations() => ExcelRegistration.GetExcelFunctions();
    }
}
