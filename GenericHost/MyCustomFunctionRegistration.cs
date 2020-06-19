using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Registration;

namespace GenericHost
{
    public class MyCustomFunctionRegistration : IFunctionRegistration
    {
        public IEnumerable<ExcelFunctionRegistration> GetFunctionRegistrations()
        {
            var allFunctions = ExcelRegistration.GetExcelFunctions().ToList();
            foreach (var reg in allFunctions)
                reg.FunctionAttribute.Name = reg.FunctionAttribute.Name + ".Custom";
            return allFunctions;
        }
    }
}
