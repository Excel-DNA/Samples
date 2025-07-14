using ExcelDna.Integration;
using System.Collections.Generic;

namespace Registration.Sample
{
    public static class MapArrayFunctionRegistration
    {
        [ExcelFunctionProcessor]
        public static IEnumerable<IExcelFunctionInfo> ProcessMapArrayFunctions(IEnumerable<IExcelFunctionInfo> registrations, IExcelFunctionRegistrationConfiguration config)
        {
            return ExcelDna.Registration.MapArrayFunctionRegistration.ProcessMapArrayFunctions(registrations, config);
        }
    }
}
