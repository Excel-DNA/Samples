using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace GenericHost
{
    public interface IFunctionRegistration
    {
        IEnumerable<ExcelFunctionRegistration> GetFunctionRegistrations();
    }

    public interface IRibbonRegistration
    {
        List<ExcelComAddIn> GetRibbonAddIns();
    }
}
