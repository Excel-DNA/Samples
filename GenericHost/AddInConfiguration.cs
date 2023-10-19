using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace GenericHost
{
    public class AddInConfiguration
    {
        public IFunctionRegistration FunctionRegistration;
        public IRibbonRegistration RibbonRegistration;

        public AddInConfiguration()
        {
            FunctionRegistration = new DefaultFunctionRegistration();
        }

        public void Register()
        {
            if (FunctionRegistration != null)
                ExcelRegistration.RegisterFunctions(FunctionRegistration.GetFunctionRegistrations());

            if (RibbonRegistration != null)
            {
                foreach (var addIn in RibbonRegistration.GetRibbonAddIns())
                {
                    ExcelComAddInHelper.LoadComAddIn(addIn);
                }
            }
        }
    }
}
