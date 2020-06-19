using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace GenericHost
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            BuildConfiguration()
                .Register();
        }

        public void AutoClose()
        {
        }

        static AddInConfiguration BuildConfiguration()
        {
            return new AddInConfiguration
            {
                FunctionRegistration = new MyCustomFunctionRegistration(),
                RibbonRegistration = new MyCustomRibbonRegistration()
            };
        }
    }
}
