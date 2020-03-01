using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace HttpClientSample
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            RegisterFunctions();
        }

        public void AutoClose()
        {
        }

        void RegisterFunctions()
        {
            ExcelRegistration.GetExcelFunctions()
                             .ProcessAsyncRegistrations(nativeAsyncIfAvailable: false)
                             .RegisterFunctions();
        }
    }
}
