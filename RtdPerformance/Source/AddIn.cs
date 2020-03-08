using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.ComInterop;
using ExcelDna.Integration;


namespace RtdPerformance
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();

            ExcelIntegration.RegisterRtdWrapper(RtdServer.ServerProgId, null,
                new ExcelFunctionAttribute 
                { 
                    Name = "rtdWrapperFast", 
                    Description = "Get the real-time data item",
                    IsExceptionSafe = true, 
                    IsThreadSafe = false 
                }, 
                new List<object>
                {
                    new ExcelArgumentAttribute
                    {
                        Name = "First topic string"
                    },
                    new ExcelArgumentAttribute
                    {
                        Name = "Second topic string"
                    },
                });

        }

        public void AutoClose()
        {
        }
    }
}
