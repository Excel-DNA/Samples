using System.Runtime.InteropServices;
using ExcelDna.Integration;

namespace ComServer
{
    [ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ExcelDna.ComInterop.ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ExcelDna.ComInterop.ComServer.DllUnregisterServer();
        }
    }
}
