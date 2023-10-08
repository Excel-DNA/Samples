using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

namespace DnaComServer
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary
    {
        string ComLibraryHello();
        double Add(double x, double y);
    }

    [ComDefaultInterface(typeof(IComLibrary))]
    public class ComLibrary
    {
        public string ComLibraryHello()
        {
            return "Hello from DnaComServer.ComLibrary";
        }

        public double Add(double x, double y)
        {
            return x + y;
        }
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary2
    {
        string ComLibrary2Hello();
        double Add2(double x, double y);
    }

    [ComDefaultInterface(typeof(IComLibrary2))]
    public class ComLibrary2
    {
        public string ComLibrary2Hello()
        {
            return "Hello from DnaComServer.ComLibrary2";
        }

        public double Add2(double x, double y)
        {
            return x + y;
        }
    }

    [ComVisible(false)]
    public class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }

    public static class Functions
    {
        [ExcelFunction]
        public static object DnaComServerHello()
        {
            return "Hello from DnaComServer!";
        }
    }
}
