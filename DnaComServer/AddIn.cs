using System.Runtime.InteropServices;
using ExcelDna.ComInterop;
using ExcelDna.Integration;

[assembly:ComVisible(false)]
    
namespace DnaComServer
{

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibraryC
    {
        string ComLibraryHello();
        double Add(double x, double y);
    }

    [ComVisible(true)]
    [ComDefaultInterface(typeof(IComLibraryC))]
    public class ComLibraryC
    {
        public ComLibraryC()
        {
        }
        
        public string ComLibraryHello()
        {
            return "Hello from DnaComServer.ComLibrary";
        }

        public double Add(double x, double y)
        {
            return x + y;
        }
    }

    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IComLibrary2C
    {
        string ComLibrary2Hello();
        double Add2(double x, double y);
    }

    [ComVisible(true)]
    [ComDefaultInterface(typeof(IComLibrary2C))]
    public class ComLibrary2C
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
