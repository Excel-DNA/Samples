using ExcelDna.Integration;
using log4net;

[assembly:log4net.Config.XmlConfigurator()]

namespace UsingLog4Net
{
    public static class MyAddIn
    {
        static ILog Logger = LogManager.GetLogger("MyAddIn");

        public static double AddThemAndLog(double x, double y)
        {
            Logger.Debug(">>>>> AddThemAndLog called.");
            return x + y;
        }
    }
}
