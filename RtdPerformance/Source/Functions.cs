using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;

namespace RtdPerformance
{
    public static class Functions
    {
        public static object rtdHello()
        {
            return "Hello from RtdPerformance Add-in!";
        }

        public static object rtdWrapperNormal(object topic1, object topic2)
        {
            return Excel(xlfRtd, RtdServer.ServerProgId, "", topic1, topic2);
        }
    }
}
