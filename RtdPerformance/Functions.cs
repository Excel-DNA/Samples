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

        public static object rtdWrapper(object topic1, object topic2)
        {
            return Excel(xlfRtd, RtdServer.ServerProgId, "", topic1, topic2);
        }

        public static object rtdWrapperTestNulls(object topicInfo)
        {
            return Excel(xlfRtd, RtdServer.ServerProgId, null, topicInfo, null, null, null);
        }
    }
}
