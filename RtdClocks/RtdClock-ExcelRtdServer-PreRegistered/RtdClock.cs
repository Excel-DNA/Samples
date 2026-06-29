using ExcelDna.Integration;

namespace RtdClock_ExcelRtdServer_PreRegistered
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock through Excel's pre-registered RTD path")]
        public static object dnaRtdClock_ExcelRtdServer_PreRegistered()
        {
            return XlCall.Excel(XlCall.xlfRtd, RtdClockServer.ServerProgId, null, "");
        }
    }
}
