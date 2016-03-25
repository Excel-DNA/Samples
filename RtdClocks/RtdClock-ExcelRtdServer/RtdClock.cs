using ExcelDna.Integration;

namespace RtdClock_ExcelRtdServer
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object dnaRtdClock_ExcelRtdServer()
        {
            // Call the Excel-DNA RTD wrapper, which does dynamic registration of the RTD server
            // Note that the topic information needs at least one string - it's not used in this sample
            return XlCall.RTD(RtdClockServer.ServerProgId, null, "");
        }
    }
}
