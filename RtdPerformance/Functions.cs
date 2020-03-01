using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using static ExcelDna.Integration.XlCall;

namespace RtdPerformance
{
    public static class Functions
    {
        static Application _xlApp;
        static Functions()
        {
            _xlApp = ExcelDnaUtil.Application as Application;
        }

        public static object SayHello() => "Hello from RtdPerformance!";

        public static object rtdClock(string topicInfo) => RTD(RtdServer.ServerProgId, null, topicInfo);

        public static object rtdClockDirect(object topicInfo)
        {
            var tis = topicInfo.ToString();
            return $"{tis}:({tis.Length}) {Excel(xlfRtd, RtdServer.ServerProgId, null, topicInfo)}";
        }

        public static object rtdClockArray(object topicInfo)
        {
            Debug.Write(DateTime.Now);
            var tis = topicInfo.ToString();
            return new object[,] {{ $"{tis}:({tis.Length}) {Excel(xlfRtd, RtdServer.ServerProgId, null, topicInfo)}", tis.Length }};
        }

        [ExcelFunction(IsThreadSafe = true)]
        public static object rtdClockThreadSafe(object topicInfo) => Excel(xlfRtd, RtdServer.ServerProgId, null, topicInfo);


        public static object rtdClockWsRtd(object topicInfo)
        {
            var tis = topicInfo.ToString();
            return $"{tis}:({tis.Length}) {Excel(xlfRtd, RtdServer.ServerProgId, null, topicInfo)}";
            _xlApp.WorksheetFunction.RTD(RtdServer.ServerProgId, null, topicInfo);
        }

        [ExcelFunction(IsThreadSafe = true)]
        public static object timeDirect(object input) => DateTime.Now;
    }
}
