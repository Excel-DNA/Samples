using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace RtdArrayTest
{
    public static class TestFunctions
    {
        public static object RtdArrayTest(string prefix)
        {
            object rtdValue = XlCall.RTD("RtdArrayTest.TestRtdServer", null, prefix);
            
            var resultString = rtdValue as string;
            if (resultString == null)
                return rtdValue;

            // We have a string value, parse and return as an 2x1 array
            var parts = resultString.Split(';');
            Debug.Assert(parts.Length == 2);
            var result = new object[2, 1];
            result[0, 0] = parts[0];
            result[1, 0] = parts[1];
            return result;
        }
    }
}
