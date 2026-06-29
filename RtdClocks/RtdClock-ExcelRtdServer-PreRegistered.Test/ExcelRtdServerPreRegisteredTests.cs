using System;
using System.Threading;
using ExcelDna.Testing;
using Microsoft.Office.Interop.Excel;
using Xunit;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;

namespace RtdClock_ExcelRtdServer_PreRegistered.Test
{
    // The path is relative to the test project's output directory.
    [ExcelTestSettings(AddIn = @"..\..\..\..\RtdClock-ExcelRtdServer-PreRegistered\bin\Debug\net472\RtdClock-ExcelRtdServer-PreRegistered-AddIn")]
    public class ExcelRtdServerPreRegisteredTests : IDisposable
    {
        readonly Workbook _workbook;

        public ExcelRtdServerPreRegisteredTests()
        {
            _workbook = Util.Application.Workbooks.Add();
            Util.Application.Run("RegisterRtdClockComServer");
        }

        public void Dispose()
        {
            _workbook.Close(SaveChanges: false);
            Util.Application.Run("UnregisterRtdClockComServer");
        }

        [ExcelFact]
        public void WorksheetFormulaCanUsePreRegisteredRtdServer()
        {
            var worksheet = (Worksheet)_workbook.Sheets[1];
            var formulaCell = worksheet.Range["A1"];

            formulaCell.Formula = "=RTD(\"RtdClock.ClockServer\",,\"\")";

            var value = WaitForClockValue(formulaCell);

            Assert.Matches(@"^\d{2}:\d{2}:\d{2}( \(ConnectData\))?$", value);
        }

        [ExcelFact]
        public void WorksheetFunctionCanUsePreRegisteredRtdServer()
        {
            var worksheet = (Worksheet)_workbook.Sheets[1];
            var formulaCell = worksheet.Range["A1"];

            formulaCell.Formula = "=dnaRtdClock_ExcelRtdServer_PreRegistered()";

            var value = WaitForClockValue(formulaCell);

            Assert.Matches(@"^\d{2}:\d{2}:\d{2}( \(ConnectData\))?$", value);
        }

        static string WaitForClockValue(ExcelRange cell)
        {
            var deadline = DateTime.UtcNow.AddSeconds(10);
            object lastValue = null;

            while (DateTime.UtcNow < deadline)
            {
                lastValue = cell.Value;
                var value = lastValue as string;

                if (IsClockValue(value))
                {
                    return value;
                }

                Thread.Sleep(250);
            }

            throw new TimeoutException($"Expected an RTD clock value, but the last cell value was '{lastValue}'.");
        }

        static bool IsClockValue(string value)
        {
            if (value == null || value.Length < 8)
            {
                return false;
            }

            var timePart = value.Substring(0, 8);
            return DateTime.TryParseExact(
                timePart,
                "HH:mm:ss",
                null,
                System.Globalization.DateTimeStyles.None,
                out _);
        }
    }
}
