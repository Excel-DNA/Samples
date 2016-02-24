using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;

namespace MasterSlave
{
    public static class SlaveDriver
    {
        // Passing this as a parameter makes it hard to call from the ribbon...
        const string path = @"C:\Work\Excel-DNA\Samples\MasterSlave\Slave\bin\Debug\Slave-AddIn.xll";
        public static void RegisterSlave()
        {
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Excel(xlcMessage, true, $"Loading {path}");
            Thread.Sleep(1000);
            ExcelIntegration.RegisterXLL(path);
        }

        public static void UnregisterSlave()
        {
            ExcelIntegration.UnregisterXLL(path);
        }
    }
}
