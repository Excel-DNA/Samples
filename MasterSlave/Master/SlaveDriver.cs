using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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

        static string GetSlavePath()
        {
            var masterPath = ExcelDnaUtil.XllPath;
            
            // On my machine we want to change
            // "C:\\Work\\Excel-DNA\\Samples\\MasterSlave\\Master\\bin\\Debug\\net472\\Master-AddIn64.xll"
            // to 
            // "C:\\Work\\Excel-DNA\\Samples\\MasterSlave\\Slave\\bin\\Debug\\net472\\Slave-AddIn64.xll"

            var slavePath = masterPath.Replace(@"\Master\", @"\Slave\");
            slavePath = slavePath.Replace(@"\Master-", @"\Slave-");
            return slavePath;
        }

        public static void RegisterSlave()
        {
            var slavePath = GetSlavePath();
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Console.Beep();
            Excel(xlcMessage, true, $"Loading {slavePath}");
            Thread.Sleep(1000);
            ExcelIntegration.RegisterXLL(slavePath);
        }

        public static void UnregisterSlave()
        {
            var slavePath = GetSlavePath();
            ExcelIntegration.UnregisterXLL(slavePath);
        }
    }
}
