using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace TestController
{
    public class Class1
    {
        public static void Main()
        {
            LoadAndUnload();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.ReadLine();
        }

        private static void Process_Exited(object sender, EventArgs e)
        {
            var process = (Process)sender;
            Console.WriteLine($"Excel exited with code {process.ExitCode}");
            Debug.Print($"Excel exited with code {process.ExitCode}");
        }

        public static void LoadAndUnload()
        {
            var basePath = @"C:\Work\Excel-DNA\Samples\MasterSlave\";

              var Application = new Application();
            var process = Process.GetProcessesByName("Excel").First();
            process.EnableRaisingEvents = true;
            process.Exited += Process_Exited;

            Debug.Print(Application.Version);
            Application.Visible = true;
            

            Application.RegisterXLL(basePath + @"Master\bin\Debug\Master-AddIn64.xll");
            foreach (AddIn addIn in Application.AddIns2)
            {
                if (addIn.IsOpen)
                    Debug.Print($"> {addIn.Name} @ {addIn.Path}");
            }
            
            // These are macros defined in Master
            Application.Run("RegisterSlave");
            Application.Run("UnregisterSlave");

            Application.Quit();
        }
    }
}
