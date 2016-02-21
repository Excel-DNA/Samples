using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace TestController
{
    public class Class1
    {
        public static void Main()
        {
            var process = LoadAndUnload();
            process.EnableRaisingEvents = true;
            process.Exited += Process_Exited;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Console.ReadLine();
        }

        private static void Process_Exited(object sender, EventArgs e)
        {
            var process = (Process)sender;
            Debug.Print($"Excel exited with code {process.ExitCode}");
        }

        public static Process LoadAndUnload()
        {
            var Application = new Application();
            var process = Process.GetProcessesByName("Excel").First();

            Debug.Print(Application.Version);
            Application.RegisterXLL(@"C:\Work\Excel-DNA\Samples\MasterSlave\Master\bin\Debug\Master-AddIn.xll");
            foreach (AddIn addIn in Application.AddIns2)
            {
                if (addIn.IsOpen)
                    Debug.Print($"> {addIn.Name} @ {addIn.Path}");
            }
            Application.Visible = true;
            //Application.Do
            //Application.Run("LoadSlave");
            //Application.Run("UnloadSlave");
            Application.Quit();

            return process;
        }
    }
}
