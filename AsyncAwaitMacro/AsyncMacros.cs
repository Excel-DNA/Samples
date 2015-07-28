using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace AsyncAwaitMacro
{
    public static class AsyncMacros
    {
        static dynamic Application = ExcelDnaUtil.Application;

        [ExcelCommand(MenuName = "AsyncAwaitMacros", MenuText = "DumpDataSlowly")]
        public static void DumpDataSlowly()
        {
            ExcelAsyncTask.Run(DumpDataSlowlyImpl);
        }

        static async Task DumpDataSlowlyImpl()
        {
            // All the code here will run on the main thread - (though any Tasks run internally may do work on separate threads)

            Debug.Print("1> {0}", Thread.CurrentThread.ManagedThreadId);
            await Task.Delay(TimeSpan.FromSeconds(5));
            Application.Range["A1"].Value = DateTime.Now.ToString("> HH:mm:ss");

            Debug.Print("2> {0}", Thread.CurrentThread.ManagedThreadId);
            await Task.Delay(TimeSpan.FromSeconds(5));
            Application.Range["A2"].Value = DateTime.Now.ToString("> HH:mm:ss");

            Debug.Print("3> {0}", Thread.CurrentThread.ManagedThreadId);
            await Task.Delay(TimeSpan.FromSeconds(5));
            Application.Range["A3"].Value = DateTime.Now.ToString("> HH:mm:ss");

            Debug.Print("4> {0}", Thread.CurrentThread.ManagedThreadId);
        }

    }
}
