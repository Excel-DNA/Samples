using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Threading.Tasks.Schedulers;
using ExcelDna.Integration;
using ExcelDna.Utils;

namespace LimitedConcurrencyAsync
{
    public static class AsyncFunctions
    {
        static TaskFactory _fourThreadFactory;

        static AsyncFunctions()
        {
            // This initialization could be lazy (and of course be any other TaskScheduler)
            var fourThreadScheduler = new LimitedConcurrencyLevelTaskScheduler(4);
            _fourThreadFactory = new TaskFactory(fourThreadScheduler);
        }

        public static object Sleep(int seconds)
        {
            Debug.Print($"Sleep Call: {seconds}");
            // The callerFunctionName and callerParameters are internally combined and used as a 'key' 
            // to link the underlying RTD calls together.
            string callerFunctionName = "Sleep";
            object callerParameters = new object[] {seconds}; // This need not be an array if it's just a single parameter

            var result = AsyncTaskUtil.RunAsTask(callerFunctionName, callerParameters, _fourThreadFactory, () =>
                {
                    Thread.Sleep(seconds * 1000);
                    return "Slept on Thread " + Thread.CurrentThread.ManagedThreadId;
                });
            Debug.Print($"Sleep Result: {result}");
            return result;
        }

        public static object SleepPerCaller(int seconds)
        {
            // Trick to get each call to be a separate instance
            // Normally you only want to add the actual parameters passed in
            object callerReference = XlCall.Excel(XlCall.xlfCaller);
            string callerFunctionName = "SleepPerCaller";
            object callerParameters = new object[] { seconds, callerReference };

            // The RunTask version (instead of RunAsTask used above) is more flexible if the Task will be created in some other way.
            return AsyncTaskUtil.RunTask(callerFunctionName, callerParameters, () =>
            {
                // The function here should return the Task to run
                return _fourThreadFactory.StartNew(() =>
                {
                    Thread.Sleep(seconds * 1000);
                    return string.Format("Slept on Thread {0}, called from {1}", Thread.CurrentThread.ManagedThreadId, callerReference);
                });
            });
        }


    }
}
