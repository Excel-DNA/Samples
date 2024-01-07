using ExcelDna.Integration;

namespace AsyncThreadSafe
{
    public static class Functions
    {
        [ExcelFunction(IsThreadSafe = true)]
        public static object SayHelloAsyncThreadSafe(string name)
        {
            return SayHelloAsync(nameof(SayHelloAsyncThreadSafe), name);
        }

        public static object SayHelloAsyncNotThreadSafe(string name)
        {
            return SayHelloAsync(nameof(SayHelloAsyncNotThreadSafe), name);
        }

        [ExcelFunction(IsThreadSafe = true)]
        public static object SayHelloAsyncFast(string name)
        {
            return ExcelAsyncUtil.Run(nameof(SayHelloAsyncFast), new object[] { name }, () => $"Hello {name}");
        }

        [ExcelFunction(IsThreadSafe = true)]
        public static object SayHelloAsyncSlow(string name)
        {
            return ExcelAsyncUtil.Run(nameof(SayHelloAsyncSlow), new object[] { name }, () => SayHelloWithDelay10(name));
        }

        private static object SayHelloAsync(string callerFunctionName, string name)
        {
            int threadId = Thread.CurrentThread.ManagedThreadId;
            return ExcelAsyncUtil.Run(callerFunctionName, new object[] { name }, () => SayHelloWithDelay(name, threadId));
        }

        private static string SayHelloWithDelay(string name, int threadId)
        {
            Thread.Sleep(2000);
            return $"Hello {name} (thread {threadId})";
        }

        private static string SayHelloWithDelay10(string name)
        {
            Thread.Sleep(10000);
            return $"Hello {name}";
        }
    }
}
