using ExcelDna.Integration;

namespace RtdClock_IExcelObservable
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object dnaRtdClock_IExcelObservable(string param)
        {
            string functionName = "dnaRtdClock_IExcelObservable";
            object paramInfo = param; // could be one parameter passed in directly, or an object array of all the parameters: new object[] {param1, param2}
            return ExcelAsyncUtil.Observe(functionName, paramInfo, () => new ExcelObservableClock());
        }

        [ExcelFunction(Description = "Provides a thread safe ticking clock", IsThreadSafe = true)]
        public static object dnaRtdClock_IExcelObservableThreadSafe(string param)
        {
            string functionName = "dnaRtdClock_IExcelObservableThreadSafe";
            object paramInfo = param; // could be one parameter passed in directly, or an object array of all the parameters: new object[] {param1, param2}
            return ExcelAsyncUtil.Observe(functionName, paramInfo, () => new ExcelObservableClock());
        }
    }
}
