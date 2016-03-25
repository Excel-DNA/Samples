using ExcelDna.Integration;

namespace RtdClock_IExcelObservable
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object dnaRtdClock_IExcelObservable()
        {
            string functionName = "dnaRtdClock_IExcelObservable";
            object paramInfo = null; // could be one parameter passed in directly, or an object array of all the parameters: new object[] {param1, param2}
            return ExcelAsyncUtil.Observe(functionName, paramInfo, () => new ExcelObservableClock());
        }
    }
}
