using System;
using System.Linq;
using System.Reactive.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration.Utils;

namespace RtdClock_Rx
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static object dnaRtdClock_Rx()
        {
            string functionName = "dnaRtdClock_Rx";
            object paramInfo = null; // could be one parameter passed in directly, or an object array of all the parameters: new object[] {param1, param2}
            return ObservableRtdUtil.Observe(functionName, paramInfo, () => GetObservableClock() );
        }

        static IObservable<string> GetObservableClock()
        {
            return Observable.Timer(dueTime: TimeSpan.Zero, period: TimeSpan.FromSeconds(1))
                             .Select(_ => DateTime.Now.ToString("HH:mm:ss"));
        }
    }
}
