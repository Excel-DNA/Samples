using System;
using System.Linq;
using System.Reactive.Linq;
using ExcelDna.Integration;

namespace RtdClock_Rx_Registration
{
    public static class RtdClock
    {
        [ExcelFunction(Description = "Provides a ticking clock")]
        public static IObservable<string> dnaRtdClock_Rx_Registration()
        {
            return Observable.Timer(dueTime: TimeSpan.Zero, period: TimeSpan.FromSeconds(1))
                             .Select(_ => DateTime.Now.ToString("HH:mm:ss"));
        }
    }
}
