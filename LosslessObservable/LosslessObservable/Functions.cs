using ExcelDna.Integration;

namespace LosslessObservable
{
    public static class Functions
    {
        public static object ObservableSequence()
        {
            return ExcelAsyncUtil.Observe("ObservableSequence", null, ExcelObservableOptions.Lossless, () => new ObservableSequence());
        }

        public static object ObservableTimedSequence()
        {
            return ExcelAsyncUtil.Observe("ObservableTimedSequence", null, ExcelObservableOptions.Lossless, () => new ObservableTimedSequence());
        }

        public static object LosslessClock()
        {
            return ExcelAsyncUtil.Observe("LosslessClock", null, ExcelObservableOptions.Lossless, () => new ObservableClock());
        }

        public static object Clock()
        {
            return ExcelAsyncUtil.Observe("Clock", null, ExcelObservableOptions.None, () => new ObservableClock());
        }

        public static int SetThrottleInterval(int interval)
        {
            Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            application.RTD.ThrottleInterval = interval;
            return application.RTD.ThrottleInterval;
        }
    }
}
