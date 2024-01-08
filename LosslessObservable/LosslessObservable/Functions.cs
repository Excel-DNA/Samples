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
            return ExcelAsyncUtil.Observe("LosslessClock", null, ExcelObservableOptions.Lossless, () => new LosslessClock());
        }
    }
}
