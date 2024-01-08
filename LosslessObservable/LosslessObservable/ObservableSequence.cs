using ExcelDna.Integration;

namespace LosslessObservable
{
    internal class ObservableSequence : IExcelObservable
    {
        public IDisposable Subscribe(IExcelObserver observer)
        {
            for (int i = 0; i <= 5; ++i)
                observer.OnNext(i.ToString());

            observer.OnCompleted();

            return new ActionDisposable();
        }

        private class ActionDisposable : IDisposable
        {
            public ActionDisposable()
            {
            }

            public void Dispose()
            {
            }
        }
    }
}
