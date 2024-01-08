using ExcelDna.Integration;

namespace LosslessObservable
{
    internal class LosslessClock : IExcelObservable
    {
        private Timer _timer;
        private List<IExcelObserver> _observers;

        public LosslessClock()
        {
            _timer = new Timer(OnTimerTick, null, 0, 1000);
            _observers = new List<IExcelObserver>();
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            _observers.Add(observer);
            observer.OnNext(DateTime.Now.ToString("HH:mm:ss.fff") + " (Subscribe)");
            return new ActionDisposable(() => _observers.Remove(observer));
        }

        private void OnTimerTick(object? _)
        {
            string now = DateTime.Now.ToString("HH:mm:ss.fff");
            foreach (var obs in _observers)
                obs.OnNext(now);
        }

        private class ActionDisposable : IDisposable
        {
            private Action _disposeAction;

            public ActionDisposable(Action disposeAction)
            {
                _disposeAction = disposeAction;
            }

            public void Dispose()
            {
                _disposeAction();
            }
        }

    }
}
