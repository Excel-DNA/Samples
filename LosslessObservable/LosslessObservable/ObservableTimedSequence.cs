using ExcelDna.Integration;

namespace LosslessObservable
{
    internal class ObservableTimedSequence : IExcelObservable
    {
        private Timer _timer;
        private List<IExcelObserver> _observers;
        private int _counter = 0;

        public ObservableTimedSequence()
        {
            _timer = new Timer(OnTimerTick, null, 500, 500);
            _observers = new List<IExcelObserver>();
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            _observers.Add(observer);
            return new ActionDisposable(() => _observers.Remove(observer));
        }

        private void OnTimerTick(object? _)
        {
            foreach (var obs in _observers)
                obs.OnNext(_counter.ToString());
            ++_counter;

            if (_counter == 5)
            {
                _timer.Dispose();
                foreach (var obs in _observers)
                    obs.OnCompleted();
            }
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
