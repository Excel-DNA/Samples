using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using ExcelDna.Integration;

namespace RtdClock_IExcelObservable
{
    class ExcelObservableClock : IExcelObservable
    {
        Timer _timer;
        List<IExcelObserver> _observers;

        public ExcelObservableClock()
        {
            _timer = new Timer(timer_tick, null, 0, 1000);
            _observers = new List<IExcelObserver>();
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            _observers.Add(observer);
            observer.OnNext(DateTime.Now.ToString("HH:mm:ss.fff") + " (Subscribe)");
            return new ActionDisposable(() => _observers.Remove(observer));
        }

        void timer_tick(object _)
        {
            string now = DateTime.Now.ToString("HH:mm:ss.fff");
            foreach (var obs in _observers)
                obs.OnNext(now);
        }

        class ActionDisposable : IDisposable
        {
            Action _disposeAction;
            public ActionDisposable(Action disposeAction)
            {
                _disposeAction = disposeAction;
            }
            public void Dispose()
            {
                _disposeAction();
                Debug.WriteLine("Disposed");
            }
        }
    }
}
