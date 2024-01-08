LosslessObservable sample
---

You can implement a RealTimeData (RTD) function using `IExcelObservable` interface and `ExcelAsyncUtil.Observe` method.

Real-time data is data that updates on its own schedule (for example, stock quotes, manufacturing statistics, Web server loads, and warehouse activity).

Excel gets updates from such function when it is in a "good state" and it waits at least the number of milliseconds specified by the RTD throttle interval. Excel does not get updates while a modal dialog box is displayed, while a cell is being edited, or while it is busy doing other things.

Generally, if more than one update comes in before Excel calls back for updates, old values are discarded and the new value is passed to Excel. But there are some instances where someone wants every update. In that case, you need to specify `ExcelObservableOptions.Lossless` option when calling `ExcelAsyncUtil.Observe`:

```c#

public static object ObservableSequence()
{
    return ExcelAsyncUtil.Observe("ObservableSequence", null, ExcelObservableOptions.Lossless, () => new ObservableSequence());
}

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

```

You can test it like this:

```

=ObservableSequence()

```

It will set the cell value with numbers from 1 to 5 with 2 seconds delay (with default Excel settings).

The delay is determined by the Excel RTD throttle interval that by default is 2000 milliseconds. It can be modified via the Excel object model or the registry, or using the provided in the sample project `SetThrottleInterval` function:

```c#

public static int SetThrottleInterval(int interval)
{
    Microsoft.Office.Interop.Excel.Application application = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
    application.RTD.ThrottleInterval = interval;
    return application.RTD.ThrottleInterval;
}

```

You can use it like this:

```

=SetThrottleInterval(500)

```

Note, that programmatically set value is still saved by Excel in registry and will persist even if you close and reopen Excel. 

The following registry key is used (it is a DWORD and is in milliseconds):

```

HKEY_CURRENT_USER\Software\Microsoft\Office\[YOUR_OFFICE_VERSION]]\Excel\Options\RTDThrottleInterval

```

The sample project also provides a sequence with numbers generated every 500 milliseconds that can be invoked with `=ObservableTimedSequence()`:

```c#
public static object ObservableTimedSequence()
{
    return ExcelAsyncUtil.Observe("ObservableTimedSequence", null, ExcelObservableOptions.Lossless, () => new ObservableTimedSequence());
}

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

```

The sample project also provides a lossless clock updated every second that can be invoked with `=LosslessClock()`. Note, that if RTD throttle interval is greater than 1 second, then displayed time will lag from real time. Thus, it is a not a good way to display correct time, but illustrates lossless behavior. 

