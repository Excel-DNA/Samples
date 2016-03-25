// This file is just a copy from the Excel-DNA Registration project: 
// https://github.com/Excel-DNA/Registration/blob/master/Source/ExcelDna.Registration/Utils/ObservableRtdUtil.cs

using System;
using ExcelDna.Integration;

namespace ExcelDna.Registration.Utils
{
    public static class ObservableRtdUtil
    {
        public static object Observe<T>(string callerFunctionName, object callerParameters, Func<IObservable<T>> observableSource)
        {
            return ExcelAsyncUtil.Observe(callerFunctionName, callerParameters, () => new ExcelObservable<T>(observableSource()));
        }

        // An IExcelObservable that wraps an IObservable
        class ExcelObservable<T> : IExcelObservable
        {
            readonly IObservable<T> _observable;

            public ExcelObservable(IObservable<T> observable)
            {
                _observable = observable;
            }

            public IDisposable Subscribe(IExcelObserver excelObserver)
            {
                var observer = new AnonymousObserver<T>(value => excelObserver.OnNext(value), excelObserver.OnError, excelObserver.OnCompleted);
                return _observable.Subscribe(observer);
            }
        }

        // An IObserver that forwards the inputs to given methods.
        class AnonymousObserver<T> : IObserver<T>
        {
            readonly Action<T> _onNext;
            readonly Action<Exception> _onError;
            readonly Action _onCompleted;

            public AnonymousObserver(Action<T> onNext, Action<Exception> onError, Action onCompleted)
            {
                if (onNext == null)
                {
                    throw new ArgumentNullException("onNext");
                }
                if (onError == null)
                {
                    throw new ArgumentNullException("onError");
                }
                if (onCompleted == null)
                {
                    throw new ArgumentNullException("onCompleted");
                }
                _onNext = onNext;
                _onError = onError;
                _onCompleted = onCompleted;
            }

            public void OnNext(T value)
            {
                _onNext(value);
            }

            public void OnError(Exception error)
            {
                _onError(error);
            }

            public void OnCompleted()
            {
                _onCompleted();
            }
        }

    }

}
