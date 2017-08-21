// To add the async / batch running into your code:
// 1. Add the whole "AsyncBatchUtil" class from this file into your project.
// 2. Create your own batch runner function, similar to the "RunBatch" method below.
// 3. Create an instance of "AsyncBatchUtil" somewhere in your code (it's called "BatchRunner" below), passing in the batch parameters and batch runner from step 2.
// 4. Create your worksheet functions like "SlowFunction" below, which call "BatchRunner.Run(...)" to run async as part of a batch.

using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using ExcelDna.Integration;
using Timer = System.Timers.Timer;
using System.Net.Http;

namespace GeneralTestsCS
{
    public static class AsyncBatchExample
    {
        // Step 3. from the list.
        static readonly AsyncBatchUtil BatchRunner = new AsyncBatchUtil(1000, TimeSpan.FromMilliseconds(250), RunBatch);

        // Step 2. from the list.
        // This function will be called for each batch, on a ThreadPool thread.
        // Each AsyncCall contains the function name and arguments passed from the function.
        // The List<object> returned by the Task must contain the results, corresponding to the calls list.
        static async Task<List<object>> RunBatch(List<AsyncBatchUtil.AsyncCall> calls)
        {
            var batchStart = DateTime.Now;
            // Simulate things taking a while...
            await Task.Delay(TimeSpan.FromSeconds(10));

            using (var httpClient = new HttpClient())
            {
                var page = await httpClient.GetStringAsync("http://www.google.com");
            }

            // Now build up the list of results...
            var results = new List<object>();
            int i = 0;
            foreach (var call in calls)
            {
                // As an example just an informative string
                var result = string.Format("{0} - {1} : {2}/{3} @ {4:HH:mm:ss.fff}", call.FunctionName, call.Arguments[0], i++, calls.Count, batchStart);
                results.Add(result);
            }

            return results;
        }

        // Step 4. from the list.
        public static object SlowFunction(string code, int value)
        {
            return BatchRunner.Run("SlowFunction", code, value);
        }
    }

    // Step 1. from the list.
    // This is the main helper class for supporting batched async calls
    public class AsyncBatchUtil
    {
        // Represents a single function call in  a batch
        public class AsyncCall
        {
            internal TaskCompletionSource<object> TaskCompletionSource;
            public string FunctionName { get; private set; }
            public object[] Arguments { get; private set; }

            public AsyncCall(TaskCompletionSource<object> taskCompletion, string functionName, object[] args)
            {
                TaskCompletionSource = taskCompletion;
                FunctionName = functionName;
                Arguments = args;
            }
        }

        // Not a hard limit
        readonly int _maxBatchSize;
        readonly Func<List<AsyncCall>, Task<List<object>>> _batchRunner;

        readonly object _lock = new object();
        readonly Timer _batchTimer;   // Timer events will fire from a ThreadPool thread
        List<AsyncCall> _currentBatch;

        public AsyncBatchUtil(int maxBatchSize, TimeSpan batchTimeout, Func<List<AsyncCall>, Task<List<object>>> batchRunner)
        {
            if (maxBatchSize < 1)
            {
                throw new ArgumentOutOfRangeException("maxBatchSize", "Max batch size must be positive");
            }
            if (batchRunner == null)
            {
                // Check early - otherwise the NullReferenceException would happen in a threadpool callback.
                throw new ArgumentNullException("batchRunner");
            }

            _maxBatchSize = maxBatchSize;
            _batchRunner = batchRunner;

            _currentBatch = new List<AsyncCall>();

            _batchTimer = new Timer(batchTimeout.TotalMilliseconds);
            _batchTimer.AutoReset = false;
            _batchTimer.Elapsed += TimerElapsed;
            // Timer is not Enabled (Started) by default
        }

        // Will only run on the main thread
        public object Run(string functionName, params object[] args)
        {
            return ExcelAsyncUtil.Observe(functionName, args, delegate
            {
                var tcs = new TaskCompletionSource<object>();
                EnqueueAsyncCall(tcs, functionName, args);
                return new TaskExcelObservable(tcs.Task);
            });
        }

        // Will only run on the main thread
        void EnqueueAsyncCall(TaskCompletionSource<object> taskCompletion, string functionName, object[] args)
        {
            lock (_lock)
            {
                _currentBatch.Add(new AsyncCall(taskCompletion, functionName, args));

                // Check if the batch size has been reached, schedule it to be run
                if (_currentBatch.Count >= _maxBatchSize)
                {
                    // This won't run the batch immediately, but will ensure that the current batch (containing this call) will run soon.
                    ThreadPool.QueueUserWorkItem(state => RunBatch((List<AsyncCall>)state), _currentBatch);
                    _currentBatch = new List<AsyncCall>();
                    _batchTimer.Stop();
                }
                else
                {
                    // We don't know if the batch containing the current call will run, 
                    // so ensure that a timer is started.
                    if (!_batchTimer.Enabled)
                    {
                        _batchTimer.Start();
                    }
                }
            }
        }

        // Will run on a ThreadPool thread
        void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            List<AsyncCall> batch;
            lock (_lock)
            {
                batch = _currentBatch;
                _currentBatch = new List<AsyncCall>();
            }
            RunBatch(batch);
        }


        // Will always run on a ThreadPool thread
        // Might be re-entered...
        // batch is allowed to be empty
        async void RunBatch(List<AsyncCall> batch)
        {
            // Maybe due to Timer re-entrancy we got an empty batch...?
            if (batch.Count == 0)
            {
                // No problem - just return
                return;
            }

            try
            {
                var resultList = await _batchRunner(batch);
                if (resultList.Count != batch.Count)
                {
                    throw new InvalidOperationException(string.Format("Batch result size incorrect. Batch Count: {0}, Result Count: {1}", batch.Count, resultList.Count));
                }

                for (int i = 0; i < resultList.Count; i++)
                {
                    batch[i].TaskCompletionSource.SetResult(resultList[i]);
                }
            }
            catch (Exception ex)
            {
                foreach (var call in batch)
                {
                    call.TaskCompletionSource.SetException(ex);
                }
            }
        }

        // Helper class to turn a task into an IExcelObservable that either returns the task result and completes, or pushes an Exception
        class TaskExcelObservable : IExcelObservable
        {
            readonly Task<object> _task;

            public TaskExcelObservable(Task<object> task)
            {
                _task = task;
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                switch (_task.Status)
                {
                    case TaskStatus.RanToCompletion:
                        observer.OnNext(_task.Result);
                        observer.OnCompleted();
                        break;
                    case TaskStatus.Faulted:
                        observer.OnError(_task.Exception.InnerException);
                        break;
                    case TaskStatus.Canceled:
                        observer.OnError(new TaskCanceledException(_task));
                        break;
                    default:
                        var task = _task;
                        // OK - the Task has not completed synchronously
                        // And handle the Task completion
                        task.ContinueWith(t =>
                        {
                            switch (t.Status)
                            {
                                case TaskStatus.RanToCompletion:
                                    observer.OnNext(t.Result);
                                    observer.OnCompleted();
                                    break;
                                case TaskStatus.Faulted:
                                    observer.OnError(t.Exception.InnerException);
                                    break;
                                case TaskStatus.Canceled:
                                    observer.OnError(new TaskCanceledException(t));
                                    break;
                            }
                        });
                        break;
                }

                return DefaultDisposable.Instance;
            }

            // Helper class to make an empty IDisposable
            sealed class DefaultDisposable : IDisposable
            {
                public static readonly DefaultDisposable Instance = new DefaultDisposable();
                // Prevent external instantiation
                DefaultDisposable() { }
                public void Dispose() { }
            }
        }
    }
}
