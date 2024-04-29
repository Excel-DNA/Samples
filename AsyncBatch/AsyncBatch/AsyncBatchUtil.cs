using System.Timers;
using Timer = System.Timers.Timer;
using ExcelDna.Integration;
using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Threading;

namespace AsyncBatch
{
    // Step 1. from the list.
    // This is the main helper class for supporting batched async calls
    internal class AsyncBatchUtil
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

        readonly SemaphoreSlim _semaphore;
        readonly bool _serializeRequests;

        public AsyncBatchUtil(int maxBatchSize, TimeSpan batchTimeout, Func<List<AsyncCall>, Task<List<object>>> batchRunner, bool serializeRequests = false)
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

            _serializeRequests = serializeRequests;
            _semaphore = _serializeRequests ? new SemaphoreSlim(1, 1) : null;
        }

        public object Run(string functionName, params object[] args)
        {
            return ExcelAsyncUtil.Observe(functionName, args, delegate
            {
                var tcs = new TaskCompletionSource<object>();
                EnqueueAsyncCall(tcs, functionName, args);
                return new TaskExcelObservable(tcs.Task);
            });
        }

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
                if (_serializeRequests)
                {
                    await _semaphore.WaitAsync();
                }

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
            finally
            {
                if (_serializeRequests)
                {
                    _semaphore.Release();
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
