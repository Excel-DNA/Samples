using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelDna.Integration
{
    public static class ExcelAsyncTask
    {
        public static Task Run(Func<Task> function)
        {
            return Task<Task>.Factory.StartNew(function, CancellationToken.None, TaskCreationOptions.DenyChildAttach, ExcelAsyncTaskScheduler.Instance).Unwrap();
        }
    }

    class ExcelAsyncTaskScheduler : TaskScheduler
    {
        protected override void QueueTask(Task task)
        {
            ExcelAsyncUtil.QueueAsMacro(PostCallback, task);
        }

        void PostCallback(object obj)
        {
            Task task = (Task)obj;
            TryExecuteTask(task);
        }

        protected override bool TryExecuteTaskInline(Task task, bool taskWasPreviouslyQueued)
        {
            // TODO: We might add a mechanism that tries this...
            return false;
        }

        protected override IEnumerable<Task> GetScheduledTasks()
        {
            return null;
        }

        internal static ExcelAsyncTaskScheduler Instance = new ExcelAsyncTaskScheduler();
    }
}
