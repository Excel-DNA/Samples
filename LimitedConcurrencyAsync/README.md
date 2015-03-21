#LimitedConcurrencyAsync sample

[Related to this discussion: https://groups.google.com/d/topic/exceldna/tCbtb2zmQrs/discussion]

This sample shows how the async function support can be customised by using the .NET 4 Task-based functionality.
In particular, we create a limited concurrency scheduler, that will restrict the number of async threads that are used to run Tasks.

Some hard-coded paths are set in the project properties in the Debug tab - the exact Excel version and command line arguments. These must be fixed before running the sample. One way is to run "Uninstall-Package Excel-DNA" and then "Install-Package Excel-DNA" in the NuGet Package Manager Console.

When running, there should be two new functions in Excel - "Sleep" and "SleepPerCaller", taking the number of seconds to sleep.

Some details on the code:

## AsyncFunctions.cs

This is the user code part of the sample. A custom TaskScheduler and related TaskFactory is initialized, and some async Excel-DNA functions defined that will create Tasks using that TaskFactory.

There are two versions of the Sleep function:
  * Sleep - different calls to Sleep with the same timeout parameter will be combined and run as the same Task.
  * SleepPerCaller - calls from different cells will create separate Task, making the concurrency behaviour easier to see.

## AsyncTaskUtil.cs

This file contains some helpers to integrate the Task-based API with Excel-DNA async support. The main helper function is AsyncTaskUtil, which takes the async call identifiers (the callerFunctionName and callerParameters) as well as an Action<Task> that will create the async Task on the first call. Internally, an ExcelTaskObservable is created, which converts the Task completion result into the appropriate IObservable interface to register with Excel-DNA.

There is also an overload that supports cancellation.

## LimitedConcurrencyLevelTaskScheduler.cs

This file is taken from the "Samples for Parallel Programming" on MSDN (https://code.msdn.microsoft.com/Samples-for-Parallel-b4b76364/sourcecode?fileId=44488&pathId=2044791305). There are a number of custom TaskScheduler samples, including a very flexible QueuedTaskScheduler. The TaskScheduler samples are discussed in detail by Stephen Taub here: http://blogs.msdn.com/b/pfxteam/archive/2010/04/09/9990424.aspx .

The LimitedConcurrencyLevelTaskScheduler uses the .NET ThreadPool threads to run the Tasks, but limits the number of concurrent Tasks that can be running.



