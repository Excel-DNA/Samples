AsyncAwaitMacro sample
---

Shows how the ExcelAsyncUtil.QueueAsMacro mechanism can be used to implement async macros with the C# async/await mechanism.

### Warning

This sample does not represent 'best practice'. It just explores how the C# async/await features interact with the Excel-DNA async mechanism, and with the Excel hosting environment.

Trying to run async macros as in this example will interfere with an interactive user busy with Excel: 
 their undo stack and copy selection will get cleared at unexpected times, and any other add-ins or macros being run might be
 interleaved with the async code.
 
#### `ExcelAsyncTask`

The sample includes a helper class called `ExcelAsyncTask` with a single `Run` method.
In turn, this starts the `Task` with a TaskScheduler which just enqueues `Task`s to run in the macro context, 
ensuring that the async/await continuations are again scheduled in a macro context.

```c#

public static void MacroToRunSlowWork()
{
    // Starts running SlowWork in a context where async/await will return to the macro context on the main Excel thread.
    ExcelTaskAsync.Run(SlowWork);
}

static async Task SlowWork()
{
    // All the code here, before and after the awaits, will run on the main thread in a macro context
    // (where C API calls and the COM object model is safe to access).
    await SomeWorkAsync();
    Application.Range["A1"].Value = "abc;
    await OtherWorkAsync();
    Application.Range["A2"].Value = "xyz;
}

```


#### `ExcelSynchronizationContext`

The first implementation I attempted was run the async/await code in a context where `SynchronizationContext.Current` was set to an `ExcelSynchronizationContext`.

There is a problem I don't yet understand when trying to use a `SynchronizationContext` in Excel-DNA.
Somehow the `SynchronizationContext.Current` is cleared in the SyncWindow or macro running process.
I have found references to `WindowsFormsSynchronizationContext.AutoInstall` causing trouble, but could not see how that applies in our case.

It could be that the unmanaged -> managed transition interferes with the thread-based context that stores the SynchronizationContext.Current.
As an alternative, we use the TaskScheduler-based approach.

