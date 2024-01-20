// To add the async / batch running into your code:
// 1. Add the whole "AsyncBatchUtil" class from this file into your project.
// 2. Create your own batch runner function, similar to the "RunBatch" method below.
// 3. Create an instance of "AsyncBatchUtil" somewhere in your code (it's called "BatchRunner" below), passing in the batch parameters and batch runner from step 2.
// 4. Create your worksheet functions like "SlowFunction" below, which call "BatchRunner.Run(...)" to run async as part of a batch.

using ExcelDna.Integration;

namespace AsyncBatch
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

        [ExcelFunction(IsThreadSafe = true)]
        public static object SlowFunctionThreadSafe(string code, int value)
        {
            return BatchRunner.Run("SlowFunctionThreadSafe", code, value);
        }
    }
}
