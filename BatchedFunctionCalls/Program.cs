using BatchedFunctionCalls;
using ExcelDna.Integration;
using Open.ChannelExtensions;
using System.Threading.Channels;
using ExcelDna.Registration;

/// <summary>
/// Demonstrates how batch function calls to a remote server.
/// </summary>
public class BatchedFunctions : IExcelAddIn
{
    private static readonly Channel<FunctionParams> c = Channel.CreateUnbounded<FunctionParams>();
    private static readonly int MaxBatchSize = 200;

    static BatchedFunctions()
    {
        c.Reader.Batch(MaxBatchSize, singleReader: true).WithTimeout(1).ReadAllAsync(async batch =>
        {
            var simulatedRequestTime = 1000 + (batch.Count * 10);// Simulate calling a remote server to get data.
            await Task.Delay(simulatedRequestTime);
            foreach (var item in batch)
            {
                item.result.SetResult(item.Year);
            }
        });
    }

    [ExcelFunction(Name = "BatchedCall", Description = "Function that will be batched")]
    public static async Task<object> BatchedCall(
        [ExcelArgument(Name = "ticker")] string ticker,
        [ExcelArgument(Name = "year")] int year)
    {        
        var param = new FunctionParams() { Ticker = ticker, Year = year };
        c.Writer.TryWrite(param);
        return await param.result.Task; ;
    }

    public void AutoOpen()
    {
        ExcelRegistration.GetExcelFunctions().ProcessAsyncRegistrations(nativeAsyncIfAvailable: false).RegisterFunctions();
    }

    public void AutoClose()
    {

    }

}
