using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace Registration.Sample
{
    public static class MyAsyncFunctions
    {
        [ExcelFunction(Description = "My first .NET function")]
        public static string SayHello(string name)
        {
            return "Hello " + name;
        }

        [ExcelFunction]
        [return: ExcelAsyncDefault("ExcelError.ExcelErrorGettingData")]
        public static async Task<string> SayHelloAsync(string name, int delayMs = 1000)
        {
            // Simulate an asynchronous operation
            await Task.Delay(delayMs);
            return "Hello (async) " + name;
        }
    }

    [AttributeUsage(AttributeTargets.ReturnValue, AllowMultiple = false)]
    public class ExcelAsyncDefaultAttribute : Attribute
    {
        internal object DefaultReturnValue;
        // This attribute can be used to mark functions that should return a default value when the result is not available
        // It can be applied to methods that return Task<T> or IObservable<T>
        public ExcelAsyncDefaultAttribute(object defaultReturnValue)
        {
            DefaultReturnValue = defaultReturnValue;
        }
    }
    internal class AsyncReturnHandler : FunctionExecutionHandler
    {
        object _defaultValue;
        public AsyncReturnHandler(object defaultValue)
        {
            _defaultValue = defaultValue;
        }

        public override void OnSuccess(FunctionExecutionArgs args)
        {
            if (args.ReturnValue.Equals(ExcelError.ExcelErrorNA))
                args.ReturnValue = _defaultValue;
        }

        // This method can be inside any class
        // It is called for every function to optionally return an additional handler for that function
        [ExcelFunctionExecutionHandlerSelector]
        public static IFunctionExecutionHandler AsyncReturnHandlerSelector(IExcelFunctionInfo functionInfo)
        {
            // We change functions that are marked with ExcelAsyncFunction.
            // By default we (initially) return #GettingData for async functions instead of #N/A.
            // This can be overridden by using the ExcelAsyncDefault attribute on the return type.
            if (!functionInfo.CustomAttributes.Any(ca => ca is ExcelAsyncFunctionAttribute))
            {
                return null; // Not an async function, no special handling
            }

            object defaultReturnValue = ExcelError.ExcelErrorGettingData;
            var defaultAtt = (ExcelAsyncDefaultAttribute)functionInfo.Return.CustomAttributes.FirstOrDefault(ca => ca is ExcelAsyncDefaultAttribute);
            // If the function has a specific default return value, use that
            if (defaultAtt != null)
            {
                // If no specific default value is set, use the default #GettingData
                defaultReturnValue = defaultAtt.DefaultReturnValue;
            }
            return new AsyncReturnHandler(defaultReturnValue);
        }
    }
}
