using ExcelDna.Integration;
using System.Threading;

namespace AsyncHelloWorld
{
    public class HelloWorld
    {
        [ExcelDna.Integration.ExcelFunction(Description = "Async Hello World")]
        public static object SayHelloAsync(string name)
        {
            return ExcelAsyncUtil.Run("RunSomethingDelay", new [] { name }, () => RunSomethingDelay(name));
        }

        public static string RunSomethingDelay(string name)
        {
            Thread.Sleep(2000);
            return $"Hello {name}";
        }
    }
}
