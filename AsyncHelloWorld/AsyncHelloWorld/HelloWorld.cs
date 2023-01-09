using ExcelDna.Integration;
using System.Threading;

namespace AsyncHelloWorld
{
    public class HelloWorld
    {
        [ExcelDna.Integration.ExcelFunction(Description = "Async Hello World")]
        public static object SayHelloAsync(string name)
        {
            return ExcelAsyncUtil.Run(nameof(SayHelloAsync), new object[] { name }, () => SayHelloWithDelay(name));
        }

        private static string SayHelloWithDelay(string name)
        {
            Thread.Sleep(2000);
            return $"Hello {name}";
        }
    }
}
