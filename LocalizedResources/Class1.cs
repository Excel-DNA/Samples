using System.Globalization;

namespace LocalizedResources
{
    public class Class1
    {
        public static string locHello()
        {
            return "Hello from LocalizedResources";
        }

        public static string locGetString1(string cultureName)
        {
            return Resource1.ResourceManager.GetString("String1", CultureInfo.GetCultureInfo(cultureName));
        }
    }
}
