using ExcelDna.Integration;

namespace Registration.Sample
{
    public class ExampleAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // First example if Instance -> Static conversion
            InstanceMemberRegistration.TestInstanceRegistration();
        }

        public void AutoClose()
        {
        }
    }
}
