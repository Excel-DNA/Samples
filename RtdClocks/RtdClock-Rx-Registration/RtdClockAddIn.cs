using ExcelDna.Integration;
using ExcelDna.Registration;

namespace RtdClock_Rx_Registration
{
    public class RtdClockAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Since we have specified ExplicitRegistration=true in the .dna file, we need to do all registration explicitly.
            // Here we add the async processing built into ExcelDna.Integration, which applies to our IObservable function.
            ExcelRegistration.GetExcelFunctions()
                             .ProcessAsyncRegistrations()
                             .RegisterFunctions();
            ExcelRegistration.GetExcelCommands().RegisterCommands();
        }

        public void AutoClose()
        {
        }
    }
}
