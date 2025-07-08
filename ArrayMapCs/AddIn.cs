using ExcelDna.Integration;
using ExcelDna.IntelliSense;

public class AddIn : IExcelAddIn
{
    public void AutoOpen()
    {
        IntelliSenseServer.Install();
    }

    public void AutoClose()
    {
        IntelliSenseServer.Uninstall();
    }
}
