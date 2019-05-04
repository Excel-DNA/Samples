using ExcelDna.Integration;

namespace CustomTaskPane
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            System.Windows.Forms.Application.EnableVisualStyles();
        }

        public void AutoClose()
        {
        }

    }
}
